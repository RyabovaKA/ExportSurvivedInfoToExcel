using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using Ascon.Pilot.SDK;

namespace PilotExcelExporter
{
    public class LoadedObjectNode
    {
        public IDataObject DataObject { get; set; }
        public string TreePath { get; set; }
    }

    public class PilotTreeLoader
    {
        private const int BatchSize = 50;
        private const int IdleTimeoutMs = 1200;     // пауза после последнего полученного объекта (если часть детей недоступна по правам)
        private const int FullBatchDelayMs = 200;   // небольшая пауза, если пришли все запрошенные объекты
        private const int MaxBatchWaitMs = 15000;   // жесткий потолок ожидания на один батч

        private readonly IObjectsRepository _objectsRepository;
        private readonly Dispatcher _uiDispatcher;

        public PilotTreeLoader(IObjectsRepository objectsRepository, Dispatcher uiDispatcher)
        {
            _objectsRepository = objectsRepository;
            _uiDispatcher = uiDispatcher;
        }

        public async Task<Dictionary<Guid, LoadedObjectNode>> LoadAccessibleTreeAsync(
            IEnumerable<Guid> startIds,
            IProgress<ExportProgressInfo> progress,
            CancellationToken token)
        {
            var result = new Dictionary<Guid, LoadedObjectNode>();
            var visited = new HashSet<Guid>();
            var queue = new Queue<KeyValuePair<Guid, string>>();

            foreach (var id in startIds.Distinct())
            {
                if (visited.Add(id))
                    queue.Enqueue(new KeyValuePair<Guid, string>(id, string.Empty));
            }

            while (queue.Count > 0)
            {
                token.ThrowIfCancellationRequested();

                var currentBatch = new List<KeyValuePair<Guid, string>>();
                while (queue.Count > 0 && currentBatch.Count < BatchSize)
                    currentBatch.Add(queue.Dequeue());

                var batchIds = currentBatch.Select(x => x.Key).ToList();
                var parentPathMap = currentBatch.ToDictionary(x => x.Key, x => x.Value);

                progress?.Report(new ExportProgressInfo
                {
                    Stage = "Сканирование доступных элементов...",
                    Details = $"Загружено: {result.Count} | В очереди: {queue.Count + currentBatch.Count}"
                });

                var loadedBatch = await FetchObjectsBatchAsync(batchIds, token);

                foreach (var obj in loadedBatch.Values)
                {
                    if (obj == null) continue;

                    string title = PilotAttributeHelper.GetObjectTitle(obj);
                    string parentPath = parentPathMap.TryGetValue(obj.Id, out var p) ? p : string.Empty;
                    string fullPath = string.IsNullOrEmpty(parentPath) ? title : $"{parentPath} / {title}";

                    result[obj.Id] = new LoadedObjectNode
                    {
                        DataObject = obj,
                        TreePath = fullPath
                    };

                    if (obj.Children != null)
                    {
                        foreach (var childId in obj.Children)
                        {
                            if (visited.Add(childId))
                                queue.Enqueue(new KeyValuePair<Guid, string>(childId, fullPath));
                        }
                    }
                }
            }

            return result;
        }

        public async Task<Dictionary<Guid, IDataObject>> FetchMissingObjectsAsync(
            IEnumerable<Guid> ids,
            IProgress<ExportProgressInfo> progress,
            CancellationToken token)
        {
            var allLoaded = new Dictionary<Guid, IDataObject>();
            var distinctIds = ids.Distinct().ToList();

            for (int i = 0; i < distinctIds.Count; i += BatchSize)
            {
                token.ThrowIfCancellationRequested();
                var batch = distinctIds.Skip(i).Take(BatchSize).ToList();

                progress?.Report(new ExportProgressInfo
                {
                    Stage = "Подгрузка связанных элементов справочников...",
                    Details = $"Обработано {Math.Min(i + BatchSize, distinctIds.Count)} из {distinctIds.Count}"
                });

                var batchResult = await FetchObjectsBatchAsync(batch, token);
                foreach (var kvp in batchResult)
                    allLoaded[kvp.Key] = kvp.Value;
            }

            return allLoaded;
        }

        private async Task<Dictionary<Guid, IDataObject>> FetchObjectsBatchAsync(List<Guid> batchIds, CancellationToken token)
        {
            var received = new ConcurrentDictionary<Guid, IDataObject>();
            if (batchIds == null || batchIds.Count == 0)
                return new Dictionary<Guid, IDataObject>();

            var targetSet = new HashSet<Guid>(batchIds);
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
            var syncLock = new object();
            Timer idleTimer = null;
            Timer maxTimer = null;
            IDisposable subscription = null;

            void Complete()
            {
                lock (syncLock)
                {
                    idleTimer?.Change(Timeout.Infinite, Timeout.Infinite);
                    maxTimer?.Change(Timeout.Infinite, Timeout.Infinite);
                    tcs.TrySetResult(true);
                }
            }

            try
            {
                idleTimer = new Timer(_ => Complete(), null, IdleTimeoutMs, Timeout.Infinite);
                maxTimer = new Timer(_ => Complete(), null, MaxBatchWaitMs, Timeout.Infinite);

                using (token.Register(() => tcs.TrySetCanceled(token)))
                {
                    await _uiDispatcher.InvokeAsync(() =>
                    {
                        if (token.IsCancellationRequested) return;

                        subscription = _objectsRepository.SubscribeObjects(batchIds).Subscribe(
                            new AnonymousObjectObserver(rawData =>
                            {
                                var objs = UnpackObjects(rawData);
                                bool anyRelevant = false;

                                foreach (var obj in objs)
                                {
                                    if (obj != null && targetSet.Contains(obj.Id))
                                    {
                                        received[obj.Id] = obj;
                                        anyRelevant = true;
                                    }
                                }

                                if (anyRelevant)
                                {
                                    lock (syncLock)
                                    {
                                        if (!tcs.Task.IsCompleted)
                                        {
                                            int delay = received.Count >= targetSet.Count ? FullBatchDelayMs : IdleTimeoutMs;
                                            idleTimer?.Change(delay, Timeout.Infinite);
                                        }
                                    }
                                }
                            }));
                    });

                    await tcs.Task;
                }
            }
            finally
            {
                idleTimer?.Dispose();
                maxTimer?.Dispose();

                if (subscription != null)
                {
                    try
                    {
                        await _uiDispatcher.InvokeAsync(() => subscription.Dispose());
                    }
                    catch { }
                }
            }

            return new Dictionary<Guid, IDataObject>(received);
        }

        private static List<IDataObject> UnpackObjects(object rawData)
        {
            var list = new List<IDataObject>();
            if (rawData is IDataObject single)
            {
                list.Add(single);
            }
            else if (rawData is System.Collections.IEnumerable enumerable)
            {
                foreach (var item in enumerable)
                {
                    if (item is IDataObject obj)
                        list.Add(obj);
                }
            }
            return list;
        }

        private class AnonymousObjectObserver : IObserver<object>
        {
            private readonly Action<object> _onNext;
            public AnonymousObjectObserver(Action<object> onNext) => _onNext = onNext;
            public void OnNext(object value) => _onNext?.Invoke(value);
            public void OnError(Exception error) { }
            public void OnCompleted() { }
        }
    }
}