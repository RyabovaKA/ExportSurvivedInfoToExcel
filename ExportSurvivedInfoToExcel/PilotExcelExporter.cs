using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Threading;
using System.Xml.Linq;
using Ascon.Pilot.SDK;
using ClosedXML.Excel;

namespace PilotExcelExporter
{
    public class ExportResult
    {
        public int FilesCreated { get; set; }
        public int ObjectsExported { get; set; }
    }

    public class PilotExcelExporter
    {
        private static readonly Guid RootObjectId = new Guid("00000001-0001-0001-0001-000000000001");
        private readonly IObjectsRepository _objectsRepository;
        private readonly Dispatcher _uiDispatcher;

        public PilotExcelExporter(IObjectsRepository objectsRepository, Dispatcher uiDispatcher)
        {
            _objectsRepository = objectsRepository;
            _uiDispatcher = uiDispatcher;
        }

        public async Task<ExportResult> ExportAsync(
            List<Guid> startIds,
            string outputFolder,
            bool separateFilePerType,
            bool includeServiceAttributes,
            bool resolveReferenceObjects,
            IProgress<ExportProgressInfo> progress,
            CancellationToken token)
        {
            // 1. Загружаем справочники (типы, оргструктуру, пользователей, состояния) на UI-потоке
            PilotLookupContext lookups = null;
            await _uiDispatcher.InvokeAsync(() =>
            {
                lookups = PilotLookupContext.Create(_objectsRepository);
            });

            // 2. Рекурсивно загружаем все доступные элементы дерева
            var loader = new PilotTreeLoader(_objectsRepository, _uiDispatcher);
            var loadedNodes = await loader.LoadAccessibleTreeAsync(startIds, progress, token);

            // Исключаем виртуальный корневой контейнер (если у него нет реального типа или это системный корень)
            var exportableNodes = loadedNodes.Values
                .Where(n => n.DataObject != null && n.DataObject.Type != null && n.DataObject.Id != RootObjectId)
                .ToList();

            if (!exportableNodes.Any())
                return new ExportResult { FilesCreated = 0, ObjectsExported = 0 };

            // Сохраняем уже загруженные объекты в словарь для расшифровки ссылок
            foreach (var kvp in loadedNodes)
            {
                if (kvp.Value.DataObject != null)
                    lookups.KnownObjects[kvp.Key] = kvp.Value.DataObject;
            }

            // 3. При необходимости подгружаем недостающие объекты справочников (ElementBook / Object)
            if (resolveReferenceObjects)
            {
                var missingRefIds = CollectMissingReferencedObjectIds(exportableNodes, lookups, includeServiceAttributes);
                if (missingRefIds.Any())
                {
                    var extraObjects = await loader.FetchMissingObjectsAsync(missingRefIds, progress, token);
                    foreach (var kvp in extraObjects)
                        lookups.KnownObjects[kvp.Key] = kvp.Value;
                }
            }

            // 4. Группируем элементы по типу
            var groupedByType = exportableNodes
                .GroupBy(n => n.DataObject.Type.Name ?? n.DataObject.Type.Title ?? "UnknownType")
                .OrderBy(g => g.First().DataObject.Type.Title ?? g.Key)
                .ToList();

            Directory.CreateDirectory(outputFolder);

            int filesCreated = 0;

            await Task.Run(() =>
            {
                if (separateFilePerType)
                {
                    var usedFileNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var group in groupedByType)
                    {
                        token.ThrowIfCancellationRequested();

                        var pilotType = group.First().DataObject.Type;
                        string typeTitle = string.IsNullOrWhiteSpace(pilotType.Title) ? pilotType.Name : pilotType.Title;

                        progress?.Report(new ExportProgressInfo
                        {
                            Stage = "Формирование файлов Excel...",
                            Details = $"Тип: {typeTitle} ({group.Count()} шт.)"
                        });

                        string baseFileName = SanitizeFileName($"{typeTitle} ({pilotType.Name})");
                        string finalFileName = MakeUniqueName(baseFileName, usedFileNames, 120) + ".xlsx";
                        string fullPath = Path.Combine(outputFolder, finalFileName);

                        using (var workbook = new XLWorkbook())
                        {
                            WriteTypeWorksheet(workbook, typeTitle, pilotType, group.ToList(), lookups, includeServiceAttributes);
                            workbook.SaveAs(fullPath);
                        }

                        filesCreated++;
                    }
                }
                else
                {
                    var usedSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    string filePath = Path.Combine(outputFolder, $"PilotExport_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.xlsx");

                    using (var workbook = new XLWorkbook())
                    {
                        foreach (var group in groupedByType)
                        {
                            token.ThrowIfCancellationRequested();

                            var pilotType = group.First().DataObject.Type;
                            string typeTitle = string.IsNullOrWhiteSpace(pilotType.Title) ? pilotType.Name : pilotType.Title;

                            progress?.Report(new ExportProgressInfo
                            {
                                Stage = "Формирование единой книги Excel...",
                                Details = $"Лист: {typeTitle} ({group.Count()} шт.)"
                            });

                            string sheetName = MakeUniqueName(SanitizeSheetName(typeTitle), usedSheetNames, 31);
                            WriteTypeWorksheet(workbook, sheetName, pilotType, group.ToList(), lookups, includeServiceAttributes);
                        }

                        workbook.SaveAs(filePath);
                    }

                    filesCreated = 1;
                }
            }, token);

            return new ExportResult
            {
                FilesCreated = filesCreated,
                ObjectsExported = exportableNodes.Count
            };
        }

        private void WriteTypeWorksheet(
            XLWorkbook workbook,
            string rawSheetName,
            IType pilotType,
            List<LoadedObjectNode> nodes,
            PilotLookupContext lookups,
            bool includeServiceAttributes)
        {
            string sheetName = SanitizeSheetName(rawSheetName);
            if (workbook.Worksheets.Contains(sheetName))
                sheetName = MakeUniqueName(sheetName, new HashSet<string>(workbook.Worksheets.Select(w => w.Name), StringComparer.OrdinalIgnoreCase), 31);

            var ws = workbook.Worksheets.Add(sheetName);

            // Определяем все поля карточки этого типа
            var typeAttributes = (pilotType.Attributes ?? Enumerable.Empty<IAttribute>())
                .Where(a => a != null && (includeServiceAttributes || !a.IsService))
                .OrderBy(a => a.DisplaySortOrder)
                .ToList();

            var knownAttrNames = new HashSet<string>(typeAttributes.Select(a => a.Name), StringComparer.OrdinalIgnoreCase);

            // На случай, если в объектах встретились дополнительные атрибуты, не описанные явно в коллекции
            var extraAttrNames = nodes
                .Where(n => n.DataObject.Attributes != null)
                .SelectMany(n => n.DataObject.Attributes.Keys)
                .Where(k => !string.IsNullOrEmpty(k) && !knownAttrNames.Contains(k))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(k => k)
                .ToList();

            // Подготовка конфигурации атрибутов и уникальных заголовков колонок
            var columns = new List<AttributeColumnInfo>();
            var usedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "ID", "Путь в дереве", "Отображаемое имя"
            };

            foreach (var attr in typeAttributes)
            {
                string header = string.IsNullOrWhiteSpace(attr.Title) ? attr.Name : attr.Title.Trim();
                if (!usedHeaders.Add(header))
                    header = $"{header} [{attr.Name}]";

                columns.Add(new AttributeColumnInfo
                {
                    AttrName = attr.Name,
                    Header = header,
                    AttributeMeta = attr,
                    Config = PilotAttributeHelper.ParsePilotConfig(attr.Configuration2())
                });
            }

            foreach (var extraName in extraAttrNames)
            {
                string header = extraName;
                if (!usedHeaders.Add(header))
                    header = $"{header} [extra]";

                columns.Add(new AttributeColumnInfo
                {
                    AttrName = extraName,
                    Header = header,
                    AttributeMeta = null,
                    Config = new PilotRefBookConfig()
                });
            }

            // Шапка таблицы
            int colIdx = 1;
            ws.Cell(1, colIdx++).Value = "ID";
            ws.Cell(1, colIdx++).Value = "Путь в дереве";
            ws.Cell(1, colIdx++).Value = "Отображаемое имя";

            foreach (var col in columns)
                ws.Cell(1, colIdx++).Value = col.Header;

            int totalCols = colIdx - 1;
            var headerRange = ws.Range(1, 1, 1, totalCols);
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E1F2");
            headerRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            headerRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            headerRange.Style.Alignment.WrapText = true;

            // Заполнение строк объектами
            int rowIdx = 2;
            foreach (var node in nodes)
            {
                var obj = node.DataObject;
                int c = 1;

                ws.Cell(rowIdx, c++).Value = obj.Id.ToString();
                ws.Cell(rowIdx, c++).Value = TruncateForExcel(node.TreePath);
                ws.Cell(rowIdx, c++).Value = TruncateForExcel(PilotAttributeHelper.GetObjectTitle(obj));

                foreach (var col in columns)
                {
                    object rawValue = null;
                    if (obj.Attributes != null)
                        obj.Attributes.TryGetValue(col.AttrName, out rawValue);

                    object formatted = PilotAttributeHelper.FormatValue(rawValue, col.AttributeMeta, col.Config, lookups);
                    SetExcelCellValue(ws.Cell(rowIdx, c++), formatted);
                }

                rowIdx++;
            }

            // Автофильтр, закрепление шапки и подбор ширины колонок
            if (rowIdx > 1)
                ws.Range(1, 1, rowIdx - 1, totalCols).SetAutoFilter();

            ws.SheetView.FreezeRows(1);
            ws.Columns(1, totalCols).AdjustToContents();

            foreach (var col in ws.ColumnsUsed())
            {
                if (col.Width < 10) col.Width = 10;
                if (col.Width > 60)
                {
                    col.Width = 60;
                    col.Style.Alignment.WrapText = true;
                }
            }
        }

        private static void SetExcelCellValue(IXLCell cell, object value)
        {
            if (value == null)
                return;

            switch (value)
            {
                case DateTime dt:
                    // В Excel поддерживаются даты начиная с 1900 года
                    if (dt.Year >= 1900 && dt.Year <= 9999)
                    {
                        cell.Value = dt;
                        cell.Style.DateFormat.Format = dt.TimeOfDay == TimeSpan.Zero ? "dd.MM.yyyy" : "dd.MM.yyyy HH:mm";
                    }
                    else
                    {
                        cell.Value = dt.ToString("dd.MM.yyyy HH:mm:ss");
                    }
                    break;

                case int i:
                    cell.Value = i;
                    break;
                case long l:
                    cell.Value = l;
                    break;
                case double d:
                    cell.Value = d;
                    break;
                case decimal dec:
                    cell.Value = (double)dec;
                    break;
                case float f:
                    cell.Value = (double)f;
                    break;
                case bool b:
                    cell.Value = b ? "Да" : "Нет";
                    break;
                default:
                    string text = TruncateForExcel(value.ToString());
                    if (!string.IsNullOrEmpty(text))
                        cell.Value = text;
                    break;
            }
        }

        private static List<Guid> CollectMissingReferencedObjectIds(
            List<LoadedObjectNode> nodes,
            PilotLookupContext lookups,
            bool includeServiceAttributes)
        {
            var refIds = new HashSet<Guid>();

            foreach (var node in nodes)
            {
                var obj = node.DataObject;
                if (obj?.Type?.Attributes == null || obj.Attributes == null)
                    continue;

                foreach (var attr in obj.Type.Attributes)
                {
                    if (!includeServiceAttributes && attr.IsService)
                        continue;

                    var config = PilotAttributeHelper.ParsePilotConfig(attr.Configuration2());
                    string typeStr = attr.Type.ToString().ToLowerInvariant();
                    bool isObjectRef = config.Kind.Equals("Object", StringComparison.OrdinalIgnoreCase) || typeStr.Contains("guid");

                    if (!isObjectRef || !obj.Attributes.TryGetValue(attr.Name, out var rawVal) || rawVal == null)
                        continue;

                    foreach (var item in PilotAttributeHelper.FlattenValue(rawVal))
                    {
                        if (item is Guid g && g != Guid.Empty && !lookups.KnownObjects.ContainsKey(g))
                            refIds.Add(g);
                        else if (item != null && Guid.TryParse(item.ToString(), out Guid parsed) && parsed != Guid.Empty && !lookups.KnownObjects.ContainsKey(parsed))
                            refIds.Add(parsed);
                    }
                }
            }

            return refIds.ToList();
        }

        private static string TruncateForExcel(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            // Удаляем недопустимые в XML управляющие символы
            string cleaned = Regex.Replace(text, @"[\x00-\x08\x0B\x0C\x0E-\x1F]", string.Empty);

            // Жесткое ограничение Excel на длину текста в ячейке — 32 767 символов
            const int maxCellLen = 32760;
            if (cleaned.Length > maxCellLen)
                return cleaned.Substring(0, maxCellLen) + "...";

            return cleaned;
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Unnamed";

            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            return name.Trim().TrimEnd('.');
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Лист";

            string cleaned = Regex.Replace(name, @"[:\\/\?\*\[\]]", "_").Trim();
            if (string.IsNullOrEmpty(cleaned))
                cleaned = "Лист";

            return cleaned.Length > 31 ? cleaned.Substring(0, 31) : cleaned;
        }

        private static string MakeUniqueName(string baseName, HashSet<string> used, int maxLength)
        {
            string candidate = baseName.Length > maxLength ? baseName.Substring(0, maxLength) : baseName;
            if (used.Add(candidate))
                return candidate;

            int idx = 2;
            while (true)
            {
                string suffix = $" ({idx})";
                int allowedLen = Math.Max(1, maxLength - suffix.Length);
                string prefix = baseName.Length > allowedLen ? baseName.Substring(0, allowedLen) : baseName;
                candidate = prefix + suffix;
                if (used.Add(candidate))
                    return candidate;
                idx++;
            }
        }

        private class AttributeColumnInfo
        {
            public string AttrName { get; set; }
            public string Header { get; set; }
            public IAttribute AttributeMeta { get; set; }
            public PilotRefBookConfig Config { get; set; }
        }
    }

    public class PilotRefBookConfig
    {
        public string Kind { get; set; } = "None";
        public Guid Source { get; set; } = Guid.Empty;
        public Guid StateMachineId { get; set; } = Guid.Empty;
    }

    public class PilotLookupContext
    {
        public Dictionary<int, string> OrgUnitDisplayMap { get; } = new Dictionary<int, string>();
        public Dictionary<Guid, string> StateDisplayMap { get; } = new Dictionary<Guid, string>();
        public Dictionary<string, string> TypeDisplayMap { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<Guid, IDataObject> KnownObjects { get; } = new Dictionary<Guid, IDataObject>();

        public static PilotLookupContext Create(IObjectsRepository repo)
        {
            var ctx = new PilotLookupContext();

            try
            {
                var people = repo.GetPeople();
                var peopleDict = people?.Where(p => p != null).ToDictionary(p => p.Id, p => p) ?? new Dictionary<int, IPerson>();

                var orgUnits = repo.GetOrganisationUnits();
                if (orgUnits != null)
                {
                    foreach (var unit in orgUnits)
                    {
                        if (unit == null) continue;
                        int personId = unit.Person();
                        if (personId != -1 && peopleDict.TryGetValue(personId, out var person) && person != null)
                        {
                            string suffix = string.IsNullOrWhiteSpace(unit.Title) ? string.Empty : $" — [{unit.Title}]";
                            ctx.OrgUnitDisplayMap[unit.Id] = $"{person.DisplayName}{suffix}";
                        }
                        else if (!string.IsNullOrWhiteSpace(unit.Title))
                        {
                            ctx.OrgUnitDisplayMap[unit.Id] = unit.Title;
                        }
                    }
                }
            }
            catch { }

            try
            {
                var states = repo.GetUserStates();
                if (states != null)
                {
                    foreach (var st in states)
                    {
                        if (st == null) continue;
                        ctx.StateDisplayMap[st.Id] = string.IsNullOrWhiteSpace(st.Title) ? st.Name : st.Title;
                    }
                }
            }
            catch { }

            try
            {
                var types = repo.GetTypes();
                if (types != null)
                {
                    foreach (var t in types)
                    {
                        if (t != null && !string.IsNullOrEmpty(t.Name))
                            ctx.TypeDisplayMap[t.Name] = string.IsNullOrWhiteSpace(t.Title) ? t.Name : t.Title;
                    }
                }
            }
            catch { }

            return ctx;
        }
    }

    public static class PilotAttributeHelper
    {
        public static object FormatValue(object rawValue, IAttribute attr, PilotRefBookConfig config, PilotLookupContext lookups)
        {
            if (rawValue == null)
                return null;

            string kind = config?.Kind ?? "None";
            string typeStr = attr?.Type.ToString().ToLowerInvariant() ?? string.Empty;
            string attrNameLower = attr?.Name?.ToLowerInvariant() ?? string.Empty;

            bool isOrgUnit = kind.Equals("OrgUnit", StringComparison.OrdinalIgnoreCase) || typeStr.Contains("org") || attrNameLower.Contains("aup");
            bool isState = kind.Equals("State", StringComparison.OrdinalIgnoreCase) || typeStr.Contains("state") || (config != null && config.StateMachineId != Guid.Empty);
            bool isType = kind.Equals("Type", StringComparison.OrdinalIgnoreCase);
            bool isObjectRef = kind.Equals("Object", StringComparison.OrdinalIgnoreCase) || typeStr.Contains("guid");

            if (rawValue is IEnumerable enumerable && !(rawValue is string))
            {
                var parts = new List<string>();
                foreach (var item in enumerable)
                {
                    if (item == null) continue;
                    object singleFormatted = FormatSingleValue(item, isOrgUnit, isState, isType, isObjectRef, lookups);
                    if (singleFormatted != null)
                    {
                        string text = singleFormatted is DateTime dt ? dt.ToString("dd.MM.yyyy HH:mm") : singleFormatted.ToString();
                        if (!string.IsNullOrWhiteSpace(text))
                            parts.Add(text);
                    }
                }
                return string.Join("; ", parts);
            }

            return FormatSingleValue(rawValue, isOrgUnit, isState, isType, isObjectRef, lookups);
        }

        private static object FormatSingleValue(
            object item,
            bool isOrgUnit,
            bool isState,
            bool isType,
            bool isObjectRef,
            PilotLookupContext lookups)
        {
            if (item == null) return null;

            if (isOrgUnit && int.TryParse(item.ToString(), out int orgId))
            {
                return lookups.OrgUnitDisplayMap.TryGetValue(orgId, out string orgDisplay)
                    ? orgDisplay
                    : orgId.ToString();
            }

            if (isState && Guid.TryParse(item.ToString(), out Guid stateGuid))
            {
                return lookups.StateDisplayMap.TryGetValue(stateGuid, out string stateDisplay)
                    ? stateDisplay
                    : stateGuid.ToString();
            }

            if (isType)
            {
                string typeKey = item.ToString();
                return lookups.TypeDisplayMap.TryGetValue(typeKey, out string typeTitle)
                    ? typeTitle
                    : typeKey;
            }

            if (isObjectRef && Guid.TryParse(item.ToString(), out Guid objGuid))
            {
                if (lookups.KnownObjects.TryGetValue(objGuid, out var refObj) && refObj != null)
                    return GetObjectTitle(refObj);
                return objGuid.ToString();
            }

            return item;
        }

        public static IEnumerable<object> FlattenValue(object rawValue)
        {
            if (rawValue == null) yield break;

            if (rawValue is IEnumerable enumerable && !(rawValue is string))
            {
                foreach (var item in enumerable)
                    yield return item;
            }
            else
            {
                yield return rawValue;
            }
        }

        public static PilotRefBookConfig ParsePilotConfig(string xml)
        {
            var config = new PilotRefBookConfig();
            if (string.IsNullOrEmpty(xml)) return config;

            try
            {
                var doc = XDocument.Parse(xml);
                var root = doc.Root;
                if (root == null) return config;

                var smAttr = root.Attribute("StateMachineId") ?? root.Attribute("StateMachine");
                if (smAttr != null && Guid.TryParse(smAttr.Value, out Guid smId))
                    config.StateMachineId = smId;

                var refBookEl = root.Name.LocalName.Equals("ReferenceBook", StringComparison.OrdinalIgnoreCase)
                    ? root
                    : root.Descendants().FirstOrDefault(d => d.Name.LocalName.Equals("ReferenceBook", StringComparison.OrdinalIgnoreCase));

                var elemBookEl = root.Name.LocalName.Equals("ElementBook", StringComparison.OrdinalIgnoreCase)
                    ? root
                    : root.Descendants().FirstOrDefault(d => d.Name.LocalName.Equals("ElementBook", StringComparison.OrdinalIgnoreCase));

                XElement targetContainer = refBookEl ?? elemBookEl ?? root;
                config.Kind = targetContainer.Attribute("Kind")?.Value ?? targetContainer.Element("Kind")?.Value;
                if (string.IsNullOrEmpty(config.Kind))
                    config.Kind = elemBookEl != null ? "Object" : "None";
            }
            catch { }

            return config;
        }

        public static string GetObjectTitle(IDataObject obj)
        {
            if (obj == null) return "Пустой объект";
            if (obj.Attributes == null || !obj.Attributes.Any())
                return $"{obj.Type?.Title ?? obj.Type?.Name ?? "Элемент"} ({obj.Id})";

            if (obj.Type?.Attributes != null)
            {
                string[] humanKeywords = { "наименование", "название", "фио", "имя", "исполнитель", "сокращ", "заголовок", "организация", "компания" };
                foreach (var keyword in humanKeywords)
                {
                    var metaAttr = obj.Type.Attributes.FirstOrDefault(a =>
                        a.Title != null && a.Title.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

                    if (metaAttr != null && obj.Attributes.TryGetValue(metaAttr.Name, out var val) && val != null)
                    {
                        string str = val.ToString().Trim();
                        if (IsValidTitleString(str)) return str;
                    }
                }

                var obligatoryNames = obj.Type.Attributes.Where(a => a.IsObligatory).Select(a => a.Name).ToList();
                foreach (var name in obligatoryNames)
                {
                    if (obj.Attributes.TryGetValue(name, out var val) && val != null)
                    {
                        string str = val.ToString().Trim();
                        if (IsValidTitleString(str)) return str;
                    }
                }
            }

            foreach (var kvp in obj.Attributes)
            {
                if (kvp.Value == null) continue;
                string str = kvp.Value.ToString().Trim();
                if (IsValidTitleString(str)) return str;
            }

            return $"{obj.Type?.Title ?? obj.Type?.Name ?? "Элемент"} ({obj.Id})";
        }

        private static bool IsValidTitleString(string str)
        {
            if (string.IsNullOrWhiteSpace(str)) return false;
            if (str.Length < 1 || str.Length > 150) return false;
            if (Guid.TryParse(str, out _)) return false;
            if (DateTime.TryParse(str, out _)) return false;
            if (str.StartsWith("{") || str.StartsWith("<")) return false;
            return true;
        }
    }
}