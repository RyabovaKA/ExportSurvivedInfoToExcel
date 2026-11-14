using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.Menu;

namespace PilotExcelExporter
{
    [Export(typeof(IMenu<ObjectsViewContext>))]
    public class ExportToExcelContextMenuPlugin : IMenu<ObjectsViewContext>
    {
        private const string EXPORT_MENU_ITEM_NAME = "ExportAccessibleToExcelMenuItem";
        private static readonly Guid RootObjectId = new Guid("00000001-0001-0001-0001-000000000001");

        private readonly IObjectsRepository _objectsRepository;

        [ImportingConstructor]
        public ExportToExcelContextMenuPlugin(IObjectsRepository objectsRepository)
        {
            _objectsRepository = objectsRepository;
        }

        public void Build(IMenuBuilder builder, ObjectsViewContext context)
        {
            if (builder.Count > 0)
                builder.AddSeparator(builder.Count);

            builder.AddItem(EXPORT_MENU_ITEM_NAME, builder.Count)
                   .WithHeader("Выгрузить доступные элементы в Excel...");
        }

        public void OnMenuItemClick(string name, ObjectsViewContext context)
        {
            if (name != EXPORT_MENU_ITEM_NAME)
                return;

            var startIds = (context?.SelectedObjects != null && context.SelectedObjects.Any())
                ? context.SelectedObjects.Where(o => o != null).Select(o => o.Id).Distinct().ToList()
                : new List<Guid> { RootObjectId };

            var optionsDialog = new ExportOptionsDialog();
            if (optionsDialog.ShowDialog() != true)
                return;

            var progressWindow = new ExportProgressWindow(
                _objectsRepository,
                startIds,
                optionsDialog.SelectedFolder,
                optionsDialog.SeparateFilePerType,
                optionsDialog.IncludeServiceAttributes,
                optionsDialog.ResolveReferenceObjects);

            progressWindow.ShowDialog();

            if (progressWindow.ExportedFilesCount > 0)
            {
                var result = MessageBox.Show(
                    $"Выгрузка завершена!\n\nЗагружено объектов: {progressWindow.TotalLoadedObjects}\nСоздано файлов Excel: {progressWindow.ExportedFilesCount}\nПапка: {optionsDialog.SelectedFolder}\n\nОткрыть папку с результатами?",
                    "Выгрузка в Excel",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        Process.Start("explorer.exe", optionsDialog.SelectedFolder);
                    }
                    catch { }
                }
            }
            else if (!progressWindow.WasCancelled && string.IsNullOrEmpty(progressWindow.ErrorMessage))
            {
                MessageBox.Show("Доступные для выгрузки элементы не найдены.", "Выгрузка в Excel", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else if (!string.IsNullOrEmpty(progressWindow.ErrorMessage))
            {
                MessageBox.Show($"Ошибка при выгрузке:\n{progressWindow.ErrorMessage}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}