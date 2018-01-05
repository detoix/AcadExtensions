using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.IO;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Globalization;
using Autodesk.AutoCAD.Geometry;

namespace AcadExtensions
{
    public class Commands
    {
        [CommandMethod("wk")]
        public void GetCoordsForCircles()
        {
            this.GetCoords<Circle>("CIRCLE");
        }

        [CommandMethod("wb")]
        public void GetCoordsForBlocks()
        {
            this.GetCoords<BlockReference>("INSERT");
        }

        private void GetCoords<T>(string type) where T : class
        {
            var document = Application.DocumentManager.MdiActiveDocument;
            var tmpFile = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".csv";

            using (Transaction t = document.Database.TransactionManager.StartTransaction())
            {
                var promptPointResult = document.Editor.GetPoint("\nSelect center point: ");
                var point = promptPointResult.Value;

                if (promptPointResult.Status == PromptStatus.Cancel) return;

                var filter = new SelectionFilter(new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, type)
                });

                var promptSelectionResult = document.Editor.GetSelection(filter);
                var selectionSet = promptSelectionResult.Value;

                if (promptSelectionResult.Status == PromptStatus.Cancel) return;

                var list = new List<T>();
                foreach (SelectedObject selectedObject in selectionSet)
                {
                    if (selectedObject == null) { continue; }
                    var item = t.GetObject(selectedObject.ObjectId, OpenMode.ForRead) as T;
                    list.Add(item);
                }

                using (var w = new StreamWriter(tmpFile))
                {
                    w.WriteLine("X,Y");

                    foreach (var item in list.OrderByDescending(e => e.GetPoint<T>().X).ThenByDescending(e => e.GetPoint<T>().Y))
                    {
                        var x = Math.Round(item.GetPoint<T>().X - point.X, 2);
                        var y = Math.Round(item.GetPoint<T>().Y - point.Y, 2);
                        w.WriteLine(string.Format("{0}{1}{2}", 
                            x.ToString(CultureInfo.CreateSpecificCulture("pl-PL")),
                            CultureInfo.CurrentCulture.TextInfo.ListSeparator, 
                            y.ToString(CultureInfo.CreateSpecificCulture("pl-PL"))));
                    }
                }
                t.Commit();
            }
            Process.Start(tmpFile);
        }
    }

    public static class ExtensionMethods
    {
        public static Point3d GetPoint<T>(this T item)
        {
            return ((item as Circle)?.Center ?? (item as BlockReference)?.Position).Value;
        }
    }

    public class AdskCommandHandler : System.Windows.Input.ICommand
    {
        public bool CanExecute(object parameter)
        {
            return true;
        }

        public event EventHandler CanExecuteChanged;
        public void Execute(object parameter)
        {
            RibbonButton rbnBtn = parameter as RibbonButton;
            if (rbnBtn != null)
            {
                //Execute command specified in ribbon button parameter
                Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute((String)rbnBtn.CommandParameter, true, false, true);
            }

        }
    }
}
