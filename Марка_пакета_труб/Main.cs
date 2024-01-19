using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Марка_пакета_труб
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            ElementCategoryFilter conduitCategoryFiltet = new ElementCategoryFilter(BuiltInCategory.OST_ConduitFitting);
            ElementClassFilter conduitInstancesFilter = new ElementClassFilter(typeof(FamilyInstance));

            LogicalAndFilter conduitilter = new LogicalAndFilter(conduitCategoryFiltet, conduitInstancesFilter);

            string user = null;

            FilteredWorksetCollector worksets
              = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset);

            WorksetTable worksetTable = doc.GetWorksetTable();
            WorksetId activeId = worksetTable.GetActiveWorksetId();
            string st = activeId.ToString();

            foreach (var workset in worksets)
            {
                if(workset.Id== activeId)
                {
                    user = workset.Owner;
                    break;
                }

            }

            var pipePackages = new FilteredElementCollector(doc)
                .WherePasses(conduitilter)
                .Where(x => x.get_Parameter(BuiltInParameter.EDITED_BY).AsString() == user)
                .Where(x => x.LookupParameter("Число труб") != null)
                .Cast<FamilyInstance>()
                .ToList();

            var cableTrays = new FilteredElementCollector(doc)
                .OfClass(typeof(CableTray))
                .Where(x => x.get_Parameter(BuiltInParameter.EDITED_BY).AsString() == user)
                .Cast<CableTray>()
                .ToList();

            string infoPipePackage = string.Empty;
            string infoCableTray = string.Empty;

            foreach (var pipePackage in pipePackages)
            {
                double zmax = Math.Round((UnitUtils.ConvertFromInternalUnits(pipePackage.get_BoundingBox(doc.ActiveView).Max.Z, UnitTypeId.Meters)), 2);
                double zmin = Math.Round((UnitUtils.ConvertFromInternalUnits(pipePackage.get_BoundingBox(doc.ActiveView).Min.Z, UnitTypeId.Meters)), 2);

                string zmaxString = $"{zmax:f3}";
                string zminString = $"{zmin:f3}";

                if (zmax >= 0)
                {
                    zmaxString = $"+{zmax:f3}";
                }
                if (zmin >= 0)
                {
                    zminString = $"+{zmin:f3}";
                }

                var name = pipePackage.Name;
                string type = null;

                switch (name)
                {
                    case string a when name.Contains("Пакет труб ду-25"):
                        type = "25";
                        break;
                    case string a when name.Contains("Пакет труб ду-40"):
                        type = "40";
                        break;
                    case string a when name.Contains("профилей 50"):
                        type = "50☐";
                        break;
                    case string a when name.Contains("ду-50"):
                        type = "50";
                        break;
                    case string a when name.Contains("профилей 60"):
                        type = "60☐";
                        break;
                    case string a when name.Contains("Гофрированная труба 25"):
                        type = "25Гф";
                        break;
                    case string a when name.Contains("Гофрированная труба 16"):
                        type = "16Гф";
                        break;
                    case string a when name.Contains("Гофрированная труба 40"):
                        type = "40Гф";
                        break;
                }
                ConnectorSet connectorSet;
                try
                {
                    connectorSet = pipePackage.MEPModel.ConnectorManager.Connectors;
                }
                catch 
                {

                    TaskDialog.Show("Ошибка", $"Проверте актуальность версии семейства \"{pipePackage.Name}\"");
                    continue;
                }

                List<Connector> connectorList = new List<Connector>();
                foreach (Connector connector in connectorSet)
                {
                    connectorList.Add(connector);
                }


                if (Math.Abs(connectorList[0].CoordinateSystem.BasisX.Z) <= 0.3 && Math.Abs(connectorList[0].CoordinateSystem.BasisY.Z) <= 0.3)
                {
                    infoPipePackage = $"{type} с отм. {zminString} на отм. {zmaxString}\n";
                }
                else
                {
                    infoPipePackage = $"{type} на отм. {zminString}\n";
                }

                if (pipePackage.LookupParameter("ADSK_Примечание") == null)
                {
                    TaskDialog.Show("Ошибка", $"Проверте наличие параметра \"ADSK_Примечание\" у пакета типа: {pipePackage.Symbol.Name}");
                    return Result.Cancelled;
                }
                using (Transaction ts = new Transaction(doc, "Set parameters"))
                {
                    ts.Start();

                    Parameter commentParameter = pipePackage.LookupParameter("ADSK_Примечание");
                    commentParameter.Set(infoPipePackage);
                    ts.Commit();

                }
            }


            foreach (var cableTray in cableTrays)
            {
                double zmax = Math.Round((UnitUtils.ConvertFromInternalUnits(cableTray.get_BoundingBox(doc.ActiveView).Max.Z, UnitTypeId.Meters)), 2);
                double zmin = Math.Round((UnitUtils.ConvertFromInternalUnits(cableTray.get_BoundingBox(doc.ActiveView).Min.Z, UnitTypeId.Meters)), 2);

                string zmaxString = $"{zmax:f3}";
                string zminString = $"{zmin:f3}";

                if (zmax >= 0)
                {
                    zmaxString = $"+{zmax:f3}";
                }
                if (zmin >= 0)
                {
                    zminString = $"+{zmin:f3}";
                }

                string size =cableTray.LookupParameter("Размер").AsString();

                ConnectorSet connectorSet= cableTray.ConnectorManager.Connectors; ;


                List<Connector> connectorList = new List<Connector>();
                foreach (Connector connector in connectorSet)
                {
                    connectorList.Add(connector);
                }


                if (Math.Abs(connectorList[0].CoordinateSystem.BasisX.Z) <= 0.3 && Math.Abs(connectorList[0].CoordinateSystem.BasisY.Z) <= 0.3)
                {
                    infoCableTray = $"{size} с отм. {zminString} на отм. {zmaxString}\n";
                }
                else
                {
                    infoCableTray = $"{size} на отм. {zminString}\n";
                }

                if (cableTray.LookupParameter("ADSK_Примечание") == null)
                {
                    TaskDialog.Show("Ошибка", $"Проверте наличие параметра \"ADSK_Примечание\" у пакета типа: {cableTray.Name}");
                    return Result.Cancelled;
                }
                using (Transaction ts = new Transaction(doc, "Set parameters"))
                {
                    ts.Start();

                    Parameter commentParameter = cableTray.LookupParameter("ADSK_Примечание");
                    commentParameter.Set(infoCableTray);
                    ts.Commit();

                }
            }


            return Result.Succeeded;
        }
    }
}
