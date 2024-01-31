using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Последовательное_подключение_светильников
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            FamilyInstance luminaire=null;
            List<FamilyInstance> luminaires=new List<FamilyInstance>();

            try
            {
                do
                {
                    luminaire = PickLuminaire(commandData);
                    luminaires.Add(luminaire);

                } while (luminaires!=null);
            }
            catch (Exception)
            { }


            if (luminaires.Count < 2)
            {
                TaskDialog.Show("Ошибка", "Не выбрано минимальное количество электрооборудования");
                return Result.Succeeded;
            }


            List<ElementId> luminairesIdList = new List<ElementId>();
            foreach (var lum in luminaires)
            {
                luminairesIdList.Add(lum.Id);
            }

            List<ElectricalSystem> electricalSystemList = new List<ElectricalSystem>();

            using (Transaction ts = new Transaction(doc, "Создание цепи и подключение к панели"))
            {
                ts.Start();
                foreach (var luminairesId in luminairesIdList)
                {
                    FamilyInstance el = doc.GetElement(luminairesId) as FamilyInstance;

                    if (el.MEPModel.ElectricalSystems != null)//Проверка  наличие цепи
                    {
                        foreach (ElectricalSystem es in el.MEPModel.ElectricalSystems)
                        {
                            electricalSystemList.Add(es);
                            break;
                        }
                        continue;

                    }
                    List<ElementId> luminairesIdList_ = new List<ElementId>();
                    luminairesIdList_.Clear();
                    luminairesIdList_.Add(luminairesId);
                    ElectricalSystem electricalSystem = ElectricalSystem.Create(doc, luminairesIdList_, ElectricalSystemType.PowerCircuit);
                    electricalSystemList.Add(electricalSystem);
                }
                int j = 0;
                for (int i = 1; i < electricalSystemList.Count; i++)
                {

                    if (electricalSystemList[i].BaseEquipment != null)
                    {
                        Element e = electricalSystemList[i].BaseEquipment;
                        ElementSet elementSet = new ElementSet();
                        elementSet.Insert(e);
                        electricalSystemList[i].RemoveFromCircuit(elementSet);
                    }
                    electricalSystemList[i].SelectPanel(luminaires[j]);
                    j++;
                }
                ts.Commit();
            }
            return Result.Succeeded;

        }

        public static List<FamilyInstance> PickLuminaires(ExternalCommandData commandData, string message = "Выберите светильники")
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            IList<Reference> elemenRef = uidoc.Selection.PickObjects(ObjectType.Element, "Выберете светильник");
            var elementList = new List<FamilyInstance>();

            foreach (var elem in elemenRef)
            {
                Element element = doc.GetElement(elem);
                if (element.Category.Name == "Электрооборудование")
                {
                    FamilyInstance fiElement = (FamilyInstance)element;
                    elementList.Add(fiElement);
                }

            }

            return elementList;
        }

        public static FamilyInstance PickLuminaire(ExternalCommandData commandData, string message = "Выберите светильники")
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            Reference selEl = null;

            FamilyInstance fiElement= null;

            selEl = uidoc.Selection.PickObject(ObjectType.Element, "Выберете светильник");
            if (selEl != null)
            {
                Element element = doc.GetElement(selEl);
                if (element.Category.Name == "Электрооборудование")
                {
                    fiElement = (FamilyInstance)element;
                }
            }
            return fiElement;
        }
    }
}
