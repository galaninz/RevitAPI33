using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI33
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            IList<Reference> selectedElementRefList = uidoc.Selection.PickObjects(ObjectType.Element, "Выберите трубы");

            double info = 0;

            var categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_PipeCurves));

            using (Transaction ts = new Transaction(doc, "Add parameter"))
            {
                ts.Start();
                CreateSharedParameter(uiapp.Application, doc, "Длина с запасом 10%", categorySet, BuiltInParameterGroup.PG_MECHANICAL, true);
                ts.Commit();
            }

            foreach (var element in selectedElementRefList)
            {
                var selectedElement = doc.GetElement(element);
                if (selectedElement is Pipe)
                {
                    Parameter lengthParameter = selectedElement.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (lengthParameter.StorageType == StorageType.Double)
                    {
                        double lengthValue = UnitUtils.ConvertToInternalUnits(lengthParameter.AsDouble(), UnitTypeId.Meters);
                        using (Transaction ts = new Transaction(doc, "Set parameters"))
                        {
                            ts.Start();
                            var familyInstance = (FamilyInstance)selectedElement;
                            Parameter parameter = familyInstance.LookupParameter("Длина с запасом 10%");
                            parameter.Set($"{lengthValue * 1.1}");
                            ts.Commit();
                        }
                        info += lengthValue;
                    }
                }

            }

            return Result.Succeeded;
        }


        private void CreateSharedParameter(Application application, Document doc, string parameterName, CategorySet categorySet, BuiltInParameterGroup builtInParameterGroup, bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return;
            }

            Definition definition = definitionFile
                .Groups.SelectMany(group => group.Definitions)
                .FirstOrDefault(def => def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден указанный параметр");
                return;
            }

            Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            map.Insert(definition, binding, builtInParameterGroup);
        }

    }
}
