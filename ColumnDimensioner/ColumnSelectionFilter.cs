using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;

namespace ColumnDimensioner
{
    class ColumnSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance
                && null != elem.Category
                && elem.Category.Id == new ElementId(BuiltInCategory.OST_StructuralColumns))
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
}
