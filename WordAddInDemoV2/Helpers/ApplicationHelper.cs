using Microsoft.Office.Interop.Word;

namespace WordAddInDemoV2.Helpers
{
    public class ApplicationHelper
    {
        public static Range GetCurrentSelectionRange()
        {
            var start = Globals.ThisAddIn.Application.Selection.Start;
            var end = Globals.ThisAddIn.Application.Selection.End;
            return Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
        }

        public static Range GetDocumentRange()
        {
            MoveCursorToEnd();
            var start = Globals.ThisAddIn.Application.Selection.Start;
            var end = Globals.ThisAddIn.Application.Selection.End;
            return Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
        }

        public static void MoveCursorToEnd()
        {
            Globals.ThisAddIn.Application.Selection.EndKey(WdUnits.wdStory);
        }
    }
}
