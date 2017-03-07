using System;

namespace WordAddInDemo
{

    public partial class ThisAddIn
    {
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        private static void OnThisAddInStartup(object sender, EventArgs e)
        {
            //TODO: user defined logic.
        }

        private static void OnThisAddInShutdown(object sender, EventArgs e)
        {
            //TODO: user defined logic.
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            Startup += OnThisAddInStartup;
            Shutdown += OnThisAddInShutdown;
        }
        
        #endregion
    }
}
