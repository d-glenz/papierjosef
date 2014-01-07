using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace DerPapierjosef
{
    public partial class ThisAddIn
    {
        private JosefPane myTaskPane;
        private Microsoft.Office.Tools.CustomTaskPane taskPane;

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return taskPane;
            }
        }

        public JosefPane MyJosefPane
        {
            get
            {
                return myTaskPane;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            myTaskPane=new JosefPane();
            taskPane = this.CustomTaskPanes.Add(myTaskPane, "Papierjosef");
            taskPane.Visible = false;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
