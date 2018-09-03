using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace high
{
    public partial class ThisAddIn
    {
        public PowerPoint.SlideRange range = null;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            this.range = SldRange;
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
