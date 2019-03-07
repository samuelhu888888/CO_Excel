using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace CO_Excel
{
    public partial class ThisAddIn
    {
        public static Microsoft.Office.Tools.CustomTaskPane MyTaskPane;
        private static TaskFrm myControl;

        public Dictionary<string, Microsoft.Office.Tools.CustomTaskPane> TaskPanels =
            new Dictionary<string, Microsoft.Office.Tools.CustomTaskPane>();


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            this.Application.WorkbookActivate += new Excel.AppEvents_WorkbookActivateEventHandler(Application_DocumentOpen2);// += new WorkbookEvents_NewEventHandler(Application_DocumentOpen);
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_DocumentOpen2);
            this.Application.WorkbookBeforeClose += new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_DocumentClose);
        }

        void Application_WorkbookActivate(Excel.Workbook Wb)
        {

        }

        public void Application_DocumentOpen()
        {

            //Excel里解决侧边栏只有一个的问题
            if (Globals.ThisAddIn.Application.ActiveWindow != null)
            {
                int hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
                if (TaskPanels.TryGetValue(hwnd.ToString(), out MyTaskPane))
                {
                    MyTaskPane = TaskPanels[hwnd.ToString()];
                    MyTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    MyTaskPane.Visible = true;
                }
                else
                {
                    myControl = new TaskFrm();
                    MyTaskPane =  this.CustomTaskPanes.Add(myControl, "测试任务空格");
                    MyTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;
                    MyTaskPane.Visible = true;

                    TaskPanels.Add(hwnd.ToString(), MyTaskPane);
                }
            }
        }
        public void Application_DocumentOpen2(Excel.Workbook wk)
        {
            //Excel里解决侧边栏只有一个的问题
            if (Globals.ThisAddIn.Application.ActiveWindow != null)
            {
                int hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
                if (TaskPanels.TryGetValue(hwnd.ToString(), out MyTaskPane))
                {
                    MyTaskPane = TaskPanels[hwnd.ToString()];
                    MyTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    MyTaskPane.Visible = true;
                }
                else
                {
                    myControl = new TaskFrm(); 
                    MyTaskPane =  this.CustomTaskPanes.Add(myControl, "测试任务空格");
                    MyTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionLeft;

                    MyTaskPane.Visible = true;
                    TaskPanels.Add(hwnd.ToString(), MyTaskPane);
                }
            }
        }
        /// <summary>
        /// 关闭前将当前文档解锁
        /// </summary>
        /// <param name="wk"></param>
        /// <param name="bl"></param>
        public void Application_DocumentClose(Excel.Workbook wk, ref bool bl)
        {
            publicFun.sc.unLockFile(publicFun.currentExcel_file);
            publicFun.currentExcel_file = "";
            //this.CustomTaskPanes.Remove(myControl);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
