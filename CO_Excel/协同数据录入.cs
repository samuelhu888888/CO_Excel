using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.ServiceModel;
using System.IO;
using CO_Excel.CO_EXCEL_Service;
using System.Windows.Forms;

namespace CO_Excel
{
    public partial class 协同数据录入
    {
        private void 协同数据录入_Load(object sender, RibbonUIEventArgs e)
        {

        }

        /// <summary>
        /// 打开或关闭任务栏
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.MyTaskPane.Visible == false)
            {
                ThisAddIn.MyTaskPane.Visible = true;
            }
            else
            {
                ThisAddIn.MyTaskPane.Visible = false;
            }
        }

        /// <summary>
        /// 刷新工具栏
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            //publicFun.getInfo_treeview(this.treeView1, this.imageList1);
        }

        /// <summary>
        /// 提交文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            if (publicFun.currentExcel_file == "")
            {
                MessageBox.Show("当前文件无法提交");
                return;
            }
            CO_Excel.CO_EXCEL_Service.IService1 clientUpload = new Service1Client();

            CO_Excel.CO_EXCEL_Service.RemoteFileInfo rfi = new CO_EXCEL_Service.RemoteFileInfo();


            rfi.FileName = publicFun.currentExcel_file;
            //打开前先关闭
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            //Microsoft.Office.Interop.Excel.Workbooks wbs = app.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook wb = app.ActiveWorkbook;
            wb.Close();
            //wbs.Close();
            //wbs.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp", "test99998888.xls"));
            //wbs.Close();
            //Microsoft.Office.Interop.Excel.Workbook workbook = Globals.ThisAddIn.Application.Workbooks.get_Item(System.IO.Path.GetFileName(publicFun.currentExcel_file));

            //workbook.Close(false, Type.Missing, Type.Missing);

            //app.Quit();
            using (rfi.FileByteStream = File.OpenRead(publicFun.Old_Excel_file))
            {
                clientUpload.upLoad(rfi);
            }

            //button2.Enabled = false;
            publicFun.currentExcel_file = "";//避免该文档二次提交,限制只有再次打开 才能提交 
        }

        private void Help_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button3_Click_1(object sender, RibbonControlEventArgs e)
        {
            //publicFun.sc.UnsetReadOnly(publicFun.sc.config_path);
        }

        /// <summary>
        /// 将新文档录入库中
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            if (publicFun.treenode_selected_path == "")
            {

                MessageBox.Show("请在点击确定后,在左侧导航栏中,通过右键菜单选择一个入库节点");
                return;
            }

            CO_Excel.CO_EXCEL_Service.IService1 clientUpload = new Service1Client();
            CO_Excel.CO_EXCEL_Service.RemoteFileInfo rfi = new CO_EXCEL_Service.RemoteFileInfo();
            //1 当前文档保存
            Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
            Microsoft.Office.Interop.Excel.Workbook wb = app.ActiveWorkbook;
            wb.Save();

            string localpath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;//要读取的本地文件的地址

            rfi.FileName = Path.Combine(publicFun.treenode_selected_path, System.IO.Path.GetFileName(localpath));//存放到数据服务器的地址
            wb.Close();
            //2 读取当前文档为stream


            using (rfi.FileByteStream = File.OpenRead(localpath))
            {
                //3 存储到服务器中

                clientUpload.upLoad(rfi);
            }
            publicFun.treenode_selected_path = "";//限制只能入库一次
        }


    }
    [MessageContract]
    public class RemoteFileInfo : IDisposable
    {
        [MessageHeader(MustUnderstand = true)]
        public string FileName { get; set; }

        [MessageBodyMember(Order = 1)]
        public System.IO.Stream FileByteStream { get; set; }

        public void Dispose()
        {
            if (FileByteStream != null)
            {
                FileByteStream.Close();
                FileByteStream = null;
            }
        }
    }
}
