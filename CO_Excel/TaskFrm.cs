using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;

using CO_Excel.CO_EXCEL_Service;

namespace CO_Excel
{
    public partial class TaskFrm : UserControl
    {


        public TaskFrm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 启动时,调用服务器程序，加载文件列表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TaskFrm_Load(object sender, EventArgs e)
        {

            publicFun.getInfo_treeview(this.treeView1, this.imageList1);

        }


        /// <summary>
        /// 双击打开文件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (publicFun.currentExcel_file != "")
            {//如果当前已经打开过文档,则关闭文档
                bool bol = false;
                Microsoft.Office.Interop.Excel.Workbook wb = null;
                MessageBox.Show("请先关闭并保存当前文档" + System.IO.Path.GetFileName(publicFun.currentExcel_file));
                //Globals.ThisAddIn.Application_DocumentClose(wb, ref bol);
                return;
            }

            if (treeView1.SelectedNode.ImageIndex == 1)//保证选中的是Excel节点且未被锁住(锁住该Index为2)
            {
                publicFun.Old_Excel_file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp", System.IO.Path.GetFileName(treeView1.SelectedNode.Tag.ToString()));
                string saveFilePath = publicFun.Old_Excel_file;
                if (!Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp")))
                {
                    Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp"));
                }

                if (File.Exists(saveFilePath))
                    File.Delete(saveFilePath);//存在即删除

                Stream sourceStream = publicFun.sc.OpenFile(treeView1.SelectedNode.Tag.ToString());//下载此文件

                if (sourceStream != null)
                {
                    if (sourceStream.CanRead)
                    {
                        using (FileStream fs = new FileStream(saveFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            const int bufferLength = 4096;//一部分一部分读取
                            byte[] myBuffer = new byte[bufferLength];
                            int count;
                            while ((count = sourceStream.Read(myBuffer, 0, bufferLength)) > 0)
                            {
                                //if (isExit == false)
                                //{
                                fs.Write(myBuffer, 0, count);
                                //}
                                //else//窗体已经关闭跳出循环
                                //{
                                //    break;
                                //}
                            }
                            fs.Close();
                            sourceStream.Close();
                        }
                    }
                }

                if (File.Exists(saveFilePath))
                {//在当前Excel中打开文件
                    Microsoft.Office.Interop.Excel.Application app = Globals.ThisAddIn.Application;
                    Microsoft.Office.Interop.Excel.Workbooks wbs = app.Workbooks;

                    wbs.Open(saveFilePath);

                    publicFun.currentExcel_file = treeView1.SelectedNode.Tag.ToString();
                    Globals.ThisAddIn.Application_DocumentOpen();

                }
            }
        }

        /// <summary>
        /// 刷新任务栏
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void 刷新RToolStripMenuItem_Click(object sender, EventArgs e)
        {
            publicFun.getInfo_treeview(this.treeView1, this.imageList1);
        }

        private void 选择SToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (treeView1.SelectedNode.ImageIndex == 0)//文件夹
            {
                publicFun.treenode_selected_path = treeView1.SelectedNode.Tag.ToString();

            }
            else if (treeView1.SelectedNode.ImageIndex == 1 || treeView1.SelectedNode.ImageIndex == 2)//文件 或锁定的文件
            {
                publicFun.treenode_selected_path = Path.GetDirectoryName(treeView1.SelectedNode.Tag.ToString());
            }
            //publicFun.treenode_selected_path = treeView1.SelectedNode.Tag.ToString();

        }

    }
}
