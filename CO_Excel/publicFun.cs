using CO_Excel.CO_EXCEL_Service;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CO_Excel
{
    public class publicFun
    {
        public static Service1Client sc = new Service1Client();
        /// <summary>
        /// 当前正在使用文档的全局变量---保存在服务器中的路径
        /// </summary>
        public static string currentExcel_file = "";
        /// <summary>
        /// 保存在当前客户端的路径
        /// </summary>
        public static string Old_Excel_file = "";
        ///// <summary>
        ///// 记录唯一的任务栏
        ///// </summary>
        //public static Microsoft.Office.Tools.CustomTaskPane onlyPane;

        /// <summary>
        /// 文件入库时,需要用户在树节点上选择入库文件的位置
        /// 入库完一次,则清空
        /// 入库时,要检查此变量是否为空,为空则不能入库
        /// </summary>
        public static string treenode_selected_path = "";

        /// <summary>
        /// 解析从服务器返回的treeview信息,并显示在treeview上
        /// </summary>
        /// <param name="treeView1"></param>
        /// <param name="imageList1"></param>
        public static void getInfo_treeview(TreeView treeView1, ImageList imageList1)
        {
            treeView1.Nodes.Clear();

            treeView1.ImageList = imageList1;
            treeView1.Nodes.Add("协同数据录入");

            string[] aa = sc.GetTreeInfo();
            TreeNode init_tn = treeView1.Nodes[0];
            for (int i = 0; i < aa.ToList().Count / 5; i++)
            {

                TreeNode tn = new TreeNode();
                tn.Text = aa[5 * i + 1];
                tn.Tag = aa[5 * i + 2];
                tn.ImageIndex = Convert.ToInt32(aa[5 * i + 3]);
                tn.SelectedImageIndex = Convert.ToInt32(aa[5 * i + 3]);
                if (init_tn.Level < Convert.ToInt32(aa[5 * i]))
                {
                    init_tn.Nodes.Add(tn);
                    if (aa[5 * i + 4] == "locked" && Convert.ToInt32(aa[5 * i + 3]) == 1)
                    {//是文件,且文件被锁住
                        tn.ImageIndex = 2;
                        tn.SelectedImageIndex = 2;
                    }
                    init_tn = tn;
                }
                else if (init_tn.Level == Convert.ToInt32(aa[5 * i]))
                {
                    init_tn.Parent.Nodes.Add(tn);
                    if (aa[5 * i + 4] == "locked" && Convert.ToInt32(aa[5 * i + 3]) == 1)
                    {//是文件,且文件被锁住
                        tn.ImageIndex = 2;
                        tn.SelectedImageIndex = 2;
                    }
                    init_tn = tn;
                }
                else
                {
                    while (init_tn.Level > Convert.ToInt32(aa[5 * i]))
                    {
                        init_tn = init_tn.Parent;
                    }
                    init_tn.Parent.Nodes.Add(tn);
                    if (aa[5 * i + 4] == "locked" && Convert.ToInt32(aa[5 * i + 3]) == 1)
                    {//是文件,且文件被锁住
                        tn.ImageIndex = 2;
                        tn.SelectedImageIndex = 2;
                    }
                    init_tn = tn;
                }

            }


            treeView1.ExpandAll();
        }


    }
}
