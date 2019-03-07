using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace CO_EXCEL_SERVICE
{

    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "Service1" in both code and config file together.
    public class Service1 : IService1
    {
        public static string config_path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "CO_Excel");
        public static List<string> list_tree_info = new List<string>();


        public string GetData(int value)
        {
            return string.Format("You entered: {0}", value);
        }

        public CompositeType GetDataUsingDataContract(CompositeType composite)
        {
            if (composite == null)
            {
                throw new ArgumentNullException("composite");
            }
            if (composite.BoolValue)
            {
                composite.StringValue += "Suffix";
            }
            return composite;
        }

        public bool lockFile(string filepath)
        {
            try
            {
                // 文件非只读
                if ((System.IO.File.GetAttributes(filepath).ToString().IndexOf("ReadOnly") != -1))
                {
                    File.SetAttributes(filepath, FileAttributes.ReadOnly); //若文件为只读状态，将其更改为非只读状态
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }

        }

        public bool unLockFile(string filepath)
        {
            try
            {
                // 文件非只读
                if ((System.IO.File.GetAttributes(filepath).ToString().IndexOf("ReadOnly") == 0))
                {
                    File.SetAttributes(filepath, FileAttributes.Normal); //若文件为只读状态，将其更改为非只读状态
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        public Stream OpenFile(string fileTag)
        {
            try
            {
                //检测文件是否锁定,是只读则不读取
                if (System.IO.File.GetAttributes(fileTag).ToString().IndexOf("ReadOnly") == 0)
                {
                    return null;
                }
                //没有,则将文件锁定

                //将文件传递到客户端 
                if (!File.Exists(fileTag))//判断文件是否存在
                {
                    return null;
                }

                Stream myStream = File.OpenRead(fileTag);

                // 文件非只读
                if ((System.IO.File.GetAttributes(fileTag).ToString().IndexOf("ReadOnly") != 0))
                {
                    File.SetAttributes(fileTag, FileAttributes.ReadOnly); //若文件为只读状态，将其更改为非只读状态
                }


                return myStream;


            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 文件入库
        /// </summary>
        /// <param name="inputFile"></param>
        /// <returns></returns>
        public bool Storage_File(RemoteFileInfo inputFile)
        {
            try
            {

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }


        /// <summary>
        /// 上传文件到服务器
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public void upLoad(RemoteFileInfo request)
        {

            string oldFilePath = request.FileName;
            Stream sourceStream = request.FileByteStream;

            try
            {
                //如果文件存在,则删除文件
                if (File.Exists(oldFilePath))//判断文件是否存在
                {
                    File.Delete(oldFilePath);
                }
                //写到文件
                if (sourceStream != null)
                {
                    if (sourceStream.CanRead)
                    {
                        using (FileStream fs = new FileStream(oldFilePath, FileMode.Create, FileAccess.Write, FileShare.None))
                        {
                            const int bufferLength = 4096;//一部分一部分读取
                            byte[] myBuffer = new byte[bufferLength];
                            int count;
                            while ((count = sourceStream.Read(myBuffer, 0, bufferLength)) > 0)
                            {
                                fs.Write(myBuffer, 0, count);
                            }
                            fs.Close();
                            sourceStream.Close();
                        }
                    }
                }
                return;
            }
            catch (Exception)
            {
                return;
            }
        }

        /// <summary>
        /// string 存储文件路径
        /// List<string> 存储文件名称 文件ImageIndex
        /// </summary>
        /// <param name="dic"></param>
        /// <returns></returns>
        public List<string> GetTreeInfo()//
        {
            try
            {
                list_tree_info = new List<string>();

                TreeView tv = new TreeView();
                TreeNode root = tv.Nodes.Add("协同数据录入");

                LoadFilesAndDirectoriesToTree(config_path, root.Nodes);

                return list_tree_info;
            }
            catch (Exception)
            {
                return new List<string>();
            }
        }

        #region 文件夹数据获取
        /// <summary>
        /// 迭归获取文件夹内文件及信息
        /// </summary>
        /// <param name="path"></param>
        /// <param name="treeNodeCollection"></param>
        private void LoadFilesAndDirectoriesToTree(string path, TreeNodeCollection treeNodeCollection)
        {
            //1.先根据路径获取所有的子文件和子文件夹
            string[] files = Directory.GetFiles(path);
            string[] dirs = Directory.GetDirectories(path);
            //2.把所有的子文件与子目录加到TreeView上。
            foreach (string item in files)
            {
                //把每一个子文件加到TreeView上
                TreeNode fnode = new TreeNode();
                fnode.Text = System.IO.Path.GetFileName(item);
                fnode.Tag = item;
                fnode.ImageIndex = 1;
                fnode.SelectedImageIndex = 1;


                treeNodeCollection.Add(fnode);

                list_tree_info.Add(fnode.Level.ToString());
                list_tree_info.Add(System.IO.Path.GetFileName(item));
                list_tree_info.Add(item);
                list_tree_info.Add("1");
                if (System.IO.File.GetAttributes(item).ToString().IndexOf("ReadOnly") != -1)
                {
                    list_tree_info.Add("locked");
                }
                else
                {
                    list_tree_info.Add("unlocked");
                }

            }
            //文件夹
            foreach (string item in dirs)
            {
                if (System.IO.Path.GetFileName(item) == "bk")
                    continue;
                TreeNode fnode = treeNodeCollection.Add(Path.GetFileName(item));
                fnode.Text = System.IO.Path.GetFileName(item);
                fnode.Tag = item;
                fnode.ImageIndex = 0;
                fnode.SelectedImageIndex = 0;

                list_tree_info.Add(fnode.Level.ToString());
                list_tree_info.Add(System.IO.Path.GetFileName(item));
                list_tree_info.Add(item);
                list_tree_info.Add("0");
                list_tree_info.Add("unlocked");

                //由于目录，可能下面还存在子目录，所以这时要对每个目录再次进行获取子目录与子文件的操作
                //这里进行了递归
                LoadFilesAndDirectoriesToTree(item, fnode.Nodes);
            }

        }


        public void UnsetReadOnly(string path)
        {
            //1.先根据路径获取所有的子文件和子文件夹
            string[] files = Directory.GetFiles(path);
            string[] dirs = Directory.GetDirectories(path);
            //2.把所有的子文件与子目录加到TreeView上。
            foreach (string item in files)
            {
                //把每一个子文件加到TreeView上 
                if (System.IO.File.GetAttributes(item).ToString().IndexOf("ReadOnly") != -1)
                {
                    //list_tree_info.Add("locked");
                    unLockFile(item);
                }
            }
            //文件夹
            foreach (string item in dirs)
            {
                File.SetAttributes(item, System.IO.FileAttributes.Normal);

                UnsetReadOnly(item);
            }

        }

        #endregion
    }
}
