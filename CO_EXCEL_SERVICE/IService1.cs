using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;
using System.Windows.Forms;

using System.ComponentModel;
using System.Data;
using System.Reflection;
using System.IO;


namespace CO_EXCEL_SERVICE
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the interface name "IService1" in both code and config file together.
    [ServiceContract]
    public interface IService1
    {


        [OperationContract]
        string GetData(int value);

        [OperationContract]
        CompositeType GetDataUsingDataContract(CompositeType composite);

        // TODO: Add your service operations here
        /// <summary>
        /// 获取树结构数据
        /// </summary>
        /// <param name="tr"></param>
        /// <returns></returns>
        [OperationContract]

        List<string> GetTreeInfo();

        /// <summary>
        /// 打开文件
        /// </summary>
        /// <param name="fileTag"></param>
        /// <returns></returns>

        [OperationContract]
        Stream OpenFile(string fileTag);

        [OperationContract]
        bool lockFile(string filepath);
        [OperationContract]
        bool unLockFile(string filepath);
        [OperationContract]
        void upLoad(RemoteFileInfo request);

        [OperationContract]
        void UnsetReadOnly(string path);
        //[OperationContract]
        //bool Storage_File(RemoteFileInfo inputFile);

    }

    [MessageContract]
    public class RemoteFileInfo : IDisposable
    {
        [MessageHeader(MustUnderstand = true)]
        public string FileName;

        [MessageBodyMember(Order = 1)]
        public System.IO.Stream FileByteStream;

        public void Dispose()
        {
            if (FileByteStream != null)
            {
                FileByteStream.Close();
                FileByteStream = null;
            }
        }
    }

    // Use a data contract as illustrated in the sample below to add composite types to service operations.
    [DataContract]
    public class CompositeType
    {
        bool boolValue = true;
        string stringValue = "Hello ";

        [DataMember]
        public bool BoolValue
        {
            get { return boolValue; }
            set { boolValue = value; }
        }

        [DataMember]
        public string StringValue
        {
            get { return stringValue; }
            set { stringValue = value; }
        }
    }
}
