using System;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using OpcRcw.Comn;
using OpcRcw.Da;

namespace OPCClientSkeleton2
{
    class Program
    {
        static void Main(string[] args)
        {
            OPCGroupSettings opcGroupSet = new OPCGroupSettings();
            OPCGroupSettings opcGroupSet2 = new OPCGroupSettings();
            OPCGroup mOPCGroup; // Объект OPC-группы
            OPCGroup mAsycOPCGroup;
            ArrayList opcItemsIDs = new ArrayList(); // Список элементов для добавления в OPC-группу 

            object m_pIfaceObj = InitServer("{3c5702a2-eb8e-11d4-83a4-00105a984cbd}"); //объект сервера 
            //Guid guid = new Guid("{f878aa1d-91c4-47d7-a880-e3ea7d07a14a}"); //EME
            //Guid guid = new Guid("{3c5702a2-eb8e-11d4-83a4-00105a984cbd}"); //iFix


            #region Пример работы синхронного чтения и записи
            opcGroupSet.Name = "MyGroup";
            opcGroupSet.pServer = (IOPCServer)m_pIfaceObj;
            mOPCGroup = OPCGroup.CreateGroup(opcGroupSet);


            if (mOPCGroup != null)
            {
                opcItemsIDs.Add("FIX.AR_2.F_CV");
                opcItemsIDs.Add("FIX.AR_3.F_CV");
                mOPCGroup.AddItems(opcItemsIDs);

                opcItemsIDs.Clear();
                opcItemsIDs.Add("FIX.AR_4.F_CV");
                opcItemsIDs.Add("FIX.AR_5.F_CV");
                mOPCGroup.AddItems(opcItemsIDs);
            }

            //mOPCGroup.EnumerateItems(true);

            Hashtable ItemValues = new Hashtable();

            ItemValues["FIX.AR_4.F_CV"] = 23;
            ItemValues["FIX.AR_5.F_CV"] = 68;

            mOPCGroup.WriteItems(ItemValues);

            mOPCGroup.ReadItems();
            #endregion

            #region Пример работы асинхронного чтения

            opcGroupSet2.Name = "MyAsyncGroup";
            opcGroupSet2.pServer = (IOPCServer)m_pIfaceObj;
            opcGroupSet2.hClientGroup = 2;
            mAsycOPCGroup = OPCGroup.CreateGroup(opcGroupSet2);
        
            opcItemsIDs.Add("FIX.AR_1.F_CV");

            mAsycOPCGroup.AddItems(opcItemsIDs);

            IConnectionPointContainer pCPC;
            IConnectionPoint m_pDataCallback;
            pCPC = (IConnectionPointContainer)mAsycOPCGroup.IFaceObj;
            Guid riid = typeof(IOPCDataCallback).GUID;
            OPCDataCallback m_pSink;
            int m_iCookie;
            try
            {
                pCPC.FindConnectionPoint(ref riid, out m_pDataCallback);
                m_pSink = new OPCDataCallback("FIX.AR_1.F_CV");
                m_pDataCallback.Advise(m_pSink, out m_iCookie);
            }
            catch (Exception ex)
            {
                /*string msg; 
                int hRes = Marshal.GetHRForException(ex);
                
                ((IOPCServer)m_pIfaceObj).GetErrorString(hRes,  0, out msg);*/
                Console.WriteLine(ex.Message);
            }

            #endregion

            Console.ReadLine();

        }

        static object InitServer(string sGuid)
        {
            Guid guid = new Guid(sGuid);
            object opcServer;

            try
            {
                Type typeOfServer = Type.GetTypeFromCLSID(guid);
                opcServer = Activator.CreateInstance(typeOfServer);
                return opcServer;
            }
            catch (ApplicationException ex)
            {
                string msg = "error";
                //Получаем HRESULT, соответствующий сгенерированному исключению 
                int hRes = Marshal.GetHRForException(ex);
                //Запрашиваем у сервера текст ошибки 
                //pServer.GetErrorString(hRes, 2, out msg);
                Console.WriteLine(msg + "Ошибка");
                return null;
            }
        }
     
    }
}
