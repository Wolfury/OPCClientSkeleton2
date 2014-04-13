using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpcRcw.Comn;
using OpcRcw.Da;
using System.Runtime.InteropServices;

namespace OPCClientSkeleton2
{
    public class OPCDataCallback : IOPCDataCallback
    {
        private string m_szItemID;

        //конструктор
        public OPCDataCallback(string szItemID) 
        {
            SetItemID(szItemID); //идентификатор текущего элемента данных
        }

        public void SetItemID(string szItemID) 
        {
            m_szItemID = szItemID;
        }

        //========Перегрузка методов интерфейса IOPCDataCallback ============ 
        public void OnCancelComplete(int dwTransid, int hGroup) { }

        public void OnDataChange(int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount, 
                                 int[] phClientItems, object[] pvValues, short[] pwQualities, OpcRcw.Da.FILETIME[] pftTimeStamps, int[] pErrors)
        {
            //Получаем тип элемента данных
            IntPtr iptrValues = Marshal.AllocCoTaskMem(2);
            Marshal.GetNativeVariantForObject(pvValues, iptrValues);
            byte[] vt = new byte[2];
            Marshal.Copy(iptrValues, vt, 0, 2);
            Marshal.FreeCoTaskMem(iptrValues);
            ushort usVt = (ushort)(vt[0] + vt[1] * 255);
            
            Console.Write(m_szItemID + "---");
            Console.Write(ToStringConverter.GetVTString(usVt) + "---");
            Console.Write(pvValues[0].ToString() + "---");
            
            Console.Write(ToStringConverter.GetFTSting(pftTimeStamps[0]) + "---");
            Console.WriteLine(ToStringConverter.GetQualityString(pwQualities[0]));
          
        }

        public void OnReadComplete(int dwTransid, int hGroup, int hrMasterquality, int hrMastererror, int dwCount, int[] phClientItems, object[] pvValues, short[] pwQualities, OpcRcw.Da.FILETIME[] pftTimeStamps, int[] pErrors)
        { }
        public void OnWriteComplete(int dwTransid, int hGroup, int hrMastererr, int dwCount, int[] pClienthandles, int[] pErrors)
        { }
    }
}
