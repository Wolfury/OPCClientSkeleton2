using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpcRcw.Comn;
using OpcRcw.Da;

namespace OPCClientSkeleton2
{
    //Параметры OPC-группы
    class OPCGroupSettings
    {
        public string Name;
        public IOPCServer pServer; //интерфейс сервера
        public int updateRate = 1000; // время опроса создаваемой группы 
        public int bActive = 1; // активность группы - активна 
        public uint hClientGroup = 1; // клиентский описатель группы
        public Guid riid = typeof(IOPCItemMgt).GUID;
        public IntPtr TimeBias = IntPtr.Zero; // смещение по времени(по умолчанию - не используется)
        public IntPtr DeadBand = IntPtr.Zero; // зона нечувствительности(по умолчанию - не используется)
    }
}
