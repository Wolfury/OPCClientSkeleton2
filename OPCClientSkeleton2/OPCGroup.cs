using System;
using System.Collections;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using OpcRcw.Comn;
using OpcRcw.Da;

namespace OPCClientSkeleton2
{
    class OPCGroup
    {
        private IOPCServer pServer; //интерфейс сервера
        private object iFaceObj; // сюда вернем интерфейс к группе 

        public object IFaceObj
        {
            get { return iFaceObj; }
        }

        private int m_hGroup; // хэндл группы
        private int ItemsCount = 0;
        private int[] hServerItems;
        private Hashtable htItemsID = new Hashtable();

        OPCGroup(IOPCServer pServer, object iFaceObj, int m_hGroup)
        {
            this.pServer = pServer;
            this.iFaceObj = iFaceObj;
            this.m_hGroup = m_hGroup;
        }

        // Статический метод, возвращающий новый объект OPC-группы или null - в случае ошибки создания группы
        static public OPCGroup CreateGroup(OPCGroupSettings groupSettings)
        {
            object iFaceObj; // сюда вернем интерфейс к группе 
            int m_hGroup; // хэндл группы

            /* Прототип AddGroup
             * void AddGroup(string szName, int bActive, int dwRequestedUpdateRate, int hClientGroup, 
            System.IntPtr pTimeBias, System.IntPtr pPercentDeadband, int dwLCID, 
            out int phServerGroup, out int pRevisedUpdateRate, ref System.Guid riid, out object ppUnk)*/

            try
            {
                groupSettings.pServer.AddGroup(groupSettings.Name, (int)groupSettings.bActive, (int)groupSettings.updateRate,
                                               (int)groupSettings.hClientGroup, groupSettings.TimeBias, groupSettings.DeadBand,
                                               2, out m_hGroup, out groupSettings.updateRate, ref groupSettings.riid, out iFaceObj);
                return new OPCGroup(groupSettings.pServer, iFaceObj, m_hGroup);
            }
            catch (Exception ex)
            {
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hRes = Marshal.GetHRForException(ex);
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT
                groupSettings.pServer.GetErrorString(hRes, 2, out msg);
                //Показываем сообщение ошибки
                Console.WriteLine(msg + "Ошибка");
                return null;
            }
        }

        // Добавление тегов в группу
        public void AddItems(ArrayList alItemsIDs)
        {
            IOPCItemMgt pItemMgt;

            // Массив описателей добавляемых элементов OPCITEMDEF 
            OPCITEMDEF[] pItems = new OPCITEMDEF[alItemsIDs.Count];

            try
            {
                pItemMgt = (IOPCItemMgt)this.iFaceObj;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }

            // Наполнение массива элементов для добавления в группу
            for (int i = 0; i < alItemsIDs.Count; i++)
            {
                    pItems[i].szItemID = alItemsIDs[i].ToString();
                    pItems[i].szAccessPath = null;
                    pItems[i].bActive = 1;
                    pItems[i].hClient = 1;
                    pItems[i].vtRequestedDataType = (short)VarEnum.VT_EMPTY;
                    pItems[i].dwBlobSize = 0;
                    pItems[i].pBlob = IntPtr.Zero; 
            }

            try
            {
                // В эти две переменные будут записаны массивы ошибок и результатов выполнения
                IntPtr iptrErrors = IntPtr.Zero;
                IntPtr iptrResults = IntPtr.Zero;

                OPCITEMRESULT pResults; // Структура с результатом выполнения
                int[] hRes = new int[1]; // HRESULT-значение(код возврата выполнения операции добавления) 

                // Добавляем элемент данных в группу 
                pItemMgt.AddItems(alItemsIDs.Count, pItems, out iptrResults, out iptrErrors);
                // Переносим результаты и ошибки из неуправляемой памяти в управляемую
                pResults = (OPCITEMRESULT)Marshal.PtrToStructure(iptrResults, typeof(OPCITEMRESULT));
                Marshal.Copy(iptrErrors, hRes, 0, 1);
                //Генерируем исключение в случае ошибки в HRESULT 
                Marshal.ThrowExceptionForHR(hRes[0]);
                Marshal.FreeCoTaskMem(iptrResults);
                Marshal.FreeCoTaskMem(iptrErrors);
                this.ItemsCount += alItemsIDs.Count;
                this.EnumerateItems();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }

        }

        // Перечислитель свойств тегов группы
        public void EnumerateItems(bool showInformationMode = false)
        {
            try
            {
                Guid riid = typeof(IEnumOPCItemAttributes).GUID;
                var pItemMgt = (IOPCItemMgt)this.iFaceObj;
                OPCITEMATTRIBUTES opcItemAttr;

                object oEnum;
                pItemMgt.CreateEnumerator(ref riid, out oEnum);
                IEnumOPCItemAttributes pEnumItemAttr = (IEnumOPCItemAttributes)oEnum;

                IntPtr iResult = IntPtr.Zero;
                int res_counter;
                pEnumItemAttr.Reset();
                hServerItems = new int[ItemsCount];
                htItemsID.Clear();

                for (int i = 0; i < this.ItemsCount; i++)
                {
                    pEnumItemAttr.Next(1, out iResult, out res_counter);
                    opcItemAttr = (OPCITEMATTRIBUTES)Marshal.PtrToStructure(iResult, typeof(OPCITEMATTRIBUTES));
                    hServerItems[i] = opcItemAttr.hServer;
                    htItemsID.Add(opcItemAttr.szItemID, opcItemAttr.hServer);//добавление хэндлов тегов в хэш-таблицу, ключ - ID

                    if(showInformationMode)
                        Console.WriteLine(opcItemAttr.hServer + "--" + opcItemAttr.szItemID);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        // Чтение тегов группы
        public OPCITEMSTATE[] ReadItems(int[] hServerItems = null, int ItemsCount = -1)
        {
            if (hServerItems == null)
                hServerItems = this.hServerItems;
            if (ItemsCount == -1)
                ItemsCount = this.ItemsCount;

            OPCITEMSTATE[] pItemState = new OPCITEMSTATE[ItemsCount];
            int[] hRes = new int[1];
            // Получаем интерфейс IOPCSyncIO для операций синхронного чтения IOPCSyncIO 
            IOPCSyncIO pSyncIO = (IOPCSyncIO)this.iFaceObj;
            // В эту переменную будут записаны результаты чтения IntPtr 
            IntPtr iptrItemState = IntPtr.Zero;
            IntPtr iptrErrors = IntPtr.Zero;

            try
            {
                //Читаем данные из сервера 
                pSyncIO.Read(OPCDATASOURCE.OPC_DS_DEVICE, ItemsCount, hServerItems, out iptrItemState, out iptrErrors);

                for (int i = 0; i < ItemsCount; i++ )
                {
                    // Переносим результаты и ошибки из неуправляемой памяти в управляемую
                    pItemState[i] = (OPCITEMSTATE)Marshal.PtrToStructure(iptrItemState, typeof(OPCITEMSTATE));
                    
                    Marshal.Copy(iptrErrors, hRes, 0, 1);
                    // Генерируем исключение в случае ошибки в HRESULT
                    Marshal.ThrowExceptionForHR(hRes[0]);

                    Console.WriteLine(pItemState[i].hClient + " --- " + pItemState[i].vDataValue);
                    iptrItemState = IntPtr.Add(iptrItemState, Marshal.SizeOf(pItemState[i]));
                    iptrErrors = IntPtr.Add(iptrErrors, Marshal.SizeOf(hRes[0]));
                }

                return pItemState;
            }
            catch (Exception ex)
            {
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hResEx = Marshal.GetHRForException(ex);
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT
                this.pServer.GetErrorString(hResEx, 2, out msg);
                //Показываем сообщение ошибки
                Console.WriteLine(msg + " Ошибка");
                return null;
            }
        }

        // Запись тегов группы
        public void WriteItems(Hashtable htItemsValues)
        {
            OPCITEMSTATE[] pItemState = new OPCITEMSTATE[ItemsCount];
            int[] hRes = new int[1];
            // Получаем интерфейс IOPCSyncIO для операций синхронного чтения IOPCSyncIO 
            IOPCSyncIO pSyncIO = (IOPCSyncIO)this.iFaceObj;
            
            // Указатель для получения результата выполнения операции записи
            IntPtr iptrErrors = IntPtr.Zero;

            int[] phServer = new int[htItemsValues.Count];
            object[] pItemValues =  new object[htItemsValues.Count];

            int j = 0;

            foreach(object currItem in htItemsValues.Keys)
            {
                if(htItemsID.ContainsKey(currItem))
                {
                    phServer[j] = (int)htItemsID[currItem];
                    pItemValues[j] = htItemsValues[currItem];
                    j++;
                }
                else
                {
                    Console.WriteLine("Ошибка! Данный тег отсутствует в группе." + currItem.ToString());
                    return;
                }
            }

            try
            {
                //Запись данных в теги сервера 
                //void Write(int dwCount, int[] phServer, object[] pItemValues, out System.IntPtr ppErrors)
                pSyncIO.Write(htItemsValues.Count, phServer, pItemValues, out iptrErrors);

                for (int i = 0; i < htItemsValues.Count; i++)
                {
                    // Переносим результаты и ошибки из неуправляемой памяти в управляемую
                    Marshal.Copy(iptrErrors, hRes, 0, 1);

                    // Генерируем исключение в случае ошибки в HRESULT
                    Marshal.ThrowExceptionForHR(hRes[0]);
                    iptrErrors = IntPtr.Add(iptrErrors, Marshal.SizeOf(hRes[0]));
                }

                return;
            }
            catch (Exception ex)
            {
                string msg;
                //Получаем HRESULT соответствующий сгененрированному исключению
                int hResEx = Marshal.GetHRForException(ex);
                //Запрашиваем у сервера текст ошибки, соответствующий текущему HRESULT
                this.pServer.GetErrorString(hResEx, 2, out msg);
                //Показываем сообщение ошибки
                Console.WriteLine(msg + " Ошибка");
                return ;
            }
        }
    }
}
