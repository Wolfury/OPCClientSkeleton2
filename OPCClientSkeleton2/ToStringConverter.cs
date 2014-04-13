using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using OpcRcw.Comn;
using OpcRcw.Da;

namespace OPCClientSkeleton2
{
    class ToStringConverter
    {
        static public string GetQualityString(short usQuality)
        {
            switch (usQuality)
            {
                case 0x00: return "Bad";
                case 0x04: return "Config Error";
                case 0x08: return "Not Connected";
                case 0x0C: return "Device Failure";
                case 0x10: return "Sensor Failure";
                case 0x14: return "Last Known";
                case 0x18: return "Comm Failure";
                case 0x1C: return "Out of Service";
                case 0x20: return "Initializing";
                case 0x40: return "Uncertain";
                case 0x44: return "Last Usable";
                case 0x50: return "Sensor Calibration";
                case 0x54: return "EGU Exceeded";
                case 0x58: return "Sub Normal";
                case 0xC0: return "Good";
                case 0xD8: return "Local Override";
                default: return "Unknown";
            }
        }

        static public string GetVTString(ushort vt)
        {
            VarEnum ee = (VarEnum)vt;
            return ((VarEnum)vt).ToString(); ;
        }

        static public string GetFTSting(OpcRcw.Da.FILETIME ft)
        {
            long lFT = (((long)ft.dwHighDateTime) << 32) + ft.dwLowDateTime;
            //DateTime dt = DateTime.FromBinary(lFT);
            //DateTime dt = DateTime.FromFileTime(lFT);
            DateTime dt = DateTime.FromFileTimeUtc(lFT);
            //DateTime dt = DateTime.Now;
            return dt.ToString();
  
        }
    }
}
