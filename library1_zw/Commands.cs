using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using zzr = ZwSoft.ZwCAD.Runtime;

namespace library1_zw
{
    public class Commands
    {
        [zzr.CommandMethod("PI_kota_konstrukcja")]

        public void WstawKoteKontrukcyjna()
        {
            Library1.Kota_Kon();
        }

        [zzr.CommandMethod("PI_kota_wykonczenie")]

        public void WstawKoteWykonczeniowa()
        {
            Library1.Kota_Wyk();
        }

        [zzr.CommandMethod("PI_viewport_skala")]

        public void Skalujviewport()
        {
            Library1.SkalujViewport();
        }

        [zzr.CommandMethod("PI_zigzag")]

        public void Zigzag()
        {
            Library1.Zigzag();
        }

        [zzr.CommandMethod("PI_izo1")]

        public void Izo1()
        {
            Library1.RysujIzo1();
        }


    }
}
