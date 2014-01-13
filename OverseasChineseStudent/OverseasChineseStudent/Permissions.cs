using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OverseasChineseStudent
{
    class Permissions
    {
        public static string 僑生統計 { get { return "OverseasChineseStudent.C7C8579F-38C8-498E-B3DA-70E1BBBE902B"; } }
        public static bool 僑生統計權限
        {
            get { return FISCA.Permission.UserAcl.Current[僑生統計].Executable; }
        }
    }
}
