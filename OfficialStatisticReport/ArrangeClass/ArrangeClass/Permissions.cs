using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ArrangeClass
{
    class Permissions
    {
        public static string 編班名冊 { get { return "ArrangeClass.2ab55894-fb6d-4d67-9d51-5b18b740b428"; } }

        public static bool 編班名冊權限
        {
            get { return FISCA.Permission.UserAcl.Current[編班名冊].Executable; }
        }
    }
}
