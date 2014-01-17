using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassAndStudentInfo
{
    class Permissions
    {
        public static string 班級及學生概況1 { get { return "ClassAndStudentInfo.2ab55894-fb6d-4d67-9d51-5b18b740b428"; } }

        public static bool 班級及學生概況1權限
        {
            get { return FISCA.Permission.UserAcl.Current[班級及學生概況1].Executable; }
        }
    }
}
