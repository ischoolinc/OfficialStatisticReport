using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForeignStudent
{
    class Permissions
    {
        public static string 外國學生統計 { get { return "ForeignStudent.64F3448F-7839-4A41-91D7-6DD89EC9B594"; } }
        public static bool 外國學生統計權限
        {
            get { return FISCA.Permission.UserAcl.Current[外國學生統計].Executable; }
        }
    }
}
