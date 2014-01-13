using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Abo
{
    class Permissions
    {
        public static string 原住民學生數及畢業生統計表 { get { return "ClassAndStudentInfo.AE30B6D4-95D4-45EB-8800-BB6F2AA00E4A"; } }
        public static bool 原住民學生數及畢業生統計表權限
        {
            get { return FISCA.Permission.UserAcl.Current[原住民學生數及畢業生統計表].Executable; }
        }
    }
}
