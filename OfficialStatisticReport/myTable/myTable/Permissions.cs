using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace myTable
{
    class Permissions
    {
        public static string 新生入學方式統計表 { get { return "ClassAndStudentInfo.2F057D88-9018-40DB-8ECA-38934CBD0F7E"; } }

        public static bool 新生入學方式統計表權限
        {
            get { return FISCA.Permission.UserAcl.Current[新生入學方式統計表].Executable; }
        }
    }
}
