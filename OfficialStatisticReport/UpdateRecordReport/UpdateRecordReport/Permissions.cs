using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateRecordReport
{
    class Permissions
    {
        public static string 高中職學校學生異動報告 { get { return "UpdateRecordReport.F6ED2FEC-D774-4329-880E-F14FBBF94D13"; } }
        public static bool 高中職學校學生異動報告權限
        {
            get { return FISCA.Permission.UserAcl.Current[高中職學校學生異動報告].Executable; }
        }
    }
}
