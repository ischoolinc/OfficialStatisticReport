using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArrangeClass.DAO
{
    // 學生學期對照
    public class SemsHistoryInfo
    {
        // 學年度
        public string SchoolYear { get; set; }
        // 學期
        public string Semester { get; set; }
        // 年級
        public string GradeYear { get; set; }
        // 班級名稱
        public string ClassName { get; set; }
        // 座號
        public string SeatNo { get; set; }        
    }
}
