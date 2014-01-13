using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassAndStudentInfo
{
    class SummaryRow
    {
        public string 科別代號 { get; set; }
        public string 科別名稱 { get; set; }

        public int 班級數_總計 { get; set; }
        public int 班級數_一年級 { get; set; }
        public int 班級數_二年級 { get; set; }
        public int 班級數_三年級 { get; set; }
        public int 班級數_四年級 { get; set; }

        public int 學生數_總計 { get; set; }
        public int 學生數_總計_男 { get; set; }
        public int 學生數_總計_女 { get; set; }

        public int 學生數_一年級_男 { get; set; }
        public int 學生數_一年級_女 { get; set; }

        public int 學生數_二年級_男 { get; set; }
        public int 學生數_二年級_女 { get; set; }

        public int 學生數_三年級_男 { get; set; }
        public int 學生數_三年級_女 { get; set; }

        public int 學生數_四年級_男 { get; set; }
        public int 學生數_四年級_女 { get; set; }

        public int 學生數_延修_男 { get; set; }
        public int 學生數_延修_女 { get; set; }


    }
}
