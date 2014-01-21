using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateRecordReport
{
    [FISCA.UDT.TableName("UpdateRecordReportUDT")]
    public class UpdateRecordReportUDT : FISCA.UDT.ActiveRecord
    {
        [FISCA.UDT.Field(Field = "target")]
        public string Target { get; set; }

        [FISCA.UDT.Field(Field = "source")]
        public string Source { get; set; }
    }
}
