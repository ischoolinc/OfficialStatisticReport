using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForeignStudent
{
    [FISCA.UDT.TableName("ForeignStudentUDT")]
    public class ForeignStudentUDT : FISCA.UDT.ActiveRecord
    {
        [FISCA.UDT.Field(Field = "target")]
        public string Target { get; set; }

        [FISCA.UDT.Field(Field = "source")]
        public string Source { get; set; }
    }
}
