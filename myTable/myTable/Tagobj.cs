using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace myTable
{
    class Tagobj
    {
        private String id, prefix, name, ref_student_id;

        public String Ref_student_id
        {
            get { return ref_student_id; }
            set { ref_student_id = value; }
        }

        public String Name
        {
            get { return name; }
            set { name = value; }
        }

        public String Prefix
        {
            get { return prefix; }
            set { prefix = value; }
        }

        public String Id
        {
            get { return id; }
            set { id = value; }
        }
        public Tagobj(String id, String prefix, String name, String ref_student_id)
        {
            this.id = id;
            this.prefix = prefix;
            this.name = name;
            this.ref_student_id = ref_student_id;
        }
    }
}
