using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Abo
{
    class GraduateStudentObj
    {
        public string ID { get; set; }
        public string Name { get; set; }
        public string Student_number { get; set; }
        public string Gender { get; set; }
        public string Dept { get; set; }
        public List<String> TagID { get; set; }

        public GraduateStudentObj(String ID, String Name, String Student_number, String Gender, String Dept, List<String> TagID)
        {
            this.ID = ID;
            this.Name = Name;
            this.Student_number = Student_number;
            this.Gender = Gender;
            this.Dept = Dept;
            this.TagID = TagID;
        }
    }
}
