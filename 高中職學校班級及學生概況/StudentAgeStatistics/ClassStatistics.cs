using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.PlugIn;
using SmartSchool.Customization.PlugIn.Report;
using SHSchool.Data;
namespace StudentAgeStatistics
{
    public class ClassStatistics
    {
        public int Level0Count24 = 0;
        public int Level0Count34 = 0;
        public int Level0Count44 = 0;
        public int Level0Count54 = 0;
        public int Level0Count55 = 0;
        
        public int Level1Count24 = 0;
        public int Level1Count34 = 0;
        public int Level1Count44 = 0;
        public int Level1Count54 = 0;
        public int Level1Count55 = 0;

        public int Level2Count24 = 0;
        public int Level2Count34 = 0;
        public int Level2Count44 = 0;
        public int Level2Count54 = 0;
        public int Level2Count55 = 0;

        public int Level3Count24 = 0;
        public int Level3Count34 = 0;
        public int Level3Count44 = 0;
        public int Level3Count54 = 0;
        public int Level3Count55 = 0;

        public int Level4Count24 = 0;
        public int Level4Count34 = 0;
        public int Level4Count44 = 0;
        public int Level4Count54 = 0;
        public int Level4Count55 = 0;

        //一年級
        public int SR15C5 = 0;
        public int SR15C6 = 0;
        public int SR15C7 = 0;
        public int SR15C8 = 0;
        public int SR15C9 = 0;
        public int SR15C10 = 0;
        public int SR15C11 = 0;
        public int SR15C12 = 0;
        public int SR15C13 = 0;
        public int SR15C14 = 0;

        public int SR16C5 = 0;
        public int SR16C6 = 0;
        public int SR16C7 = 0;
        public int SR16C8 = 0;
        public int SR16C9 = 0;
        public int SR16C10 = 0;
        public int SR16C11 = 0;
        public int SR16C12 = 0;
        public int SR16C13 = 0;
        public int SR16C14 = 0;

        //二年級
        public int SR17C5 = 0;
        public int SR17C6 = 0;
        public int SR17C7 = 0;
        public int SR17C8 = 0;
        public int SR17C9 = 0;
        public int SR17C10 = 0;
        public int SR17C11 = 0;
        public int SR17C12 = 0;
        public int SR17C13 = 0;
        public int SR17C14 = 0;

        public int SR18C5 = 0;
        public int SR18C6 = 0;
        public int SR18C7 = 0;
        public int SR18C8 = 0;
        public int SR18C9 = 0;
        public int SR18C10 = 0;
        public int SR18C11 = 0;
        public int SR18C12 = 0;
        public int SR18C13 = 0;
        public int SR18C14 = 0;

        //三年級
        public int SR19C5 = 0;
        public int SR19C6 = 0;
        public int SR19C7 = 0;
        public int SR19C8 = 0;
        public int SR19C9 = 0;
        public int SR19C10 = 0;
        public int SR19C11 = 0;
        public int SR19C12 = 0;
        public int SR19C13 = 0;
        public int SR19C14 = 0;

        public int SR20C5 = 0;
        public int SR20C6 = 0;
        public int SR20C7 = 0;
        public int SR20C8 = 0;
        public int SR20C9 = 0;
        public int SR20C10 = 0;
        public int SR20C11 = 0;
        public int SR20C12 = 0;
        public int SR20C13 = 0;
        public int SR20C14 = 0;

        //四年級
        public int SR21C5 = 0;
        public int SR21C6 = 0;
        public int SR21C7 = 0;
        public int SR21C8 = 0;
        public int SR21C9 = 0;
        public int SR21C10 = 0;
        public int SR21C11 = 0;
        public int SR21C12 = 0;
        public int SR21C13 = 0;
        public int SR21C14 = 0;

        public int SR22C5 = 0;
        public int SR22C6 = 0;
        public int SR22C7 = 0;
        public int SR22C8 = 0;
        public int SR22C9 = 0;
        public int SR22C10 = 0;
        public int SR22C11 = 0;
        public int SR22C12 = 0;
        public int SR22C13 = 0;
        public int SR22C14 = 0;

        // 延修生
        public int SR23C5 = 0;
public int SR23C6 = 0;
public int SR23C7 = 0;
public int SR23C8 = 0;
public int SR23C9 = 0;
public int SR23C10 = 0;
public int SR23C11 = 0;
public int SR23C12 = 0;
public int SR23C13 = 0;
public int SR23C14 = 0;

public int SR24C5 = 0;
public int SR24C6 = 0;
public int SR24C7 = 0;
public int SR24C8 = 0;
public int SR24C9 = 0;
public int SR24C10 = 0;
public int SR24C11 = 0;
public int SR24C12 = 0;
public int SR24C13 = 0;
public int SR24C14 = 0;

        public ClassStatistics()
        {
 
        }


        private bool IsMatchClassCondition(int GradeYear,int ClassCount,int ClassLevel,int Min,int Max)
        {
            if (ClassLevel == 0)
                return (ClassCount >= Min && ClassCount <= Max)?true:false;
            else
                return (GradeYear==ClassLevel && ClassCount >= Min && ClassCount <= Max) ? true : false;
        }

        private bool IsMathStudentCondition(int StudentGradeYear,int StudentAge,string StudentGender,int ConditionGradeYear,int MinAge,int MaxAge,string ConditionGender)
        {
            if (ConditionGender.Equals(""))
                return (StudentAge>=MinAge && StudentAge<=MaxAge  && StudentGradeYear==ConditionGradeYear)?true:false;
            else
                return (StudentAge >= MinAge && StudentAge <= MaxAge && StudentGender.Equals(ConditionGender) && StudentGradeYear == ConditionGradeYear) ? true : false; 
        }

        private int GetStudentAge(string Birthday)
        {
            int retVal = 0;

            DateTime DateBirthday;
            if (DateTime.TryParse(Birthday, out DateBirthday))
            {
                if (DateBirthday.Month >= 1 && DateBirthday.Month <= 8)
                    retVal=DateTime.Now.Year - DateBirthday.Year;
                else
                    retVal=DateTime.Now.Year - DateBirthday.Year - 1;
            }
            return retVal;
        }

        public void StartStatistics(List<string> type)
        {
            AccessHelper Helper = new AccessHelper();

            List<ClassRecord> Classes = Helper.ClassHelper.GetAllClass();
            List<SHDepartmentRecord> DeptList = new List<SHDepartmentRecord>();
            DeptList = SHDepartment.SelectAll();
            List<string> deptContains = new List<string>();
            foreach (SHDepartmentRecord dp in DeptList)
            {
                if (type.Contains(dp.RefDeptGroupID.ToString()))
                    deptContains.Add(dp.FullName);
            }
            try
            {
                foreach (ClassRecord ClassRec in Classes)
                {
                    if (deptContains.Contains(ClassRec.Department))
                    {
                        // 沒有班級年級跳出(有延修生的班級會無法統計，故無法跳出by王金鳳)
                        int grYear = 0;
                        int.TryParse(ClassRec.GradeYear, out grYear);
                        //if (grYear < 1)
                        //    continue;
                        foreach (StudentRecord Student in ClassRec.Students)
                        {
                            int StudentAge = GetStudentAge(Student.Birthday);

                            int GradeYear = 0;
                            if (Student.Status == "延修")
                                GradeYear = 99;
                            else
                                int.TryParse(ClassRec.GradeYear, out GradeYear);
                            // 年級解析錯誤
                            if (GradeYear == 0)
                                continue;

                            if (StudentAge == 0)
                            {
                                Program.ErrorList.Add("學號：" + Student.StudentNumber + ", 班級：" + Student.RefClass.ClassName + ", 座號：" + Student.SeatNo + ", 學生姓名：" + Student.StudentName + ", 生日無法辨識，不被列入計算");
                                continue;
                            }

                            //一年級
                            SR15C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 0, 13, "男") ? SR15C5 + 1 : SR15C5;
                            SR15C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 14, 14, "男") ? SR15C6 + 1 : SR15C6;
                            SR15C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 15, 15, "男") ? SR15C7 + 1 : SR15C7;
                            SR15C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 16, 16, "男") ? SR15C8 + 1 : SR15C8;
                            SR15C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 17, 17, "男") ? SR15C9 + 1 : SR15C9;
                            SR15C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 18, 18, "男") ? SR15C10 + 1 : SR15C10;
                            SR15C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 19, 19, "男") ? SR15C11 + 1 : SR15C11;
                            SR15C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 20, 20, "男") ? SR15C12 + 1 : SR15C12;
                            SR15C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 21, 21, "男") ? SR15C13 + 1 : SR15C13;
                            SR15C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 22, 200, "男") ? SR15C14 + 1 : SR15C14; ;

                            SR16C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 0, 13, "女") ? SR16C5 + 1 : SR16C5;
                            SR16C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 14, 14, "女") ? SR16C6 + 1 : SR16C6;
                            SR16C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 15, 15, "女") ? SR16C7 + 1 : SR16C7;
                            SR16C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 16, 16, "女") ? SR16C8 + 1 : SR16C8;
                            SR16C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 17, 17, "女") ? SR16C9 + 1 : SR16C9;
                            SR16C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 18, 18, "女") ? SR16C10 + 1 : SR16C10;
                            SR16C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 19, 19, "女") ? SR16C11 + 1 : SR16C11;
                            SR16C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 20, 20, "女") ? SR16C12 + 1 : SR16C12;
                            SR16C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 21, 21, "女") ? SR16C13 + 1 : SR16C13;
                            SR16C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 1, 22, 200, "女") ? SR16C14 + 1 : SR16C14; ;

                            //二年級
                            SR17C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 0, 13, "男") ? SR17C5 + 1 : SR17C5;
                            SR17C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 14, 14, "男") ? SR17C6 + 1 : SR17C6;
                            SR17C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 15, 15, "男") ? SR17C7 + 1 : SR17C7;
                            SR17C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 16, 16, "男") ? SR17C8 + 1 : SR17C8;
                            SR17C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 17, 17, "男") ? SR17C9 + 1 : SR17C9;
                            SR17C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 18, 18, "男") ? SR17C10 + 1 : SR17C10;
                            SR17C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 19, 19, "男") ? SR17C11 + 1 : SR17C11;
                            SR17C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 20, 20, "男") ? SR17C12 + 1 : SR17C12;
                            SR17C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 21, 21, "男") ? SR17C13 + 1 : SR17C13;
                            SR17C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 22, 200, "男") ? SR17C14 + 1 : SR17C14; ;

                            SR18C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 0, 13, "女") ? SR18C5 + 1 : SR18C5;
                            SR18C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 14, 14, "女") ? SR18C6 + 1 : SR18C6;
                            SR18C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 15, 15, "女") ? SR18C7 + 1 : SR18C7;
                            SR18C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 16, 16, "女") ? SR18C8 + 1 : SR18C8;
                            SR18C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 17, 17, "女") ? SR18C9 + 1 : SR18C9;
                            SR18C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 18, 18, "女") ? SR18C10 + 1 : SR18C10;
                            SR18C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 19, 19, "女") ? SR18C11 + 1 : SR18C11;
                            SR18C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 20, 20, "女") ? SR18C12 + 1 : SR18C12;
                            SR18C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 21, 21, "女") ? SR18C13 + 1 : SR18C13;
                            SR18C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 2, 22, 200, "女") ? SR18C14 + 1 : SR18C14; ;

                            //三年級
                            SR19C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 0, 13, "男") ? SR19C5 + 1 : SR19C5;
                            SR19C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 14, 14, "男") ? SR19C6 + 1 : SR19C6;
                            SR19C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 15, 15, "男") ? SR19C7 + 1 : SR19C7;
                            SR19C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 16, 16, "男") ? SR19C8 + 1 : SR19C8;
                            SR19C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 17, 17, "男") ? SR19C9 + 1 : SR19C9;
                            SR19C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 18, 18, "男") ? SR19C10 + 1 : SR19C10;
                            SR19C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 19, 19, "男") ? SR19C11 + 1 : SR19C11;
                            SR19C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 20, 20, "男") ? SR19C12 + 1 : SR19C12;
                            SR19C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 21, 21, "男") ? SR19C13 + 1 : SR19C13;
                            SR19C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 22, 200, "男") ? SR19C14 + 1 : SR19C14; ;

                            SR20C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 0, 13, "女") ? SR20C5 + 1 : SR20C5;
                            SR20C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 14, 14, "女") ? SR20C6 + 1 : SR20C6;
                            SR20C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 15, 15, "女") ? SR20C7 + 1 : SR20C7;
                            SR20C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 16, 16, "女") ? SR20C8 + 1 : SR20C8;
                            SR20C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 17, 17, "女") ? SR20C9 + 1 : SR20C9;
                            SR20C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 18, 18, "女") ? SR20C10 + 1 : SR20C10;
                            SR20C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 19, 19, "女") ? SR20C11 + 1 : SR20C11;
                            SR20C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 20, 20, "女") ? SR20C12 + 1 : SR20C12;
                            SR20C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 21, 21, "女") ? SR20C13 + 1 : SR20C13;
                            SR20C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 3, 22, 200, "女") ? SR20C14 + 1 : SR20C14; ;

                            //四年級
                            SR21C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 0, 13, "男") ? SR21C5 + 1 : SR21C5;
                            SR21C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 14, 14, "男") ? SR21C6 + 1 : SR21C6;
                            SR21C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 15, 15, "男") ? SR21C7 + 1 : SR21C7;
                            SR21C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 16, 16, "男") ? SR21C8 + 1 : SR21C8;
                            SR21C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 17, 17, "男") ? SR21C9 + 1 : SR21C9;
                            SR21C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 18, 18, "男") ? SR21C10 + 1 : SR21C10;
                            SR21C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 19, 19, "男") ? SR21C11 + 1 : SR21C11;
                            SR21C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 20, 20, "男") ? SR21C12 + 1 : SR21C12;
                            SR21C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 21, 21, "男") ? SR21C13 + 1 : SR21C13;
                            SR21C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 22, 200, "男") ? SR21C14 + 1 : SR21C14; ;

                            SR22C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 0, 13, "女") ? SR22C5 + 1 : SR22C5;
                            SR22C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 14, 14, "女") ? SR22C6 + 1 : SR22C6;
                            SR22C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 15, 15, "女") ? SR22C7 + 1 : SR22C7;
                            SR22C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 16, 16, "女") ? SR22C8 + 1 : SR22C8;
                            SR22C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 17, 17, "女") ? SR22C9 + 1 : SR22C9;
                            SR22C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 18, 18, "女") ? SR22C10 + 1 : SR22C10;
                            SR22C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 19, 19, "女") ? SR22C11 + 1 : SR22C11;
                            SR22C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 20, 20, "女") ? SR22C12 + 1 : SR22C12;
                            SR22C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 21, 21, "女") ? SR22C13 + 1 : SR22C13;
                            SR22C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 4, 22, 200, "女") ? SR22C14 + 1 : SR22C14;

                            //延修生
                            SR23C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 0, 13, "男") ? SR23C5 + 1 : SR23C5;
                            SR23C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 14, 14, "男") ? SR23C6 + 1 : SR23C6;
                            SR23C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 15, 15, "男") ? SR23C7 + 1 : SR23C7;
                            SR23C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 16, 16, "男") ? SR23C8 + 1 : SR23C8;
                            SR23C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 17, 17, "男") ? SR23C9 + 1 : SR23C9;
                            SR23C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 18, 18, "男") ? SR23C10 + 1 : SR23C10;
                            SR23C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 19, 19, "男") ? SR23C11 + 1 : SR23C11;
                            SR23C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 20, 20, "男") ? SR23C12 + 1 : SR23C12;
                            SR23C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 21, 21, "男") ? SR23C13 + 1 : SR23C13;
                            SR23C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 22, 200, "男") ? SR23C14 + 1 : SR23C14; ;

                            SR24C5 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 0, 13, "女") ? SR24C5 + 1 : SR24C5;
                            SR24C6 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 14, 14, "女") ? SR24C6 + 1 : SR24C6;
                            SR24C7 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 15, 15, "女") ? SR24C7 + 1 : SR24C7;
                            SR24C8 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 16, 16, "女") ? SR24C8 + 1 : SR24C8;
                            SR24C9 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 17, 17, "女") ? SR24C9 + 1 : SR24C9;
                            SR24C10 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 18, 18, "女") ? SR24C10 + 1 : SR24C10;
                            SR24C11 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 19, 19, "女") ? SR24C11 + 1 : SR24C11;
                            SR24C12 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 20, 20, "女") ? SR24C12 + 1 : SR24C12;
                            SR24C13 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 21, 21, "女") ? SR24C13 + 1 : SR24C13;
                            SR24C14 = IsMathStudentCondition(GradeYear, StudentAge, Student.Gender, 99, 22, 200, "女") ? SR24C14 + 1 : SR24C14;
                        }

                        try
                        {
                            //一年級班級數總計
                            Level1Count24 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 1, 0, 24) ? Level1Count24 + 1 : Level1Count24;
                            Level1Count34 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 1, 25, 34) ? Level1Count34 + 1 : Level1Count34;
                            Level1Count44 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 1, 35, 44) ? Level1Count44 + 1 : Level1Count44;
                            Level1Count54 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 1, 45, 54) ? Level1Count54 + 1 : Level1Count54;
                            Level1Count55 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 1, 55, 99999) ? Level1Count55 + 1 : Level1Count55;

                            //二年級班級數總計
                            Level2Count24 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 2, 0, 24) ? Level2Count24 + 1 : Level2Count24;
                            Level2Count34 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 2, 25, 34) ? Level2Count34 + 1 : Level2Count34;
                            Level2Count44 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 2, 35, 44) ? Level2Count44 + 1 : Level2Count44;
                            Level2Count54 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 2, 45, 54) ? Level2Count54 + 1 : Level2Count54;
                            Level2Count55 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 2, 55, 99999) ? Level2Count55 + 1 : Level2Count55;

                            //三年級班級數總計
                            Level3Count24 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 3, 0, 24) ? Level3Count24 + 1 : Level3Count24;
                            Level3Count34 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 3, 25, 34) ? Level3Count34 + 1 : Level3Count34;
                            Level3Count44 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 3, 35, 44) ? Level3Count44 + 1 : Level3Count44;
                            Level3Count54 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 3, 45, 54) ? Level3Count54 + 1 : Level3Count54;
                            Level3Count55 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 3, 55, 99999) ? Level3Count55 + 1 : Level3Count55;

                            //四年級班級數總計
                            Level4Count24 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 4, 0, 24) ? Level4Count24 + 1 : Level4Count24;
                            Level4Count34 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 4, 25, 34) ? Level4Count34 + 1 : Level4Count34;
                            Level4Count44 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 4, 35, 44) ? Level4Count44 + 1 : Level4Count44;
                            Level4Count54 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 4, 45, 54) ? Level4Count54 + 1 : Level4Count54;
                            Level4Count55 = IsMatchClassCondition(grYear, ClassRec.Students.Count, 4, 55, 99999) ? Level4Count55 + 1 : Level4Count55;
                        }

                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.ErrorList.Add(ex.Message);
            }
        }
        
    }
}