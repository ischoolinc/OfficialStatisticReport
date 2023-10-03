using System;
using System.Collections.Generic;
using System.Text;
using SmartSchool.Customization.PlugIn;
using FISCA.Permission;

namespace 高中職學校班級及學生概況
{
    //ischool plugin需為static class
    public static class Program
    {

        [FISCA.MainMethod]
        public static void main()
        {   
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "資料統計"];            
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["高中職學校班級及學生概況"].Enable = UserAcl.Current["SH.School.SchoolStatistics"].Executable;
            item1["報表"]["高中職學校班級及學生概況"].Click += delegate
            {
                Form1 f1 = new Form1();
                f1.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature("SH.School.SchoolStatistics", "高中職學校班級及學生概況"));
        }



        ////需宣告plugin主要方法為[MainMethod]
        //[MainMethod]
        //public static void Main()
        //{
        //    new SchoolStatistics();
        //}

        public static List<string> ErrorList = new List<string>();
    }    
}