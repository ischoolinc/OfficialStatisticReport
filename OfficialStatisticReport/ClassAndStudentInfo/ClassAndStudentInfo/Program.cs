using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Permission;
using FISCA;
using K12.Data;
using FISCA.Presentation;

namespace ClassAndStudentInfo
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            var key = "6D095778-8617-4DEB-A457-7E0E642E765A";
            RoleAclSource.Instance["教務作業"]["功能按鈕"].Add(new RibbonFeature(key, "高中職學校班級及學生概況（一）權限"));
            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高中職學校班級及學生概況（一）"].Enable = FISCA.Permission.UserAcl.Current[key].Executable;

            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高中職學校班級及學生概況（一）"].Click += delegate
            {
                PrintBranch PrintBranch = new PrintBranch();
                PrintBranch.Show();
            };

        }
    }
}
