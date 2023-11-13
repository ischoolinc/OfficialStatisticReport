
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA;
using FISCA.Presentation;

namespace OverseasChineseStudent
{
    public class Program
    {
        [MainMethod]
        public static void main()
        {
            var key = "AED503DF-5F0F-435E-A6BE-E118E2FE14EF";
            RoleAclSource.Instance["教務作業"]["功能按鈕"].Add(new RibbonFeature(key, "高級中等學校僑生及大陸和港澳地區學生統計權限"));
            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校僑生及大陸和港澳地區學生統計"].Enable = FISCA.Permission.UserAcl.Current[key].Executable;

            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校僑生及大陸和港澳地區學生統計"].Click += delegate
            {
                PrintBranch PrintBranch = new PrintBranch();
                PrintBranch.Show();
            };         
        }
    }
}
