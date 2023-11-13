using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Permission;
using FISCA;
using FISCA.Presentation;

namespace StudentAgeStatistics
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {
            var key = "47D51D77-3E76-46C8-A5B8-BA65EAF6A5F3";
            RoleAclSource.Instance["教務作業"]["功能按鈕"].Add(new RibbonFeature(key, "高級中等學校班級及學生概況(二)權限"));
            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校班級及學生概況(二)"].Enable = FISCA.Permission.UserAcl.Current[key].Executable;

            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校班級及學生概況(二)"].Click += delegate
            {
                PrintSet PrintSet = new PrintSet();
                PrintSet.Show();
            };

        }
        public static List<string> ErrorList = new List<string>();
    }
}
