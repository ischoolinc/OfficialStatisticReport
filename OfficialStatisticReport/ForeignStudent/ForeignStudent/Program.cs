using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Internal;
using FISCA.Permission;
using FISCA;
using FISCA.Presentation;
namespace ForeignStudent
{
    public class Program
    {
        [MainMethod]
        public static void main()
        {
            var key = "D9150AC8-0E45-4973-9110-A3FDF87459FD";
            RoleAclSource.Instance["教務作業"]["功能按鈕"].Add(new RibbonFeature(key, "高級中等學校外國學生統計權限"));
            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校外國學生統計"].Enable = FISCA.Permission.UserAcl.Current[key].Executable;

             MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校外國學生統計"].Click += delegate
            {
                PrintBranch PrintBranch = new PrintBranch();
                PrintBranch.Show();
            };
        }
    }
}
