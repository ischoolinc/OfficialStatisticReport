using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
namespace Abo
{
    public class Program
    {
        [MainMethod]
        public static void main()
        {
            var key = "F1009EA4-2071-4CE9-8EAB-929864A682C0";
            RoleAclSource.Instance["教務作業"]["功能按鈕"].Add(new RibbonFeature(key, "高級中等學校原住民學生數及畢業生數權限"));
            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校原住民學生數及畢業生數"].Enable = FISCA.Permission.UserAcl.Current[key].Executable;

            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校原住民學生數及畢業生數"].Click += delegate
            {
                PrintBranch PrintBranch = new PrintBranch();
                PrintBranch.Show();
            };
        }

    }
}
