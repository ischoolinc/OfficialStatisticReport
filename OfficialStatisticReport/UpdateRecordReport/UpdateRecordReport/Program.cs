using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA;
using FISCA.Presentation;
namespace UpdateRecordReport
{
    public class Program
    {
        [MainMethod]
        public static void main()
        {
            var key = "F6AD4FC5-7792-47F8-8714-9D7F1D9D441A";
            RoleAclSource.Instance["教務作業"]["功能按鈕"].Add(new RibbonFeature(key, "高級中等學校學生異動概況權限"));
            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校學生異動概況"].Enable = FISCA.Permission.UserAcl.Current[key].Executable;

            MotherForm.RibbonBarItems["教務作業", "資料統計"]["報表"]["高級中等學校學生異動概況"].Click += delegate
            {
                PrintBranch PrintBranch = new PrintBranch();
                PrintBranch.Show();
            };
        }
    }
}
