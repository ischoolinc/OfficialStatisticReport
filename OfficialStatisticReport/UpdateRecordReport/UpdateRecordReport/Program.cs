using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UpdateRecordReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "公務統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["高中職學校學生異動報告"].Enable = Permissions.高中職學校學生異動報告權限;
            item1["報表"]["高中職學校學生異動報告"].Click += delegate
            {
                Form1 fm = new Form1();
                fm.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.高中職學校學生異動報告, "高中職學校學生異動報告"));
        }
    }
}
