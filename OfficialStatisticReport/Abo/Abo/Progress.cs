using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Abo
{
    public class Progress
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "公務統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["原住民學生數及畢業生統計表"].Enable = Permissions.原住民學生數及畢業生統計表權限;
            item1["報表"]["原住民學生數及畢業生統計表"].Click += delegate
            {
                Form1 form = new Form1();
                form.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.原住民學生數及畢業生統計表, "原住民學生數及畢業生統計表"));
        }
    }
}
