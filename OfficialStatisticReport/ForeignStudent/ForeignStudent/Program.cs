using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Internal;
using FISCA.Permission;

namespace ForeignStudent
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "公務統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["外國學生統計"].Enable = Permissions.外國學生統計權限;
            item1["報表"]["外國學生統計"].Click += delegate
            {
                Form1 form = new Form1();
                form.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.外國學生統計, "外國學生統計"));
        }
    }
}
