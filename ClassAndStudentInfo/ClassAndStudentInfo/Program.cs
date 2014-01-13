using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Permission;
using FISCA;
using K12.Data;

namespace ClassAndStudentInfo
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {

            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "公務統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["班級及學生概況1"].Enable = Permissions.班級及學生概況1權限;
            item1["報表"]["班級及學生概況1"].Click += delegate
            {
                Printer printer = new Printer();
                printer.Start();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.班級及學生概況1, "班級及學生概況1"));

        }
    }
}
