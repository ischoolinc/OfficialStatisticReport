using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Permission;
using FISCA;
using K12.Data;
using FISCA.Presentation;

namespace ArrangeClass
{
    public class Program
    {
        [MainMethod()]
        public static void Main()
        {

            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "資料統計"];   
            item1["報表"]["編班名冊"].Enable = Permissions.編班名冊權限;
            item1["報表"]["編班名冊"].Click += delegate
            {
                Printer printer = new Printer();
                printer.Start();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.編班名冊, "編班名冊"));

        }
    }
}
