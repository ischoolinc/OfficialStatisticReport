using FISCA.Presentation.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace ArrangeClass
{
    public partial class PrintForm : BaseForm
    {
        Dictionary<string, string> _ClassNameMappingDict;

        public PrintForm()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PrintForm_Load(object sender, EventArgs e)
        {
            //取得系統內班級名稱與設定檔對照
            _ClassNameMappingDict = QueryTransfer.GetClassNameList();

            //填入dgClass
            dgClass.Rows.Clear();
            foreach (string name in _ClassNameMappingDict.Keys)
            {
                int rowIdx = dgClass.Rows.Add();
                dgClass.Rows[rowIdx].Cells[colClassName.Index].Value = name;
                dgClass.Rows[rowIdx].Cells[colExpClassName.Index].Value = _ClassNameMappingDict[name];              
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {

            btnExport.Enabled = false;
            
            //儲存設定檔
            // 建立 XmlDocument 物件
            XmlDocument xmlDoc = new XmlDocument();

            // 建立 Configurations 元素
            XmlElement configurationsElement = xmlDoc.CreateElement("Configurations");

            // 建立 ClassName 元素
            XmlElement classNameElement = xmlDoc.CreateElement("ClassName");

            // 建立 Item 元素並加入屬性
            if (dgClass.Rows.Count > 0)
            {
                foreach (DataGridViewRow dr in dgClass.Rows)
                {
                    if (dr.IsNewRow)
                        continue;

                    string name = "", value = "";
                    if (dr.Cells[colClassName.Index].Value != null)
                        name = dr.Cells[colClassName.Index].Value.ToString();

                    if (dr.Cells[colExpClassName.Index].Value != null)
                        value = dr.Cells[colExpClassName.Index].Value.ToString();

                    if (name != "" && value != "")
                    {
                        XmlElement itemElement = xmlDoc.CreateElement("Item");
                        itemElement.SetAttribute("name", name);
                        itemElement.SetAttribute("value", value);

                        // 把 Item 元素加入 ClassName 元素
                        classNameElement.AppendChild(itemElement);
                    }
                }
            }
            // 把 ClassName 元素加入 Configurations 元素
            configurationsElement.AppendChild(classNameElement);

            // 把 Configurations 元素加入 XmlDocument 物件
            xmlDoc.AppendChild(configurationsElement);

            // 將 XmlDocument 物件轉成字串
            string xmlString = xmlDoc.OuterXml;

            if (!QueryTransfer.SaveConfigure(xmlString))
            {
                MsgBox.Show("編班名冊設定檔儲存失敗。");
                btnExport.Enabled = true;
                return;
            }

            Printer printer = new Printer();
            printer.Start();
            btnExport.Enabled = true;
        }
    }
}
