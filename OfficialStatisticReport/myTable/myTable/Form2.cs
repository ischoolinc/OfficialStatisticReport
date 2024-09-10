using Aspose.Cells;
using FISCA.Data;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using SmartSchool.Feature.Basic;
using K12.Data.Configuration;

namespace myTable
{
    public partial class Form2 : BaseForm
    {
        String _SchoolYear; //當前學年度
        private BackgroundWorker _BGWClassStudentAbsenceDetail; //背景模式

        Dictionary<String, String> _column3Items; //全部類別對照表,key=TagId,value=prefix+":"+name

        Dictionary<String, List<String>> _mappingData;//mapping資料


        Dictionary<String, List<String>> XML_mappingData;//mapping資料 (以XML 來儲存)

        Filter filter;
        List<String> AboList; //原住民生清單
        String dept;
        Workbook _wk;

        //Dictionary<string, string> _ClassTypeCodeDic;

        public Form2()
        {
            InitializeComponent();

            // 入學方式
            Column2Prepare();

            // 入學身分
            dataGridViewComboBoxExColumn2Prepare();

            //新生中具原住民身分者
            dataGridViewComboBoxExColumn3Prepare();

            //入學方式、入學身分、新生中具原住民身分者 來源欄位填值
            Column3Prepare();

            SchoolYearItem();

            //LoadLastRecord(); //舊CODE 方法，不再使用，故註解。

            LoadConfigXml();

            LoadClassTypeCodeDic();
        }

        private void LoadClassTypeCodeDic()
        {
            //_ClassTypeCodeDic.Clear();
            //_ClassTypeCodeDic.Add("1", "日間部");
            //_ClassTypeCodeDic.Add("2", "夜間部");
            //_ClassTypeCodeDic.Add("3", "實用技能學程(一般班)");
            //_ClassTypeCodeDic.Add("4", "建教班");
            //_ClassTypeCodeDic.Add("6", "產學訓合作計畫班(產學合作班)");
            //_ClassTypeCodeDic.Add("7", "重點產業班/台德菁英班/雙軌旗艦訓練計畫專班");
            //_ClassTypeCodeDic.Add("8", "建教僑生專班");
            //_ClassTypeCodeDic.Add("9", "實驗班");
            //_ClassTypeCodeDic.Add("01", "進修部(核定班)");
            //_ClassTypeCodeDic.Add("02", "編制班");
            //_ClassTypeCodeDic.Add("03", "自給自足班");
            //_ClassTypeCodeDic.Add("04", "員工進修班");
            //_ClassTypeCodeDic.Add("05", "重點產業班");
            //_ClassTypeCodeDic.Add("06", "產業人力套案專班");
        }
        private void LoadConfigXml()
        {
            ConfigData cd = K12.Data.School.Configuration["新生入學統計報表_來源目標設定Config"];

            XmlElement config = cd.GetXml("XmlData", null);

            // 假如要刪除所有舊資料Config 資料，可以啟用下面三行。
            //config.RemoveAll(); 
            //cd.SetXml("XmlData", config);
            //cd.Save();

            if (config != null) //如果不是空的
            {
                XmlElement EnterSchool_Way = (XmlElement)config.SelectSingleNode("入學方式");

                XmlElement EnterSchool_identity = (XmlElement)config.SelectSingleNode("入學身分");

                XmlElement FreshMenWith_Aboriginal_Identity = (XmlElement)config.SelectSingleNode("新生中具原住民身分者");

                XmlNodeList EnterSchool_WayList;

                XmlNodeList EnterSchool_identityList;

                XmlNodeList FreshMenWith_Aboriginal_IdentityList;

                DataGridViewRow row;

                if (EnterSchool_Way != null) //  Config內有設定才做讀取
                {

                    EnterSchool_WayList = EnterSchool_Way.SelectNodes("item");

                    foreach (XmlElement item in EnterSchool_WayList)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX1);
                        row.Cells[0].Value = item.HasAttribute("target") ? item.GetAttribute("target") : "";
                        row.Cells[1].Value = item.HasAttribute("source") ? item.GetAttribute("source") : "";
                        dataGridViewX1.Rows.Add(row);
                    }
                }
                else
                {
                    //Config無資料則提供預設標記
                    for (int i = 0; i < Column2.Items.Count; i++)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX1);
                        row.Cells[0].Value = Column2.Items[i];//target
                        dataGridViewX1.Rows.Add(row);
                    }
                }

                if (EnterSchool_identity != null) //  Config內有設定才做讀取
                {
                    EnterSchool_identityList = EnterSchool_identity.SelectNodes("item");

                    foreach (XmlElement item in EnterSchool_identityList)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX2);
                        row.Cells[0].Value = item.HasAttribute("target") ? item.GetAttribute("target") : "";
                        row.Cells[1].Value = item.HasAttribute("source") ? item.GetAttribute("source") : "";
                        dataGridViewX2.Rows.Add(row);
                    }
                }
                else
                {
                    //Config無資料則提供預設標記
                    for (int i = 0; i < dataGridViewComboBoxExColumn1.Items.Count; i++)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX2);
                        row.Cells[0].Value = dataGridViewComboBoxExColumn1.Items[i];//target
                        dataGridViewX2.Rows.Add(row);
                    }
                }

                if (FreshMenWith_Aboriginal_Identity != null) //  Config內有設定才做讀取
                {
                    FreshMenWith_Aboriginal_IdentityList = FreshMenWith_Aboriginal_Identity.SelectNodes("item");

                    foreach (XmlElement item in FreshMenWith_Aboriginal_IdentityList)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX3);
                        row.Cells[0].Value = item.HasAttribute("target") ? item.GetAttribute("target") : "";
                        row.Cells[1].Value = item.HasAttribute("source") ? item.GetAttribute("source") : "";
                        dataGridViewX3.Rows.Add(row);
                    }
                }
                else
                {
                    //Config無資料則提供預設標記
                    for (int i = 0; i < dataGridViewComboBoxExColumn3.Items.Count; i++)
                    {
                        row = new DataGridViewRow();
                        row.CreateCells(dataGridViewX3);
                        row.Cells[0].Value = dataGridViewComboBoxExColumn3.Items[i];//target
                        dataGridViewX3.Rows.Add(row);
                    }
                }
            }
            else
            {
                #region 產生空白設定檔

                config = new XmlDocument().CreateElement("新生入學統計報表_來源目標設定Config");

                XmlElement EnterSchool_Way = config.OwnerDocument.CreateElement("入學方式");

                List<string> EnterSchoolWays = new List<string>();

                #region 1.入學方式
                //九種入學方式
                EnterSchoolWays.Add("入學方式:免試入學--校內直升");
                EnterSchoolWays.Add("入學方式:免試入學--就學區免試(含共同就學區)");
                EnterSchoolWays.Add("入學方式:免試入學--技優甄審");
                EnterSchoolWays.Add("入學方式:免試入學--免試獨招");
                EnterSchoolWays.Add("入學方式:免試入學--其他");
                EnterSchoolWays.Add("入學方式:特色招生--考試分發");
                EnterSchoolWays.Add("入學方式:特色招生--甄選入學");
                EnterSchoolWays.Add("入學方式:適性輔導安置(十二年安置)");
                EnterSchoolWays.Add("入學方式:其他");

                int i = 1;

                foreach (string way in EnterSchoolWays)
                {
                    XmlElement EnterSchool_Way_Item = EnterSchool_Way.OwnerDocument.CreateElement("item");

                    EnterSchool_Way_Item.SetAttribute("ID", "" + i);

                    EnterSchool_Way_Item.SetAttribute("target", way);

                    EnterSchool_Way_Item.SetAttribute("source", "");

                    EnterSchool_Way.AppendChild(EnterSchool_Way_Item);

                    i++;
                }
                #endregion

                #region 2.入學身分
                XmlElement EnterSchool_identity = config.OwnerDocument.CreateElement("入學身分");

                XmlElement EnterSchool_identity1 = EnterSchool_identity.OwnerDocument.CreateElement("item");

                EnterSchool_identity1.SetAttribute("ID", "1");

                EnterSchool_identity1.SetAttribute("target", "入學身份:一般生(非外加錄取)");

                EnterSchool_identity1.SetAttribute("source", "");

                EnterSchool_identity.AppendChild(EnterSchool_identity1);

                XmlElement EnterSchool_identity2 = EnterSchool_identity.OwnerDocument.CreateElement("item");

                EnterSchool_identity2.SetAttribute("ID", "2");

                EnterSchool_identity2.SetAttribute("target", "入學身份:外加錄取--原住民生");

                EnterSchool_identity2.SetAttribute("source", "");

                EnterSchool_identity.AppendChild(EnterSchool_identity2);

                XmlElement EnterSchool_identity3 = EnterSchool_identity.OwnerDocument.CreateElement("item");

                EnterSchool_identity3.SetAttribute("ID", "3");

                EnterSchool_identity3.SetAttribute("target", "入學身份:外加錄取--身心障礙生");

                EnterSchool_identity3.SetAttribute("source", "");

                EnterSchool_identity.AppendChild(EnterSchool_identity3);

                XmlElement EnterSchool_identity4 = EnterSchool_identity.OwnerDocument.CreateElement("item");

                EnterSchool_identity4.SetAttribute("ID", "4");

                EnterSchool_identity4.SetAttribute("target", "入學身份:外加錄取--其他");

                EnterSchool_identity4.SetAttribute("source", "");

                EnterSchool_identity.AppendChild(EnterSchool_identity4);

                config.AppendChild(EnterSchool_identity);
                #endregion

                #region 3.新生中具原住民身分者
                XmlElement FreshMenWith_Aboriginal_Identity = config.OwnerDocument.CreateElement("新生中具原住民身分者");

                XmlElement FreshMenWith_Aboriginal_Identity1 = FreshMenWith_Aboriginal_Identity.OwnerDocument.CreateElement("item");

                FreshMenWith_Aboriginal_Identity1.SetAttribute("ID", "1");

                FreshMenWith_Aboriginal_Identity1.SetAttribute("target", "新生中具原住民身分者");

                FreshMenWith_Aboriginal_Identity1.SetAttribute("source", "");

                FreshMenWith_Aboriginal_Identity.AppendChild(FreshMenWith_Aboriginal_Identity1);

                config.AppendChild(FreshMenWith_Aboriginal_Identity);
                #endregion

                cd.SetXml("XmlData", config);

                #endregion
            }
            cd.Save();
        }

        ////Column2的選單產生  (1.入學方式)
        private void Column2Prepare()
        {
            //List<String> prefix = new List<string>();
            //List<String> name = new List<string>();
            //prefix.Add("甄選入學");
            //prefix.Add("申請入學");
            //prefix.Add("登記分發");
            //prefix.Add("直升入學");
            //prefix.Add("免試入學");
            //prefix.Add("其他");
            //name.Add("一般生");
            //name.Add("原住民生");
            //name.Add("身心障礙生");
            //name.Add("其他");

            //foreach (String a in prefix)
            //{
            //    foreach (String b in name)
            //    {
            //        Column2.Items.Add(a + ":" + b);
            //    }
            //}

            Column2.Items.Add("入學方式:免試入學--校內直升");
            Column2.Items.Add("入學方式:免試入學--就學區免試(含共同就學區)");
            Column2.Items.Add("入學方式:免試入學--技優甄審");
            Column2.Items.Add("入學方式:免試入學--免試獨招");
            Column2.Items.Add("入學方式:免試入學--其他");
            Column2.Items.Add("入學方式:特色招生--考試分發");
            Column2.Items.Add("入學方式:特色招生--甄選入學");
            Column2.Items.Add("入學方式:適性輔導安置(十二年安置)");
            Column2.Items.Add("入學方式:其他");
        }

        //入學方式、入學身分、新生中具原住民身分者 來源欄位填值
        // 來源的選擇 是取 所有屬於 "學生" 的"類別"
        private void Column3Prepare()
        {
            _column3Items = new Dictionary<String, String>();
            QueryHelper _Q = new QueryHelper();

            DataTable dt = _Q.Select("select * from tag where category='Student' order by prefix,name");
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String prefix = row["prefix"].ToString();
                String name = row["name"].ToString();
                if (!_column3Items.ContainsKey(id))
                {
                    _column3Items.Add(id, prefix + ":" + name);
                }
            }
            foreach (KeyValuePair<String, String> k in _column3Items)
            {
                String item = k.Value;
                if (item.Substring(0, 1) == ":") item = item.Substring(1); //若選項開頭為":",擷取第二字元到結尾

                Column3.Items.Add(item); //建立Column3的來源選單

                dataGridViewComboBoxExColumn2.Items.Add(item);//建立dataGridViewComboBoxExColumn2的來源選單

                dataGridViewComboBoxExColumn4.Items.Add(item);//建立dataGridViewComboBoxExColumn4的來源選單
            }
        }

        //// 2.入學身分
        private void dataGridViewComboBoxExColumn2Prepare()
        {
            dataGridViewComboBoxExColumn1.Items.Add("入學身份:一般生(非外加錄取)");
            dataGridViewComboBoxExColumn1.Items.Add("入學身份:外加錄取--原住民生");
            dataGridViewComboBoxExColumn1.Items.Add("入學身份:外加錄取--身心障礙生");
            dataGridViewComboBoxExColumn1.Items.Add("入學身份:外加錄取--其他");
        }

        //3.新生中具原住民身分者
        private void dataGridViewComboBoxExColumn3Prepare()
        {
            dataGridViewComboBoxExColumn3.Items.Add("新生中具原住民身分者");
        }

        // 列印
        private void buttonX1_Click(object sender, EventArgs e)
        {
            if (tryConvert(comboBoxEx2.Text))
            {
                try
                {
                    // 儲存現在的Setting，供下次使用可直接利用
                    //SaveMappingRecord();
                    SaveMappingXmlRecord();

                    //ReadMappingData();
                    ReadXMLMappingData();
                    DataSetting();
                }
                catch
                {
                    MessageBox.Show("網路或資料庫異常,請稍後再試...");
                    this.buttonX1.Enabled = true;
                    this.linkLabel1.Enabled = true;
                    this.dataGridViewX1.Enabled = true;
                    this.comboBoxEx1.Enabled = true;
                    this.comboBoxEx2.Enabled = true;
                }
            }
        }

        //關閉
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //讀取DataGridView資料，以建立XML_mappingData
        void ReadXMLMappingData()
        {
            SetXMLMappingDataKey(); //初始化_mappingData

            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Cells[0].Value != null && r.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    String id = "";
                    foreach (KeyValuePair<String, String> k in _column3Items) //尋找選項的TagID
                    {
                        String item = r.Cells[1].Value.ToString();
                        if (!item.Contains(":")) //若選項無":"字串代表建立時prefix為空白,查詢時需補上":"
                        {
                            item = ":" + item;
                        }
                        if (item == k.Value)
                        {
                            id = k.Key;
                        }
                    }

                    if (id != "") //找不到對應ID不執行
                    {
                        if (!XML_mappingData.ContainsKey(r.Cells[0].Value.ToString())) //建立目標對應ID的字典
                        {
                            XML_mappingData.Add(r.Cells[0].Value.ToString(), new List<string>());
                        }
                        XML_mappingData[r.Cells[0].Value.ToString()].Add(id); //收集Mapping的TagId
                    }
                }
            }
            foreach (DataGridViewRow dgvR in dataGridViewX2.Rows)
            {
                if (dgvR.Cells[0].Value != null && dgvR.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    String id = "";
                    foreach (KeyValuePair<String, String> keyPair in _column3Items) //尋找選項的TagID
                    {
                        String item = dgvR.Cells[1].Value.ToString();
                        if (!item.Contains(":")) //若選項無":"字串代表建立時prefix為空白,查詢時需補上":"
                        {
                            item = ":" + item;
                        }
                        if (item == keyPair.Value)
                        {
                            id = keyPair.Key;
                        }
                    }

                    if (id != "") //找不到對應ID不執行
                    {
                        if (!XML_mappingData.ContainsKey(dgvR.Cells[0].Value.ToString())) //建立目標對應ID的字典
                        {
                            XML_mappingData.Add(dgvR.Cells[0].Value.ToString(), new List<string>());
                        }
                        XML_mappingData[dgvR.Cells[0].Value.ToString()].Add(id); //收集Mapping的TagId
                    }
                }
            }

            foreach (DataGridViewRow dgvR in dataGridViewX3.Rows)
            {
                if (dgvR.Cells[0].Value != null && dgvR.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    String id = "";
                    foreach (KeyValuePair<String, String> k in _column3Items) //尋找選項的TagID
                    {
                        String item = dgvR.Cells[1].Value.ToString();
                        if (!item.Contains(":")) //若選項無":"字串代表建立時prefix為空白,查詢時需補上":"
                        {
                            item = ":" + item;
                        }
                        if (item == k.Value)
                        {
                            id = k.Key;
                        }
                    }

                    if (id != "") //找不到對應ID不執行
                    {
                        if (!XML_mappingData.ContainsKey(dgvR.Cells[0].Value.ToString())) //建立目標對應ID的字典
                        {
                            XML_mappingData.Add(dgvR.Cells[0].Value.ToString(), new List<string>());
                        }
                        XML_mappingData[dgvR.Cells[0].Value.ToString()].Add(id); //收集Mapping的TagId
                    }
                }
            }
            foreach (KeyValuePair<String, List<String>> k in XML_mappingData) //刪去value中的重複ID
            {
                for (int i = 0; i < k.Value.Count; i++)
                {
                    String s = k.Value[i];
                    int count = 0; //重複次數
                    for (int j = 0; j < k.Value.Count; j++)
                    {
                        if (s == k.Value[j]) //發現相同ID
                        {
                            count++;
                            if (count > 1) //只允許大於1
                            {
                                k.Value.Remove(s);
                                j--;
                                count--;
                            }
                        }
                    }
                }
            }
        }

        //讀取DataGridView資料， 舊方法，已不使用
        void ReadMappingData()
        {
            SetMappingDataKey(); //初始化_mappingData
            foreach (DataGridViewRow r in dataGridViewX1.Rows)
            {
                if (r.Cells[0].Value != null && r.Cells[1].Value != null)  //欄位有空值跳下一行
                {
                    String id = "";
                    foreach (KeyValuePair<String, String> k in _column3Items) //尋找選項的TagID
                    {
                        String item = r.Cells[1].Value.ToString();
                        if (!item.Contains(":")) //若選項無":"字串代表建立時prefix為空白,查詢時需補上":"
                        {
                            item = ":" + item;
                        }
                        if (item == k.Value)
                        {
                            id = k.Key;
                        }
                    }

                    if (id != "") //找不到對應ID不執行
                    {
                        if (!_mappingData.ContainsKey(r.Cells[0].Value.ToString())) //建立目標對應ID的字典
                        {
                            _mappingData.Add(r.Cells[0].Value.ToString(), new List<string>());
                        }
                        _mappingData[r.Cells[0].Value.ToString()].Add(id); //收集Mapping的TagId
                    }
                }
            }

            foreach (KeyValuePair<String, List<String>> k in _mappingData) //刪去value中的重複ID
            {
                for (int i = 0; i < k.Value.Count; i++)
                {
                    String s = k.Value[i];
                    int count = 0; //重複次數
                    for (int j = 0; j < k.Value.Count; j++)
                    {
                        if (s == k.Value[j]) //發現相同ID
                        {
                            count++;
                            if (count > 1) //只允許大於1
                            {
                                k.Value.Remove(s);
                                j--;
                                count--;
                            }
                        }
                    }
                }
            }
        }

        void DataSetting()
        {
            dept = this.comboBoxEx1.Text;
            _SchoolYear = comboBoxEx2.Text;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("正在產生新生入學方式統計表...");
            this.buttonX1.Enabled = false;
            this.linkLabel1.Enabled = false;
            this.dataGridViewX1.Enabled = false;
            this.comboBoxEx1.Enabled = false;
            this.comboBoxEx2.Enabled = false;
            _BGWClassStudentAbsenceDetail = new BackgroundWorker();
            _BGWClassStudentAbsenceDetail.DoWork += new DoWorkEventHandler(_BGWClassStudentAbsenceDetail_DoWork);
            _BGWClassStudentAbsenceDetail.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWClassStudentAbsenceDetail_Completed);
            this.picLoding.Visible = true;
            _BGWClassStudentAbsenceDetail.RunWorkerAsync();

        }

        private void _BGWClassStudentAbsenceDetail_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            this.picLoding.Visible = false;
            this.buttonX1.Enabled = true;
            this.linkLabel1.Enabled = true;
            this.dataGridViewX1.Enabled = true;
            this.comboBoxEx1.Enabled = true;
            this.comboBoxEx2.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("產生 新生入學方式統計表 已完成");

            SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
            sd.Title = "另存新檔";
            sd.FileName = "新生入學方式統計表.xls";
            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
            if (sd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _wk.Save(sd.FileName);
                    if (filter.error_list.Count > 0)
                    {
                        MessageBox.Show("發現" + filter.error_list.Count + "筆異常資料未列入統計\r\n詳細資料請確認報表中的[異常資料表]");
                    }
                    System.Diagnostics.Process.Start(sd.FileName);
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    this.Enabled = true;
                    return;
                }
            }
        }

        private void _BGWClassStudentAbsenceDetail_DoWork(object sender, DoWorkEventArgs e)
        {

            Dictionary<String, myStudent> myDic = new Dictionary<string, myStudent>();
            List<myStudent> mylist = new List<myStudent>();
            QueryHelper _Q = new QueryHelper();

            //SQL查詢要求的年級資料
            //DataTable dt = _Q.Select("select student.id,student.name,student.gender,student.permanent_address,student.ref_class_id,student.status,class.class_name,class.grade_year,dept.name as dept_name,tag_student.ref_tag_id from student left join class on student.ref_class_id=class.id left join dept on class.ref_dept_id=dept.id left join tag_student on student.id= tag_student.ref_student_id where student.status in ('1','4','16') and class.grade_year='1'");

            //2017/1/17 穎驊改寫修正， 上面的SQL 僅會抓取學生 所屬班級上的 科別 ，而不會抓學生身上自己設定的科別 ， 現在該改SQL，  優先先抓取學生自己的科別。 若無 則以 班級設定的科別帶入。
            //2022-07-12 Cynthia  將新生異動的班別也讀出來 
            //DataTable dt = _Q.Select("select student.id,student.name,student.gender,student.permanent_address,student.ref_class_id,student.status,student.ref_dept_id as student_ref_dept_id,class.ref_dept_id as class_ref_dept_id ,class.class_name,class.grade_year,dept.name as dept_name,tag_student.ref_tag_id from student left join class on student.ref_class_id=class.id  left join dept on  case    when student.ref_dept_id is null   then class.ref_dept_id=dept.id   else student.ref_dept_id=dept.id  end left join tag_student on student.id= tag_student.ref_student_id where student.status in ('1','4','16') and class.grade_year='1'");
            string sql = @"WITH update_record_info AS(
	SELECT 
		ref_student_id
		, school_year
		, update_desc
		, code
		, array_to_string(xpath('//ClassType/text()', ClassTypesEle), '') AS Class_Type
	FROM 
		(
			SELECT *
				, CAST(update_code AS INT) AS code
				, unnest(xpath('//ContextInfo/ClassType', xmlparse(content context_info))) AS ClassTypesEle
			FROM 
				update_record 
			WHERE
				CAST(update_code AS INT) <100 AND school_year ={0} 
		) AS record
)
SELECT 
student.id,student.name,student.gender,student.permanent_address,student.ref_class_id,student.status
,student.ref_dept_id as student_ref_dept_id,class.ref_dept_id as class_ref_dept_id 
,class.class_name,class.grade_year,dept.name as dept_name,tag_student.ref_tag_id 
, update_record_info.code
, TRIM(update_record_info.Class_Type) AS Class_Type
FROM student 
left join class on student.ref_class_id=class.id  
left join dept on  case  when student.ref_dept_id is null   then class.ref_dept_id=dept.id   else student.ref_dept_id=dept.id  end 
left join tag_student on student.id= tag_student.ref_student_id 
left join update_record_info on student.id=update_record_info.ref_student_id 
WHERE student.status in ('1','4','16') and class.grade_year='1'
ORDER BY dept_name, TRIM(update_record_info.Class_Type) ";
            sql = string.Format(sql, _SchoolYear);
            DataTable dt = _Q.Select(sql);


            int num = 0;
            //建立myStuden物件放至List中
            foreach (DataRow row in dt.Rows)
            {
                String id = row["id"].ToString();
                String name = row["name"].ToString();
                String gender = row["gender"].ToString();
                String ref_class_id = row["ref_class_id"].ToString();
                String class_name = row["class_name"].ToString();
                String grade_year = row["grade_year"].ToString();
                String dept_name = row["dept_name"].ToString();
                String ref_tag_id = row["ref_tag_id"].ToString();
                String classType = row["Class_Type"].ToString(); //班別

                //戶籍 縣市
                String Before_School_Location = "";


                // 選取此學生 的前期畢業學校 資訊(用來取得 前學校所在地)
                K12.Data.BeforeEnrollmentRecord ber = K12.Data.BeforeEnrollment.SelectByStudentID(id);

                if (ber != null)
                {
                    Before_School_Location = ber.SchoolLocation;

                }

                System.Xml.Linq.XDocument XD;
                //戶籍 縣市
                String County = "";
                string permanent_address = "" + row["permanent_address"];

                if (permanent_address != "") // 如果戶籍地不是空值
                {
                    XD = System.Xml.Linq.XDocument.Parse(permanent_address);
                    System.Xml.Linq.XElement element = XD.Element("AddressList");

                    if (element.Element("Address") != null)
                    {
                        if (element.Element("Address").Element("County") != null)
                        {
                            County = element.Element("Address").Element("County").Value;
                        };
                    };
                }
                else
                {
                    County = "";
                }

                if (!myDic.ContainsKey(id)) //ID當key,不存在就建立
                {
                    //{
                    //    string location = Before_School_Location;
                    //}
                    myDic.Add(id, new myStudent(id, name, gender, ref_class_id, class_name, grade_year, dept_name, County, Before_School_Location, classType, new List<string>()));
                }
                myDic[id].Tag.Add(ref_tag_id);
                num++;
            }

            //取得新生異動資料清單
            List<K12.Data.UpdateRecordRecord> records = K12.Data.UpdateRecord.SelectByStudentIDs(myDic.Keys.ToList());
            foreach (KeyValuePair<String, myStudent> kvp in myDic)
            {
                if (CheckStudentStatus(records, kvp.Key)) //檢查新生異動資料是否符合
                {
                    mylist.Add(kvp.Value);  //符合者加入mylist清單
                }
            }

            filter = new Filter(mylist, dept);
            Export();

        }

        //確認學生為一般新生,排除重讀生等其他狀態
        public bool CheckStudentStatus(List<K12.Data.UpdateRecordRecord> records, String id)
        {
            foreach (K12.Data.UpdateRecordRecord record in records)
            {
                if (record.StudentID == id)  //找到符合的ID開始後續比對
                {
                    if (record.SchoolYear.ToString() == _SchoolYear) //確認學年度為當前學年度
                    {
                        //2017/1/19 穎驊筆記，經過詢問恩正後 各個種類的異動代碼 有屬於自己的區間， 
                        //以新生異動而言，是001~ 100 之間 ，又內容順序幾乎不會變所以他才敢以 Code <100 作為判別是 合法新生的依據
                        //為了避免後人產生太多疑惑，我詳列各個異動代碼的全意
                        // 001 持國民中學畢業證書者(含國中補校)
                        // 002 持國民中學補習學校資格證明書者
                        // 003 持國民中學補習學校結(修)業證明書者
                        // 004 持國民中學修業證明書者(修習三年級課程)
                        // 005 持國民中學畢業程度學力鑑定考詴及格證明書者
                        // 006 回國僑生(專案核准)
                        // 007 持大陸學歷者(需附證明文件)
                        // 008 特殊教育學校學生(需附證明文件)
                        // 009 持國外學歷者(需附證明文件)
                        // 010 取得相當於丙級(含)以上技術士證之資格者(需附證明文件)
                        // 011 持香港澳門學歷者(需附證明文件)
                        // 099 其他(需附證明文件

                        if (Convert.ToInt16(record.UpdateCode) < 100) //異動代碼小於100
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        //輸出至Excel
        public void Export()
        {
            _wk = new Workbook();
            Worksheet ws;
            Cells cs;
            int index, row;
            _wk.Worksheets.Add();
            ////該年級學生總表
            //_wk.Worksheets.Add();
            //ws = _wk.Worksheets[2];
            //ws.Name = "All_List";
            //cs = ws.Cells;
            //cs["A1"].PutValue("ID");
            //cs["B1"].PutValue("Name");
            //cs["C1"].PutValue("Gender");
            //cs["D1"].PutValue("Ref_Class_Id");
            //cs["E1"].PutValue("Class_Name");
            //cs["F1"].PutValue("Grade_Year");
            //cs["G1"].PutValue("Dept_Name");
            //cs["H1"].PutValue("ref_tag_id");
            //index = 1;
            //foreach (myStudent s in filter.clean_list)
            //{
            //    cs[index, 0].PutValue(s.Id);
            //    cs[index, 1].PutValue(s.Name);
            //    cs[index, 2].PutValue(s.Gender);
            //    cs[index, 3].PutValue(s.Ref_class_id);
            //    cs[index, 4].PutValue(s.Class_name);
            //    cs[index, 5].PutValue(s.Grade_year);
            //    cs[index, 6].PutValue(s.Dept_name);
            //    String column7 = "";
            //    foreach (String l in s.Tag)
            //    {
            //        column7 += l + ",";
            //    }
            //    cs[index, 7].PutValue(column7);
            //    index++;
            //}

            //該科別異常的學生資料表


            ws = _wk.Worksheets[1];
            ws.Name = "異常資料表";
            cs = ws.Cells;
            cs["A1"].PutValue("系統編號");
            cs["B1"].PutValue("姓名");
            cs["C1"].PutValue("性別");
            //cs["D1"].PutValue("Ref_Class_Id");
            cs["D1"].PutValue("班級名稱");
            cs["E1"].PutValue("年級");
            cs["F1"].PutValue("科別名稱");
            cs["G1"].PutValue("班別");
            //cs["H1"].PutValue("ref_tag_id");
            index = 1;
            foreach (myStudent s in filter.error_list)
            {
                cs[index, 0].PutValue(s.Id);
                cs[index, 1].PutValue(s.Name);
                cs[index, 2].PutValue(s.Gender);
                //cs[index, 3].PutValue(s.Ref_class_id);
                cs[index, 3].PutValue(s.Class_name);
                cs[index, 4].PutValue(s.Grade_year);
                cs[index, 5].PutValue(s.Dept_name);
                cs[index, 6].PutValue(s.Class_Type);
                //String column7 = "";
                //foreach (String l in s.Tag)
                //{
                //    column7 += l + ",";
                //}
                //cs[index, 7].PutValue(column7);
                index++;
            }


            #region 1.入學方式 TagID 整理
            //下面共九種入學方式

            // 入學方式:免試入學--校內直升 ，所標記類別 tag ID List
            List<string> enter_Way_NoExam_SchoolPromote_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:免試入學--校內直升")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_NoExam_SchoolPromote_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:免試入學--就學區免試(含共同就學區) ，所標記類別 tag ID List
            List<string> enter_Way_NoExam_SchoolArea_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:免試入學--就學區免試(含共同就學區)")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_NoExam_SchoolArea_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:免試入學--技優甄審 ，所標記類別 tag ID List
            List<string> enter_Way_NoExam_GoodSkill_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:免試入學--技優甄審")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_NoExam_GoodSkill_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:免試入學--免試獨招 ，所標記類別 tag ID List
            List<string> enter_Way_NoExam_NoExam_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:免試入學--免試獨招")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_NoExam_NoExam_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:免試入學--其他 ，所標記類別 tag ID List
            List<string> enter_Way_NoExam_Other_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:免試入學--其他")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_NoExam_Other_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:特色招生--考試分發 ，所標記類別 tag ID List
            List<string> enter_Way_SpecialRecuit_ExamAtribute_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:特色招生--考試分發")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_SpecialRecuit_ExamAtribute_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:特色招生--甄選入學 ，所標記類別 tag ID List
            List<string> enter_Way_SpecialRecuit_Selection_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:特色招生--甄選入學")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_SpecialRecuit_Selection_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:適性輔導安置(十二年安置) ，所標記類別 tag ID List
            List<string> enter_Way_SafelySet_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:適性輔導安置(十二年安置)")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_SafelySet_ID_list.Add(s);
                    }
                }
            }

            // 入學方式:其他 ，所標記類別 tag ID List
            List<string> enter_Way_Other_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學方式:其他")
                {
                    foreach (String s in map.Value)
                    {
                        enter_Way_Other_ID_list.Add(s);
                    }
                }
            }

            List<List<string>> EnterWayTagsID_Mapping_List = new List<List<string>>();

            EnterWayTagsID_Mapping_List.Add(enter_Way_NoExam_SchoolPromote_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_NoExam_SchoolArea_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_NoExam_GoodSkill_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_NoExam_NoExam_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_NoExam_Other_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_SpecialRecuit_ExamAtribute_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_SpecialRecuit_Selection_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_SafelySet_ID_list);
            EnterWayTagsID_Mapping_List.Add(enter_Way_Other_ID_list);
            #endregion


            #region 2.入學身分 TagID 整理
            // 入學身份:一般生(非外加錄取) ，所標記類別 tag ID List
            List<string> enter_identity_normal_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學身份:一般生(非外加錄取)")
                {
                    foreach (String s in map.Value)
                    {
                        enter_identity_normal_ID_list.Add(s);
                    }
                }
            }

            // 入學身份:外加錄取--原住民生 ，所標記類別 tag ID List
            List<string> enter_identity_aboriginal_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學身份:外加錄取--原住民生")
                {
                    foreach (String s in map.Value)
                    {
                        enter_identity_aboriginal_ID_list.Add(s);
                    }
                }
            }

            // 入學身份:外加錄取--身心障礙生 ，所標記類別 tag ID List
            List<string> enter_identity_IEP_ID_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學身份:外加錄取--身心障礙生")
                {
                    foreach (String s in map.Value)
                    {
                        enter_identity_IEP_ID_list.Add(s);
                    }
                }
            }


            // 入學身份:外加錄取--其他 ，所標記類別 tag ID List
            List<string> enter_identity_Other_list = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "入學身份:外加錄取--其他")
                {
                    foreach (String s in map.Value)
                    {
                        enter_identity_Other_list.Add(s);
                    }
                }
            }

            List<List<string>> EnterIdentityTagsID_Mapping_List = new List<List<string>>();

            EnterIdentityTagsID_Mapping_List.Add(enter_identity_normal_ID_list);
            EnterIdentityTagsID_Mapping_List.Add(enter_identity_aboriginal_ID_list);
            EnterIdentityTagsID_Mapping_List.Add(enter_identity_IEP_ID_list);
            EnterIdentityTagsID_Mapping_List.Add(enter_identity_Other_list);
            #endregion


            #region 3.新生中具原住民身分者 TagID 整理
            // 新生具有原住民身分者 ，所標記類別 tag ID List
            List<string> aboIDList = new List<string>();

            foreach (KeyValuePair<String, List<String>> map in XML_mappingData)
            {
                if ("" + map.Key == "新生中具原住民身分者")
                {
                    foreach (String s in map.Value)
                    {
                        aboIDList.Add(s);
                    }
                }
            }
            #endregion


            //新生入學方式統計表-- 填值
            Workbook wk2 = new Workbook();
            wk2.Open(new MemoryStream(Properties.Resources.template_105_7_ver_)); //開啟範本文件 // 2017/1/17 穎驊筆記，在此載入105/7 最新版

            _wk.Worksheets[0].Copy(wk2.Worksheets[0]); //複製範本文件
            ws = _wk.Worksheets[0];
            ws.Name = "新生入學方式統計表";
            cs = ws.Cells;

            index = 12;
            //todo 
            int col = 10;

            int flexInsex = 1; //因為樣板不是每一項都是佔1格 ，有些有二合一合併，所以靠一個參數彈性調整(可以自行 去看 Resource/ template(105.7ver) 內有許多 合併欄位)

            List<myStudent> summary = new List<myStudent>(); //建立summary清單收集dic_byDept的展開學生物件

            #region 每一科別+班別的整理 
            //dept_ClassTypeDic //dic_byDept
            foreach (KeyValuePair<String, List<myStudent>> k in filter.dept_ClassTypeDic)
            {
                string[] keyArray = k.Key.Split('⊕');
                col = 10;
                //Table1 Left
                cs[index, 1].PutValue(filter.getDeptCode(k.Key)); //科別代碼
                cs[index, 2].PutValue(keyArray[0]); //科別名稱
                if (keyArray.Length >= 2)
                    if (filter.ClassTypeCodeDic.ContainsKey(keyArray[1]))
                        cs[index, 3].PutValue(filter.ClassTypeCodeDic[keyArray[1]]); //班別名稱
                cs[index, 6].PutValue(filter.getClassCount(k.Value)); //實際招生班數
                cs[index, 7].PutValue(k.Value.Count); //學生總計數
                cs[index, 8].PutValue(filter.getGenderCount(k.Value, "1")); //男生總數
                cs[index, 9].PutValue(filter.getGenderCount(k.Value, "0")); //女生總數

                foreach (myStudent s in k.Value)
                {
                    summary.Add(s); //展開dic_byDept,收集內容的myStudent物件
                }

                #region 學生 依入學九大方式 做的分類
                //入學方式:免試入學--校內直升 ，Student List           
                List<myStudent> enter_Way_NoExam_SchoolPromote_Student_list = new List<myStudent>();
                enter_Way_NoExam_SchoolPromote_Student_list = filter.getListByTagId(enter_Way_NoExam_SchoolPromote_ID_list, k.Value);

                //入學方式:免試入學--就學區免試(含共同就學區) ，Student List           
                List<myStudent> enter_Way_NoExam_SchoolArea_Student_list = new List<myStudent>();
                enter_Way_NoExam_SchoolArea_Student_list = filter.getListByTagId(enter_Way_NoExam_SchoolArea_ID_list, k.Value);

                //入學方式:免試入學--技優甄審 ，Student List           
                List<myStudent> enter_Way_NoExam_GoodSkill_Student_list = new List<myStudent>();
                enter_Way_NoExam_GoodSkill_Student_list = filter.getListByTagId(enter_Way_NoExam_GoodSkill_ID_list, k.Value);

                //入學方式:免試入學--免試獨招 ，Student List           
                List<myStudent> enter_Way_NoExam_NoExam_Student_list = new List<myStudent>();
                enter_Way_NoExam_NoExam_Student_list = filter.getListByTagId(enter_Way_NoExam_NoExam_ID_list, k.Value);

                //入學方式:免試入學--其他 ，Student List           
                List<myStudent> enter_Way_NoExam_Other_Student_list = new List<myStudent>();
                enter_Way_NoExam_Other_Student_list = filter.getListByTagId(enter_Way_NoExam_Other_ID_list, k.Value);

                //入學方式:特色招生--考試分發         
                List<myStudent> enter_Way_SpecialRecuit_ExamAtribute_Student_list = new List<myStudent>();
                enter_Way_SpecialRecuit_ExamAtribute_Student_list = filter.getListByTagId(enter_Way_SpecialRecuit_ExamAtribute_ID_list, k.Value);

                //入學方式:特色招生--甄選入學 ，Student List           
                List<myStudent> enter_Way_SpecialRecuit_Selection_Student_list = new List<myStudent>();
                enter_Way_SpecialRecuit_Selection_Student_list = filter.getListByTagId(enter_Way_SpecialRecuit_Selection_ID_list, k.Value);

                //入學方式:適性輔導安置(十二年安置) ，Student List           
                List<myStudent> enter_Way_SafelySet_Student_list = new List<myStudent>();
                enter_Way_SafelySet_Student_list = filter.getListByTagId(enter_Way_SafelySet_ID_list, k.Value);

                //入學方式:其他 ，Student List           
                List<myStudent> enter_Way_Other_Student_list = new List<myStudent>();
                enter_Way_Other_Student_list = filter.getListByTagId(enter_Way_Other_ID_list, k.Value);

                List<List<myStudent>> All_EnterWays_StudentList = new List<List<myStudent>>();

                All_EnterWays_StudentList.Add(enter_Way_NoExam_SchoolPromote_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_NoExam_SchoolArea_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_NoExam_GoodSkill_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_NoExam_NoExam_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_NoExam_Other_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_SpecialRecuit_ExamAtribute_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_SpecialRecuit_Selection_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_SafelySet_Student_list);
                All_EnterWays_StudentList.Add(enter_Way_Other_Student_list);
                #endregion

                #region 填值， 以入學方式、入學身分 二因素 做 網狀Mapping

                //  其中 入學方式:免試入學--技優甄審  、 其他  因欄位與其他不同 需要做另外處理(可以自行 去看 Resource/ template(105.7ver) 內有許多 合併欄位)
                foreach (List<myStudent> EnterWays_StudentList in All_EnterWays_StudentList)
                {
                    if (EnterWays_StudentList == enter_Way_NoExam_GoodSkill_Student_list)
                    {
                        cs[index, col].PutValue(filter.getListByTagId(enter_identity_normal_ID_list, EnterWays_StudentList).Count);

                        //原住民生、身心障礙生、其他 一起都算在 "外加錄取" 欄位中
                        int extraChoose = 0;

                        extraChoose += filter.getListByTagId(enter_identity_aboriginal_ID_list, EnterWays_StudentList).Count;

                        extraChoose += filter.getListByTagId(enter_identity_IEP_ID_list, EnterWays_StudentList).Count;

                        extraChoose += filter.getListByTagId(enter_identity_Other_list, EnterWays_StudentList).Count;

                        cs[index, col + 2].PutValue(extraChoose);

                        col = col + 4;

                        continue;
                    }
                    else if (EnterWays_StudentList == enter_Way_SafelySet_Student_list)
                    {
                        //只有 一般生、身心障礙生
                        cs[index, col].PutValue(filter.getListByTagId(enter_identity_normal_ID_list, EnterWays_StudentList).Count);
                        cs[index, col + 1].PutValue(filter.getListByTagId(enter_identity_IEP_ID_list, EnterWays_StudentList).Count);

                        col = col + 2;
                        continue;
                    }

                    // 其他入學方式的規則都一樣，  九種入學方式對應 四種入學身分
                    else
                    {
                        foreach (List<string> EnterWaysTagID in EnterIdentityTagsID_Mapping_List)
                        {
                            List<myStudent> EnterIdentityTagsID_Mapping_StudentList_collect__enterWay_Student_list = new List<myStudent>();

                            EnterIdentityTagsID_Mapping_StudentList_collect__enterWay_Student_list = filter.getListByTagId(EnterWaysTagID, EnterWays_StudentList);

                            cs[index, col].PutValue(EnterIdentityTagsID_Mapping_StudentList_collect__enterWay_Student_list.Count);

                            if (col == 18 || col == 20)
                            {
                                col = col + 2;
                            }
                            else
                            {
                                col = col + flexInsex;
                            }
                        }
                    }
                }
                #endregion

                //Table1 Right
                //row = 10;
                //foreach (KeyValuePair<String, List<String>> map in _mappingData) //Form2傳入的Mapping資料
                //{
                //    if (map.Value.Count > 0)
                //    {
                //        List<myStudent> list = new List<myStudent>();
                //        list = filter.getListByTagId(map.Value, k.Value); //list收集符合的TagId學生物件
                //        cs[index, row].PutValue(list.Count); //列出符合的TagId學生物件總數
                //    }
                //    row++; //換欄

                //}

                index++; //每做完一次k.value即換行
            }
            #endregion

            //Table2 Left
            //Dictionary<String, List<String>> table2Left = new Dictionary<string, List<String>>();

            //foreach (KeyValuePair<String, List<String>> map in _mappingData)
            //{
            //    String[] key = map.Key.Split(':');
            //    if (!table2Left.ContainsKey(key[1]))
            //    {
            //        table2Left.Add(key[1], new List<String>());
            //    }
            //    foreach (String s in map.Value)
            //    {
            //        if (map.Key.Split(':')[1] == key[1])
            //        {
            //            table2Left[key[1]].Add(s);
            //        }
            //    }
            //}




            // 入學身份:一般生(非外加錄取) ，Student List           
            List<myStudent> enter_identity_normal_Student_list = new List<myStudent>();
            enter_identity_normal_Student_list = filter.getListByTagId(enter_identity_normal_ID_list, summary);

            // 入學身份:外加錄取--原住民生 ，Student List           
            List<myStudent> enter_identity_aboriginal_Student_list = new List<myStudent>();
            enter_identity_aboriginal_Student_list = filter.getListByTagId(enter_identity_aboriginal_ID_list, summary);

            // 入學身份:外加錄取--身心障礙生 ，Student List           
            List<myStudent> enter_identity_IEP_Student_list = new List<myStudent>();
            enter_identity_IEP_Student_list = filter.getListByTagId(enter_identity_IEP_ID_list, summary);

            // 入學身份:外加錄取--其他 ，Student List           
            List<myStudent> enter_identity_other_Student_list = new List<myStudent>();
            enter_identity_other_Student_list = filter.getListByTagId(enter_identity_Other_list, summary);


            cs[25, 6].PutValue(enter_identity_normal_Student_list.Count); //入學身份:一般生(非外加錄取)總數
            cs[26, 6].PutValue(enter_identity_aboriginal_Student_list.Count); //入學身份:一般生(非外加錄取)總數
            cs[27, 6].PutValue(enter_identity_IEP_Student_list.Count); //入學身份:一般生(非外加錄取)總數
            cs[28, 6].PutValue(enter_identity_other_Student_list.Count); //入學身份:一般生(非外加錄取)總數


            cs[25, 7].PutValue(filter.getGenderCount(enter_identity_normal_Student_list, "1")); //入學身份:一般生(非外加錄取) 男生總數
            cs[25, 8].PutValue(filter.getGenderCount(enter_identity_normal_Student_list, "0")); //入學身份:一般生(非外加錄取) 女生總數

            cs[26, 7].PutValue(filter.getGenderCount(enter_identity_aboriginal_Student_list, "1")); // 入學身份:外加錄取--原住民生 男生總數
            cs[26, 8].PutValue(filter.getGenderCount(enter_identity_aboriginal_Student_list, "0")); // 入學身份:外加錄取--原住民生 女生總數

            cs[27, 7].PutValue(filter.getGenderCount(enter_identity_IEP_Student_list, "1")); //入學身份:外加錄取--身心障礙生 男生總數
            cs[27, 8].PutValue(filter.getGenderCount(enter_identity_IEP_Student_list, "0")); //入學身份:外加錄取--身心障礙生 女生總數

            cs[28, 7].PutValue(filter.getGenderCount(enter_identity_other_Student_list, "1")); //入學身份:外加錄取--其他 男生總數
            cs[28, 8].PutValue(filter.getGenderCount(enter_identity_other_Student_list, "0")); //入學身份:外加錄取--其他 女生總數


            col = 9;

            flexInsex = 1; //因為樣板不是每一項都是佔1格 ，有些有二合一合併，所以靠一個參數彈性調整

            foreach (List<string> EnterWaysTagID in EnterWayTagsID_Mapping_List)
            {

                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__enter_identity_normal_Student_list = new List<myStudent>();

                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__enter_identity_aboriginal_Student_list = new List<myStudent>();

                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__enter_identity_IEP_Student_list = new List<myStudent>();

                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__enter_identity_other_Student_list = new List<myStudent>();




                EnterWaysTagID_Mapping_StudentList_collect__enter_identity_normal_Student_list = filter.getListByTagId(EnterWaysTagID, enter_identity_normal_Student_list);

                EnterWaysTagID_Mapping_StudentList_collect__enter_identity_aboriginal_Student_list = filter.getListByTagId(EnterWaysTagID, enter_identity_aboriginal_Student_list);

                EnterWaysTagID_Mapping_StudentList_collect__enter_identity_IEP_Student_list = filter.getListByTagId(EnterWaysTagID, enter_identity_IEP_Student_list);

                EnterWaysTagID_Mapping_StudentList_collect__enter_identity_other_Student_list = filter.getListByTagId(EnterWaysTagID, enter_identity_other_Student_list);



                cs[25, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_normal_Student_list, "1")); //入學身份:一般生(非外加錄取) 男生總數 in EnterWayTagsID_Mapping_List
                cs[25, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_normal_Student_list, "0")); //入學身份:一般生(非外加錄取) 女生總數 in EnterWayTagsID_Mapping_List

                cs[26, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_aboriginal_Student_list, "1")); //入學身份:外加錄取--原住民生 男生總數 in EnterWayTagsID_Mapping_List
                cs[26, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_aboriginal_Student_list, "0")); //入學身份:外加錄取--原住民生 女生總數 in EnterWayTagsID_Mapping_List

                cs[27, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_IEP_Student_list, "1")); //入學身份:外加錄取--身心障礙生 男生總數 in EnterWayTagsID_Mapping_List
                cs[27, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_IEP_Student_list, "0")); //入學身份:外加錄取--身心障礙生 女生總數 in EnterWayTagsID_Mapping_List

                cs[28, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_other_Student_list, "1")); //入學身份:外加錄取--其他 男生總數 in EnterWayTagsID_Mapping_List
                cs[28, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__enter_identity_other_Student_list, "0")); //入學身份:外加錄取--其他 女生總數 in EnterWayTagsID_Mapping_List



                if (col == 9)
                {
                    col = 12;
                    flexInsex = 2;
                }
                else
                {
                    if (col == 36)
                    {
                        col = col + 3;
                        flexInsex = 1;
                    }
                    else
                    {
                        col = col + 4;
                    }
                }
            }




            ////收集原住民生
            //foreach (KeyValuePair<String, List<String>> k in table2Left)
            //{
            //    if (k.Key == "原住民生")
            //    {
            //        AboList = k.Value; //收入TagID
            //    }
            //}

            //index = 32;
            //foreach (KeyValuePair<String, List<String>> k in table2Left)
            //{
            //    List<myStudent> list = new List<myStudent>();
            //    list = filter.getListByTagId(k.Value, summary);

            //    cs[index, 4].PutValue(list.Count);
            //    cs[index, 6].PutValue(filter.getGenderCount(list, "1"));
            //    cs[index, 8].PutValue(filter.getGenderCount(list, "0"));
            //    index++;
            //}


            ////Table2 Right
            //index = 32;
            //row = 10;
            //foreach (KeyValuePair<String, List<String>> map in _mappingData)
            //{
            //    if (index > 35) { index = 32; row += 2; } //換行換欄
            //    if (map.Value.Count > 0)
            //    {
            //        List<myStudent> list = new List<myStudent>();
            //        list = filter.getListByTagId(map.Value, summary);

            //        cs[index, row].PutValue(filter.getGenderCount(list, "1"));
            //        cs[index, row + 1].PutValue(filter.getGenderCount(list, "0"));
            //    }
            //    index++;


            //}

            #region 按國中畢/修業年度分
            //Table3 Left

            List<myStudent> collect__LastGrade = new List<myStudent>();  //應屆畢業的收集清單
            List<myStudent> collect__LastComplete = new List<myStudent>();  //應屆結業的收集清單
            List<myStudent> collect__LastOther = new List<myStudent>();  //應屆其他的收集清單

            List<myStudent> collect__abo_LastGrade = new List<myStudent>();  //應屆畢業的收集清單
            List<myStudent> collect__abo_LastComplete = new List<myStudent>();  //應屆結業的收集清單
            List<myStudent> collect__abo_LastOther = new List<myStudent>();  //應屆其他的收集清單

            List<myStudent> collect__LastGradeT = new List<myStudent>();  //應屆的收集清單
            List<myStudent> collect__LastGradeF = new List<myStudent>();  //非應屆的收集清單


            List<String> collect_List = new List<string>(); //收集學生ID的清單

            foreach (myStudent student in summary) //收集summary所有學生ID
            {
                collect_List.Add(student.Id);
            }

            // 所有學生的異動資料
            List<K12.Data.UpdateRecordRecord> UpdateRecord_records = K12.Data.UpdateRecord.SelectByStudentIDs(collect_List);

            //傳入學生ID清單供查詢
            List<SHSchool.Data.SHBeforeEnrollmentRecord> recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
            {
                foreach (myStudent student in summary)
                {
                    if (rec.RefStudentID == student.Id) //找到對應ID後,判斷前級畢業年度
                    {
                        String last_grade_year = rec.GraduateSchoolYear;
                        if (last_grade_year == "") last_grade_year = "0"; //空值填方便後續計算

                        //int year = Convert.ToInt16(last_grade_year) + 1912; //學年度+1912若等於現在年份則判斷為應屆生

                        int year = Convert.ToInt16(last_grade_year);


                        if ((year + 1).ToString() == K12.Data.School.DefaultSchoolYear)
                        {
                            if (CheckStudentBeforeStatus(UpdateRecord_records, student.Id) == "當年畢業")
                            {
                                collect__LastGrade.Add(student);
                            }
                            if (CheckStudentBeforeStatus(UpdateRecord_records, student.Id) == "當年修業")
                            {
                                collect__LastComplete.Add(student);
                            }
                            if (CheckStudentBeforeStatus(UpdateRecord_records, student.Id) == "其他(含領結業證書)")
                            {
                                collect__LastOther.Add(student);
                            }

                            //collect__LastGradeT.Add(student); //收入應屆清單
                        }
                        else
                        {
                            collect__LastOther.Add(student);
                            //collect__LastGradeF.Add(student); //收入非應屆清單
                        }
                    }
                }

                //整理新生中具原住民身分者
                List<myStudent> aboStudentlist = new List<myStudent>();
                aboStudentlist = filter.getListByTagId(aboIDList, summary);


                foreach (myStudent Ms in aboStudentlist)
                {
                    if (rec.RefStudentID == Ms.Id) //找到對應ID後,判斷前級畢業年度
                    {
                        String last_grade_year = rec.GraduateSchoolYear;
                        if (last_grade_year == "") last_grade_year = "0"; //空值填方便後續計算

                        //int year = Convert.ToInt16(last_grade_year) + 1912; //學年度+1912若等於現在年份則判斷為應屆生

                        int year = Convert.ToInt16(last_grade_year);


                        if ((year + 1).ToString() == K12.Data.School.DefaultSchoolYear)
                        {
                            if (CheckStudentBeforeStatus(UpdateRecord_records, Ms.Id) == "當年畢業")
                            {
                                collect__abo_LastGrade.Add(Ms);
                            }
                            if (CheckStudentBeforeStatus(UpdateRecord_records, Ms.Id) == "當年修業")
                            {
                                collect__abo_LastComplete.Add(Ms);
                            }
                            if (CheckStudentBeforeStatus(UpdateRecord_records, Ms.Id) == "其他(含領結業證書)")
                            {
                                collect__abo_LastOther.Add(Ms);
                            }


                        }
                        else
                        {
                            collect__abo_LastOther.Add(Ms);

                        }
                    }
                }


            }


            cs[29, 6].PutValue(collect__LastGrade.Count); //應屆畢業總數
            cs[30, 6].PutValue(collect__LastComplete.Count); //應屆修業總數
            cs[31, 6].PutValue(collect__LastOther.Count); //其他種入學

            cs[29, 7].PutValue(filter.getGenderCount(collect__LastGrade, "1")); //應屆畢業男生總數
            cs[29, 8].PutValue(filter.getGenderCount(collect__LastGrade, "0")); //應屆畢業女生總數

            cs[30, 7].PutValue(filter.getGenderCount(collect__LastComplete, "1")); //應屆結業男生總數
            cs[30, 8].PutValue(filter.getGenderCount(collect__LastComplete, "0")); //應屆結業女生總數

            cs[31, 7].PutValue(filter.getGenderCount(collect__LastOther, "1")); //其他種入學男生總數
            cs[31, 8].PutValue(filter.getGenderCount(collect__LastOther, "0")); //其他種入學女生總數



            col = 9;

            flexInsex = 1; //因為樣板不是每一項都是佔1格 ，有些有二合一合併，所以靠一個參數彈性調整

            foreach (List<string> EnterWaysTagID in EnterWayTagsID_Mapping_List)
            {
                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__LastGrade = new List<myStudent>();

                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__LastComplete = new List<myStudent>();

                List<myStudent> EnterWaysTagID_Mapping_StudentList_collect__collect__LastOther = new List<myStudent>();

                EnterWaysTagID_Mapping_StudentList_collect__LastGrade = filter.getListByTagId(EnterWaysTagID, collect__LastGrade);

                EnterWaysTagID_Mapping_StudentList_collect__LastComplete = filter.getListByTagId(EnterWaysTagID, collect__LastComplete);

                EnterWaysTagID_Mapping_StudentList_collect__collect__LastOther = filter.getListByTagId(EnterWaysTagID, collect__LastOther);

                cs[29, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__LastGrade, "1")); //應屆畢業男生總數 in EnterWayTagsID_Mapping_List
                cs[29, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__LastGrade, "0")); //應屆畢業女生總數 in EnterWayTagsID_Mapping_List

                cs[30, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__LastComplete, "1")); //應屆結業男生總數 in EnterWayTagsID_Mapping_List
                cs[30, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__LastComplete, "0")); //應屆結業女生總數 in EnterWayTagsID_Mapping_List

                cs[31, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__collect__LastOther, "1")); //其他種入學男生總數 in EnterWayTagsID_Mapping_List
                cs[31, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_collect__collect__LastOther, "0")); //其他種入學女生總數 in EnterWayTagsID_Mapping_List

                if (col == 9)
                {
                    col = 12;
                    flexInsex = 2;
                }
                else
                {
                    if (col == 36)
                    {
                        col = col + 3;
                        flexInsex = 1;
                    }
                    else
                    {
                        col = col + 4;
                    }
                }
            }

            cs[38, 6].PutValue(collect__abo_LastGrade.Count); //新生具有原住民身分者應屆畢業總數
            cs[39, 6].PutValue(collect__abo_LastComplete.Count); //新生具有原住民身分者應屆修業總數
            cs[40, 6].PutValue(collect__abo_LastOther.Count); //新生具有原住民身分者其他種入學

            cs[38, 7].PutValue(filter.getGenderCount(collect__abo_LastGrade, "1")); //新生具有原住民身分者應屆畢業男生總數
            cs[38, 8].PutValue(filter.getGenderCount(collect__abo_LastGrade, "0")); //新生具有原住民身分者應屆畢業女生總數

            cs[39, 7].PutValue(filter.getGenderCount(collect__abo_LastComplete, "1")); //新生具有原住民身分者應屆結業男生總數
            cs[39, 8].PutValue(filter.getGenderCount(collect__abo_LastComplete, "0")); //新生具有原住民身分者應屆結業女生總數

            cs[40, 7].PutValue(filter.getGenderCount(collect__abo_LastOther, "1")); //新生具有原住民身分者其他種入學男生總數
            cs[40, 8].PutValue(filter.getGenderCount(collect__abo_LastOther, "0")); //新生具有原住民身分者其他種入學女生總數



            //cs[36, 4].PutValue(collect__LastGradeT.Count); //應屆畢業總數
            //cs[37, 4].PutValue(collect__LastGradeF.Count); //非應屆畢業總數
            //cs[36, 6].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆畢業男生總數
            //cs[36, 8].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆畢業女生總數
            //cs[37, 6].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆畢業男生總數
            //cs[37, 8].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆畢業女生總數 
            #endregion



            ////Table3 Right
            //Dictionary<String, List<String>> ndic = new Dictionary<string, List<string>>(); //為綜合入學方式,建立字典
            //foreach (KeyValuePair<String, List<String>> map in _mappingData)
            //{
            //    String key = map.Key.Substring(0, 2); //建立key為前面兩個字串:甄選,申請,登記,直升,免試,其他
            //    if (!ndic.ContainsKey(key))
            //    {
            //        ndic.Add(key, new List<string>()); //key不存在即建立
            //    }
            //    foreach (String s in map.Value)
            //    {
            //        if (map.Key.Contains(key)) //針對符合的key做TagID的收集
            //        {
            //            ndic[key].Add(s);
            //        }
            //    }
            //}

            //index = 36;
            //row = 10;
            //foreach (KeyValuePair<String, List<String>> nmap in ndic)
            //{
            //    if (index > 36) { index = 36; row += 2; } //換行換欄
            //    if (nmap.Value.Count == 0)  //遇到空值index++並繼續迴圈
            //    {
            //        index++;
            //        continue;
            //    }
            //    List<myStudent> list = new List<myStudent>();
            //    list = filter.getListByTagId(nmap.Value, summary); //收集符合TagID的學生物件
            //    collect_List = new List<string>(); //清空之前的清單
            //    collect__LastGradeT = new List<myStudent>(); //清空之前的清單
            //    collect__LastGradeF = new List<myStudent>(); //清空之前的清單
            //    foreach (myStudent student in list)
            //    {
            //        collect_List.Add(student.Id); //收集學生ID
            //    }

            //    recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            //    foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
            //    {
            //        foreach (myStudent student in list)
            //        {
            //            if (rec.RefStudentID == student.Id)
            //            {
            //                String last_grade_year = rec.GraduateSchoolYear;
            //                if (last_grade_year == "") last_grade_year = "0";
            //                int year = Convert.ToInt16(last_grade_year) + 1912;
            //                if (year.ToString() == DateTime.Now.Year.ToString())
            //                {
            //                    collect__LastGradeT.Add(student); //收入應屆清單
            //                }
            //                else
            //                {
            //                    collect__LastGradeF.Add(student); //收入非應屆清單
            //                }
            //            }
            //        }
            //    }
            //    cs[index, row].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆男生數
            //    cs[index, row + 1].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆女生數
            //    cs[index + 1, row].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆男生數
            //    cs[index + 1, row + 1].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆女生數
            //    index++; //換行
            //}

            ////Table3 End
            //collect_List = new List<string>(); //清空之前的清單
            //collect__LastGradeT = new List<myStudent>(); //清空之前的清單
            //collect__LastGradeF = new List<myStudent>(); //清空之前的清單
            //List<myStudent> AboStudent = filter.getListByTagId(AboList, summary);
            //foreach (myStudent student in AboStudent)
            //{
            //    collect_List.Add(student.Id);
            //}
            //recl = SHSchool.Data.SHBeforeEnrollment.SelectByStudentIDs(collect_List);
            //foreach (SHSchool.Data.SHBeforeEnrollmentRecord rec in recl)
            //{
            //    foreach (myStudent student in AboStudent)
            //    {
            //        if (rec.RefStudentID == student.Id)
            //        {
            //            String last_grade_year = rec.GraduateSchoolYear;
            //            if (last_grade_year == "") last_grade_year = "0";
            //            int year = Convert.ToInt16(last_grade_year) + 1912;
            //            if (year.ToString() == DateTime.Now.Year.ToString())
            //            {
            //                collect__LastGradeT.Add(student); //收入應屆清單
            //            }
            //            else
            //            {
            //                collect__LastGradeF.Add(student); //收入非應屆清單
            //            }
            //        }
            //    }
            //}

            //cs[36, 22].PutValue(filter.getGenderCount(collect__LastGradeT, "1")); //應屆原住民男生數
            //cs[36, 23].PutValue(filter.getGenderCount(collect__LastGradeT, "0")); //應屆原住民女生數
            //cs[37, 22].PutValue(filter.getGenderCount(collect__LastGradeF, "1")); //非應屆原住民男生數
            //cs[37, 23].PutValue(filter.getGenderCount(collect__LastGradeF, "0")); //非應屆原住民女生數

            List<myStudent> collect__LocalCounty = new List<myStudent>();  //戶籍位於本縣市
            List<myStudent> collect__OtherCounty = new List<myStudent>();  //戶籍非位於本縣市

            #region 按戶籍地分
            XmlElement Element = Config.GetSchoolInfo();

            string LocalCounty = getNodeData("County", Element, "SchoolInformation");

            foreach (myStudent student in summary)
            {
                if (student.County == LocalCounty)
                {
                    collect__LocalCounty.Add(student);
                }
                else
                {
                    collect__OtherCounty.Add(student);
                }
            }

            cs[32, 6].PutValue(collect__LocalCounty.Count); //戶籍位於本縣市
            cs[33, 6].PutValue(collect__OtherCounty.Count); //戶籍非位於本縣市

            cs[32, 7].PutValue(filter.getGenderCount(collect__LocalCounty, "1")); //戶籍位於本縣市男生總數
            cs[32, 8].PutValue(filter.getGenderCount(collect__LocalCounty, "0")); //戶籍位於本縣市女生總數

            cs[33, 7].PutValue(filter.getGenderCount(collect__OtherCounty, "1")); //戶籍非位於本縣市男生總數
            cs[33, 8].PutValue(filter.getGenderCount(collect__OtherCounty, "0")); //戶籍非位於本縣市女生總數 

            col = 9;

            flexInsex = 1; //因為樣板不是每一項都是佔1格 ，有些有二合一合併，所以靠一個參數彈性調整

            foreach (List<string> EnterWaysTagID in EnterWayTagsID_Mapping_List)
            {
                List<myStudent> EnterWaysTagID_Mapping_StudentList_LocalCounty = new List<myStudent>();

                List<myStudent> EnterWaysTagID_Mapping_StudentList_OtherCounty = new List<myStudent>();

                EnterWaysTagID_Mapping_StudentList_LocalCounty = filter.getListByTagId(EnterWaysTagID, collect__LocalCounty);

                EnterWaysTagID_Mapping_StudentList_OtherCounty = filter.getListByTagId(EnterWaysTagID, collect__OtherCounty);

                cs[32, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_LocalCounty, "1")); //戶籍位於本縣市男生總數 in EnterWayTagsID_Mapping_List
                cs[32, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_LocalCounty, "0")); //戶籍位於本縣市女生總數 in EnterWayTagsID_Mapping_List

                cs[33, col].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_OtherCounty, "1")); //戶籍非位於本縣市男生總數 in EnterWayTagsID_Mapping_List
                cs[33, col + flexInsex].PutValue(filter.getGenderCount(EnterWaysTagID_Mapping_StudentList_OtherCounty, "0")); //戶籍非位於本縣市女生總數 in EnterWayTagsID_Mapping_List

                if (col == 9)
                {
                    col = 12;
                    flexInsex = 2;
                }
                else
                {
                    if (col == 36)
                    {
                        col = col + 3;
                        flexInsex = 1;
                    }
                    else
                    {
                        col = col + 4;
                    }
                }
            }
            #endregion


            #region 按畢業國中所在地分

            #region 整理所在地 Dict
            Dictionary<string, List<myStudent>> Collect_BeforeSchoolLocationList = new Dictionary<string, List<myStudent>>();

            Collect_BeforeSchoolLocationList.Add("總計", new List<myStudent>());

            Collect_BeforeSchoolLocationList.Add("新北市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("臺北市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("臺中市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("臺南市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("高雄市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("宜蘭縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("桃園市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("新竹縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("苗栗縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("彰化縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("南投縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("雲林縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("嘉義縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("屏東縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("臺東縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("花蓮縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("澎湖縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("基隆市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("新竹市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("嘉義市", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("金門縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("連江縣", new List<myStudent>());
            Collect_BeforeSchoolLocationList.Add("其他", new List<myStudent>());

            foreach (myStudent student in summary)
            {
                if (student.Before_School_Location == "新北市")
                {
                    Collect_BeforeSchoolLocationList["新北市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "臺北市")
                {
                    Collect_BeforeSchoolLocationList["臺北市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "臺中市")
                {
                    Collect_BeforeSchoolLocationList["臺中市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "臺南市")
                {
                    Collect_BeforeSchoolLocationList["臺南市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "高雄市")
                {
                    Collect_BeforeSchoolLocationList["高雄市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "宜蘭縣")
                {
                    Collect_BeforeSchoolLocationList["宜蘭縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "桃園市")
                {
                    Collect_BeforeSchoolLocationList["桃園市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "新竹縣")
                {
                    Collect_BeforeSchoolLocationList["新竹縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "苗栗縣")
                {
                    Collect_BeforeSchoolLocationList["苗栗縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "彰化縣")
                {
                    Collect_BeforeSchoolLocationList["彰化縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "南投縣")
                {
                    Collect_BeforeSchoolLocationList["南投縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "雲林縣")
                {
                    Collect_BeforeSchoolLocationList["雲林縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "嘉義縣")
                {
                    Collect_BeforeSchoolLocationList["嘉義縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "屏東縣")
                {
                    Collect_BeforeSchoolLocationList["屏東縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "臺東縣")
                {
                    Collect_BeforeSchoolLocationList["臺東縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "花蓮縣")
                {
                    Collect_BeforeSchoolLocationList["花蓮縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "澎湖縣")
                {
                    Collect_BeforeSchoolLocationList["澎湖縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "基隆市")
                {
                    Collect_BeforeSchoolLocationList["基隆市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "新竹市")
                {
                    Collect_BeforeSchoolLocationList["新竹市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "嘉義市")
                {
                    Collect_BeforeSchoolLocationList["嘉義市"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "金門縣")
                {
                    Collect_BeforeSchoolLocationList["金門縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else if (student.Before_School_Location == "連江縣")
                {
                    Collect_BeforeSchoolLocationList["連江縣"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
                else
                {
                    Collect_BeforeSchoolLocationList["其他"].Add(student);
                    Collect_BeforeSchoolLocationList["總計"].Add(student);
                }
            }
            #endregion

            #region 填值

            cs[35, 6].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["總計"], "1")); // 總計 男
            cs[36, 6].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["總計"], "0")); // 總計 女

            cs[35, 7].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["新北市"], "1")); // 新北市 男
            cs[36, 7].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["新北市"], "0")); // 新北市 女

            cs[35, 8].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺北市"], "1")); // 臺北市 男
            cs[36, 8].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺北市"], "0")); // 臺北市 女

            cs[35, 9].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺中市"], "1")); // 臺中市 男
            cs[36, 9].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺中市"], "0")); // 臺中市 女

            cs[35, 10].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺南市"], "1")); // 臺南市 男
            cs[36, 10].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺南市"], "0")); // 臺南市 女

            cs[35, 12].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["高雄市"], "1")); // 高雄市 男
            cs[36, 12].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["高雄市"], "0")); // 高雄市 女

            cs[35, 14].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["宜蘭縣"], "1")); // 宜蘭縣 男
            cs[36, 14].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["宜蘭縣"], "0")); // 宜蘭縣 女

            cs[35, 16].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["桃園市"], "1")); // 桃園市 男
            cs[36, 16].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["桃園市"], "0")); // 桃園市 女

            cs[35, 18].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["新竹縣"], "1")); // 新竹縣 男
            cs[36, 18].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["新竹縣"], "0")); // 新竹縣 女

            cs[35, 20].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["苗栗縣"], "1")); // 苗栗縣 男
            cs[36, 20].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["苗栗縣"], "0")); // 苗栗縣 女

            cs[35, 22].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["彰化縣"], "1")); // 彰化縣 男
            cs[36, 22].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["彰化縣"], "0")); // 彰化縣 女

            cs[35, 24].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["南投縣"], "1")); // 南投縣 男
            cs[36, 24].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["南投縣"], "0")); // 南投縣 女

            cs[35, 26].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["雲林縣"], "1")); // 雲林縣 男
            cs[36, 26].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["雲林縣"], "0")); // 雲林縣 女

            cs[35, 28].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["嘉義縣"], "1")); // 嘉義縣 男
            cs[36, 28].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["嘉義縣"], "0")); // 嘉義縣 女

            cs[35, 30].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["屏東縣"], "1")); // 屏東縣 男
            cs[36, 30].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["屏東縣"], "0")); // 屏東縣 女

            cs[35, 32].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺東縣"], "1")); // 臺東縣 男
            cs[36, 32].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["臺東縣"], "0")); // 臺東縣 女

            cs[35, 34].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["花蓮縣"], "1")); // 花蓮縣 男
            cs[36, 34].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["花蓮縣"], "0")); // 花蓮縣 女

            cs[35, 36].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["澎湖縣"], "1")); // 澎湖縣 男
            cs[36, 36].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["澎湖縣"], "0")); // 澎湖縣 女

            cs[35, 38].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["基隆市"], "1")); // 基隆市 男
            cs[36, 38].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["基隆市"], "0")); // 基隆市 女

            cs[35, 39].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["新竹市"], "1")); // 新竹市 男
            cs[36, 39].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["新竹市"], "0")); // 新竹市 女

            cs[35, 40].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["嘉義市"], "1")); // 嘉義市 男
            cs[36, 40].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["嘉義市"], "0")); // 嘉義市 女

            cs[35, 42].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["金門縣"], "1")); // 金門縣 男
            cs[36, 42].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["金門縣"], "0")); // 金門縣 女

            cs[35, 44].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["連江縣"], "1")); // 連江縣 男
            cs[36, 44].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["連江縣"], "0")); // 連江縣 女

            cs[35, 46].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["其他"], "1")); // 其他 男
            cs[36, 46].PutValue(filter.getGenderCount(Collect_BeforeSchoolLocationList["其他"], "0")); // 其他 女  
            #endregion

            #endregion

            cs["U5"].PutValue(_SchoolYear);
        }

        public void SaveMappingRecord() //儲存上次Mapping紀錄
        {
            AccessHelper _A = new AccessHelper();
            List<myTableUDT> UDTlist = _A.Select<myTableUDT>();
            _A.DeletedValues(UDTlist); //清除UDT資料

            UDTlist = new List<myTableUDT>(); //清空UDTlist
            foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
            {
                if (row.Cells[0].Value == null) //遇到空白的Target即跳到下個loop
                {
                    continue;
                }

                String target = row.Cells[0].Value.ToString();
                String source = "";
                if (row.Cells[1].Value != null) { source = row.Cells[1].Value.ToString(); }

                myTableUDT obj = new myTableUDT();
                obj.Target = target;
                obj.Source = source;
                UDTlist.Add(obj);
            }

            _A.InsertValues(UDTlist); //回存到UDT
        }


        // 儲存 Mapping 的Xml 資料
        public void SaveMappingXmlRecord()
        {
            ConfigData cd = K12.Data.School.Configuration["新生入學統計報表_來源目標設定Config"];

            XmlElement config = cd.GetXml("XmlData", null);

            if (config != null) //如果不是空的
            {
                #region 表1--入學方式 Xml 紀錄 儲存
                XmlElement EnterSchool_Way = (XmlElement)config.SelectSingleNode("入學方式");

                if (EnterSchool_Way != null) //  Config內有設定才做讀取
                {
                    // 移除舊資料    
                    EnterSchool_Way.RemoveAll();

                    int i = 1;

                    foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
                    {

                        XmlElement EnterSchool_Way_item = EnterSchool_Way.OwnerDocument.CreateElement("item");

                        if (row.Cells[0].Value == null || "" + row.Cells[0].Value == "") //遇到空白的Target即跳到下個loop
                        {
                            continue;
                        }

                        if (row.Cells[1].Value == null || "" + row.Cells[1].Value == "") //遇到空白的Source即跳到下個loop
                        {
                            continue;
                        }

                        String target = row.Cells[0].Value.ToString();
                        String source = "";
                        if (row.Cells[1].Value != null)
                        {
                            source = row.Cells[1].Value.ToString();
                        }

                        EnterSchool_Way_item.SetAttribute("ID", "" + i);
                        EnterSchool_Way_item.SetAttribute("target", target);
                        EnterSchool_Way_item.SetAttribute("source", source);

                        EnterSchool_Way.AppendChild(EnterSchool_Way_item);

                        i++;

                    }
                }
                else
                {
                    EnterSchool_Way = config.OwnerDocument.CreateElement("入學方式");

                    int i = 1;

                    foreach (DataGridViewRow row in dataGridViewX1.Rows) //取得DataDataGridViewRow資料
                    {

                        XmlElement EnterSchool_Way_item = EnterSchool_Way.OwnerDocument.CreateElement("item");

                        if (row.Cells[0].Value == null || "" + row.Cells[0].Value == "") //遇到空白的Target即跳到下個loop
                        {
                            continue;
                        }

                        if (row.Cells[1].Value == null || "" + row.Cells[1].Value == "") //遇到空白的Source即跳到下個loop
                        {
                            continue;
                        }

                        String target = row.Cells[0].Value.ToString();
                        String source = "";
                        if (row.Cells[1].Value != null)
                        {
                            source = row.Cells[1].Value.ToString();
                        }

                        EnterSchool_Way_item.SetAttribute("ID", "" + i);
                        EnterSchool_Way_item.SetAttribute("target", target);
                        EnterSchool_Way_item.SetAttribute("source", source);

                        EnterSchool_Way.AppendChild(EnterSchool_Way_item);

                        i++;

                    }

                }

                config.AppendChild(EnterSchool_Way);
                #endregion

                #region 表2--入學身份 Xml 紀錄 儲存
                XmlElement EnterSchool_identity = (XmlElement)config.SelectSingleNode("入學身分");

                if (EnterSchool_identity != null) //  Config內有設定才做讀取
                {
                    // 移除舊資料    
                    EnterSchool_identity.RemoveAll();

                    int i = 1;

                    foreach (DataGridViewRow row in dataGridViewX2.Rows) //取得DataDataGridViewRow資料
                    {

                        XmlElement EnterSchool_identity_item = EnterSchool_identity.OwnerDocument.CreateElement("item");

                        if (row.Cells[0].Value == null || "" + row.Cells[0].Value == "") //遇到空白的Target即跳到下個loop
                        {
                            continue;
                        }

                        if (row.Cells[1].Value == null || "" + row.Cells[1].Value == "") //遇到空白的Source即跳到下個loop
                        {
                            continue;
                        }

                        String target = row.Cells[0].Value.ToString();
                        String source = "";
                        if (row.Cells[1].Value != null)
                        {
                            source = row.Cells[1].Value.ToString();
                        }

                        EnterSchool_identity_item.SetAttribute("ID", "" + i);
                        EnterSchool_identity_item.SetAttribute("target", target);
                        EnterSchool_identity_item.SetAttribute("source", source);

                        EnterSchool_identity.AppendChild(EnterSchool_identity_item);

                        i++;

                    }
                }
                else
                {
                    EnterSchool_identity = config.OwnerDocument.CreateElement("入學身分");

                    int i = 1;

                    foreach (DataGridViewRow row in dataGridViewX2.Rows) //取得DataDataGridViewRow資料
                    {

                        XmlElement EnterSchool_identity_item = EnterSchool_identity.OwnerDocument.CreateElement("item");

                        if (row.Cells[0].Value == null || "" + row.Cells[0].Value == "") //遇到空白的Target即跳到下個loop
                        {
                            continue;
                        }

                        if (row.Cells[1].Value == null || "" + row.Cells[1].Value == "") //遇到空白的Source即跳到下個loop
                        {
                            continue;
                        }

                        String target = row.Cells[0].Value.ToString();
                        String source = "";
                        if (row.Cells[1].Value != null)
                        {
                            source = row.Cells[1].Value.ToString();
                        }

                        EnterSchool_identity_item.SetAttribute("ID", "" + i);
                        EnterSchool_identity_item.SetAttribute("target", target);
                        EnterSchool_identity_item.SetAttribute("source", source);

                        EnterSchool_identity.AppendChild(EnterSchool_identity_item);

                        i++;

                    }

                }

                config.AppendChild(EnterSchool_identity);
                #endregion

                #region 新生中具原住民身分者 Xml 紀錄儲存
                XmlElement FreshMenWith_Aboriginal_Identity = (XmlElement)config.SelectSingleNode("新生中具原住民身分者");

                if (FreshMenWith_Aboriginal_Identity != null) //  Config內有設定才做讀取
                {
                    // 移除舊資料 
                    FreshMenWith_Aboriginal_Identity.RemoveAll();

                    int i = 1;

                    foreach (DataGridViewRow row in dataGridViewX3.Rows) //取得DataDataGridViewRow資料
                    {

                        XmlElement FreshMenWith_Aboriginal_Identity_item = FreshMenWith_Aboriginal_Identity.OwnerDocument.CreateElement("item");

                        if (row.Cells[0].Value == null || "" + row.Cells[0].Value == "") //遇到空白的Target即跳到下個loop
                        {
                            continue;
                        }

                        if (row.Cells[1].Value == null || "" + row.Cells[1].Value == "") //遇到空白的Source即跳到下個loop
                        {
                            continue;
                        }

                        String target = row.Cells[0].Value.ToString();
                        String source = "";
                        if (row.Cells[1].Value != null)
                        {
                            source = row.Cells[1].Value.ToString();
                        }

                        FreshMenWith_Aboriginal_Identity_item.SetAttribute("ID", "" + i);
                        FreshMenWith_Aboriginal_Identity_item.SetAttribute("target", target);
                        FreshMenWith_Aboriginal_Identity_item.SetAttribute("source", source);

                        FreshMenWith_Aboriginal_Identity.AppendChild(FreshMenWith_Aboriginal_Identity_item);

                        i++;

                    }

                }
                else
                {
                    FreshMenWith_Aboriginal_Identity = config.OwnerDocument.CreateElement("新生中具原住民身分者");

                    int i = 1;

                    foreach (DataGridViewRow row in dataGridViewX3.Rows) //取得DataDataGridViewRow資料
                    {

                        XmlElement FreshMenWith_Aboriginal_Identity_item = FreshMenWith_Aboriginal_Identity.OwnerDocument.CreateElement("item");

                        if (row.Cells[0].Value == null || "" + row.Cells[0].Value == "") //遇到空白的Target即跳到下個loop
                        {
                            continue;
                        }

                        if (row.Cells[1].Value == null || "" + row.Cells[1].Value == "") //遇到空白的Source即跳到下個loop
                        {
                            continue;
                        }

                        String target = row.Cells[0].Value.ToString();
                        String source = "";
                        if (row.Cells[1].Value != null)
                        {
                            source = row.Cells[1].Value.ToString();
                        }

                        FreshMenWith_Aboriginal_Identity_item.SetAttribute("ID", "" + i);
                        FreshMenWith_Aboriginal_Identity_item.SetAttribute("target", target);
                        FreshMenWith_Aboriginal_Identity_item.SetAttribute("source", source);

                        FreshMenWith_Aboriginal_Identity.AppendChild(FreshMenWith_Aboriginal_Identity_item);

                        i++;

                    }
                }

                config.AppendChild(FreshMenWith_Aboriginal_Identity);
                #endregion

            }
            else
            {

            }
            cd.SetXml("XmlData", config);
            cd.Save();
        }


        public void LoadLastRecord() //讀取上次Mapping設定
        {
            AccessHelper _A = new AccessHelper();
            List<myTableUDT> UDTlist = _A.Select<myTableUDT>(); //檢查UDT並回傳資料
            DataGridViewRow row;
            if (UDTlist.Count > 0) //UDT內有設定才做讀取
            {
                for (int i = 0; i < UDTlist.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = UDTlist[i].Target;
                    row.Cells[1].Value = UDTlist[i].Source;
                    dataGridViewX1.Rows.Add(row);
                }
            }
            else
            {
                //UDT無資料則提供預設標記
                for (int i = 0; i < Column2.Items.Count; i++)
                {
                    row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1);
                    row.Cells[0].Value = Column2.Items[i];
                    dataGridViewX1.Rows.Add(row);
                }
            }

        }

        //初始化MappingData資料
        public void SetMappingDataKey()
        {
            _mappingData = new Dictionary<string, List<string>>();
            foreach (String s in Column2.Items)
            {
                _mappingData.Add(s, new List<string>());
            }
        }

        //初始化XMLMappingData資料
        public void SetXMLMappingDataKey()
        {
            XML_mappingData = new Dictionary<string, List<string>>();

            // 加入 2.入學身分 的TargetList
            foreach (String s in dataGridViewComboBoxExColumn1.Items)
            {
                XML_mappingData.Add(s, new List<string>());
            }

            // 加入 3.新生中具原住民身分者 的TargetList
            foreach (String s in dataGridViewComboBoxExColumn3.Items)
            {
                XML_mappingData.Add(s, new List<string>());
            }
        }

        private void SchoolYearItem() //建立comboBoxEx2的下拉清單
        {
            int school_year = Convert.ToInt16(K12.Data.School.DefaultSchoolYear);
            for (int i = -3; i < 4; i++)
            {
                comboBoxEx2.Items.Add(school_year + i);
            }
            comboBoxEx2.Text = K12.Data.School.DefaultSchoolYear;
        }

        private bool tryConvert(String str) //避免comboBoxEx2被輸入非數字字串
        {
            try
            {
                Convert.ToInt16(str);
                return true;
            }
            catch
            {
                MessageBox.Show("學年度請確認輸入正確數值");
                return false;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                AccessHelper _A = new AccessHelper();
                List<myTableUDT> UDTlist = _A.Select<myTableUDT>();
                _A.DeletedValues(UDTlist); //清除UDT資料
                dataGridViewX1.Rows.Clear();  //清除datagridview資料

                //LoadLastRecord(); //再次讀入Mapping設定
                LoadConfigXml();
            }
            catch
            {
                MessageBox.Show("網路或資料庫異常,請稍後再試...");
            }

        }

        //確認學生為拿 一般畢業證書、修業證書、還是其他
        public string CheckStudentBeforeStatus(List<K12.Data.UpdateRecordRecord> records, String id)
        {

            foreach (K12.Data.UpdateRecordRecord record in records)
            {
                if (record.StudentID == id)  //找到符合的ID開始後續比對
                {
                    if (record.SchoolYear.ToString() == _SchoolYear) //確認學年度為當前學年度
                    {
                        //2017/1/19 穎驊筆記，經過詢問恩正後 各個種類的異動代碼 有屬於自己的區間， 
                        //以新生異動而言，是001~ 100 之間 ，又內容順序幾乎不會變所以他才敢以 Code <100 作為判別是 合法新生的依據
                        //為了避免後人產生太多疑惑，我詳列各個異動代碼的全意
                        // 001 持國民中學畢業證書者(含國中補校)
                        // 002 持國民中學補習學校資格證明書者
                        // 003 持國民中學補習學校結(修)業證明書者
                        // 004 持國民中學修業證明書者(修習三年級課程)
                        // 005 持國民中學畢業程度學力鑑定考詴及格證明書者
                        // 006 回國僑生(專案核准)
                        // 007 持大陸學歷者(需附證明文件)
                        // 008 特殊教育學校學生(需附證明文件)
                        // 009 持國外學歷者(需附證明文件)
                        // 010 取得相當於丙級(含)以上技術士證之資格者(需附證明文件)
                        // 011 持香港澳門學歷者(需附證明文件)
                        // 099 其他(需附證明文件

                        if (Convert.ToInt16(record.UpdateCode) == 001 || Convert.ToInt16(record.UpdateCode) == 002) //異動代碼 = 001 、002 就當成一般畢業證書
                        {
                            return "當年畢業";
                        }
                        if (Convert.ToInt16(record.UpdateCode) == 004) //異動代碼 =004 就當拿修業證書
                        {
                            return "當年修業";
                        }
                        else
                        {
                            return "其他(含領結業證書)";
                        }
                    }
                }
            }
            return "其他(含領結業證書)";
        }

        private string getNodeData(string nodeName, XmlElement Element, string nodesName)
        {
            string value = "";
            foreach (XmlElement xe in Element.SelectNodes(nodesName))
            {
                if (xe.SelectSingleNode(nodeName) != null)
                    value = xe.SelectSingleNode(nodeName).InnerText;
            }
            return value;
        }


    }

}




