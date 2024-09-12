using FISCA.Data;
using K12.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ArrangeClass
{
    public class QueryTransfer
    {
        /// <summary>
        /// 取得系統班級名稱與設定檔對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetClassNameList()
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();

            // Print.cs 濾出 狀態 為 1.一般、4.休學、2.延修 的學生，休學、延修的班級名稱不用填寫
            QueryHelper qh = new QueryHelper();
            string query = @"WITH classTable AS(
SELECT DISTINCT 
	class.id AS class_id
	, class_name
	, class.grade_year 
	, display_order
FROM 
	class 
INNER JOIN 
	student 
	ON class.id=student.ref_class_id 
	AND student.status IN (1,2,4) 
ORDER BY 
	class.grade_year 
	, display_order
	, class_name
), config_table AS (
    SELECT     
      (array_to_string(xpath('//Item/@name', _xml), '')::text) AS name    
      , (array_to_string(xpath('//Item/@value', _xml), '')::text) AS value    
      FROM (    
        SELECT  
		unnest(xpath('//ClassName/Item', xmlparse(content '<root>'||content||'</root>')))  AS _xml
        FROM list WHERE name='編班名冊設定檔'    
        ) AS XML   
)
	SELECT classTable.*
	, value
	FROM classTable
	LEFT JOIN config_table ON config_table.name=classTable.class_name";
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                string className = dr["class_name"].ToString();
                string value = dr["value"].ToString();
                retVal.Add(className, value);
            }
            return retVal;
        }

        /// <summary>
        /// 取得編班名冊設定檔: 班級名稱對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetConfigure()
        {
            QueryHelper qh = new QueryHelper();
            string query = @"SELECT * FROM list WHERE name = '編班名冊設定檔' ";
            DataTable dt = qh.Select(query);
            string content = "";
            foreach (DataRow dr in dt.Rows)
            {
                content = dr["content"] + "";
            }
            Dictionary<string, string> dict = new Dictionary<string, string>();

            // 創建XmlDocument對象並加載XML
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(content);

            // 獲取所有Item元素
            XmlNodeList itemList = xmlDoc.SelectNodes("//Item");

            // 遍歷所有Item元素，將其name屬性和value屬性加入字典中
            foreach (XmlNode itemNode in itemList)
            {
                XmlAttribute nameAttribute = itemNode.Attributes["name"];
                XmlAttribute valueAttribute = itemNode.Attributes["value"];

                if (nameAttribute != null && valueAttribute != null)
                {
                    dict[nameAttribute.Value] = valueAttribute.Value;
                }
            }


            return dict;
        }

        /// <summary>
        /// 儲存設定檔
        /// </summary>
        /// <param name="xml"></param>
        /// <returns>成功/失敗</returns>
        public static bool SaveConfigure(string xml)
        {
            bool successed = true;

            UpdateHelper updateHelper = new UpdateHelper();
            string sql = @"  INSERT INTO
	                                        list (name, content)
	                                        VALUES
	                                        ('編班名冊設定檔', '" + xml + @"')
	                                        ON CONFLICT(name) DO UPDATE
	                                        SET content ='" + xml + "'";
            try
            {
                updateHelper.Execute(sql);
            }
            catch (Exception ex)
            {
                successed = false;
            }

            return successed;
        }

        // 取得 教務作業>批次作業/檢視>異動作業>核班人數維護 資料內容
        public static Dictionary<string, string> GetClassTyepUDict(string SchoolYear)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string query = string.Format(@"
                SELECT
                    class_type,
                    class_typeu
                FROM
                    $campus.updaterecord.govapprovednumofclass
                WHERE
                    schoolyear = {0}
                ", SchoolYear);

                DataTable dt = qh.Select(query);
                foreach(DataRow dr in dt.Rows)
                {
                    string classType = dr["class_type"] + "";
                    if (!value.ContainsKey(classType))
                        value.Add(classType, dr["class_typeu"] + "");
                }                
            }
            catch (Exception ex)
            {
                Console.WriteLine("取得核班人數維護錯誤," + ex.Message);
            }

            return value;
        }

    }
}
