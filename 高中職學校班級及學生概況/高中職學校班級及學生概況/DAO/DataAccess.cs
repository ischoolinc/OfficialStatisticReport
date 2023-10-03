using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;

namespace 高中職學校班級及學生概況.DAO
{
    public class DataAccess
    {
        // 取得系統內部別與科別資料
        public static List<DeptGroupInfo> GetDeptGroupInfo()
        {
            List<DeptGroupInfo> retVal = new List<DeptGroupInfo>();

            string sql = @"
            SELECT
                dept_group.name AS dept_group_name,
                dept.name AS dept_name,
                dept.code AS dept_code
            FROM
                dept_group
                INNER JOIN dept ON dept_group.id = ref_dept_group_id
            ORDER BY
                dept_group.code,
                dept.code
             ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sql);
            foreach(DataRow dr in dt.Rows)
            {
                DeptGroupInfo data = new DeptGroupInfo();
                data.DeptGroupName = dr["dept_group_name"] + "";
                data.DeptName = dr["dept_name"] + "";
                data.DeptCode = dr["dept_code"] + "";
                retVal.Add(data);
            }

            return retVal;
        }
    }
}
