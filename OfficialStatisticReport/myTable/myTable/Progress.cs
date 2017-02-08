using FISCA.Permission;
using FISCA.UDT;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using K12.Data;

namespace myTable
{
    public class Progress
    {
        [FISCA.MainMethod]
        public static void main()
        {
            FISCA.Presentation.RibbonBarItem item1 = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "資料統計"];
            item1["報表"].Image = Properties.Resources.Report;
            item1["報表"].Size = FISCA.Presentation.RibbonBarButton.MenuButtonSize.Large;
            item1["報表"]["新生入學方式統計表"].Enable = Permissions.新生入學方式統計表權限;
            item1["報表"]["新生入學方式統計表"].Click += delegate
            {
                Form2 form = new Form2();
                form.ShowDialog();
            };

            //權限設定
            Catalog permission = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.新生入學方式統計表, "新生入學方式統計表"));


            //2017/2/8 穎驊註解，原本因應公務統計報表 "新生入學方式統計報表"項目 而新增提供的預設類別項目，
            //為了防止 所有使用本類別模組的學校(包含 沒有使用公務統計報表模組的學校)  一併被提供到預設定類別項目
            //將下面的CODE  自MOD_Tagging模組 Tagging 專案 其 Program   移轉過來  
            //如此一來就只有 使用公務統計報表模組的學校會被新增類別項目

            #region 加入預設的入學方式、入學身分、原住民類別

            List<string> EnterSchoolWays = new List<string>();
            List<string> EnterSchoolIdentities = new List<string>();

            List<string> aboList = new List<string>();
            List<string> aboList2 = new List<string>();

            //九種入學方式
            EnterSchoolWays.Add("免試入學--校內直升");
            EnterSchoolWays.Add("免試入學--就學區免試(含共同就學區)");
            EnterSchoolWays.Add("免試入學--技優甄審");
            EnterSchoolWays.Add("免試入學--免試獨招");
            EnterSchoolWays.Add("免試入學--其他");
            EnterSchoolWays.Add("特色招生--考試分發");
            EnterSchoolWays.Add("特色招生--甄選入學");
            EnterSchoolWays.Add("適性輔導安置(十二年安置)");
            EnterSchoolWays.Add("其他");

            //四種入學身分
            EnterSchoolIdentities.Add("一般生(非外加錄取)");
            EnterSchoolIdentities.Add("外加錄取--原住民生");
            EnterSchoolIdentities.Add("外加錄取--身心障礙生");
            EnterSchoolIdentities.Add("外加錄取--其他");

            //十七種 原住民身分
            aboList.Add("阿美族");
            aboList.Add("泰雅族");
            aboList.Add("排灣族");
            aboList.Add("布農族");
            aboList.Add("卑南族");
            aboList.Add("鄒(曹)族");
            aboList.Add("魯凱族");
            aboList.Add("賽夏族");
            aboList.Add("雅美族或達悟族");
            aboList.Add("卲族");
            aboList.Add("噶瑪蘭族");
            aboList.Add("太魯閣族(含 德魯固族)");
            aboList.Add("撒奇萊雅族");
            aboList.Add("賽德克族");
            aboList.Add("拉阿魯哇族");
            aboList.Add("卡那卡那富族");
            aboList.Add("其他");

            aboList2.Add("阿美族");
            aboList2.Add("泰雅族");
            aboList2.Add("排灣族");
            aboList2.Add("布農族");
            aboList2.Add("卑南族");
            aboList2.Add("鄒(曹)族");
            aboList2.Add("魯凱族");
            aboList2.Add("賽夏族");
            aboList2.Add("雅美族或達悟族");
            aboList2.Add("卲族");
            aboList2.Add("噶瑪蘭族");
            aboList2.Add("太魯閣族(含 德魯固族)");
            aboList2.Add("撒奇萊雅族");
            aboList2.Add("賽德克族");
            aboList2.Add("拉阿魯哇族");
            aboList2.Add("卡那卡那富族");
            aboList2.Add("其他");

            // 若學校本來自己就有"原住民" Tag ，則以加入他的原住民項目 為主，幫他補齊。
            bool Tag_Prefix原校內為原住民 = false;

            //排除已加入的名單，避免重覆insert會爆掉
            foreach (TagConfigRecord each in TagConfig.SelectAll())
            {
                if (each.Prefix == "入學方式")
                {
                    if (EnterSchoolWays.Contains(each.Name))
                    {
                        EnterSchoolWays.Remove(each.Name);
                    }
                }
                if (each.Prefix == "入學身分")
                {
                    if (EnterSchoolIdentities.Contains(each.Name))
                    {
                        EnterSchoolIdentities.Remove(each.Name);
                    }
                }
                if (each.Prefix == "原住民")
                {
                    if (aboList.Contains(each.Name))
                    {
                        aboList.Remove(each.Name);
                    }
                    Tag_Prefix原校內為原住民 = true;
                }
                if (each.Prefix == "原住民族別")
                {
                    if (aboList.Contains(each.Name))
                    {
                        aboList2.Remove(each.Name);
                    }
                }
            }

            // 加入 入學方式 Tag
            foreach (string aboRaceName in EnterSchoolWays)
            {
                TagConfigRecord _current_tag;

                _current_tag = new TagConfigRecord();
                _current_tag.Category = TagCategory.Student.ToString();
                _current_tag.Prefix = "入學方式";
                _current_tag.Name = aboRaceName;
                _current_tag.Color = System.Drawing.Color.White;

                TagConfig.Insert(_current_tag);
            }

            // 加入 入學身分 Tag
            foreach (string aboRaceName in EnterSchoolIdentities)
            {
                TagConfigRecord _current_tag;

                _current_tag = new TagConfigRecord();
                _current_tag.Category = TagCategory.Student.ToString();
                _current_tag.Prefix = "入學身分";
                _current_tag.Name = aboRaceName;
                _current_tag.Color = System.Drawing.Color.White;

                TagConfig.Insert(_current_tag);
            }

            //加入 原住民Tag
            if (Tag_Prefix原校內為原住民)
            {
                foreach (string aboRaceName in aboList)
                {
                    TagConfigRecord _current_tag;

                    _current_tag = new TagConfigRecord();
                    _current_tag.Category = TagCategory.Student.ToString();
                    _current_tag.Prefix = "原住民";
                    _current_tag.Name = aboRaceName;
                    _current_tag.Color = System.Drawing.Color.White;

                    TagConfig.Insert(_current_tag);
                }

            }
            else
            {
                foreach (string aboRaceName in aboList2)
                {
                    TagConfigRecord _current_tag;

                    _current_tag = new TagConfigRecord();
                    _current_tag.Category = TagCategory.Student.ToString();
                    _current_tag.Prefix = "原住民族別";
                    _current_tag.Name = aboRaceName;
                    _current_tag.Color = System.Drawing.Color.White;

                    TagConfig.Insert(_current_tag);
                }
            }
            #endregion

        }
    }
}
