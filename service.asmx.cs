using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Services;
using System.Xml;

namespace FCPortal
{
    /// <summary>
    /// service 的摘要说明
    /// </summary>
    [WebService(Namespace = "http://10.137.12.32/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    // 若要允许使用 ASP.NET AJAX 从脚本中调用此 Web 服务，请取消注释以下行。 
    // [System.Web.Script.Services.ScriptService]
    public class service : System.Web.Services.WebService
    {

        [WebMethod]
        public string GetInformation(string year, string month)
        {
            XmlTextWriter xmlWriter;
            //string strFilename = Server.MapPath("data.xml");
            MemoryStream m = new MemoryStream();
            xmlWriter = new XmlTextWriter(m, Encoding.GetEncoding("GB2312"));//创建一个xml文档
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("DATA");
            

            try
            {
                DataTable dtAllContracts = data.msSQL.GetDataTable("select distinct(FO_NO) from FC_SESRelatedData where Expire_Date>='" + (new DateTime(int.Parse(year), int.Parse(month), 1)).ToString("yyyy-MM-dd") + "'");
                DateTime dtStart = DateTime.Parse(year + "-" + month + "-16");
                if (dtAllContracts != null && dtAllContracts.Rows.Count > 0)
                {
                    for (int j = 0; j < dtAllContracts.Rows.Count; j++)
                    {
                        string sFO = data.object2Str(dtAllContracts.Rows[j][0]);
                        //SELECT FO_NO,Contract_Title,Contractor,Original_WC,Contract_Admin,Main_Coordinator,dbo.getScore(FO_NO,'2016-01-01',1,'11',3) as 'ScoreAll',dbo.getScore(FO_NO,'2016-01-01',1,'11',1) as 'ScoreUser',dbo.getScore(FO_NO,'2016-01-01',1,'11',2) as 'ScoreDep',dbo.getScore(FO_NO,'2016-01-01',1,'11',-171) as 'ScoreTimely',dbo.getScore(FO_NO,'2016-01-01',1,'11',-172) as 'ScoreHonesty' from FCMain WHERE FO_NO='4915791920'
                        DataTable dtAvg = data.msSQL.GetDataTable("SELECT FO_NO,Contract_Title,Contractor,Original_WC,Contract_Admin,Main_Coordinator,dbo.getScore(FO_NO,'" + dtStart.AddMonths(1).ToString("yyyy-MM-01") + "',1,'11',3) as 'ScoreAll',dbo.getScore(FO_NO,'" + dtStart.AddMonths(1).ToString("yyyy-MM-01") + "',1,'11',1) as 'ScoreUser',dbo.getScore(FO_NO,'" + dtStart.AddMonths(1).ToString("yyyy-MM-01") + "',1,'11',2) as 'ScoreDep',dbo.getScore(FO_NO,'" + dtStart.AddMonths(1).ToString("yyyy-MM-01") + "',1,'11',-171) as 'ScoreTimely',dbo.getScore(FO_NO,'" + dtStart.AddMonths(1).ToString("yyyy-MM-01") + "',1,'11',-172) as 'ScoreHonesty' from FCMain WHERE FO_NO='" + sFO + "'");
                        if (dtAvg != null && dtAvg.Rows.Count > 0 && !string.IsNullOrEmpty(data.object2Str(dtAvg.Rows[0][6])))
                        {
                            xmlWriter.WriteStartElement("ROW");
                            xmlWriter.WriteStartElement("contractno");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][0]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("contracttitle");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][1]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("suppliername");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][2]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("suppliercode");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][3]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("year");
                            xmlWriter.WriteString(year);
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("month");
                            xmlWriter.WriteString(month);
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("contractadmin");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][4]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("maincoordinator");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][5]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("score");
                            xmlWriter.WriteString(data.object2Str(dtAvg.Rows[0][6]));
                            xmlWriter.WriteEndElement();
                            xmlWriter.WriteStartElement("detail");
                                xmlWriter.WriteStartElement("item");
                                    xmlWriter.WriteStartElement("itemname");
                                        xmlWriter.WriteString("用户");
                                    xmlWriter.WriteEndElement();
                                    xmlWriter.WriteStartElement("itemscore");
                                    xmlWriter.WriteString(getPercentage(data.object2Str(dtAvg.Rows[0][7]),70.0));
                                    xmlWriter.WriteEndElement();
                                xmlWriter.WriteEndElement();
                                xmlWriter.WriteStartElement("item");
                                    xmlWriter.WriteStartElement("itemname");
                                    xmlWriter.WriteString("职能部门");
                                    xmlWriter.WriteEndElement();
                                    xmlWriter.WriteStartElement("itemscore");
                                    xmlWriter.WriteString(getPercentage(data.object2Str(dtAvg.Rows[0][8]),30.0));
                                    xmlWriter.WriteEndElement();
                                xmlWriter.WriteEndElement();
                                xmlWriter.WriteStartElement("item");
                                    xmlWriter.WriteStartElement("itemname");
                                    xmlWriter.WriteString("及时性");
                                    xmlWriter.WriteEndElement();
                                    xmlWriter.WriteStartElement("itemscore");
                                    xmlWriter.WriteString(getPercentage(data.object2Str(dtAvg.Rows[0][9]),5.0));
                                    xmlWriter.WriteEndElement();
                                xmlWriter.WriteEndElement();
                                xmlWriter.WriteStartElement("item");
                                    xmlWriter.WriteStartElement("itemname");
                                    xmlWriter.WriteString("诚实度");
                                    xmlWriter.WriteEndElement();
                                    xmlWriter.WriteStartElement("itemscore");
                                    xmlWriter.WriteString(getPercentage(data.object2Str(dtAvg.Rows[0][10]),5.0));
                                    xmlWriter.WriteEndElement();
                                xmlWriter.WriteEndElement();
                            xmlWriter.WriteEndElement();

                            xmlWriter.WriteEndElement();
                        }

                    }


                }
            }
            catch (Exception)
            {

                //throw;
            }

            xmlWriter.WriteEndElement();

            xmlWriter.Flush();// 确保书写器更新到Stream中
            m.Position = 0;//重置流的位置，以便我们可以从头读取  
            string str = System.Text.Encoding.GetEncoding("GB2312").GetString(m.ToArray());

            xmlWriter.Close();
            m.Close();

            return str;//File.ReadAllText(Server.MapPath("data.xml"), Encoding.GetEncoding("GB2312"));
        }

        private string getPercentage(string sScore,double dAllScore)
        {
            try
            {
                if (string.IsNullOrEmpty(sScore))
                {
                    sScore = (dAllScore * 0.6).ToString();
                }
                double dScore = double.Parse(sScore.Trim());
                return (dScore/ dAllScore).ToString("0.00%");
            }
            catch (Exception)
            {
                //throw;
            }
            return "";
        }

        [WebMethod]
        public string GetCPTInformation(DateTime fromDate, DateTime endDate)
        {
            XmlTextWriter xmlWriter;
            //string strFilename = Server.MapPath("data.xml");
            MemoryStream m = new MemoryStream();
            xmlWriter = new XmlTextWriter(m, Encoding.GetEncoding("GB2312"));//创建一个xml文档
            xmlWriter.Formatting = Formatting.Indented;
            xmlWriter.WriteStartDocument();
            xmlWriter.WriteStartElement("DATA");
            try
            {
                DataTable dt = data.msSQL.GetDataTable("SELECT Contract_No,FC_Desctription,CPT_No,Net_Amount,Tax_Amount,Tax,Currency,Report_Date from CPTList where Report_Date>='" + fromDate.ToString("yyyy-MM-dd") + "' and Report_Date<='"+endDate.ToString("yyyy-MM-dd")+"'");
                if (dt!=null&&dt.Rows.Count>0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        xmlWriter.WriteStartElement("ROW");

                        xmlWriter.WriteStartElement("contractno");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][0]));
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("contractTitle");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][1]));
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("CPTNo");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][2]));
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("Amounttobepaid");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][3]));
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("AmounttobepaidTax");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][4]));
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("Tax");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][5]));
                        xmlWriter.WriteEndElement();

                        xmlWriter.WriteStartElement("Currency");
                        xmlWriter.WriteString(data.object2Str(dt.Rows[i][6]));
                        xmlWriter.WriteEndElement();

                        DataTable dtSES = data.msSQL.GetDataTable("SELECT SES,Short_Description,Budget,Quotation,Net_Value,Tax_Value,Con_Days,Submit_Date from SESList where CPT_No='" + data.object2Str(dt.Rows[i][2]) + "'");
                        
                        if (dtSES!=null&&dtSES.Rows.Count>0)
                        {
                            for (int j = 0; j < dtSES.Rows.Count; j++)
                            {
                                xmlWriter.WriteStartElement("detail");

                                xmlWriter.WriteStartElement("SESNo");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][0]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteStartElement("ShortDescription");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][1]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteStartElement("Budget");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][2]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteStartElement("Quotation");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][3]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteStartElement("NetValue");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][4]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteStartElement("InclTax");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][5]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteStartElement("Con");
                                xmlWriter.WriteString(data.object2Str(dtSES.Rows[j][6]));
                                xmlWriter.WriteEndElement();

                                xmlWriter.WriteEndElement();

                            }
                        }
                       


                        xmlWriter.WriteEndElement();
                    }
                }
            }
            catch (Exception)
            {
                //throw;
            }
            xmlWriter.WriteEndElement();

            xmlWriter.Flush();// 确保书写器更新到Stream中
            m.Position = 0;//重置流的位置，以便我们可以从头读取  
            string str = System.Text.Encoding.GetEncoding("GB2312").GetString(m.ToArray());

            xmlWriter.Close();
            m.Close();

            return str;//File.ReadAllText(Server.MapPath("data.xml"), Encoding.GetEncoding("GB2312"));
        }
    }
}
