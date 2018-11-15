using Ext.Net;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Web;

namespace FCPortal
{
    /// <summary>
    /// data 的摘要说明
    /// </summary>
    public class data : IHttpHandler
    {
        //public static clsMSSQL.MsSql msSQL = new clsMSSQL.MsSql("Persist Security Info=False;User ID=sa;pwd=documentum;Initial Catalog=CPT;Data Source=10.137.12.32");//10.137.12.32

        public static clsMSSQL.MsSql msSQL = new clsMSSQL.MsSql("Persist Security Info=False;User ID=sa;pwd=administrator123;Initial Catalog=CPT;Data Source=127.0.0.1");//10.137.12.32

        public static string sSQLRecent = " AND (Deleted is NULL and Blocked is NULL) AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND datediff(day, B.DateIn,GETDATE())>16 )";
        public static int iStartDay = 11;
        public void ProcessRequest(HttpContext context)
        {
            //refine the SQL
            DateTime d=DateTime.Now;
            //if (d.Day > iStartDay)
            //{
            //    sSQLRecent = " AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<='"+d.Year.ToString()+"-"+d.Month.ToString("00")+"-"+(iStartDay+1).ToString()+"' )";
            //}
            //else
            //{
            //    d = d.AddMonths(-1);
            //    sSQLRecent = " AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<='" + d.Year.ToString() + "-" + d.Month.ToString("00") + "-"+(iStartDay+1).ToString() + "' )";
            //}
            sSQLRecent = " AND (Deleted is NULL and Blocked is NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'" + d.ToString("yyyy-MM-01") + "' )";

            StoreRequestParameters prms = new StoreRequestParameters(context);
            
            string sSource = context.Request.Params["object"];
            string sNO=context.Request.Params["NO"];
            string sExportExcel = context.Request.Params["exportExcel"];
            string sExportReportExcel = context.Request.Params["exportReportExcel"];
            string sExportFileID = context.Request.Params["fileID"];
            string sAction = context.Request.Params["action"];
            string sOutPut = "";
            if (sAction=="statistics")
            {
                string sDateNow = context.Request.Params["date"];
                DateTime dtTime = DateTime.Now;
                if (!string.IsNullOrEmpty(sDateNow))
                {
                    dtTime = DateTime.Parse(sDateNow);
                }
                //newly added 2014-03-10
                //SELECT DISTINCT(Requisitioner) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'2014-03-01') AND Requisitioner is not NULL and Requisitioner NOT in (SELECT DISTINCT([By]) FROM FC_Score WHERE DateIn>='2014-03-01' AND Role='User')
                //User
                string sOut = "Null";
                DataTable dt = msSQL.GetDataTable("SELECT count(DISTINCT(Requisitioner)) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL)  AND (Deleted is NULL and Blocked is NULL)   AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'" + dtTime.ToString("yyyy-MM-01") + "') AND Requisitioner is not NULL");
                if (dt!=null&&dt.Rows.Count==1)
                {
                    int iAll = int.Parse(dt.Rows[0][0].ToString());
                    dt = msSQL.GetDataTable("SELECT count(DISTINCT(Requisitioner)) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) AND (Deleted is NULL and Blocked is NULL)   AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'" + dtTime.ToString("yyyy-MM-01") + "') AND Requisitioner is not NULL and Requisitioner NOT in (SELECT DISTINCT([By]) FROM FC_Score WHERE DateIn>='" + dtTime.ToString("yyyy-MM-01") + "' AND Role='User' AND Remark<>'Auto')");
                    if (dt != null && dt.Rows.Count == 1)
                    {
                        int iLeft = int.Parse(dt.Rows[0][0].ToString());
                        sOut = "User:" + (iAll - iLeft).ToString() + "/" + iAll.ToString() + "≈" + ((double)(iAll - iLeft)/(double)iAll).ToString("##.##%");
                    }
                }
                context.Response.Write(sOut+"<br/>Dep.:"+((dtTime.Month==DateTime.Now.Month)?getDepStatistics():"Null"));
                return;
            }
            if (context.Request["js"] != null)
            {
                context.Response.ContentType = "text/javascript";
                string sJsFilePath = Path.Combine(HttpContext.Current.Request.PhysicalApplicationPath, context.Request["js"]);
                if (File.Exists(sJsFilePath))
                {
                    context.Response.Write(File.ReadAllText(sJsFilePath));
                }
                return;
            } 
            //for exporting excel functions
            if (!string.IsNullOrEmpty(sExportReportExcel))
            {
                exportReportExcel(context,DateTime.Now);
                return;
            }
            if (!string.IsNullOrEmpty(sExportExcel))
            {
                exportExcel(context);
                return;
            }
            if (!string.IsNullOrEmpty(sExportFileID))
            {
                string sPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("Files"), sExportFileID);
                string sFileName = sExportFileID;
                DataTable dt = msSQL.GetDataTable("select FileName from FC_File where FileID='"+sExportFileID+"'");
                if (dt!=null&&dt.Rows.Count>0)
                {
                    sFileName = dt.Rows[0][0].ToString();
                }
                if (!File.Exists(sPath))
                {
                    context.Response.Write("File Not Found!");
                    return;
                }
                context.Response.Clear();
                context.Response.ContentType = "application/octet-stream";
                context.Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(sFileName));
                context.Response.TransmitFile(sPath);
                context.Response.End();
                return;
            }
            context.Response.ContentType = "application/json";
            //if (sSource == "overallKPI")
            //{
            //    //getting the pricing type
            //    //List<Ext.Net.Examples.ChartData> data=Ext.Net.Examples.ChartData.GenerateData(6);

            //    List<clsFCDataPie> data = KPIData(sNO);
            //    sOutPut = JSON.Serialize(data);
            //    context.Response.Write(sOutPut);
            //    return;
            //}
            if (sSource == "scatterKPI" || sSource == "scatterKPI2" || sSource == "overallKPI")
            {
                string sDateStart = context.Request.Params["DateStart"];
                DateTime dStart = DateTime.Now;
                //if (sSource=="scatterKPI2")
                //{
                //    dStart = DateTime.Now.AddYears(-1);
                //}
                if (!string.IsNullOrEmpty(sDateStart) && sDateStart.Contains("T"))
                {
                    dStart = DateTime.Parse(sDateStart.Replace("T"," "));
                }
                string sDateEnd = context.Request.Params["DateEnd"];
                DateTime dEnd = DateTime.Now;
                if (!string.IsNullOrEmpty(sDateEnd) && sDateEnd.Contains("T"))
                {
                    dEnd = DateTime.Parse(sDateEnd.Replace("T", " "));
                }
                string sOverAll = context.Request.Params["OverAll"];
                if (string.IsNullOrEmpty(sOverAll))
                {
                    sOverAll = "false";
                }
                List<clsFCDataPie> data = null;
                if (sSource != "overallKPI")
                {
                    data = (sSource == "scatterKPI") ? KPIScatter(dStart.ToString("yyyy-MM"), dEnd.ToString("yyyy-MM"), bool.Parse(sOverAll)) : KPIScatter2(sNO, dStart.ToString("yyyy-MM"), dEnd.ToString("yyyy-MM"));
                }
                else
                {
                    data = KPIData(sNO, dStart.ToString("yyyy-MM"), dEnd.ToString("yyyy-MM"));
                }
                sOutPut = JSON.Serialize(data);
                context.Response.Write(sOutPut);
                return;
            }
            if (sSource == "userKPI" || sSource == "userSample" || sSource == "contractorKPI" || sSource == "userPro" || sSource == "DisDecKPI")
            {
                sNO = "test";
            }
            if (string.IsNullOrEmpty(sSource))
            {                
                sOutPut = JSON.Serialize(this.DataBack(prms,context));
                context.Response.Write(sOutPut);
                return;
            }
            else if (!string.IsNullOrEmpty(sNO))
            {
                if (sSource == "pie1" || sSource == "pie3" || sSource == "userKPI" || sSource == "userSample" || sSource == "contractorKPI"||sSource== "userPro" || sSource == "DisDecKPI")
                {
                    string sOverAll = context.Request.Params["OverAll"];
                    if (string.IsNullOrEmpty(sOverAll))
                    {
                        sOverAll = "false";
                    }
                    //getting the pricing type
                    //List<Ext.Net.Examples.ChartData> data=Ext.Net.Examples.ChartData.GenerateData(6);
                    string sDateStart = context.Request.Params["DateStart"];
                    DateTime dStart = DateTime.Now.AddMonths(-12);
                    if (!string.IsNullOrEmpty(sDateStart) && sDateStart.Contains("T"))
                    {
                        dStart = DateTime.Parse(sDateStart.Replace("T", " "));
                    }
                    string sDateEnd = context.Request.Params["DateEnd"];
                    DateTime dEnd = DateTime.Now;
                    if (!string.IsNullOrEmpty(sDateEnd) && sDateEnd.Contains("T"))
                    {
                        dEnd = DateTime.Parse(sDateEnd.Replace("T", " "));
                    }
                    if (sSource == "userPro")
                    {
                        List<clsFCDataPie> dataAll = PieData1(sNO, dStart.ToString("yyyy-MM"), dEnd.ToString("yyyy-MM"), sSource, false, sOverAll);
                        //[{"Section":2005,"STS":24810000},{"Section":2008,"LIF":38910000,"STS":56070000}]
                        //List<object> data = new List<object> 
                        //{ 
                        //    new {Section = 2005, STS = 24810000},
                        //    new {Section = 2008, LIF = 38910000, STS = 56070000}
                        //};
                        sOutPut = "[";//JSON.Serialize(data);
                        string sNowSection = "", sAllDis = "BCM,CAD,CAL,CCS,CIV,CLE,EHR,EIC,ENG,EPC,EPM,EPR,ERM,FFS,FSY,HVA,INS,IPV,IVR,LES,LIF,MAS,MOR,NDE,PAI,PQS,QSS,RBC,ROC,SCA,SLP,STS,TRA,CFM,HOT,OEM";//"ENG,STS,SLP,SCA,QSS,PQS,PAI,OEM,NDE,MOR,MAS,LIF,LES,IVR,IPV,INS,HVA,FFS,ERM,EPR,BCM";
                        ArrayList alName = new ArrayList(), alValue = new ArrayList();
                        for (int i = 0; i < dataAll.Count; i++)
                        {
                            if (!string.IsNullOrEmpty(dataAll[i].Name) && sAllDis.Contains(dataAll[i].MOM))
                            {
                                int iIndex = alName.IndexOf(dataAll[i].Name);
                                if (iIndex >= 0)
                                {
                                    alValue[iIndex] = ((double)alValue[iIndex]) + dataAll[i].Data1;
                                }
                                else
                                {
                                    alName.Add(dataAll[i].Name);
                                    alValue.Add(dataAll[i].Data1);
                                }
                            }
                        }
                        for (int i = 0; i < dataAll.Count; i++)
                        {
                            if (!string.IsNullOrEmpty(dataAll[i].Name) && sAllDis.Contains(dataAll[i].MOM))
                            {
                                if (string.IsNullOrEmpty(sNowSection))
                                {
                                    sNowSection = dataAll[i].Name;
                                    sOutPut += "{\"Section\":\"" + sNowSection + "\",\"" + dataAll[i].MOM + "\":" + (100.0 * dataAll[i].Data1 / (double)alValue[alName.IndexOf(dataAll[i].Name)]).ToString("0.0");
                                }
                                else
                                {
                                    if (sNowSection == dataAll[i].Name)
                                    {
                                        sOutPut += ",\"" + dataAll[i].MOM + "\":" + (100.0 * dataAll[i].Data1 / (double)alValue[alName.IndexOf(dataAll[i].Name)]).ToString("0.0");
                                    }
                                    else
                                    {
                                        sNowSection = dataAll[i].Name;
                                        sOutPut += "},{\"Section\":\"" + sNowSection + "\",\"" + dataAll[i].MOM + "\":" + (100.0 * dataAll[i].Data1 / (double)alValue[alName.IndexOf(dataAll[i].Name)]).ToString("0.0");

                                    }
                                }
                            }
                        }
                        sOutPut += "}]";
                        context.Response.Write(sOutPut);
                        return;
                    }
                    else
                    {
                        List<clsFCDataPie> data = PieData1(sNO, dStart.ToString("yyyy-MM"), dEnd.ToString("yyyy-MM"), sSource, bool.Parse(sOverAll));
                        sOutPut = JSON.Serialize(data);
                        context.Response.Write(sOutPut);
                    }
                    return;
                }
                if (sSource == "pie2")
                {
                    string sMonths = context.Request.Params["Months"];
                    string sEndDate = context.Request.Params["EndDate"];
                    if (!string.IsNullOrEmpty(sMonths))
                    {
                        List<clsFCDataPie2> data = PieData2(sNO, sMonths, sEndDate);
                        sOutPut = JSON.Serialize(data);
                        context.Response.Write(sOutPut);
                        return; 
                    }
                }
                if (sSource == "evaluation")
                {
                    string sUser = context.Request.Params["User"];
                    string sisUser = context.Request.Params["isUser"];
                    if (!string.IsNullOrEmpty(sUser))
                    {
                        List<clsFCSESReport> data = getEvaluation(sNO, sUser, (sisUser.ToLower() == "true"));
                        sOutPut = JSON.Serialize(data);
                        context.Response.Write(sOutPut);
                        return;
                    }
                }
                if (sSource == "getfiles")
                {
                    List<clsContractFile> data = getContractFiles(sNO);
                    sOutPut = JSON.Serialize(data);
                    context.Response.Write(sOutPut);
                    return;
                }

            }
            else
            {
                if (sSource == "evaluationContract")
                {
                    string sUser = context.Request.Params["User"];
                    string sisUser = context.Request.Params["isUser"];
                    if (!string.IsNullOrEmpty(sUser))
                    {
                        List<clsFCEvaluation> data = getEvaluationContract(sUser,(sisUser.ToLower()=="true"));
                        sOutPut = JSON.Serialize(data);
                        context.Response.Write(sOutPut);
                        return;
                    }
                    
                }
            }
            context.Response.Write("[]");
            return;
        }

        private Paging<clsFCData> DataBack(StoreRequestParameters prms,HttpContext context)
        {            
            //now getting data
            string sSort = "[Un-used_Percentage]";
            string sDir = "ASC";
            string sSearchFilter = " where DATEDIFF(MONTH, GETDATE(), Expire_Date)>-24";
            if (prms.Sort.Length>0&&!string.IsNullOrEmpty(prms.Sort[0].Property)&&!string.IsNullOrEmpty(prms.Sort[0].Direction.ToString()))
            {
                sSort = getDBName(prms.Sort[0].Property);
                sDir = prms.Sort[0].Direction.ToString();
            }
            //now for search filters
            string sKeyword = context.Request.Params["searchKey"].Replace("null", "");
            string sNO = context.Request.Params["searchNO"].Replace("null", "");
            string sCA = context.Request.Params["CA"].Replace("null", "");
            string sPS = context.Request.Params["PS"].Replace("null", "");
            string sValid = context.Request.Params["isValid"].Replace("null", "");
            //if (!string.IsNullOrEmpty(sNO))
            //{
            //    sSearchFilter += " where FO_NO like '%"+sNO+"%'";
            //}
            if (!string.IsNullOrEmpty(sKeyword)&&sKeyword.ToLower()!="all")
            {
                //get Chinese and English out
                string ptn = "[\u4e00-\u9fa5]+|[a-zA-Z\\s]+";
                System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(ptn);
                System.Text.RegularExpressions.MatchCollection ms = reg.Matches(sKeyword);
                if (ms.Count>0)
                {
                    sKeyword = reg.Matches(sKeyword)[0].Value; 
                }
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where ([FO_NO] LIKE N'%" + sKeyword + "%' OR [Actual_Budget] LIKE N'%" + sKeyword + "%' OR [Net_Amount] LIKE N'%" + sKeyword + "%' OR [Un-used_Budget] LIKE N'%" + sKeyword + "%' OR [Un-used_Percentage] LIKE N'%" + sKeyword + "%' OR [Contract_Title] LIKE N'%" + sKeyword + "%' OR [Contractor] LIKE N'%" + sKeyword + "%' OR [Pricing_Scheme] LIKE N'%" + sKeyword + "%' OR [Original_WC] LIKE N'%" + sKeyword + "%' OR [Type] LIKE N'%" + sKeyword + "%' OR [Contract_Admin] LIKE N'%" + sKeyword + "%' OR [Buyer] LIKE N'%" + sKeyword + "%' OR [Main_Coordinator] LIKE N'%" + sKeyword + "%' OR [User_Representative] LIKE N'%" + sKeyword + "%' OR [Validate_Date] LIKE N'%" + sKeyword + "%' OR [Expire_Date] LIKE N'%" + sKeyword + "%' OR [Contract_Person] LIKE N'%" + sKeyword + "%' OR [Contract_Tel] LIKE N'%" + sKeyword + "%' OR [Score] LIKE N'%" + sKeyword + "%' )";
                }
                else
                {
                    sSearchFilter += " and ([FO_NO] LIKE N'%" + sKeyword + "%' OR [Actual_Budget] LIKE N'%" + sKeyword + "%' OR [Net_Amount] LIKE N'%" + sKeyword + "%' OR [Un-used_Budget] LIKE N'%" + sKeyword + "%' OR [Un-used_Percentage] LIKE N'%" + sKeyword + "%' OR [Contract_Title] LIKE N'%" + sKeyword + "%' OR [Contractor] LIKE N'%" + sKeyword + "%' OR [Pricing_Scheme] LIKE N'%" + sKeyword + "%' OR [Original_WC] LIKE N'%" + sKeyword + "%' OR [Type] LIKE N'%" + sKeyword + "%' OR [Contract_Admin] LIKE N'%" + sKeyword + "%' OR [Buyer] LIKE N'%" + sKeyword + "%' OR [Main_Coordinator] LIKE N'%" + sKeyword + "%' OR [User_Representative] LIKE N'%" + sKeyword + "%' OR [Validate_Date] LIKE N'%" + sKeyword + "%' OR [Expire_Date] LIKE N'%" + sKeyword + "%' OR [Contract_Person] LIKE N'%" + sKeyword + "%' OR [Contract_Tel] LIKE N'%" + sKeyword + "%' OR [Score] LIKE N'%" + sKeyword + "%' )";
                }
            }
            if (!string.IsNullOrEmpty(sNO))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where FO_NO like '%" + sNO + "%'";
                }
                else
                {
                    sSearchFilter += " and FO_NO like '%" + sNO + "%'";
                }
            }
            if (!string.IsNullOrEmpty(sCA))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where Contract_Admin like '%" + sCA + "%'";
                }
                else
                {
                    sSearchFilter += " and Contract_Admin like '%" + sCA + "%'";
                }
            }
            if (!string.IsNullOrEmpty(sPS))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where Pricing_Scheme like '%" + sPS + "%'";
                }
                else
                {
                    sSearchFilter += " and Pricing_Scheme like '%" + sPS + "%'";
                }
            }
            if (!string.IsNullOrEmpty(sValid)&&sValid.ToLower()=="true")
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where Expire_Date>getDate()";
                }
                else
                {
                    sSearchFilter += " and Expire_Date>getDate()";
                }
            }

            List<clsFCData> data = new List<clsFCData>();
            if (prms.Limit < 0)
            {
                return new Paging<clsFCData>(data, 0);
            }
            DataTable dt = msSQL.GetDataTable("SELECT TOP  " + prms.Limit.ToString() + " * from FCMain WHERE " + ((string.IsNullOrEmpty(sSearchFilter) ? "" : sSearchFilter.Replace("where", "") + " and ")) + " FO_NO " + " NOT IN  ( SELECT TOP " + ((prms.Page - 1) * prms.Limit).ToString() + " " + " FO_NO " + " FROM FCMain " + sSearchFilter + " ORDER BY " + sSort + " " + sDir + ") ORDER BY " + sSort + " " + sDir);
            if (dt!=null&&dt.Rows.Count>0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //now getting the values
                    clsFCData fd = new clsFCData(object2Str(dt.Rows[i]["FO_NO"]));
                    fd.bugdet = object2Float(dt.Rows[i]["Actual_Budget"]);
                    fd.buyer = object2Str(dt.Rows[i]["Buyer"]);
                    fd.contact = object2Str(dt.Rows[i]["Contract_Person"]);
                    fd.contractAdmin = object2Str(dt.Rows[i]["Contract_Admin"]);
                    fd.contractor = refineContractName(object2Str(dt.Rows[i]["Contractor"]));
                    fd.expireDate = object2Date(dt.Rows[i]["Expire_Date"]);
                    fd.fcDescription = refineContractName(object2Str(dt.Rows[i]["Contract_Title"]));
                    fd.mainCoordinator = object2Str(dt.Rows[i]["Main_Coordinator"]);
                    fd.netAmount = object2Float(dt.Rows[i]["Net_Amount"]);
                    fd.score = object2Float(dt.Rows[i]["score"]);
                    fd.telephone = object2Str(dt.Rows[i]["Contract_Tel"]);
                    fd.unusedBudget = object2Float(dt.Rows[i]["Un-used_Budget"]);
                    fd.unusedPercentage = object2Float(dt.Rows[i]["Un-used_Percentage"]);
                    fd.remainingPercentage = object2Float(dt.Rows[i]["Remaining_Duration"]); if (fd.remainingPercentage < 0) { fd.remainingPercentage = 0; }
                    fd.userRepresentative = object2Str(dt.Rows[i]["User_Representative"]);
                    fd.validateDate = object2Date(dt.Rows[i]["Validate_Date"]);
                    fd.workCenter = object2Str(dt.Rows[i]["Original_WC"]);
                    fd.pricingScheme = object2Str(dt.Rows[i]["Pricing_Scheme"]);
                    fd.remark = "";

                    data.Add(fd);
                }
            } 
            //getting the total number
            int iCount = dt.Rows.Count;
            dt = msSQL.GetDataTable("select count(*) from FCMain"+sSearchFilter);            
            if (dt!=null&&dt.Rows.Count==1)
            {
                iCount = int.Parse(dt.Rows[0][0].ToString());
            }
            return new Paging<clsFCData>(data, iCount);
        }
        private List<clsFCDataPie2> PieData2(string sNO, string sMonths,string sEndDatetime)
        {
            List<clsFCDataPie2> data = new List<clsFCDataPie2>();
            //DataTable dt = msSQL.GetDataTable("SELECT dbo.getScore('A546059088',DATEADD(MONTH, 1, GETDATE()),'10',3)");
            for (int i = int.Parse(sMonths); i <= 0; i++)
            {
                clsFCDataPie2 pie2 = new clsFCDataPie2();
                DataTable dt = msSQL.GetDataTable("SELECT dbo.getScore('" + sNO + "',DATEADD(MONTH, " + i.ToString() + ", convert(datetime,'"+sEndDatetime+"')),1,'" + iStartDay.ToString() + "',1)");
                if (dt!=null&&dt.Rows.Count>0&&!string.IsNullOrEmpty(object2Str(dt.Rows[0][0])))
                {
                    pie2.userScore = (int)(double.Parse(object2Str(dt.Rows[0][0])));
                }
                dt = msSQL.GetDataTable("SELECT dbo.getScore('" + sNO + "',DATEADD(MONTH, " + i.ToString() + ", convert(datetime,'" + sEndDatetime + "')),1,'" + iStartDay.ToString() + "',2)");
                if (dt != null && dt.Rows.Count > 0 && !string.IsNullOrEmpty(object2Str(dt.Rows[0][0])))
                {
                    pie2.depScore = (int)double.Parse(object2Str(dt.Rows[0][0]));
                }
                pie2.Contract_NO = sNO;
                DateTime d = DateTime.Parse(sEndDatetime).AddMonths(i-1);
                //if (d.Day<iStartDay)
                //{
                //    d = d.AddMonths(-1);
                //}
                pie2.Month = d.Year.ToString()+"."+d.Month.ToString("00");
                data.Add(pie2);
            }
            
            return data;
        }
        private List<clsFCDataPie> PieData1(string sNO, string sDateMonthStart, string sDateMonthEnd, string sType, bool isDis)
        {
            return PieData1(sNO,sDateMonthStart,sDateMonthEnd,sType,isDis,"");
        }
        private List<clsFCDataPie> PieData1(string sNO, string sDateMonthStart, string sDateMonthEnd,string sType,bool isDis,string sDep)
        {
            List<clsFCDataPie> data = new List<clsFCDataPie>();
            DataTable dt = null,dtAll=null;
            int iAllValue = 0,iTopValue=0;
            if (sType == "pie3" || sType == "userKPI")
            {
                string sContractNOSQL = (sType=="userKPI")?"":("Contract_No='" + sNO + "' AND");
                string sSum = (sType == "userKPI") ? "Deduction" : "Net_Value";
                dt = msSQL.GetDataTable("SELECT [Section],SUM(" + sSum + "),SUM(Deduction)/SUM(Quotation) from ContractorKPI0 where " + sContractNOSQL + " DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' and [Section] is not NULL GROUP BY [Section] ORDER BY [Section] ASC");
            }
            else if (sType == "DisDecKPI")
            {
                dt = msSQL.GetDataTable("SELECT [Dis],SUM(Deduction),SUM(Deduction)/SUM(Quotation) from ContractorKPI0 where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' and [Dis] is not NULL GROUP BY [Dis] ORDER BY [Dis] ASC");
            }
            else if (sType == "userSample")
            {
                dt = msSQL.GetDataTable("SELECT [Section],COUNT(SES),SUM(Net_Value) from ContractorKPI0 where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' GROUP BY [Section] ORDER BY [Section] ASC");
            }
            else if (sType == "contractorKPI")
            {
                //now getting top 20 values and all values
                dtAll = msSQL.GetDataTable("SELECT SUM(Net_Value) from ContractorKPI0 where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "'");
                if (dtAll != null && dtAll.Rows.Count == 1)
                {
                    iAllValue = (int)(double.Parse(dtAll.Rows[0][0].ToString()));
                }
                if (!isDis)
                {
                    dt = msSQL.GetDataTable("SELECT TOP 20 Contractor,SUM(Net_Value) from ContractorKPI0 where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' GROUP BY Contractor ORDER BY SUM(Tax_Value) DESC");
                }
                else
                {
                    dt = msSQL.GetDataTable("SELECT Dis,SUM(Net_Value) from ContractorKPI0 where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' GROUP BY Dis ORDER BY SUM(Tax_Value) DESC");
                }
            }
            else if (sType == "contractorKPI")
            {
                dt = msSQL.GetDataTable("SELECT Dis,SUM(Net_Value) from ContractorKPI0 where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' where Dis is not null GROUP BY Dis ORDER BY SUM(Tax_Value) DESC");
            }
            else if (sType == "userPro")
            {
                string sSQLSearch = "";
                sDep = sDep.Replace("false","All");
                if (!string.IsNullOrEmpty(sDep)&&sDep!="All")
                {
                    sSQLSearch = " AND [Section] LIKE '"+sDep+"%'";
                    if (sDep=="CTE")
                    {
                        sSQLSearch = " AND ([Section] LIKE 'CTE%' or [Section] LIKE 'PS%')";
                    }
                }
                dt = msSQL.GetDataTable("SELECT [Section],Dis,SUM(Net_Value) from ContractorKPI0 A where [Section] is not NULL AND DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' " + sSQLSearch + " GROUP BY Dis,[Section] ORDER BY [Section],SUM(Net_Value) DESC");
            }
            else
            {
                dt = msSQL.GetDataTable("SELECT Contract_No,Contractor,DateMonth,Deduction,Rate,MOM FROM ContractorKPI where Contract_No='" + sNO + "' and DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "'");
            }
            if (dt!=null&&dt.Rows.Count>0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clsFCDataPie clsMain = new clsFCDataPie();
                    if (sType == "pie3" || sType == "userKPI" || sType == "DisDecKPI")
                    {
                        clsMain.Name = dt.Rows[i][0].ToString().Trim().Replace("-", "");
                        clsMain.Data1 = double.Parse(dt.Rows[i][1].ToString().Trim());
                        clsMain.Data2 = (int)(100* double.Parse(dt.Rows[i][2].ToString().Trim()));
                        clsMain.MOM = "";// object2Str(dt.Rows[i][5]);
                        //clsMain.MOM = "MOM"+DateTime.Now.Ticks.ToString();  
                    }
                    else if(sType == "userSample")
                    {
                        clsMain.Name = dt.Rows[i][0].ToString().Trim().Replace("-", "");
                        if (string.IsNullOrEmpty(clsMain.Name))
                        {
                            clsMain.Name = "None";
                        }
                        clsMain.Data1 = double.Parse(dt.Rows[i][2].ToString().Trim());
                        clsMain.Data2 = int.Parse(dt.Rows[i][1].ToString().Trim());
                        clsMain.MOM = "";
                        iAllValue += (int)(double.Parse(dt.Rows[i][2].ToString().Trim()));
                        iTopValue += int.Parse(dt.Rows[i][1].ToString().Trim());
                    }
                    else if (sType == "contractorKPI")
                    {
                        clsMain.Name = dt.Rows[i][0].ToString().Trim().Replace("-", "");                        
                        clsMain.Data1 = (int)(double.Parse(dt.Rows[i][1].ToString().Trim())/1000);
                        if (isDis&&i<7)
                        {
                            iTopValue += (int)(double.Parse(dt.Rows[i][1].ToString().Trim())); 
                        }
                        if (!isDis)
                        {
                            iTopValue += (int)(double.Parse(dt.Rows[i][1].ToString().Trim())); 
                        }
                    }
                    else if (sType == "userPro")
                    {
                        clsMain.Name = dt.Rows[i][0].ToString().Trim().Replace("-", "");
                        clsMain.Data1 = double.Parse(dt.Rows[i][2].ToString().Trim());
                        //clsMain.Data2 = double.Parse(dt.Rows[i][3].ToString().Trim());
                        clsMain.MOM = dt.Rows[i][1].ToString().Trim();                        
                    }
                    else
                    {
                        clsMain.Name = dt.Rows[i][2].ToString().Trim().Replace("-", "");
                        clsMain.Data1 = double.Parse(dt.Rows[i][3].ToString().Trim());
                        clsMain.Data2 = (int)(100.0 * double.Parse(dt.Rows[i][4].ToString().Trim()));
                        clsMain.MOM = object2Str(dt.Rows[i][5]);
                        //clsMain.MOM = "MOM"+DateTime.Now.Ticks.ToString();  
                    }
                    if (sType == "contractorKPI")
                    {
                        if (i < 15 && !isDis)
                        {
                            data.Add(clsMain); 
                        }
                        if (isDis)
                        {
                            data.Add(clsMain); 
                        }
                    }
                    else
                    {
                        data.Add(clsMain); 
                    }
                }
                if (sType == "contractorKPI")
                {
                    if (!isDis)
                    {
                        data[0].MOM = "<center>Contractor Cost<br/>All  Contractor Cost: " + iAllValue.ToString() + " RMB/The TOP 20 Contractor Cost: " + iTopValue.ToString() + " RMB</center>";
                    }
                    else
                    {
                        data[0].MOM = "<center>Discipline Cost<br/>All  Discipline Cost: " + iAllValue.ToString() + " RMB/The TOP 7 Discipline Cost: " + iTopValue.ToString() + " RMB</center>"; 
                    }
                }
                if (sType == "userSample")
                {
                    clsFCDataPie clsTmp = new clsFCDataPie();
                    clsTmp = data[0];
                    data.RemoveAt(0);
                    data.Add(clsTmp);
                    data[0].MOM = "<center>User Sample Analyse<br/>SSR Cost: " + iAllValue.ToString() + " RMB/SSR Count: " + iTopValue.ToString() + "</center>";
                }
            }
            //DataTable dt = msSQL.GetDataTable("select CT_Director,SUM(cast(REPLACE(Net_Amount, ',', '') as FLOAT)) from CPTList where Contract_No='" + sNO + "' GROUP BY CT_Director");
            //if (dt != null && dt.Rows.Count > 0)
            
            //    for (int i = 0; i < dt.Rows.Count; i++)
            //    {
            //        clsFCDataPie clsMain = new clsFCDataPie();
            //        clsMain.Name = dt.Rows[i][0].ToString().Replace("\r\n", "");
            //        clsMain.Data1 = double.Parse(dt.Rows[i][1].ToString());
            //        data.Add(clsMain);
            //    }
            //}
            return data;
        }
        private List<clsFCDataPie> KPIData(string sNO, string sDateMonthStart, string sDateMonthEnd)
        {
            
            List<clsFCDataPie> data = new List<clsFCDataPie>();
            DataTable dt = null;
            if (string.IsNullOrEmpty(sNO)||sNO== "true")
            {
                dt = msSQL.GetDataTable("SELECT SUM(Deduction),SUM(Deduction)/SUM(Quotation),DateMonth from ContractorKPI where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' GROUP BY DateMonth ");
            }
            else if (sNO.Contains(","))
            {
                string sNOPath = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("/"), "contractNO.txt");
                File.WriteAllText(sNOPath, sNO);
                int iQuarter = DateTime.Now.Month/3;
                string sStartDate = "",sEndDate="";
                if (iQuarter == 0)
                {
                    sStartDate = (DateTime.Now.Year - 1).ToString() + "-10";
                    sEndDate = (DateTime.Now.Year - 1).ToString() + "-12";
                }
                else
                {
                    sStartDate = DateTime.Now.Year.ToString() + "-"+((iQuarter-1)*3+1).ToString();
                    sEndDate = DateTime.Now.Year.ToString() + "-"+ (iQuarter * 3).ToString();
                }
                string[] sNOs = sNO.Split(new char[] {','},StringSplitOptions.RemoveEmptyEntries);
                string sSQLNOs = "";
                for (int i = 0; i < sNOs.Length; i++)
                {
                    if (string.IsNullOrEmpty(sSQLNOs))
                    {
                        sSQLNOs += "Contract_No='" + sNOs[i].Trim() + "'";
                    }
                    else
                    {
                        sSQLNOs += " or Contract_No='" + sNOs[i].Trim() + "'";
                    }
                }
                sSQLNOs ="("+sSQLNOs.Trim() + ")";
                dt = msSQL.GetDataTable("SELECT SUM(Deduction),SUM(Deduction)/SUM(Quotation),Contract_No from ContractorKPI where " + sSQLNOs + " and DateMonth>='" + sStartDate + "' and DateMonth<='" + sEndDate + "' GROUP BY Contract_No ");
            }
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clsFCDataPie clsMain = new clsFCDataPie();
                    clsMain.Name = dt.Rows[i][2].ToString().Trim().Replace("-", "");
                    clsMain.Data1 = double.Parse(dt.Rows[i][0].ToString().Trim());
                    clsMain.Data2 = 100.0 * double.Parse(dt.Rows[i][1].ToString().Trim());
                    //clsMain.MOM = object2Str(dt.Rows[i][5]);
                    //clsMain.MOM = "MOM"+DateTime.Now.Ticks.ToString();
                    data.Add(clsMain);
                }
            }            
            return data;
        }
        private List<clsFCDataPie> KPIScatter(string sDateMonthStart, string sDateMonthEnd,bool isOverall)
        {

            List<clsFCDataPie> data = new List<clsFCDataPie>();
            DataTable dt = null;
            double dSumQuotation = 1.0,dSumDeduction=1.0,dAverageDeductionRate=0.0;
            //if (isOverall)
            {
                dt = msSQL.GetDataTable("SELECT SUM(Quotation) from ContractorKPI where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "'");
                if (dt!=null&&dt.Rows.Count==1&&!string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                {
                    dSumQuotation = double.Parse(dt.Rows[0][0].ToString());
                }
                dt = msSQL.GetDataTable("SELECT SUM(Deduction) from ContractorKPI where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "'");
                if (dt != null && dt.Rows.Count == 1&&!string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                {
                    dSumDeduction = double.Parse(dt.Rows[0][0].ToString());
                }
                dt = msSQL.GetDataTable("SELECT AVG(Deduction/Quotation) from ContractorKPI where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "'");
                if (dt != null && dt.Rows.Count == 1&&!string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                {
                    dAverageDeductionRate = 100.0*double.Parse(dt.Rows[0][0].ToString());
                }
            }
            dt = msSQL.GetDataTable("SELECT SUM(Deduction),SUM(Deduction)/SUM(Quotation),Contract_No,Contractor from ContractorKPI where DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' GROUP BY Contract_No,Contractor ");
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clsFCDataPie clsMain = new clsFCDataPie();
                    clsMain.Name = dt.Rows[i][2].ToString().Trim() + "<br/>" + dt.Rows[i][3].ToString().Trim();
                    clsMain.Data1 = isOverall ? (100.0*double.Parse(dt.Rows[i][0].ToString().Trim())/dSumQuotation) : double.Parse(dt.Rows[i][0].ToString().Trim());
                    clsMain.Data2 = 100.0 * double.Parse(dt.Rows[i][1].ToString().Trim());
                    if (i==0)
                    {
                        clsMain.MOM = "Overall Dedcution Cost:" + dSumDeduction.ToString("#.##") + ",Overall Deduction Rate:" + (100.0 * dSumDeduction / dSumQuotation).ToString("#.##") + "%,Average Deduction Rate:" + dAverageDeductionRate.ToString("#.##") + "%"; 
                    }
                    data.Add(clsMain);
                }
            }
            return data;
        }
        private List<clsFCDataPie> KPIScatter2(string sContractNO,string sDateMonthStart, string sDateMonthEnd)
        {
            List<clsFCDataPie> data = new List<clsFCDataPie>(); 
            if (!string.IsNullOrEmpty(sContractNO))
            {
                DataTable dt = null;
                dt = msSQL.GetDataTable("SELECT SUM(Quotation),SUM(Deduction)/SUM(Quotation),SES from ContractorKPI0 where Contract_No='"+sContractNO+"' and DateMonth>='" + sDateMonthStart + "' and DateMonth<='" + sDateMonthEnd + "' GROUP BY SES ");
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        clsFCDataPie clsMain = new clsFCDataPie();
                        clsMain.Name = dt.Rows[i][2].ToString().Trim();
                        clsMain.Data1 = 1*double.Parse(dt.Rows[i][0].ToString().Trim());
                        clsMain.Data2 = 100.00 * double.Parse(dt.Rows[i][1].ToString().Trim());
                        //clsMain.MOM = object2Str(dt.Rows[i][5]);
                        //clsMain.MOM = "MOM"+DateTime.Now.Ticks.ToString();
                        data.Add(clsMain);
                    }
                }
            }                       
            return data;
        }
        private List<clsFCSESReport> getEvaluation(string sNO, string sUser,bool isUser)
        {
            List<clsFCSESReport> data = new List<clsFCSESReport>();
            DataTable dt = msSQL.GetDataTable("select SES_No,Short_Descrption,Start_Date,End_Date,TECO_Date,TECO_Format,Requisitioner,SES_Confirmed_on,SES_CONF_Format,FO from FC_SESReport A where (TECO_Format is not null or SES_CONF_Format is not NULL) " + (isUser ? ("and UPPER(Requisitioner)='" + sUser.ToUpper() + "'") : "") + " and FO='" + sNO + "'" + sSQLRecent);
            //DataTable dt = msSQL.GetDataTable("select SES_No,Short_Descrption,Start_Date,End_Date,TECO_Date,TECO_Format,Requisitioner,SES_Confirmed_on,SES_CONF_Format,FO from FC_SESReport A where (TECO_Format is not null or SES_CONF_Format is not NULL) "+(isUser?("and UPPER(Requisitioner)='" + sUser.ToUpper() + "'"):"")+" and FO='" + sNO + "'"+sSQLRecent);
            if (dt != null && dt.Rows.Count > 0)
            {                
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clsFCSESReport clsMain = new clsFCSESReport();
                    clsMain.FO = sNO;
                    clsMain.SES_No = object2Str(dt.Rows[i][0]);
                    clsMain.Short_Descrption = object2Str(dt.Rows[i][1]);
                    clsMain.Start_Date = object2Str(dt.Rows[i][2]);
                    clsMain.End_Date = object2Str(dt.Rows[i][3]);
                    clsMain.TECO_Date = object2Str(dt.Rows[i][4]);
                    clsMain.TECO_Format = object2Date(dt.Rows[i][5]);
                    clsMain.Requisitioner = object2Str(dt.Rows[i][6]);
                    clsMain.SES_Confirmed_on = object2Str(dt.Rows[i][7]);
                    clsMain.SES_CONF_Format = object2Date(dt.Rows[i][8]);
                    data.Add(clsMain);
                }
            }            
            return data;
        }
        private List<clsFCEvaluation> getEvaluationContract(string sUser,bool isUser)
        {
            List<clsFCEvaluation> data = new List<clsFCEvaluation>();
            string sFCRecent = " and DateIn<'" + DateTime.Now.ToString("yyyy-MM-"+(iStartDay+1).ToString()) + "' and DateIn>='" + DateTime.Now.ToString("yyyy-MM-01") + "'";
            if (isUser)
            {
                DataTable dt = msSQL.GetDataTable("SELECT DISTINCT(FO) from FC_SESReport A where (SES_CONF_Format is not NULL or TECO_Format is not NULL) AND UPPER(Requisitioner)='" + sUser.ToUpper() + "'"+sSQLRecent);
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        clsFCEvaluation clsMain = new clsFCEvaluation();
                        clsMain.Contract_No = dt.Rows[i][0].ToString();
                        //now checking if this FO has been rated by this User
                        DataTable dtTmp = msSQL.GetDataTable("select Score1,Score2,Score3,Score4,Score5,Score6,Score7 from FC_Score where Contract_No='" + clsMain.Contract_No + "' and [By]='" + sUser + "' and Role='User' "+sFCRecent);
                        if (dtTmp != null && dtTmp.Rows.Count > 0)
                        {
                            clsMain.Score1 = double.Parse(dtTmp.Rows[0][0].ToString()) / 2.0;
                            clsMain.Score2 = double.Parse(dtTmp.Rows[0][1].ToString()) / 2.0;
                            clsMain.Score3 = double.Parse(dtTmp.Rows[0][2].ToString()) / 2.0;
                            clsMain.Score4 = double.Parse(dtTmp.Rows[0][3].ToString()) / 2.0;
                            clsMain.Score5 = double.Parse(dtTmp.Rows[0][4].ToString()) / 2.0;
                            clsMain.Score6 = double.Parse(dtTmp.Rows[0][5].ToString()) / 2.0;
                            clsMain.Score7 = double.Parse(dtTmp.Rows[0][6].ToString()) / 2.0;
                            //now for 0 points
                            if (clsMain.Score1<0) {clsMain.Score1 = 0; }
                            if (clsMain.Score2 < 0) { clsMain.Score2 = 0; }
                            if (clsMain.Score3 < 0) { clsMain.Score3 = 0; }
                            if (clsMain.Score4 < 0) { clsMain.Score4 = 0; }
                            if (clsMain.Score5 < 0) { clsMain.Score5 = 0; }
                            if (clsMain.Score6 < 0) { clsMain.Score6 = 0; }
                            if (clsMain.Score7 < 0) { clsMain.Score7 = 0; }
                            clsMain.Evaluated = true;
                        }
                        else
                        {
                            clsMain.Score1 = clsMain.Score2 = clsMain.Score3 = clsMain.Score4 = clsMain.Score5 = clsMain.Score6 = clsMain.Score7 = 0;
                            clsMain.Evaluated = false;
                        }
                        data.Add(clsMain);
                    }
                }
            }
            else
            { 
                //now reading form FC_Score
                DataTable dt = msSQL.GetDataTable("SELECT Contract_No,Score1,Score2,Score3,Score4,Score5,Score6 from FC_Score WHERE [By]='" + sUser + "' and Role='Dep' " + sFCRecent);//and DATEDIFF(DAY, DateIn, GETDATE())<30");
                if (dt != null && dt.Rows.Count > 0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        clsFCEvaluation clsMain = new clsFCEvaluation();
                        clsMain.Contract_No = dt.Rows[i][0].ToString();
                        clsMain.Score1 = double.Parse(dt.Rows[i][1].ToString()) * 2.0;
                        clsMain.Score2 = double.Parse(dt.Rows[i][2].ToString()) * 2.0;
                        clsMain.Score3 = double.Parse(dt.Rows[i][3].ToString()) * 2.0;
                        clsMain.Score4 = double.Parse(dt.Rows[i][4].ToString()) * 2.0;
                        clsMain.Score5 = double.Parse(dt.Rows[i][5].ToString()) * 2.0;
                        clsMain.Score6 = double.Parse(dt.Rows[i][6].ToString()) * 2.0;
                        clsMain.Evaluated = true;
                        //now for 0 points
                        if (clsMain.Score1 < 0) { clsMain.Score1 = -1; }
                        if (clsMain.Score2 < 0) { clsMain.Score2 = -1; }
                        if (clsMain.Score3 < 0) { clsMain.Score3 = -1; }
                        if (clsMain.Score4 < 0) { clsMain.Score4 = -1; }
                        if (clsMain.Score5 < 0) { clsMain.Score5 = -1; }
                        if (clsMain.Score6 < 0) { clsMain.Score6 = -1; }
                        if (clsMain.Score7 < 0) { clsMain.Score7 = -1; }
                        data.Add(clsMain);
                    }
                }
            }
            return data;
        }
        private List<clsContractFile> getContractFiles(string sNO)
        {
            List<clsContractFile> data = new List<clsContractFile>();
            DataTable dt = msSQL.GetDataTable("SELECT FileID,FileName,FileLength,DateIn,UploadUser,Contract_NO,Remark,FileType FROM FC_File where Contract_NO='"+sNO+"'");
            if (dt != null && dt.Rows.Count > 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    clsContractFile clsMain = new clsContractFile();
                    clsMain.FileID = object2Str(dt.Rows[i][0]);
                    clsMain.FileName = object2Str(dt.Rows[i][1]);
                    clsMain.FileLength = int.Parse(dt.Rows[i][2].ToString());                    
                    clsMain.DateIn = object2Date(dt.Rows[i][3]);
                    clsMain.UploadUser = object2Str(dt.Rows[i][4]);
                    clsMain.Contract_NO = object2Str(dt.Rows[i][5]);
                    clsMain.Remark = object2Str(dt.Rows[i][6]);
                    clsMain.Type = object2Str(dt.Rows[i][7]);
                    data.Add(clsMain);
                }
            }
            return data;
        }
        public string getDBName(string sClsName)
        {
            string[] sDBNames = new string[] { "Contract_Title", "Contractor", "FO_NO", "Actual_Budget", "Net_Amount", "Un-used_Budget", "Un-used_Percentage", "Remaining_Duration", "score", "Pricing_Scheme", "Contract_Admin" };
            string[] sClassNames = new string[] { "fcDescription", "contractor", "NO", "bugdet", "netAmount", "unusedBudget", "unusedPercentage", "remainingPercentage", "score", "pricingScheme", "contractAdmin" };
            for (int i = 0; i < sDBNames.Length; i++)
            {
                if (sClsName==sClassNames[i])
                {
                    return "["+sDBNames[i]+"]";
                }
            }
            return "[Un-used_Percentage]";
        }
        public static string object2Str(object dbObject)
        {
            return object2Str(dbObject, "");
        }
        public static string object2Str(object dbObject,string sNullReplace)
        {
            if (dbObject == null || string.IsNullOrEmpty(dbObject.ToString()))
            {
                return sNullReplace;
            }
            return dbObject.ToString();
        }
        public DateTime? object2Date(object dbObject)
        {
            if (dbObject == null)
            {
                return null;
            }
            if (dbObject.GetType().Name.ToLower().Contains("time"))
            {
                return (DateTime?)dbObject;
            }
            return null;
        }
        private float object2Float(object dbObject)
        {
            if (dbObject == null || string.IsNullOrEmpty(dbObject.ToString()))
            {
                return 0.0f;
            }
            return float.Parse(dbObject.ToString());
        }

        public static string FormatBytes(double dbytes)
        {
            const long scale = 1024;
            decimal ddbytes = (decimal)dbytes;
            string[] orders = new string[] { "PB", "TB", "GB", "MB", "KB", "Bytes" };
            var max = (long)Math.Pow(scale, (orders.Length - 1));
            foreach (string order in orders)
            {
                if (ddbytes > max)
                    return string.Format("{0:##.##} {1}", Decimal.Divide(ddbytes, max), order);
                max /= scale;
            }
            return "0 Bytes";
        }
        //convert 03.07.2013 to Datetime
        public static DateTime? getSAPDateTime(string sSAPDate)
        {
            if (!string.IsNullOrEmpty(sSAPDate) && sSAPDate.Length > 0)
            {
                string[] ss = sSAPDate.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                if (ss.Length == 3 && ss[0].Length == 2 & ss[1].Length == 2 && ss[2].Length == 4)
                {
                    return DateTime.Parse(ss[2] + "-" + ss[1] + "-" + ss[0]);
                }
            }
            return null;
        }
        /// <summary> 
        /// MD5 16位加密 
        /// </summary> 
        /// <param name="ConvertString">要加密的字符串</param> 
        /// <returns>返回加密后的字符串</returns> 
        public static string Md516(string ConvertString)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            string t2 = BitConverter.ToString(md5.ComputeHash(UTF8Encoding.Default.GetBytes(ConvertString)), 4, 8);
            t2 = t2.Replace("-", "");
            return t2;
        }
        /// <summary> 
        /// MD5　32位加密 
        /// </summary> 
        /// <param name="str">要加密的字符串</param> 
        /// <returns>返回加密后的字符串</returns> 
        public static string Md532(string str)
        {
            string cl = str;
            string pwd = "";
            MD5 md5 = MD5.Create();//实例化一个md5对像 
            // 加密后是一个字节类型的数组，这里要注意编码UTF8/Unicode等的选择　 
            byte[] s = md5.ComputeHash(Encoding.UTF8.GetBytes(cl));
            // 通过使用循环，将字节类型的数组转换为字符串，此字符串是常规字符格式化所得 
            for (int i = 0; i < s.Length; i++)
            {
                // 将得到的字符串使用十六进制类型格式。格式后的字符是小写的字母，如果使用大写（X）则格式后的字符是大写字符 

                pwd = pwd + s[i].ToString("X");

            }
            return pwd;
        }
        public static void LogFile(string sMsg, string sFileName)
        {
            try
            {
                string sPath = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("/"), sFileName);
                if (File.Exists(sPath))
                {
                    string sAll = File.ReadAllText(sPath);
                    if (sAll.Length > 60000)
                    {
                        //File.WriteAllText(sPath, "");
                        File.Delete(sPath);
                    }
                }
                File.AppendAllText(sPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ":" + sMsg + "\r\n"); 
            }
            catch (Exception)
            {
                //throw;
            }
        }

        private void exportExcel(HttpContext context)
        {
            context.Response.Clear();
            string sKeyword = context.Request.Params["searchKey"].Replace("null", "");
            string sNO=context.Request.Params["NO"].Replace("null","");
            string sCA = context.Request.Params["CA"].Replace("null", "");
            string sPS = context.Request.Params["PS"].Replace("null", "");
            string sValid = context.Request.Params["isValid"].Replace("null", "");
            string sSearchFilter = " where DATEDIFF(MONTH, GETDATE(), Expire_Date)>-12";
            if (!string.IsNullOrEmpty(sKeyword))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where ([FO_NO] LIKE N'%" + sKeyword + "%' OR [Actual_Budget] LIKE N'%" + sKeyword + "%' OR [Net_Amount] LIKE N'%" + sKeyword + "%' OR [Un-used_Budget] LIKE N'%" + sKeyword + "%' OR [Un-used_Percentage] LIKE N'%" + sKeyword + "%' OR [Contract_Title] LIKE N'%" + sKeyword + "%' OR [Contractor] LIKE N'%" + sKeyword + "%' OR [Pricing_Scheme] LIKE N'%" + sKeyword + "%' OR [Original_WC] LIKE N'%" + sKeyword + "%' OR [Type] LIKE N'%" + sKeyword + "%' OR [Contract_Admin] LIKE N'%" + sKeyword + "%' OR [Buyer] LIKE N'%" + sKeyword + "%' OR [Main_Coordinator] LIKE N'%" + sKeyword + "%' OR [User_Representative] LIKE N'%" + sKeyword + "%' OR [Validate_Date] LIKE N'%" + sKeyword + "%' OR [Expire_Date] LIKE N'%" + sKeyword + "%' OR [Contract_Person] LIKE N'%" + sKeyword + "%' OR [Contract_Tel] LIKE N'%" + sKeyword + "%' OR [Score] LIKE N'%" + sKeyword + "%' )";
                }
                else
                {
                    sSearchFilter += " and ([FO_NO] LIKE N'%" + sKeyword + "%' OR [Actual_Budget] LIKE N'%" + sKeyword + "%' OR [Net_Amount] LIKE N'%" + sKeyword + "%' OR [Un-used_Budget] LIKE N'%" + sKeyword + "%' OR [Un-used_Percentage] LIKE N'%" + sKeyword + "%' OR [Contract_Title] LIKE N'%" + sKeyword + "%' OR [Contractor] LIKE N'%" + sKeyword + "%' OR [Pricing_Scheme] LIKE N'%" + sKeyword + "%' OR [Original_WC] LIKE N'%" + sKeyword + "%' OR [Type] LIKE N'%" + sKeyword + "%' OR [Contract_Admin] LIKE N'%" + sKeyword + "%' OR [Buyer] LIKE N'%" + sKeyword + "%' OR [Main_Coordinator] LIKE N'%" + sKeyword + "%' OR [User_Representative] LIKE N'%" + sKeyword + "%' OR [Validate_Date] LIKE N'%" + sKeyword + "%' OR [Expire_Date] LIKE N'%" + sKeyword + "%' OR [Contract_Person] LIKE N'%" + sKeyword + "%' OR [Contract_Tel] LIKE N'%" + sKeyword + "%' OR [Score] LIKE N'%" + sKeyword + "%' )";
                }
            }
            if (!string.IsNullOrEmpty(sNO))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where FO_NO like '%" + sNO + "%'";
                }
                else
                {
                    sSearchFilter += " and FO_NO like '%" + sNO + "%'";
                }
            }
            if (!string.IsNullOrEmpty(sCA))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where Contract_Admin like '%" + sCA + "%'";
                }
                else
                {
                    sSearchFilter += " and Contract_Admin like '%" + sCA + "%'";
                }
            }
            if (!string.IsNullOrEmpty(sPS))
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where Pricing_Scheme like '%" + sPS + "%'";
                }
                else
                {
                    sSearchFilter += " and Pricing_Scheme like '%" + sPS + "%'";
                }
            }
            if (!string.IsNullOrEmpty(sValid) && sValid.ToLower() == "true")
            {
                if (string.IsNullOrEmpty(sSearchFilter))
                {
                    sSearchFilter += " where Expire_Date>getDate()";
                }
                else
                {
                    sSearchFilter += " and Expire_Date>getDate()";
                }
            }
            DataTable dt = msSQL.GetDataTable("select FO_NO,Contract_Title,Contractor,Contract_Admin,Pricing_Scheme,Actual_Budget,Net_Amount,[Un-used_Budget],[Un-used_Percentage],Validate_Date,Expire_Date from FCMain" + sSearchFilter);
            if (dt!=null&&dt.Rows.Count>0)
            {
                //now exporting
                HSSFWorkbook workbook = new HSSFWorkbook();
                MemoryStream ms = new MemoryStream();
                HSSFSheet sheet = workbook.CreateSheet("Contract List") as HSSFSheet;

                HSSFCellStyle style1 = workbook.CreateCellStyle() as HSSFCellStyle;
                style1.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.GREY_40_PERCENT.index;                
                HSSFFont font1 = (HSSFFont)workbook.CreateFont();
                font1.Color = NPOI.HSSF.Util.HSSFColor.BLACK.index;
                font1.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.BOLD;
                font1.FontHeightInPoints = 25;
                style1.SetFont(font1);

                IRow row = sheet.CreateRow(0);
                row.RowStyle = style1;
                sheet.SetColumnWidth(1, 100 * 256);
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    if (i == 1 || i == 2)
                    {
                        sheet.SetColumnWidth(i, 60 * 256);
                    }
                    else
                    {
                        sheet.SetColumnWidth(i, 10 * 256);
                    }
                    row.CreateCell(i, CellType.STRING).SetCellValue(dt.Columns[i].ColumnName.Replace("_"," "));                    
                }
                
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //sheet.CreateRow(0).CreateCell(0).SetCellValue("0"); 
                    row=sheet.CreateRow(i+1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row.CreateCell(j).SetCellValue(object2Str(dt.Rows[i][j]));
                    }                    
                }
                workbook.Write(ms);
                context.Response.ContentType = "application/octet-stream";
                context.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=ContractList_"+DateTime.Now.ToString("yyyy-MM-dd")+".xls"));
                context.Response.BinaryWrite(ms.ToArray());                 
                workbook = null; 
                ms.Close(); 
                ms.Dispose();
            }

        }
        public static void exportReportExcel(HttpContext context,DateTime dtNow)
        {
            //added 2014-11-12 for progress show
            string sProgressPath = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("/"), "progress.txt");
            //if (File.Exists(sProgressPath))
            //{
            //    File.Delete(sProgressPath);
            //}
            File.WriteAllText(sProgressPath,"0");            

            context.Response.Clear();

            //FileStream file = new FileStream(strTemplateFileName, FileMode.Open, FileAccess.Read);//读入excel模板
            //HSSFWorkbook workbook = new HSSFWorkbook(file);

            string sSQLRecent2 = " AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND ( B.DateIn<'" + dtNow.ToString("yyyy-MM-01") + "' OR  B.DateIn>='" + dtNow.ToString("yyyy-MM-" + (iStartDay + 1).ToString()) + "') )";
            string sSQL = "SELECT FO_NO,Contract_Title,Contractor,Contract_Admin,Main_Coordinator,User_Representative from FC_SESRelatedData where FO_NO IN(SELECT DISTINCT(FO) from FC_SESReport A where A.Requisitioner is not null  AND (Deleted is NULL and Blocked is NULL)  and (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + sSQLRecent2 + ")";
            DataTable dt = msSQL.GetDataTable(sSQL);
            //now exporting
            byte[] bs = FCPortal.Properties.Resources.Template_Details;
            MemoryStream msFile = new MemoryStream(bs);
            IWorkbook workbook = NPOI.SS.UserModel.WorkbookFactory.Create(msFile);
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.GetSheet("Overview");
            if (dt!=null&&dt.Rows.Count>0)
            {               
                ICellStyle style1 = workbook.CreateCellStyle();
                style1.Alignment = HorizontalAlignment.CENTER;
                style1.VerticalAlignment = VerticalAlignment.CENTER;
                //style1.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");  format 
                int iCellIndex = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    File.WriteAllText(sProgressPath, (i * 20 / dt.Rows.Count).ToString());
                    iCellIndex = 0;
                    string sFO=dt.Rows[i][0].ToString();
                    string sMainCoordinator = dt.Rows[i][4].ToString();
                    IRow row = sheet.CreateRow(i+2);                    
                    row.CreateCell(0).SetCellValue(sFO);
                    row.CreateCell(1).SetCellValue(refineContractName(dt.Rows[i][1].ToString()));
                    row.CreateCell(2).SetCellValue(refineContractName(dt.Rows[i][2].ToString()));
                    row.CreateCell(3).SetCellValue(sMainCoordinator);
                    row.CreateCell(5).SetCellFormula("(H"+(i+3).ToString()+"+L"+(i+3).ToString()+"+P"+(i+3).ToString()+"+T"+(i+3).ToString()+"+X"+(i+3).ToString()+"+AB"+(i+3).ToString()+")/(J"+(i+3).ToString()+"+N"+(i+3).ToString()+"+R"+(i+3).ToString()+"+V"+(i+3).ToString()+"+Z"+(i+3).ToString()+"+AD"+(i+3).ToString()+")");
                    //now getting the Total Score
                    
                    DataTable dtTmp = msSQL.GetDataTable("EXEC getScoreDetails2 '"+sFO+"',0,6,0,'" + dtNow.ToString("yyyy-MM-dd")+"'");
                    if (dtTmp!=null&&dtTmp.Rows.Count>0)
                    {
                        //row.CreateCell(3).SetCellValue(object2Str(dtTmp.Rows[0][0]));
                        string sValue = object2Str(dtTmp.Rows[0][0]);
                        if (!string.IsNullOrEmpty(sValue))
                        {
                            row.CreateCell(4).SetCellValue(double.Parse(sValue));
                        }
                        else
                        {
                            row.CreateCell(4).SetCellValue(sValue);
                        }
                    }
                    //now for Score1 to Score6
                    for (int j = 0; j < 6; j++)
                    {
                        for (int x = 0; x < 4; x++)
                        {
                            iCellIndex = (j+1)* 4+x+2;
                            //now getting separate Score
                            dtTmp = msSQL.GetDataTable("EXEC getScoreDetails2 '" + sFO + "',"+(j+1).ToString()+","+x.ToString()+",1,'" + dtNow.ToString("yyyy-MM-dd")+"'");
                            if (dtTmp != null && dtTmp.Rows.Count > 0)
                            {
                                ICell cell = row.CreateCell(iCellIndex);
                                string sValue = object2Str(dtTmp.Rows[0][0]);
                                if (!string.IsNullOrEmpty(sValue))
                                {
                                    cell.SetCellValue(double.Parse(sValue));
                                }
                                else
                                {
                                    cell.SetCellValue(sValue);    
                                }
                                
                                cell.CellStyle = style1;
                            }
                        }
                    }
                    //now for Timely and honesty
                    for (int j = 4; j < 6; j++)
                    {
                        iCellIndex += 1;
                        dtTmp = msSQL.GetDataTable("EXEC getScoreDetails2 '"+sFO+"',0,"+j.ToString()+",1,'" + dtNow.ToString("yyyy-MM-dd")+"'");
                        if (dtTmp != null && dtTmp.Rows.Count > 0)
                        {
                            ICell cell = row.CreateCell(iCellIndex);
                            string sValue = object2Str(dtTmp.Rows[0][0]);
                            if (!string.IsNullOrEmpty(sValue))
                            {
                                cell.SetCellValue(double.Parse(sValue));
                            }
                            else
                            {
                                cell.SetCellValue(sValue);
                            }
                            cell.CellStyle = style1;
                        }
                    }
                    //for Dep AVG
                    for (int j = 1; j < 7; j++)
                    {
                        iCellIndex += 1;
                        dtTmp = msSQL.GetDataTable("EXEC getScoreDetails2 '" + sFO + "'," + j.ToString() + ",0,0,'" + dtNow.ToString("yyyy-MM-dd")+"'");
                        if (dtTmp != null && dtTmp.Rows.Count > 0)
                        {
                            ICell cell = row.CreateCell(iCellIndex);
                            string sValue = object2Str(dtTmp.Rows[0][0]);
                            if (!string.IsNullOrEmpty(sValue))
                            {
                                cell.SetCellValue(double.Parse(sValue));
                            }
                            else
                            {
                                cell.SetCellValue(sValue);
                            }
                            cell.CellStyle = style1;
                        }
                    }
                }
                //now for extreme scores
                sheet = workbook.GetSheet("Extreme Scores");
                iCellIndex = 1;
                //                                    0        1      2      3      4      5       6     7     8     9       10           11        12    
                dt = msSQL.GetDataTable("select Contract_No,Score1,Score2,Score3,Score4,Score5,Score6,Score7,[By],Role,Contract_Title,Contractor,Contract_Admin from extremeScores where DateIn>='" + dtNow.ToString("yyyy-MM-01") + "' AND DateIn<'" + dtNow.AddMonths(1).ToString("yyyy-MM-01") + "' order by Contract_No ASC");
                if (dt!=null&&dt.Rows.Count>0)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        File.WriteAllText(sProgressPath, (20+i * 20 / dt.Rows.Count).ToString());
                        string sRole = object2Str(dt.Rows[i][9]);
                        iCellIndex += 1;
                        IRow row = sheet.CreateRow(iCellIndex);
                        ICell cell = row.CreateCell(0);
                        cell.SetCellValue(object2Str(dt.Rows[i][0]));
                        cell = row.CreateCell(1);
                        cell.SetCellValue(object2Str(dt.Rows[i][11]));
                        cell = row.CreateCell(2);
                        cell.SetCellValue(object2Str(dt.Rows[i][10]));
                        cell = row.CreateCell(3);
                        cell.SetCellValue(object2Str(dt.Rows[i][8]));
                        cell = row.CreateCell(4);
                        cell.SetCellValue(object2Str(dt.Rows[i][12]));
                        if (sRole == "User")
                        {
                            for (int j = 0; j < 7; j++)
                            {
                                string sValueTmp = object2Str(dt.Rows[i][j + 1]);
                                if (!string.IsNullOrEmpty(sValueTmp))
                                {
                                    double d = double.Parse(sValueTmp);
                                    if (d<=2||d==10)
                                    {
                                        if (j==6&&d==10)
                                        {
                                            continue;
                                        }
                                        cell=row.CreateCell(5 + j);
                                        cell.SetCellValue(d);
                                    }
                                }
                                
                            }
                        }
                        else
                        {
                            for (int j = 0; j < 6; j++)
                            {
                                string sValueTmp = object2Str(dt.Rows[i][j + 1]);
                                if (!string.IsNullOrEmpty(sValueTmp))
                                {
                                    double d = double.Parse(sValueTmp);
                                    if (d<=1 || d == 5)
                                    {
                                        cell=row.CreateCell(12 + j);
                                        cell.SetCellValue(d);
                                    }
                                }

                            }
                        }

                        //now getting attachments
                        DataTable dtFile = msSQL.GetDataTable("SELECT FileID,FileName from FC_File WHERE DateIn>='" + dtNow.ToString("yyyy-MM-01") + "' AND DateIn<'" + dtNow.AddMonths(1).ToString("yyyy-MM-01") + "' AND UploadUser='" + object2Str(dt.Rows[i][8]) + "'");
                        if (dtFile != null)
                        {
                            for (int n = 0; n < dtFile.Rows.Count; n++)
                            {
                                //object2Str(dtFile.Rows[i][0])
                                cell = row.CreateCell(n + 18);
                                cell.SetCellValue(object2Str(dtFile.Rows[n][1]));
                                //cell.CellType = CellType.NUMERIC;

                                HSSFHyperlink link = new HSSFHyperlink(HyperlinkType.URL);
                                link.Address = "http://" + context.Request.Url.Host + context.Request.ApplicationPath + "/data.ashx?fileID=" + object2Str(dtFile.Rows[n][0]);
                                cell.Hyperlink = link;

                            }
                        }
                        
                    }
                }
                //now for User Evaluation Status
                //first of all, for Users
                DataTable dtUsers = msSQL.GetDataTable("SELECT [By],Contract_No from FC_SESRecord A where DateIn>='" + dtNow.ToString("yyyy-MM-01") + "' AND DateIn<'" + dtNow.ToString("yyyy-MM-" + (iStartDay + 1).ToString()) + "' AND Remark='Auto' AND not EXISTS (SELECT * FROM FC_SESRecord B WHERE A.[By]=B.[By] AND DateIn>='" + dtNow.ToString("yyyy-MM-01") + "' AND DateIn<'" + dtNow.ToString("yyyy-MM-" + (iStartDay + 1).ToString()) + "' AND Remark!='Auto')");
                ArrayList alUser = new ArrayList();
                if (dtUsers!=null&&dtUsers.Rows.Count>0)
                {
                    //added on 2014-07-21 for cotract NO display
                    ArrayList alUserTmp = new ArrayList();
                    for (int x = 0; x < dtUsers.Rows.Count; x++)
                    {
                        string sUserNow = object2Str(dtUsers.Rows[x][0]);
                        if (!string.IsNullOrEmpty(sUserNow))
                        {
                            if (!alUserTmp.Contains(sUserNow))
                            {
                                alUserTmp.Add(sUserNow);
                            }
                        }                       
                    }
                    File.WriteAllText(sProgressPath, "45");
                    sheet = workbook.GetSheet("NO Evaluation Users");
                    iCellIndex = 0;
                    for (int x = 0; x < alUserTmp.Count; x++)
                    {
                        string sUserNow = alUserTmp[x].ToString();
                        if (!string.IsNullOrEmpty(sUserNow))
                        {
                            string sFullName = getUserDetails(sUserNow,0);
                            
                            if (!string.IsNullOrEmpty(sFullName)&&!alUser.Contains(sUserNow))
                            {
                                alUser.Add(sUserNow);
                                iCellIndex += 1;
                                IRow row = sheet.CreateRow(iCellIndex);
                                ICell cell = row.CreateCell(0);
                                cell.SetCellValue(sFullName);
                                cell = row.CreateCell(1);
                                cell.SetCellValue(sUserNow);
                                cell = row.CreateCell(2);
                                cell.SetCellValue("否");
                                cell = row.CreateCell(3);
                                cell.SetCellValue("User");
                                cell = row.CreateCell(4);
                                cell.SetCellValue(getUserDetails(sUserNow,2));
                                //added on 2014-07-21 for cotract NO display,now display Contract NOs
                                int iCol = 0;
                                ArrayList alContractTmp = new ArrayList();
                                for (int y = 0; y < dtUsers.Rows.Count; y++)
                                {
                                    string sUserTmp = object2Str(dtUsers.Rows[y][0]);
                                    if (sUserTmp==sUserNow)
                                    {
                                        string sContractNow = object2Str(dtUsers.Rows[y][1]);
                                        if (!string.IsNullOrEmpty(sContractNow) && !alContractTmp.Contains(sContractNow))
                                        {
                                            alContractTmp.Add(sContractNow);
                                            cell = row.CreateCell(iCol+5);
                                            cell.SetCellValue(sContractNow);
                                            iCol += 1;
                                        } 
                                    }
                                }
                                iCol = 0; alContractTmp = new ArrayList();
                            }
                        }
                    }
                    File.WriteAllText(sProgressPath, "50");
                    alUser = new ArrayList();
                    //now for Dep
                    for (int i = 0; i < 6; i++)
                    {
                        File.WriteAllText(sProgressPath, (50+i*10/6).ToString());
                        string[] sReturns = getEmails(i);
                        if (sReturns!=null&&sReturns.Length>0)
                        {
                            foreach (string sReturn in sReturns)
                            {
                                string[] ss = sReturn.Split(new string[] { "[@]" }, StringSplitOptions.None);
                                if (ss.Length==2)
                                {
                                    DataTable dtTmp = msSQL.GetDataTable("select ID from FC_Score where [By]='" + ss[0] + "'  AND Role='Dep'  and DateIn>='" + dtNow.ToString("yyyy-MM-01") + "' AND DateIn<'" + dtNow.ToString("yyyy-MM-" + (iStartDay + 1).ToString()) + "'");
                                    if (dtTmp!=null&&dtTmp.Rows.Count==0)
                                    {
                                        string sFullName = getUserDetails(ss[0], 0);

                                        if (!string.IsNullOrEmpty(sFullName)&&!alUser.Contains(ss[0]))
                                        {
                                            alUser.Add(ss[0]);
                                            iCellIndex += 1;
                                            IRow row = sheet.CreateRow(iCellIndex);
                                            ICell cell = row.CreateCell(0);
                                            cell.SetCellValue(sFullName);
                                            cell = row.CreateCell(1);
                                            cell.SetCellValue(ss[0]);
                                            cell = row.CreateCell(2);
                                            cell.SetCellValue("否");
                                            cell = row.CreateCell(3);
                                            cell.SetCellValue("Dep");
                                            cell = row.CreateCell(4);
                                            cell.SetCellValue(getUserDetails(ss[0], 2));
                                            //added on 2014-07-24 for cotract NO display,now display Contract NOs
                                            string sContracts = getContracts(i, ss[0]);
                                            if (!string.IsNullOrEmpty(sContracts))
                                            {
                                                string[] sCons = sContracts.Split(',');
                                                for (int z = 0; z < sCons.Length; z++)
                                                {
                                                    cell = row.CreateCell(5+z);
                                                    cell.SetCellValue(sCons[z]);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                //now for up to date scores
                sheet = workbook.GetSheet("Year_to_date Report");
                iCellIndex = 1;
                //DataTable dtAvg = msSQL.GetDataTable("SELECT FO_NO,Contract_Title,Contractor,ScoreAll,ScoreUser1,ScoreUser2,ScoreUser3,ScoreUser4,ScoreUser5,ScoreUser6,ScoreUser7,ScoreDep1,ScoreDep2,ScoreDep3,ScoreDep4,ScoreDep5,ScoreDep6 from getDetails('" + (DateTime.Now.Year-1).ToString() + "-01-01'," + (DateTime.Now.Month+12).ToString() + ") WHERE ScoreUser1 is not NULL");
                //if (dtAvg!=null&&dtAvg.Rows.Count>0)
                //{
                //    for (int i = 0; i < dtAvg.Rows.Count; i++)
                //    {
                //        IRow row = sheet.CreateRow(i+1);
                //        for (int j = 0; j < 17; j++)
                //        {
                //            ICell cell = row.CreateCell(j);
                //            string sValue = object2Str(dtAvg.Rows[i][j]);
                //            if (!string.IsNullOrEmpty(sValue)&&j>=3)
                //            {
                //                cell.SetCellValue(double.Parse(sValue));
                //            }
                //            else
                //            {
                //                cell.SetCellValue(sValue);
                //            }
                //        }
                //    }                    
                //}
                
                ////new year to date report
                //start date 2013-08-01
                DateTime dtStart = DateTime.Now.AddMonths(-24);
                DateTime dtSysStart=DateTime.Parse("2013-09-01");
                if ((dtSysStart-dtStart).TotalDays>0)
                {
                    dtStart = dtSysStart;
                }
                DataTable dtAllContracts = msSQL.GetDataTable("select distinct(Contract_No) from FC_Score");
                if (dtAllContracts!=null&&dtAllContracts.Rows.Count>0)
                {
                    for (int i = 0; i < 24; i++)
                    {
                        File.WriteAllText(sProgressPath, (65+i).ToString());
                        if ((dtStart.AddMonths(i)-DateTime.Now).TotalDays>0)
                        {
                            break;
                        }
                        //now marking the time
                        IRow row = sheet.GetRow(0);
                        ICell cell = row.CreateCell(5+i);
                        cell.SetCellValue(dtStart.AddMonths(i - 1).ToString("yyyy-MM"));
                        for (int j = 0; j < dtAllContracts.Rows.Count; j++)
                        {
                            DataTable dtAvg = msSQL.GetDataTable("SELECT FO_NO,Contract_Title,Contractor,Contract_Admin,Main_Coordinator,dbo.getScore(FO_NO,'" + (dtStart.AddMonths(i)).ToString("yyyy-MM-01") + "',1,'11',3) as 'ScoreAll' from FCMain where FO_NO='" + object2Str(dtAllContracts.Rows[j][0]) + "'");
                            if (i == 0)
                            {
                                row = sheet.CreateRow(j + 1);
                                cell = row.CreateCell(0);
                                cell.SetCellValue(object2Str(dtAllContracts.Rows[j][0]));
                                if (dtAvg!=null&&dtAvg.Rows.Count>0)
                                {
                                    cell = row.CreateCell(1);
                                    cell.SetCellValue(data.refineContractName(object2Str(dtAvg.Rows[0][1])));
                                    cell = row.CreateCell(2);
                                    cell.SetCellValue(data.refineContractName(object2Str(dtAvg.Rows[0][2])));
                                    cell = row.CreateCell(3);
                                    cell.SetCellValue(object2Str(dtAvg.Rows[0][3]));
                                    cell = row.CreateCell(4);
                                    cell.SetCellValue(object2Str(dtAvg.Rows[0][4])); 
                                }
                            }
                            if (dtAvg != null && dtAvg.Rows.Count > 0)
                            {
                                row = sheet.GetRow(j + 1);
                                cell = row.CreateCell(5 + i);
                                string sValue= object2Str(dtAvg.Rows[0][5],"-");
                                if (!string.IsNullOrEmpty(sValue) && sValue != "-")
                                {
                                    cell.SetCellValue(double.Parse(sValue));
                                }
                                else
                                {
                                    cell.SetCellValue(sValue);
                                }
                            }
                            else if (dtAvg != null && dtAvg.Rows.Count == 0)
                            {
                                dtAvg = msSQL.GetDataTable("SELECT dbo.getScore('" + object2Str(dtAllContracts.Rows[j][0]) + "','" + (dtStart.AddMonths(i)).ToString("yyyy-MM-01") + "',1,'11',3) as 'ScoreAll'");
                                if (dtAvg != null && dtAvg.Rows.Count > 0)
                                {
                                    row = sheet.GetRow(j + 1);
                                    cell = row.CreateCell(5 + i);
                                    string sValue = object2Str(dtAvg.Rows[0][0], "-");
                                    if (!string.IsNullOrEmpty(sValue) && sValue != "-")
                                    {
                                        cell.SetCellValue(double.Parse(sValue));
                                    }
                                    else
                                    {
                                        cell.SetCellValue(sValue);
                                    }
                                }
                            }
                        }

                    } 
                }

            }
            File.WriteAllText(sProgressPath, "100");
            context.Session["progress"] = 100;
            workbook.Write(ms);
            context.Response.ContentType = "application/octet-stream";
            context.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=EvaluationReport_" + dtNow.ToString("yyyy-MM-dd") + ".xls"));
            context.Response.BinaryWrite(ms.ToArray());
            workbook = null;
            msFile.Close();
            ms.Close();
            ms.Dispose();

        }

        public static string refineContractName(string sContractName)
        {
            try
            {
                //get the first Chinese word
                sContractName = sContractName.Replace("\r", "").Replace("\n", "").Replace("\t", "").Trim();
                int iPos = -1, iPosEn = -1;
                for (int i = 0; i < sContractName.Length; i++)
                {
                    byte[] byte_len = System.Text.Encoding.Default.GetBytes(sContractName.Substring(i, 1));
                    if (byte_len.Length > 1)
                    {
                        if (iPos < 0)
                        {
                            iPos = i;
                        }
                    }
                    else if (iPosEn < 0)
                    {
                        iPosEn = i;
                    }
                }
                if (iPos >= sContractName.Length || iPos < 0)
                {
                    return sContractName.Trim();
                }
                else
                {
                    string sChinese = "", sEnglish = "";
                    if (iPos > 5)
                    {
                        sChinese = sContractName.Substring(iPos, sContractName.Length - iPos).Trim();
                    }
                    if (iPos <= 2)
                    {
                        sChinese = sContractName.Substring(0, iPosEn).Trim();
                    }
                    sEnglish = sContractName.Replace(sChinese, "").Trim();
                    return sChinese + " " + sEnglish;
                }
            }
            catch (Exception)
            {
                //throw;
            }
            return sContractName;
        }
        public static void exportTemplate(HttpContext context, string sDataSheetPath)
        {
            context.Response.Clear();
            string sSQL = "SELECT FO_NO,Contract_Title,Contractor,Contract_Admin,Main_Coordinator,User_Representative from FC_SESRelatedData where FO_NO IN(SELECT DISTINCT(FO) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + sSQLRecent + ")";
            DataTable dt = msSQL.GetDataTable(sSQL);
            if (dt!=null&&dt.Rows.Count>0)
            {
                //now exporting
                FileStream file = new FileStream(sDataSheetPath, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = NPOI.SS.UserModel.WorkbookFactory.Create(file);
                MemoryStream ms = new MemoryStream();
                ISheet sheet = workbook.GetSheetAt(0);                              

                //IRow row = sheet.GetRow(1);
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    //sheet.CreateRow(0).CreateCell(0).SetCellValue("0"); 
                    IRow row = sheet.CreateRow(i + 1);
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        row.CreateCell(j).SetCellValue(object2Str(dt.Rows[i][j]));
                    }
                }
                workbook.Write(ms);
                context.Response.ContentType = "application/octet-stream";
                context.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=Evaluation_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx"));
                context.Response.BinaryWrite(ms.ToArray());
                workbook = null;
                ms.Close();
                ms.Dispose();
                context.Response.End();

            }
        }
        public static string getUserDetails(string sUserID, int iType)//0 full name;1 email;2 Department
        {
            string sAllData = FCPortal.Properties.Resources.userDetails;
            string sReturn = "";
            string[] sLines = sAllData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            for (int j = 0; j < sLines.Length; j++)
            {
                string[] ss = sLines[j].Split(new string[] { "[@]" }, StringSplitOptions.None);
                if (ss.Length >= 4)
                {
                    if (ss[0].Trim().ToLower() == sUserID.Trim().ToLower())
                    {
                        if (iType == 0)
                        {
                            sReturn = ss[2].Split(new char[] { '-' })[0] + " " + ss[1].Split(new char[] { '-' })[0];
                        }
                        else if (iType == 1)
                        {
                            sReturn = ss[3];
                        }
                        else
                        {
                            if (ss.Length==5)
                            {
                                return ss[4];
                            }
                            return "";
                        }
                        return sReturn;
                        break;
                    }
                }
            }
            return sReturn;
        }
        public static string[] getEmails(int iType)
        {
            try
            {
                ArrayList al = new ArrayList();
                //FO_NO,Contract_Title,Contractor,Contract_Admin,Main_Coordinator,User_Representative
                string sDBName = "Contract_Admin";
                if (iType==0)//cte/d
                {
                    sDBName = "Contract_Admin";
                }
                if (iType == 1)
                {
                    sDBName = "Main_Coordinator";
                }
                if (iType == 2)
                {
                    sDBName = "User_Representative";
                }
                if (iType == 3)//CTS
                {
                    return new string[] { "liuj31[@]JinLing.Liu@basf-ypc.com.cn" };
                }
                if (iType == 4)//CHA
                {
                    return new string[] { "zhengyj2[@]zhengyj@basf-ypc.com.cn" };
                }
                if (iType == 5)//CTM/T
                {
                    return new string[] { "linm[@]linm@basf-ypc.com.cn", "zhouhx2[@]zhouhx@basf-ypc.com.cn" };
                }
                string sSQL = "SELECT DISTINCT(" + sDBName + ") from FC_SESRelatedData where FO_NO IN(SELECT DISTINCT(FO) from FC_SESReport A where A.Requisitioner is not NULL and (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + sSQLRecent + ")";
                DataTable dt = msSQL.GetDataTable(sSQL);
                if (dt!=null&&dt.Rows.Count>0)
                {
                    //now getting emails
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i]!=null&&!string.IsNullOrEmpty(dt.Rows[i][0].ToString()))
                        {
                            string sUserName = dt.Rows[i][0].ToString();
                            //BLJ=Bao Linjin YangJ=Yang Jing
                            if (sUserName.ToLower()=="blj")
                            {
                                sUserName = "Sun Yan";//"Bao Lijin";
                            }
                            if (sUserName.ToLower() == "suny2" || sUserName.ToLower() == "suny" || sUserName.ToLower() == "sun yan" || sUserName.ToLower() == "sy")
                            {
                                sUserName = "Sun Yan";
                            }
                            if (sUserName.ToLower() == "yangj")
                            {
                                sUserName = "Yang Jing";
                            }
                            if (sUserName.ToLower() == "wang jin")
                            {
                                sUserName = "Wang Jin-Ade";
                            }
                            string sAllData = FCPortal.Properties.Resources.userDetails;
                            string[] sLines = sAllData.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
                            for (int j = 0; j < sLines.Length; j++)
                            {
                                string[] ss = sLines[j].Split(new string[] { "[@]" }, StringSplitOptions.None);
                                if (ss.Length >= 4)
                                {
                                    if ((ss[2] + ss[1]).ToLower().Replace(" ","").Contains(sUserName.ToLower().Replace(" ","")))
                                    {
                                        al.Add(ss[0] + "[@]" + ss[3]);
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (al.Count>0)
                    {
                        string[] ss=new string[al.Count];
                        al.CopyTo(ss);
                        return ss;
                    }
                }
            }
            catch (Exception)
            {

                //throw;
            }
            return null;
        }
        public static string getContracts(int iType,string sUserID)
        {
            try
            {
                string sDBName = "Contract_Admin";
                if (iType == 0)//cte/d
                {
                    sDBName = "Contract_Admin";
                }
                if (iType == 1)
                {
                    sDBName = "Main_Coordinator";
                }
                if (iType == 2)
                {
                    sDBName = "User_Representative";
                }

                string sSQL = "SELECT DISTINCT(FO_NO) from FC_SESRelatedData where LOWER(" + sDBName + ")='" + getUserDetails(sUserID, 0).ToLower() + "' and FO_NO IN(SELECT DISTINCT(FO) from FC_SESReport A where A.Requisitioner is not NULL and (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + sSQLRecent + ")";
                DataTable dt = msSQL.GetDataTable(sSQL);
                if (dt!=null&&dt.Rows.Count>0)
                {
                    string sReturn = "";
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        sReturn += dt.Rows[i][0].ToString()+",";
                    }
                    return sReturn.TrimEnd(',');
                }
            }
            catch (Exception)
            {
                //throw;
            }
            return null;
        }

        public static string sendEmail(string sEMailTo, string sTile, string sContent,string sLogPath)
        {            
            try
            {
                string sFilePath= Path.Combine(sLogPath, "email.log");
                //now checking if this email has been sent this month
                int iCount = 0;
                if (File.Exists(sFilePath) && !sEMailTo.ToLower().Contains("gang.ji"))
                {
                    string[] sAll = File.ReadAllLines(sFilePath);
                    foreach (string s in sAll)
                    {
                        if (s.StartsWith(DateTime.Now.ToString("yyyy-MM-"))&&s.Contains(sEMailTo))
                        {
                            //return null;
                            iCount += 1;
                        }
                    }
                }
                if (iCount>4)
                {
                    return null;
                }              

                ServiceReference1.Service1SoapClient client = new ServiceReference1.Service1SoapClient();
                string sResult = client.sendEmail("jigang", sEMailTo, sTile, sContent);
                LogFile(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ":Send Email to " + sEMailTo + ",result:" + sResult + "\r\n",sFilePath);
                return sResult;
            }
            catch (Exception)
            {

                //throw;
            }
            return null;
        }

        private string getDepStatistics()
        {
            try
            {
                //for departments
                string sInfo = "";
                int iNum = 0, iAll = 0;
                for (int i = 0; i < 6; i++)
                {
                    string[] sResults = data.getEmails(i);
                    if (sResults != null)
                    {
                        for (int j = 0; j < sResults.Length; j++)
                        {
                            
                            string[] ss = sResults[j].Split(new string[] { "[@]" }, StringSplitOptions.RemoveEmptyEntries);
                            if (ss.Length == 2)
                            {
                                if (!sInfo.Contains(ss[1]))
                                {
                                    iAll += 1;
                                    string sUserName = data.getUserDetails(ss[0], 0);
                                    //for resending purposes 2013-09-05
                                    DataTable dtTmp = data.msSQL.GetDataTable("SELECT ID from FC_Score where LOWER([By])='" + ss[0].ToLower() + "' AND DATEDIFF(month, DateIn, GETDATE())<1");
                                    if (dtTmp != null && dtTmp.Rows.Count > 0)
                                    {
                                        continue;
                                    }
                                    sInfo += sUserName + "/" + ss[1] + " As Dep.<br/>";
                                    iNum += 1;
                                }
                            }
                        }
                    }

                }
                return (iAll - iNum).ToString() + "/" + iAll.ToString() + "≈" + ((double)(iAll - iNum)/(double)iAll).ToString("##.##%");
            }
            catch (Exception)
            {

                //throw;
            }
            return "Null";
        }

        //public static void LogFile(string sMsg, string sPath)
        //{
        //    try
        //    {
        //        //string sPath = Path.Combine(System.Web.HttpContext.Current.Server.MapPath("Files"), sFileName);
        //        File.AppendAllText(sPath, sMsg);
        //    }
        //    catch (Exception)
        //    {

        //        //throw;
        //    }
        //}

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}