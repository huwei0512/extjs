using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FCPortal
{
    public partial class Report : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string sContract_NO = Request.Params["FO"];
                string sStartDate = Request.Params["startDate"];
                string sEndDate = Request.Params["endDate"];
                string sExportType = Request.Params["export"];
                string CurrentHost = Request.Url.Host + ((Request.Url.Port == 80) ? "" : ":" + Request.Url.Port.ToString()) +"/"+Request.Url.Segments[Request.Url.Segments.Length - 2].Replace("/","");
                CurrentHost = CurrentHost.TrimEnd(new char[]{'/'});
                if (string.IsNullOrEmpty(sContract_NO))
                {
                    sContract_NO = "A546059087";
                }
                if (string.IsNullOrEmpty(sExportType))
                {
                    sExportType = "overall";
                }
                if (string.IsNullOrEmpty(sStartDate))
                {
                    sStartDate = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-dd");
                }
                if (string.IsNullOrEmpty(sEndDate))
                {
                    sEndDate = DateTime.Now.ToString("yyyy-MM-dd");
                }                
                DateTime dtStart = DateTime.Parse(sStartDate);
                DateTime dtEnd = DateTime.Parse(sEndDate);
                if (dtStart.Day < 10)
                {
                    dtStart = dtStart.AddMonths(-1);
                    dtStart = DateTime.Parse(dtStart.ToString("yyyy-MM-11"));
                }
                else
                {
                    dtStart = DateTime.Parse(dtStart.ToString("yyyy-MM-11"));
                }
                if (dtEnd.Day < 10)
                {
                    dtEnd = DateTime.Parse(dtEnd.ToString("yyyy-MM-10"));
                }
                else
                {
                    dtEnd = dtEnd.AddMonths(1);
                    dtEnd = DateTime.Parse(dtEnd.ToString("yyyy-MM-10"));
                }
                sStartDate = dtStart.ToString("yyyy-MM-dd");
                sEndDate = dtEnd.ToString("yyyy-MM-dd");
                if (sExportType == "monthdata")
                {
                    //added on 2015-10-10 by request of Caixd
                    exportMonthData(Context,dtStart,dtEnd);
                    return;
                }
                int iMonthCount=(int)((dtEnd-dtStart).TotalDays/30);

                reportViewer1.ShowPrintButton = false; reportViewer1.LocalReport.EnableHyperlinks = true; 
                //绑定报表
                reportViewer1.LocalReport.ReportPath = MapPath("App_Data\\Report1.rdlc");
                //reportViewer1.LocalReport.SetParameters(new Microsoft.Reporting.WebForms.ReportParameter("dateStart", sStartDate));
                //reportViewer1.LocalReport.SetParameters(new Microsoft.Reporting.WebForms.ReportParameter("dateEnd", sEndDate));   
             
                 //List<Microsoft.Reporting.WebForms.ReportParameter> listParam = new List<Microsoft.Reporting.WebForms.ReportParameter>();
                 //listParam.Add(new Microsoft.Reporting.WebForms.ReportParameter("currentHost", CurrentHost));
                reportViewer1.LocalReport.SetParameters(new Microsoft.Reporting.WebForms.ReportParameter("currentHost", CurrentHost));
                
                DataTable dt = new DataTable();//data.msSQL.GetDataTable("select *,'test' as 'Remark' from getDetails('"+sStartDate+"',1) where FO_NO='" + sContract_NO + "'");
                //now generating different reports according to export type:overall month quarter year
                switch(sExportType)
                {
                    case "overall":
                        dt = data.msSQL.GetDataTable("select *,'" + dtStart.AddMonths(-1).ToString("yyyy-MM-dd") + "/" + dtEnd.AddMonths(-1).ToString("yyyy-MM-dd") + "' as 'TimeInterval','" + sStartDate + "' as 'startDate','" + sEndDate + "' as 'endDate' from getDetails('" + sStartDate + "'," + iMonthCount.ToString() + ") where FO_NO='" + sContract_NO + "'");
                        break;
                    case "month":
                        for (int i = 0; i < iMonthCount; i++)
                        {
                            sStartDate = dtStart.AddMonths(i-1).ToString("yyyy-MM-dd");
                            sEndDate = dtStart.AddMonths(i-1).AddMonths(1).ToString("yyyy-MM-10");
                            DataTable dttmp = data.msSQL.GetDataTable("select *,'" + sStartDate + "/" + sEndDate + "' as 'TimeInterval','" + sStartDate + "' as 'startDate','" + sEndDate + "' as 'endDate' from getDetails('" + dtStart.AddMonths(i).ToString("yyyy-MM-dd") + "',1) where FO_NO='" + sContract_NO + "'");
                            if (dt.Rows.Count==0)
                            {
                                dt = dttmp.Clone();
                            }
                            dt.Rows.Add(dttmp.Rows[0].ItemArray);
                        }
                        break;
                    case "quarter":
                        for (int i = 0; i < (iMonthCount/3)+1; i++)
                        {
                            if (dtStart.AddMonths(i*3)>=dtEnd)
                            {
                                continue;
                            }
                            sStartDate = dtStart.AddMonths(i*3-1).ToString("yyyy-MM-dd");
                            sEndDate = dtStart.AddMonths(i*3+2-1).AddMonths(1).ToString("yyyy-MM-10");
                            DataTable dttmp = data.msSQL.GetDataTable("select *,'" + sStartDate + "/" + sEndDate + "' as 'TimeInterval','" + sStartDate + "' as 'startDate','" + sEndDate + "' as 'endDate' from getDetails('" + dtStart.AddMonths(i * 3).ToString("yyyy-MM-dd") + "',3) where FO_NO='" + sContract_NO + "'");
                            if (dt.Rows.Count == 0)
                            {
                                dt = dttmp.Clone();
                            }
                            dt.Rows.Add(dttmp.Rows[0].ItemArray);
                        }
                        break;
                    case "year":
                        for (int i = 0; i < (iMonthCount / 12) + 1; i++)
                        {
                            if (dtStart.AddMonths(i * 12) >= dtEnd)
                            {
                                continue;
                            }
                            sStartDate = dtStart.AddMonths(i * 12-1).ToString("yyyy-MM-dd");
                            sEndDate = dtStart.AddMonths(i * 12 + 11-1).AddMonths(1).ToString("yyyy-MM-10");
                            DataTable dttmp = data.msSQL.GetDataTable("select *,'" + sStartDate + "/" + sEndDate + "' as 'TimeInterval','" + sStartDate + "' as 'startDate','" + sEndDate + "' as 'endDate' from getDetails('" + dtStart.AddMonths(i * 12).ToString("yyyy-MM-dd") + "',12) where FO_NO='" + sContract_NO + "'");
                            if (dt.Rows.Count == 0)
                            {
                                dt = dttmp.Clone();
                            }
                            dt.Rows.Add(dttmp.Rows[0].ItemArray);
                        }
                        break;
                    default:
                        break;
                }

                //为报表浏览器指定报表文件
                //this.reportViewer1.LocalReport.ReportEmbeddedResource = "Report1.rdlc";
                //指定数据集,数据集名称后为表,不是DataSet类型的数据集
                this.reportViewer1.LocalReport.DataSources.Clear();
                this.reportViewer1.LocalReport.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("fc", dt));
                this.reportViewer1.LocalReport.SubreportProcessing += LocalReport_SubreportProcessing;
                //显示报表                
                this.reportViewer1.LocalReport.Refresh();
                
            }
        }

        void LocalReport_SubreportProcessing(object sender, Microsoft.Reporting.WebForms.SubreportProcessingEventArgs e)
        {
            DataTable dtFiles = new DataTable();
            try
            {
                string sContract_NO = e.Parameters["FO"].Values[0];
                string sStartDate = e.Parameters["startDate"].Values[0];
                string sEndDate = e.Parameters["endDate"].Values[0];
                dtFiles = data.msSQL.GetDataTable("select top 6 * from FC_File where DateIn>='" + sStartDate + "' and DateIn<='"+sEndDate+"' and Contract_NO='" + sContract_NO + "'");
            }
            catch (Exception)
            {
                //throw;
            }
            e.DataSources.Add(new Microsoft.Reporting.WebForms.ReportDataSource("files", dtFiles));
        }

        private void exportMonthData(HttpContext context, DateTime dtStart,DateTime dtEnd)
        {
            byte[] bs = FCPortal.Properties.Resources.Overall_Month_Data;
            MemoryStream msFile = new MemoryStream(bs);
            IWorkbook workbook = NPOI.SS.UserModel.WorkbookFactory.Create(msFile);
            MemoryStream ms = new MemoryStream();
            ISheet sheet = workbook.GetSheet("Overall Month Report");
            DateTime dtSysStart = DateTime.Parse("2013-09-01");
            if ((dtSysStart - dtStart).TotalDays > 0)
            {
                dtStart = dtSysStart;
            }
            int iMonths=(int)(dtEnd-dtStart).TotalDays/30;
            DataTable dtAllContracts = data.msSQL.GetDataTable("select distinct(Contract_No) from FC_Score");
            if (dtAllContracts != null && dtAllContracts.Rows.Count > 0)
            {
                for (int i = 0; i < iMonths+1; i++)
                {
                    if ((dtStart.AddMonths(i) - DateTime.Now).TotalDays > 0)
                    {
                        break;
                    }
                    //now marking the time
                    IRow row = sheet.GetRow(0);
                    ICell cell = row.CreateCell(5 + i);
                    cell.SetCellValue(dtStart.AddMonths(i - 1).ToString("yyyy-MM"));
                    for (int j = 0; j < dtAllContracts.Rows.Count; j++)
                    {
                        DataTable dtAvg = data.msSQL.GetDataTable("SELECT FO_NO,Contract_Title,Contractor,Contract_Admin,Main_Coordinator,dbo.getScore(FO_NO,'" + (dtStart.AddMonths(i)).ToString("yyyy-MM-01") + "',1,'11',3) as 'ScoreAll' from FCMain where FO_NO='" + data.object2Str(dtAllContracts.Rows[j][0]) + "'");
                        if (i == 0)
                        {
                            row = sheet.CreateRow(j + 1);
                            cell = row.CreateCell(0);
                            cell.SetCellValue(data.object2Str(dtAllContracts.Rows[j][0]));
                            if (dtAvg != null && dtAvg.Rows.Count > 0)
                            {
                                cell = row.CreateCell(1);
                                cell.SetCellValue(data.refineContractName(data.object2Str(dtAvg.Rows[0][1])));
                                cell = row.CreateCell(2);
                                cell.SetCellValue(data.refineContractName(data.object2Str(dtAvg.Rows[0][2])));
                                cell = row.CreateCell(3);
                                cell.SetCellValue(data.object2Str(dtAvg.Rows[0][3]));
                                cell = row.CreateCell(4);
                                cell.SetCellValue(data.object2Str(dtAvg.Rows[0][4]));
                            }
                        }
                        if (dtAvg != null && dtAvg.Rows.Count > 0)
                        {
                            row = sheet.GetRow(j + 1);
                            cell = row.CreateCell(5 + i);
                            string sValue = data.object2Str(dtAvg.Rows[0][5], "-");
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
                            dtAvg = data.msSQL.GetDataTable("SELECT dbo.getScore('" + data.object2Str(dtAllContracts.Rows[j][0]) + "','" + (dtStart.AddMonths(i)).ToString("yyyy-MM-01") + "',1,'11',3) as 'ScoreAll'");
                            if (dtAvg != null && dtAvg.Rows.Count > 0)
                            {
                                row = sheet.GetRow(j + 1);
                                cell = row.CreateCell(5 + i);
                                string sValue = data.object2Str(dtAvg.Rows[0][0], "-");
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
            workbook.Write(ms);
            context.Response.ContentType = "application/octet-stream";
            context.Response.AddHeader("Content-Disposition", string.Format("attachment; filename=OverallMonthReport_" + DateTime.Now.ToString("yyyy-MM-dd") + ".xls"));
            context.Response.BinaryWrite(ms.ToArray());
            workbook = null;
            msFile.Close();
            ms.Close();
            ms.Dispose();

        }

    }
}