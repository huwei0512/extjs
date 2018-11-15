using Ext.Net;
using Ext.Net.Utilities;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace FCPortal
{
    public partial class _Default : System.Web.UI.Page
    {
        public static Thread nt = null;
        public static string sPath = "";
        protected void Page_Load(object sender, EventArgs e)
        {    
            if (!X.IsAjaxRequest)
            {
                initComponents();

                //Test Update 2017-09-30
                //updateData();

                //

                if (!Request.Url.AbsoluteUri.ToLower().Contains("localhost")&&!Request.Url.AbsoluteUri.ToLower().Contains("test"))
                {                    
                    updateData();
                    //auto Score
                    AutoScore();
                    batchEmail(); 
                }
            }
        }

        protected void btnLogin_Click(object sender, DirectEventArgs e)
        {
            if (txtPassword.Text=="jigang")
            {
                hiddenUser.Text = txtUsername.Text.Trim().ToLower();//liuj31 YANSH
                btnLogin.Text = hiddenUser.Text;
                storeEvaluationContract.Reload();
                return;
            }           

            // Do some Authentication...
            string sUser = txtUsername.Text.Trim().ToLower();            
            BYC.clsBYCUser bycUser = new BYC.clsBYCUser(sUser, txtPassword.Text.Trim());            
            if (bycUser.checkUser())
            {
                //X.Msg.Alert("Welcome:", bycUser.FullName).Show();
                string sFullName=data.getUserDetails(sUser,0);
                string sEmail = data.getUserDetails(sUser, 1);

                btnLogin.Icon = Icon.LockOpen;
                btnLogin.Text = string.IsNullOrEmpty(sFullName) ? bycUser.FullName : sFullName;
                hiddenUser.Text = sUser;
                
                //now logging the user infomation for next auto login
                string sLoginCookies = penlau.clsEncode.Encrypt(btnLogin.Text + "[@]" + sUser + "[@]" + rdUser.Checked.ToString(), "12345678");
                Response.Cookies["UserInfo"].Value = sLoginCookies;
                Response.Cookies["UserInfo"].Expires = DateTime.Now.AddDays(3650);

                Notification.Show(new NotificationConfig
                {
                    Title = "Information",
                    Icon = Icon.Information,
                    Html = btnLogin.Text+ ":<br>" + (string.IsNullOrEmpty(sEmail)?bycUser.EMail:sEmail) + "<br>" + bycUser.Department
                });

                storeEvaluationContract.Reload(); 
                //now checking if this User has contract to evaluate
                //DataTable dt = data.msSQL.GetDataTable("SELECT DISTINCT(FO) from FC_SESReport A WHERE (CS_REC_Format is not null or TECO_Format is not NULL) and UPPER(Requisitioner)='"+sUser.ToUpper()+"' and (not EXISTS(SELECT * from FC_Score B where A.FO=B.Contract_No and B.[By]='"+sUser+"'))");
                //if (dt!=null&&dt.Rows.Count>0)
                //{
                //    //now hiligting
                //    X.AddScript("hilightEvaluation();");                    
                //}
                //now send emails for testing
                //data.sendEmail(sEmail, "Email form FCPortal", "<P>Dear <STRONG>"+sFullName+"</STRONG>:</P><P>Here is a test email sent from<STRONG> <A href=\"http://10.137.12.32/fcl/\" target=_blank>FCPortal</A>&nbsp;</STRONG><EM><FONT size=4>for test</FONT></EM>.</P>");
                //sendEmail(btnLogin.Text,sUser,sEmail,true,"");
            }
            else
            {
                Notification.Show(new NotificationConfig
                {
                    Title = "Error",
                    Icon = Icon.Error,
                    Html = bycUser.LastError
                });
            }
            bycUser.close();
        }
        private void checkLogin()
        {
            //if (Request.Cookies["User"] != null)
            //{
            //    Request.Cookies["User"].Value = null;
            //    Response.Cookies["User"].Expires = DateTime.Now.AddMinutes(-10);
            //}
            if (Request.Cookies["UserInfo"] != null)
            {
                //Response.Write(Server.HtmlEncode(Request.Cookies["username"].Value));
                string sValue = penlau.clsEncode.Decrypt(Request.Cookies["UserInfo"].Value, "12345678");
                if (!string.IsNullOrEmpty(sValue))
                {
                    string[] ss = sValue.Split(new string[] { "[@]" }, StringSplitOptions.RemoveEmptyEntries);
                    if (ss.Length==3)
                    {
                        btnLogin.Icon = Icon.LockOpen;
                        btnLogin.Text = ss[0];
                        hiddenUser.Text = ss[1].ToLower().Trim();
                        if (ss[2].ToLower()=="false")
                        {
                            rdDep.SetValue(true);
                            rdUser.SetValue(false);
                        }
                    }
                }
            }
        }

        private void initComponents()
        {
            checkLogin();
            sPath = MapPath(".");

            //for status text indicating the time updated
            string sDateRecordPath = Path.Combine(MapPath("."), "timeRecord.txt");
            DateTime dt = DateTime.Now;
            FileInfo fi = new FileInfo(sDateRecordPath);
            if (fi.Exists)
            {
                dt = fi.LastWriteTime;
                statusBar1.SetStatus("Updated on "+dt.ToString("yyyy-MM-dd") );
            }
            //for search Pricing Scheme lists
            DataTable dtData = data.msSQL.GetDataTable("select distinct(Pricing_Scheme) from FC_SESRelatedData");
            if (dtData!=null&&dtData.Rows.Count>0)
            {
                for (int i = 0; i < dtData.Rows.Count; i++)
                {
                    string sTmp = dtData.Rows[i][0].ToString();
                    SearchPringScheme.Items.Add(new Ext.Net.ListItem(sTmp));
                }
            }
            dtData = data.msSQL.GetDataTable("select distinct(Contract_Admin) from FC_SESRelatedData");
            if (dtData != null && dtData.Rows.Count > 0)
            {
                for (int i = 0; i < dtData.Rows.Count; i++)
                {
                    string sTmp = dtData.Rows[i][0].ToString();
                    searchCA.Items.Add(new Ext.Net.ListItem(sTmp));
                }
            }

            //newly added 2014-03-21
            cbContractStatus.Items.Add(new Ext.Net.ListItem("Valid Contracts"));
            cbContractStatus.Items.Add(new Ext.Net.ListItem("All Contracts"));
            cbContractStatus.Select(0);
            dtData = data.msSQL.GetDataTable("select distinct(Contract_Title) from FCMain");
            if (dtData != null && dtData.Rows.Count > 0)
            {
                cbContractDescription.Items.Add(new Ext.Net.ListItem("All"));
                for (int i = 0; i < dtData.Rows.Count; i++)
                {
                    string sTmp = dtData.Rows[i][0].ToString();
                    cbContractDescription.Items.Add(new Ext.Net.ListItem(sTmp));
                }
            }
            dtData = data.msSQL.GetDataTable("select distinct(Contractor) from FCMain");
            if (dtData != null && dtData.Rows.Count > 0)
            {
                cbContractor.Items.Add(new Ext.Net.ListItem("All"));
                for (int i = 0; i < dtData.Rows.Count; i++)
                {
                    string sTmp = dtData.Rows[i][0].ToString();
                    cbContractor.Items.Add(new Ext.Net.ListItem(sTmp));
                }
            }

            //for startDate and Enddate
            this.DateField1.SetValue(DateTime.Now.AddMonths(-11));
            this.DateField2.SetValue(DateTime.Now);
            this.DateReport.SetValue(DateTime.Now);

            string sNOPath = Path.Combine(MapPath("/"), "contractNO.txt");
            if (File.Exists(sNOPath))
            {
                txtContracts.SetValue(File.ReadAllText(sNOPath));
            }

            dfKPIStart.SetValue(DateTime.Now);
            dfKPIEnd.SetValue(DateTime.Now);
            dfKPI1Start.SetValue(DateTime.Now.AddMonths(-11));
            dfKPI1End.SetValue(DateTime.Now);
            dfKPI2Start.SetValue(DateTime.Now);
            dfKPI2End.SetValue(DateTime.Now);
            dfKPILineStart.SetValue(DateTime.Now.AddMonths(-11));
            dfKPILineEnd.SetValue(DateTime.Now);
            dfKPI3Start.SetValue(DateTime.Now.AddMonths(-11));
            dfKPI3End.SetValue(DateTime.Now);
            dfUserKPIStart.SetValue(DateTime.Now.AddMonths(-11));
            dfUserKPIEnd.SetValue(DateTime.Now);
            dfUserSampleStart.SetValue(DateTime.Now.AddMonths(-2));
            dfUserSampleEnd.SetValue(DateTime.Now);
            dfContractorKPIStart.SetValue(DateTime.Now.AddMonths(-11));
            dfContractorKPIEnd.SetValue(DateTime.Now);
            dfUserProStart.SetValue(DateTime.Now.AddMonths(-11));
            dfUserProEnd.SetValue(DateTime.Now);
        }

        protected void exportExcel(object sender, DirectEventArgs e)
        {
            //X.Msg.Notify("The Server Time is: ", DateTime.Now.ToLongTimeString()).Show();
            //updateData();
            if (sender.GetType().ToString().Contains("Menu"))
            {
                string sUser = hiddenUser.Text.ToLower();
                string[] sUsers = new string[] { "suhy", "suhy", "jig1", "suny2", "zhux3", "wangja", "caixd", "yangj47", "wangda", "jansenb4", "wieselal", "xuz12", "coneno1" };
                bool bContains = false;
                foreach (string s in sUsers)
                {
                    if (sUser==s)
                    {
                        bContains = true;
                        break;
                    }
                }
                if (bContains)
                {
                    //now exporting
                    X.AddScript("window.open('data.ashx?exportReportExcel=true','_blank')");
                }
                else
                {
                    X.AddScript("showMsg(\"Information\", \"You are not authorized!\");");
                }
                return;
            }

            X.AddScript("window.open('data.ashx?exportExcel=true&searchKey=" + e.ExtraParams["searchKey"] + "&NO=" + e.ExtraParams["NO"] + "&CA=" + e.ExtraParams["CA"] + "&PS=" + e.ExtraParams["PS"] + "&isValid=" + e.ExtraParams["isValid"] + "','_blank')");
        }
        protected void storeWinFile_Submit(object sender, StoreSubmitDataEventArgs e)
        {
            string sPath = "",sFileName="";
            Store s = (Store)sender;
            if (s.ID == "storeMain")
            {
                sFileName = "Evaluation.xlsx";
                sPath = Path.Combine(MapPath("DataSheet"), "Evaluation Template.xlsx");
                data.exportTemplate(Context, sPath);
                return;
            }
            else if (s.ID == "StoreSub")
            {
                string sUser = hiddenUser.Text.ToLower();
                string[] sUsers = new string[] { "suhy", "jig1", "wangja", "suny2", "zhux3", "caixd", "yangj47", "wangda2", "linm", "zhouhx2", "jansenb4", "wieselal", "xuz12", "coneno1" };
                bool bContains = false;
                foreach (string seach in sUsers)
                {
                    if (sUser == seach)
                    {
                        bContains = true;
                        break;
                    }
                }


                if (bContains)
                {
                    //now exporting
                    //X.AddScript("window.open('data.ashx?exportReportExcel=true','_blank')");                    
                }
                else
                {
                    X.AddScript("showMsg(\"Information\", \"You are not authorized!\");"); 
                    return;
                }
                //X.AddScript("showMsg(\"Information\", \"" + ((DateTime)DateReport.Value).ToString("yyyy-MM-dd") + "\");"); return;
                DateTime dtNow = DateTime.Now;
                try
                {
                    dtNow = (DateTime)DateReport.Value;
                }
                catch (Exception)
                {                    
                    //throw;
                }
                data.exportReportExcel(Context, dtNow);                
                return;
            }
            else
            {
                string sAll = hiddenFileID.Text;
                string[] ss = sAll.Split(new string[] { "[@]" }, StringSplitOptions.RemoveEmptyEntries);
                sFileName = ss[1];
                sPath = Path.Combine(MapPath("Files"), ss[0]);
            }            


            Response.Clear();
            Response.ContentType = "application/octet-stream";
            Response.AddHeader("Content-Disposition", "attachment;filename=" + HttpUtility.UrlEncode(sFileName));
            Response.TransmitFile(sPath);            
            Response.End();
        }
        private void updateData()
        {
            try
            {
                if (nt==null||nt.IsAlive==false)
                {
                    nt = new Thread(new ThreadStart(updateDataThread));
                    //nt.IsBackground = true;
                    nt.Start(); 
                }
            }
            catch (Exception)
            {
                //throw;
            }
        }
        private void updateDataThread()
        {
            try
            {
                //now first of all,check if it necessary to update the records
                string sHostPath = MapPath(".");
                string sDateRecordPath = Path.Combine(sHostPath, "timeRecord.txt");
                FileInfo fi = new FileInfo(sDateRecordPath);
                string sDataSheetPath = Path.Combine(sHostPath + "\\DataSheet\\", "SES related data.xlsx");
                FileInfo fiExcel = new FileInfo(sDataSheetPath);
                if (!fiExcel.Exists)
                {
                    //X.Msg.Alert("Error", "FC Datasheet does not exist!").Show();
                    return;
                }
                bool bNeedUpdate = false;
                if (fi.Exists)
                {
                    if (fi.LastWriteTime!=fiExcel.LastWriteTime)
                    {
                        bNeedUpdate = true;
                    }
                }
                else
                {                    
                    bNeedUpdate = true;
                }
                if (bNeedUpdate)
                {
                    File.WriteAllText(sDateRecordPath, fiExcel.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    fi.LastWriteTime = fiExcel.LastWriteTime;
                    updateDB(sDataSheetPath);
                }

                //******************************************************************
                sDateRecordPath = Path.Combine(sHostPath, "timeRecordSES.txt");
                fi = new FileInfo(sDateRecordPath);
                sDataSheetPath = Path.Combine(sHostPath + "\\DataSheet\\", "SES Report.xlsx");
                fiExcel = new FileInfo(sDataSheetPath);
                if (!fiExcel.Exists)
                {
                    //X.Msg.Alert("Error", "FC Datasheet does not exist!").Show();
                    return;
                }
                bNeedUpdate = false;
                if (fi.Exists)
                {
                    if (fi.LastWriteTime != fiExcel.LastWriteTime)
                    {
                        bNeedUpdate = true;
                    }
                }
                else
                {
                    bNeedUpdate = true;
                }
                if (bNeedUpdate)
                {
                    File.WriteAllText(sDateRecordPath, fiExcel.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"));
                    fi.LastWriteTime = fiExcel.LastWriteTime;
                    updateDBSES_Report(sDataSheetPath);
                }



               //X.Msg.Notify("The Server Time is: ", dt.ToString("yyyy-MM-dd HH:mm:ss")).Show();
            }
            catch (Exception ex)
            {
                //X.Msg.Alert("Error", ex.Message).Show();
                //throw;
            }
        }
        private void updateDB(string sDataSheetPath)
        {
            //sDataSheetPath = @"C:\Users\W766CWG88X\Desktop\SES related data.xlsx";
            int iCount = 0;
            string[] sExcelNames = new string[] { "Contract Title", "Contractor", "Pricing Scheme", "COBALT FO No.", "Item", "Material Group", "Purchase Group", "Cost Element", "Vendor No", "Currency", "Original W/C(DO not use in Cobalt)", "Type", "Contract Admin.", "Buyer", "Main Coordinator", "User Representative", "Applicant", "Validate Date", "Expire Date", "FC Status", "Contact Person", "Tel.", "Total Budget", "Proportion of FC Definition", "Actual Budget" };
            string[] sDBNames = new string[] { "Contract_Title", "Contractor", "Pricing_Scheme", "FO_NO", "Item", "Material_Group", "Purchase_Group", "Cost_Element", "Vendor_NO", "Currency", "Original_WC", "Type", "Contract_Admin", "Buyer", "Main_Coordinator", "User_Representative", "Applicant", "Validate_Date", "Expire_Date", "FC_Status", "Contract_Person", "Contract_Tel", "Total_Budget", "Proportion_of_FC_Definition", "Actual_Budget" };
            using (FileStream file = new FileStream(sDataSheetPath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook hssfworkbook = NPOI.SS.UserModel.WorkbookFactory.Create(file);
                iCount = hssfworkbook.NumberOfSheets;
                for (int i = 0; i < iCount; i++)
                {
                    ISheet sheet = hssfworkbook.GetSheetAt(i);
                    //int iColumns = 27;//sheet.GetRow(0).LastCellNum;
                    for (int j = 1; j < sheet.LastRowNum + 1; j++)
                    {
                        try
                        {
                            IRow row = sheet.GetRow(j);
                            IRow rowTitle = sheet.GetRow(0);
                            //if (row.LastCellNum < iColumns)
                            //{
                            //    continue;
                            //}
                            string sSQLName = "", sSQLValue = "", sFONO = "", sSQLUpdate = "";
                            for (int n = 0; n < row.Cells.Count; n++)
                            {
                                //now getting the tilte
                                //if ((row.Cells[n].ColumnIndex - rowTitle.FirstCellNum) > rowTitle.Cells.Count - 1)
                                //{
                                //    continue;
                                //}
                                //string sTile = rowTitle.Cells[row.Cells[n].ColumnIndex - rowTitle.FirstCellNum].StringCellValue;
                                ICell cell = rowTitle.GetCell(row.Cells[n].ColumnIndex);
                                if (cell == null)
                                {
                                    continue;
                                }
                                string sTile = cell.StringCellValue;
                                //now generating SQLs
                                for (int m = 0; m < sExcelNames.Length; m++)
                                {
                                    string sExcelName = sExcelNames[m];
                                    if (sExcelName.Trim() == sTile.Trim())
                                    {
                                        if (row.Cells[n].CellType.ToString() == "ERROR")
                                        {
                                            continue;
                                        }
                                        //now getting values
                                        if (sDBNames[m] == "FO_NO")
                                        {
                                            sFONO = row.Cells[n].ToString();
                                            if (string.IsNullOrEmpty(sFONO) || sFONO.Length > 12)
                                            {
                                                continue;
                                            }
                                        }

                                        if (sExcelName.ToLower().Contains("date"))
                                        {
                                            if (row.Cells[n].CellType.ToString() != "NUMERIC" && !(row.Cells[n].CellType.ToString() == "STRING" && row.Cells[n].StringCellValue.ToLower().Trim() == "unlimited"))
                                            {
                                                continue;
                                            }
                                            sSQLName += "[" + sDBNames[m] + "],";
                                            DateTime dt = ((row.Cells[n].CellType.ToString() == "STRING" && row.Cells[n].StringCellValue.ToLower().Trim() == "unlimited")) ? (new DateTime(2099, 01, 01)) : row.Cells[n].DateCellValue;
                                            sSQLValue += "'" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                            sSQLUpdate += sDBNames[m] + "='" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                        }
                                        else if (sExcelName.ToLower().Contains("budget"))//|| sExcelName.ToLower().Contains("definition")
                                        {
                                            if (row.Cells[n].CellType.ToString() != "NUMERIC" && !(row.Cells[n].CellType.ToString() == "FORMULA" && row.Cells[n].CachedFormulaResultType.ToString() == "NUMERIC"))
                                            {
                                                continue;
                                            }
                                            sSQLName += "[" + sDBNames[m] + "],";
                                            sSQLValue += row.Cells[n].NumericCellValue + ",";
                                            sSQLUpdate += sDBNames[m] + "=" + row.Cells[n].NumericCellValue + ",";
                                        }
                                        else
                                        {
                                            string sTmp = row.Cells[n].ToString();
                                            if (sDBNames[m] == "FC_Status")
                                            {
                                                sTmp = row.Cells[n].StringCellValue;
                                            }
                                            sSQLName += "[" + sDBNames[m] + "],";
                                            sSQLValue += "'" + sTmp.Replace("'", " ").Trim() + "',";
                                            sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ").Trim() + "',";
                                        }
                                        
                                    }
                                }
                            }
                            //now insert into the DB;
                            sSQLName += "[DateIn]";
                            sSQLValue += "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                            sSQLUpdate += "DateIn='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                            if (!string.IsNullOrEmpty(sFONO)&&sFONO.Length<=12)
                            {
                                DataTable dt = data.msSQL.GetDataTable("select * from [FC_SESRelatedData] where FO_NO='" + sFONO + "'");
                                if (dt.Rows.Count > 0)
                                {
                                    //now updating
                                    bool b = data.msSQL.ExecuteQuery("UPDATE [FC_SESRelatedData] SET " + sSQLUpdate + " WHERE FO_NO='" + sFONO + "'");
                                }
                                else
                                {
                                    //now inserting
                                    bool b = data.msSQL.ExecuteQuery("insert into [FC_SESRelatedData](" + sSQLName + ") Values (" + sSQLValue + ")");
                                }
                                //at last update the Old SAP Data
                                double dOldSAPData = getOldSAPData(sFONO);
                                if (dOldSAPData>0)
                                {
                                    //now updating
                                    bool b = data.msSQL.ExecuteQuery("UPDATE [FC_SESRelatedData] SET OldSAPData="+dOldSAPData.ToString()+" WHERE FO_NO='" + sFONO + "'");
                                }
                            }

                        }
                        catch (Exception ex)
                        {
                            
                            //throw;
                        }
                    }
                }               
                
            }
        }
        //modified on 2013-08-07 for importing data directly form SAP
        private void updateDBSES_Report(string sDataSheetPath)
        {
            //string sDataSheetPath = @"C:\Documents and Settings\Administrator\桌面\temp\01 SES Report 2013.06.xlsx";
            using (FileStream file = new FileStream(sDataSheetPath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook hssfworkbook = NPOI.SS.UserModel.WorkbookFactory.Create(file);
                ISheet sheet = hssfworkbook.GetSheetAt(0);
                if (sheet != null)
                {
                    //now reading data into the database
                    string[] sExcelNames = new string[] { "SES No.", "Accepted", "Deleted", "Blocked", "Short Descrption", "Start Date", "End Date", "Created by", "Created on", "TECO Date", "TECO Format", "Requisitioner", "FO", "Item", "Vendor Name", "Discipline", " SSR budget", "Currency", "Contractor quotation", "SSR Actual cost", "Cost checker", "Tax rate", "Changed by", "   Deviation", "Deviation Percentage", "Overrun", "Long text", "Work Order", "Function location", "Main work center", "Work Center", "Dep.", "Section", "Cost center", "WBS", "Network", "TODAY", "Claim sheets receive", "CS REC Format", "No SUBM To today", "Contractor duration", "Engineer confirmed o", "ENG CONF Format", "No BoQ CONF to today", "BoQ confirmation dur", "SES Confirmed on", "SES CONF Format", "No SES to today", "Settlement duration", "Invoiced on", "Invoice duration", "Payment made on", "Payment duration" };
                    string[] sDBNames = new string[] { "SES_No", "Accepted", "Deleted", "Blocked", "Short_Descrption", "Start_Date", "End_Date", "Created_by", "Created_on", "TECO_Date", "TECO_Format", "Requisitioner", "FO", "Item", "Vendor_Name", "Discipline", "SSR_budget", "Currency", "Contractor_quotation", "SSR_Actual_cost", "Cost_checker", "Tax_rate", "Changed_by", "Deviation", "Deviation_Percentage", "Overrun", "Long_text", "Work_Order", "Function_location", "Main_work_center", "Work_Center", "Dep", "Section", "Cost_center", "WBS", "Network", "TODAY", "Claim_sheets_receive", "CS_REC_Format", "No_SUBM_To_today", "Contractor_duration", "Engineer_confirmed_o", "ENG_CONF_Format", "No_BoQ_CONF_to_today", "BoQ_confirmation_dur", "SES_Confirmed_on", "SES_CONF_Format", "No_SES_to_today", "Settlement_duration", "Invoiced_on", "Invoice_duration", "Payment_made_on", "Payment_duration" };
                    //int iColumns = sheet.GetRow(0).LastCellNum;
                    for (int j = 1; j < sheet.LastRowNum + 1; j++)
                    {
                        try
                        {
                            IRow row = sheet.GetRow(j);
                            IRow rowTitle = sheet.GetRow(0);
                            //if (row.LastCellNum < iColumns)
                            //{
                            //    continue;
                            //}
                            string sSQLName = "", sSQLValue = "", sSESNO = "", sSQLUpdate = "",sValuesMD5="";
                            bool isSectionFilled = false;
                            string sTmpSectionSQLName = "", sTmpSectionSQLValue = "", sTmpSectionSQLUpdate = "";
                            for (int n = 0; n < row.Cells.Count; n++)                           
                            {
                                //now getting the tilte
                                //if ((row.Cells[n].ColumnIndex - rowTitle.FirstCellNum) > rowTitle.Cells.Count - 1)
                                //{
                                //    continue;
                                //}
                                ICell cell= rowTitle.GetCell(row.Cells[n].ColumnIndex);
                                if (cell==null)
                                {
                                    continue;
                                }
                                string sTile = cell.StringCellValue;
                                //now generating SQLs
                                for (int m = 0; m < sExcelNames.Length; m++)
                                {
                                    string sExcelName = sExcelNames[m];
                                    if (sExcelName.Trim() == sTile.Trim())
                                    {
                                        if (row.Cells[n].CellType.ToString() == "ERROR")
                                        {
                                            continue;
                                        }
                                        //now getting values
                                        if (sDBNames[m] == "SES_No")
                                        {
                                            sSESNO = row.Cells[n].ToString();
                                            if (string.IsNullOrEmpty(sSESNO) || sSESNO.Length > 10||sSESNO=="1138024899")//added 2016-12-16 for wangj's mistake
                                            {
                                                continue;
                                            }
                                        }

                                        if (sExcelName.ToLower().Contains("format") || sExcelName.ToLower().Trim() == "today")
                                        {
                                            if (row.Cells[n].CellType.ToString() != "NUMERIC" && !(row.Cells[n].CellType.ToString() == "STRING" && row.Cells[n].StringCellValue.ToLower().Trim() == "unlimited"))
                                            {
                                                continue;
                                            }
                                            sSQLName += "[" + sDBNames[m] + "],";
                                            DateTime dt = ((row.Cells[n].CellType.ToString() == "STRING" && row.Cells[n].StringCellValue.ToLower().Trim() == "unlimited")) ? (new DateTime(2099, 01, 01)) : row.Cells[n].DateCellValue;
                                            sSQLValue += "'" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                            sSQLUpdate += sDBNames[m] + "='" + dt.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                        }
                                        else
                                        {
                                            string sTmp = row.Cells[n].ToString();
                                            if (sTmp==null)
                                            {
                                                sTmp = "";
                                            }
                                            sTmp = sTmp.Trim();
                                            if (sTile.ToLower().Trim()=="claim sheets receive")
                                            {
                                                DateTime? dtTmp = data.getSAPDateTime(sTmp.Replace("'", " "));
                                                if (dtTmp!=null)
                                                {
                                                    sSQLName += "[CS_REC_Format],";
                                                    sSQLValue += "'" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                    sSQLUpdate += "CS_REC_Format='" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                }
                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }
                                            else if (sTile.ToLower().Trim() == "engineer confirmed o")
                                            {
                                                DateTime? dtTmp = data.getSAPDateTime(sTmp.Replace("'", " "));
                                                if (dtTmp != null)
                                                {
                                                    sSQLName += "[ENG_CONF_Format],";
                                                    sSQLValue += "'" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                    sSQLUpdate += "ENG_CONF_Format='" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                }
                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }
                                            else if (sTile.ToLower().Trim() == "ses confirmed on")
                                            {
                                                DateTime? dtTmp = data.getSAPDateTime(sTmp.Replace("'", " "));
                                                if (dtTmp != null)
                                                {
                                                    sSQLName += "[SES_CONF_Format],";
                                                    sSQLValue += "'" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                    sSQLUpdate += "SES_CONF_Format='" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                }
                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }
                                            else if (sTile.ToLower().Trim() == "teco date")
                                            {
                                                DateTime? dtTmp = data.getSAPDateTime(sTmp.Replace("'", " "));
                                                if (dtTmp != null)
                                                {
                                                    sSQLName += "[TECO_Format],";
                                                    sSQLValue += "'" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                    sSQLUpdate += "TECO_Format='" + dtTmp.Value.ToString("yyyy-MM-dd HH:mm:ss") + "',";
                                                }
                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }
                                            else if (sTile.ToLower().Trim() == "wbs")//added for c/t's charts
                                            {
                                                if (!string.IsNullOrEmpty(sTmp))
                                                {
                                                    sTmpSectionSQLName = "[Section],";
                                                    sTmpSectionSQLValue = "'PS',";
                                                    sTmpSectionSQLUpdate = "[Section]='PS',";
                                                    isSectionFilled = true;
                                                }                                                
                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }
                                            else if (sTile.ToLower().Trim() == "work center")//added for c/t's charts
                                            {
                                                if (sTmp.Length>4&&!isSectionFilled)
                                                {
                                                    string sTmpSection = sTmp.Substring(0, 5);
                                                    if (sTmp.Length > 5)
                                                    {
                                                        if (sTmpSection=="CTA/L")
                                                        {
                                                            if (sTmp.Substring(0, 6) == "CTA/LV" || sTmp.Substring(0, 6) == "CTA/LW")
                                                            {

                                                            }
                                                            else
                                                            {
                                                                sTmpSection = "CTA/P";
                                                            } 
                                                        }
                                                    }
                                                    if (sTmpSection == "CTM/L")
                                                    {
                                                        sTmpSection = "CTM/P";
                                                    }
                                                    sTmpSectionSQLName = "[Section],";
                                                    sTmpSectionSQLValue = "'" + sTmpSection + "',";
                                                    sTmpSectionSQLUpdate = "[Section]='" + sTmpSection + "',";
                                                }
                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }
                                            else if (sTile.ToLower().Trim() == "section")//added for c/t's charts
                                            {
                                                //no action has been taken as has been decided by other parts
                                            }
                                            else
                                            {                                              

                                                sSQLName += "[" + sDBNames[m] + "],";
                                                sSQLValue += "'" + sTmp.Replace("'", " ") + "',";
                                                sSQLUpdate += sDBNames[m] + "='" + sTmp.Replace("'", " ") + "',";
                                            }

                                        }


                                    }
                                }
                            }
                            if (!string.IsNullOrEmpty(sTmpSectionSQLName))
                            {
                                sSQLName += sTmpSectionSQLName;
                                sSQLValue += sTmpSectionSQLValue;
                                sSQLUpdate += sTmpSectionSQLUpdate;
                            }
                            //now insert into the DB;
                            sValuesMD5 = data.Md516(sSQLValue);
                            sSQLName += "[Remark1],";
                            sSQLValue += "'" + sValuesMD5 + "',";
                            sSQLUpdate += "Remark1='" + sValuesMD5 + "',";
                            sSQLName += "[DateIn]";
                            sSQLValue += "'" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                            sSQLUpdate += "DateIn='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'";
                            if (!string.IsNullOrEmpty(sSESNO))
                            {
                                DataTable dt = data.msSQL.GetDataTable("select SES_No,Remark1 from [FC_SESReport] where SES_No='" + sSESNO + "'");
                                if (dt.Rows.Count > 0)
                                {
                                    if (dt.Rows[0][1] != null && dt.Rows[0][1].ToString() == sValuesMD5)
                                    {
                                        //continue;
                                    }
                                    else
                                    {
                                        //now updating
                                        bool b = data.msSQL.ExecuteQuery("UPDATE [FC_SESReport] SET " + sSQLUpdate + " WHERE SES_No='" + sSESNO + "'"); 
                                    }
                                }
                                else
                                {
                                    //now inserting
                                    bool b = data.msSQL.ExecuteQuery("insert into [FC_SESReport](" + sSQLName + ") Values (" + sSQLValue + ")");
                                }

                   
                            }

                        }
                        catch (Exception)
                        {
                            //throw;
                        }
                    }
                }

            }
        }
        private double getOldSAPData(string sContractNO)
        {
            try
            {
                string sAllData = FCPortal.Properties.Resources.OldSAPData;
                int iStart = sAllData.IndexOf(sContractNO + "@");
                if (iStart >= 0)
                {
                    int iEnd = sAllData.IndexOf("\r\n", iStart);
                    if (iEnd > 0)
                    {
                        string sValue = sAllData.Substring(iStart + sContractNO.Length + 1, iEnd - iStart - sContractNO.Length - 1);
                        if (!string.IsNullOrEmpty(sValue))
                        {
                            double dValue = double.Parse(sValue);
                            //added 2013-05-31 for those 6% 17% Tax deleted 2013-06-07
                            //if (dTax == 0.06 || dTax == 0.17)
                            //{
                            //    dValue = dValue / (dTax + 1);
                            //}
                            return dValue;
                        }
                    }
                }
            }
            catch (Exception)
            {

                //throw;
            }
            return 0.0;
        }
        //private string getUserDetails(string sUserID,int iType)//0 for email,1 for Full Name
        //{
        //    try
        //    {
        //        string sAllData = FCPortal.Properties.Resources.userDetails;
        //        string[] sLines = sAllData.Split(new string[]{"\r\n"},StringSplitOptions.RemoveEmptyEntries);
        //        for (int i = 0; i < sLines.Length; i++)
        //        {
        //            if (sLines[i].ToLower().StartsWith(sUserID.ToLower()))
        //            {
        //                string[] ss = sLines[i].Split(new string[]{"[@]"},StringSplitOptions.None);
        //                if (ss.Length>=4)
        //                {
        //                    if (iType == 1)
        //                    {
        //                        if (!string.IsNullOrEmpty(ss[2])&&!string.IsNullOrEmpty(ss[3]))
        //                        {
        //                            return ss[2].Split(new char[] { '-' })[0] + " " + ss[1].Split(new char[] { '-' })[0];
        //                        }
        //                    }
        //                    else 
        //                    {
        //                        return ss[3]; 
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        //throw;
        //    }
        //    return null;
        //}

        protected void gridMianSelected(object sender, DirectEventArgs e)
        {
            string sNO= e.ExtraParams["NO"];
            return;
        }

        protected void Evaluation_Click(object sender, DirectEventArgs e)
        {
            string sContract_No = e.ExtraParams["Contract_No"];
            string sScore1 = e.ExtraParams["Score1"], sScore2 = e.ExtraParams["Score2"], sScore3 = e.ExtraParams["Score3"],
                sScore4=e.ExtraParams["Score4"],sScore5=e.ExtraParams["Score5"],sScore6=e.ExtraParams["Score6"];
            string sUser=e.ExtraParams["User"].ToLower();
            string sRole = "User";
            //if (!rdUser.Checked)
            //{
            //    sRole = "Dep";
            //}
            if (!string.IsNullOrEmpty(sContract_No)&&!string.IsNullOrEmpty(sScore1))
            {
                //now inserting or updating
                //DataTable dt = data.msSQL.GetDataTable("select * from FC_Score where Contract_No='" + sContract_No + "' and [By]='" + sUser + "'");
                bool b = updateInsertEvaluation(sContract_No,sUser,sRole,null,sScore1,sScore2,sScore3,sScore4,sScore5,sScore6);
                if (b)
                {                    
                    //added on 2013-11-15
                    if (sScore1 == "0" || sScore2 == "0" || sScore3 == "0" || sScore4 == "0" || sScore5 == "0" || sScore6 == "0")
                    {
                        X.AddScript("showInformation('Please Note','There is Zero point,please confirm!<br/>您的评分中包含零分，请确认！如果不能确定或者不适用请给及格分！');");
                    }
                    else
                    {
                        Notification.Show(new NotificationConfig
                        {
                            Title = "Information",
                            Icon = Icon.Information,
                            Html = sContract_No + " Evaluated successfully!<br/>谢谢您的参与!"
                        });
                    }
                    //storeMain.Reload();
                    X.AddScript("searchAutomatic('" + sContract_No + "')");
                }
                else
                {
                    Notification.Show(new NotificationConfig
                    {
                        Title = "Error",
                        Icon = Icon.Error,
                        Html = "Error in evaluating "+sContract_No+"!"
                    });
                }
            }
        }
        private bool updateInsertEvaluation(string sContract_No, string sUser,string sRole,string sRemark, string sScore1, string sScore2, string sScore3, string sScore4, string sScore5, string sScore6)
        {
            bool b = false;
            if (string.IsNullOrEmpty(sRemark))
            {
                sRemark = "";
            }
            try
            {
                sUser = sUser.ToLower();
                //now inserting or updating
                string sFCRecent = " and DateIn<'" + DateTime.Now.ToString("yyyy-MM-" + (data.iStartDay + 1).ToString()) + "' and DateIn>='" + DateTime.Now.ToString("yyyy-MM-01") + "'";
                DataTable dt = data.msSQL.GetDataTable("select * from FC_Score where Contract_No='" + sContract_No + "' and [By]='" + sUser + "' and Role='"+sRole+"'" + sFCRecent);
                if (dt != null && dt.Rows.Count > 0)
                {
                    if (sRemark=="Auto")
                    {
                        DataTable dtRecord = data.msSQL.GetDataTable("SELECT SES_No,FO,Requisitioner from FC_SESReport A where FO='" + sContract_No + "' and (ISNULL(Requisitioner, 'Null'))='" + sUser + "' AND (Deleted is NULL and Blocked is NULL)  AND (TECO_Format is not NULL or SES_CONF_Format is not NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No)");
                        //now recording into the FC_SESRecord DB
                        if (dtRecord != null && dtRecord.Rows.Count > 0)
                        {
                            for (int i = 0; i < dtRecord.Rows.Count; i++)
                            {
                                data.msSQL.ExecuteQuery("insert into FC_SESRecord(SES_No,Contract_No,[By],Role,Remark,DateIn) Values('" + dtRecord.Rows[i][0].ToString() + "','" + dtRecord.Rows[i][1].ToString() + "','" + sUser + "','" + sRole + "','" + sRemark + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
                            }
                        }
                        //now rechecking if update FC_Score
                        dtRecord = data.msSQL.GetDataTable("select * from FC_Score where Role='User' and Contract_No='" + sContract_No + "' and [By]='" + sUser + "' AND DateIn>='"+DateTime.Now.ToString("yyyy-MM-01")+"'");
                        if (dtRecord!=null&&dtRecord.Rows.Count==0)
                        {
                            b = data.msSQL.ExecuteQuery("insert into FC_Score(Contract_No,Score1,Score2,Score3,Score4,Score5,Score6,[By],Role,Remark,DateIn) Values('" + sContract_No + "'," + sScore1 + "," + sScore2 + "," + sScore3 + "," + sScore4 + "," + sScore5 + "," + sScore6 + ",'" + sUser + "','" + sRole + "','" + sRemark + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
                            return true;
                        }
                        return false;
                    }
                    //updating
                    b = data.msSQL.ExecuteQuery("UPDATE FC_Score SET Score1=" + sScore1 + ",Score2=" + sScore2 + ",Score3=" + sScore3 + ",Score4=" + sScore4 + ",Score5=" + sScore5 + ",Score6=" + sScore6 + ",Role='"+sRole+"',Remark='" + sRemark + "',DateIn='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where Contract_No='" + sContract_No + "' and [By]='" + sUser + "' AND DateIn>='" + DateTime.Now.ToString("yyyy-MM-01") + "'");
                }
                else
                {
                    //inserting
                    b = data.msSQL.ExecuteQuery("insert into FC_Score(Contract_No,Score1,Score2,Score3,Score4,Score5,Score6,[By],Role,Remark,DateIn) Values('" + sContract_No + "'," + sScore1 + "," + sScore2 + "," + sScore3 + "," + sScore4 + "," + sScore5 + "," + sScore6 + ",'" + sUser + "','" + sRole + "','"+sRemark+"','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
                    DataTable dtRecord = data.msSQL.GetDataTable("SELECT SES_No,FO,Requisitioner from FC_SESReport A where FO='" + sContract_No + "' and ISNULL(Requisitioner, 'Null')='" + sUser + "' AND (Deleted is NULL and Blocked is NULL)  AND (TECO_Format is not NULL or SES_CONF_Format is not NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No)");
                    //now recording into the FC_SESRecord DB
                    if (b && dtRecord != null && dtRecord.Rows.Count > 0)
                    {
                        for (int i = 0; i < dtRecord.Rows.Count; i++)
                        {
                            data.msSQL.ExecuteQuery("insert into FC_SESRecord(SES_No,Contract_No,[By],Role,Remark,DateIn) Values('" + dtRecord.Rows[i][0].ToString() + "','" + dtRecord.Rows[i][1].ToString() + "','" + sUser + "','" + sRole + "','"+sRemark+"','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
                        }
                    }
                }
            }
            catch (Exception)
            {

                //throw;
            }
            return b;
        }

        //public static string GetUserInfo(string sUserName, string sPass)
        //{
        //    //{"errors":{"reason":"username or password is not corect!"},"success":false}
        //    //{"loginid":"jig1","group":"cte-d","success":true}
        //    System.Net.WebClient webc = new System.Net.WebClient();
        //    string sCon = webc.DownloadString("http://10.137.1.73/portal/DocbaseService?method=login&loginid=" + sUserName + "&password=" + sPass);
        //    return sCon;
        //}

        protected void UploadClick(object sender, DirectEventArgs e)
        {
            Ext.Net.Button btn = (Ext.Net.Button)sender;
            if (btn.ID == "btnUploadFile")
            {
                if (DateTime.Now.Day>=data.iStartDay)
                {
                    X.Msg.Show(new MessageBoxConfig
                    {
                        Buttons = MessageBox.Button.OK,
                        Icon = MessageBox.Icon.ERROR,
                        Title = "Error",
                        Message = "You can only evaluate from first to tenth!\r\n您只能在每月的一号到十号间评价!"
                    });
                    return;
                }
                fileTemplateManage(hiddenUser.Text);
                return;
            }
            string tpl = "Uploaded file: {0}<br/>Size: {1}";
            
            if (this.FileUploadField1.HasFile)
            {                
                //now generating Unique ID
                string sFileID = System.Guid.NewGuid().ToString().ToUpper().Replace("-", "").Trim() + Path.GetExtension(FileUploadField1.PostedFile.FileName);
                string sFileName = Path.GetFileName(FileUploadField1.PostedFile.FileName);
                string sUser = hiddenUser.Text;
                string sFileLength = FileUploadField1.PostedFile.ContentLength.ToString();
                string sRemark = fileRemark.Text;
                string sContract_NO=fileContractNO.Text;
                string sFileType = cbFileType.Text;
                //now saving files
                string sPath = Path.Combine(MapPath("Files"), sFileID);
                DirectoryInfo di = new DirectoryInfo(MapPath("Files"));
                if (!di.Exists)
                {
                    di.Create();
                }
                FileUploadField1.PostedFile.SaveAs(sPath);
                //now storing into DB
                bool b = data.msSQL.ExecuteQuery("insert into FC_File(FileID,FileName,Remark,DateIn,FileLength,UploadUser,Contract_NO,FileType) Values('" + sFileID + "','" + sFileName + "','" + sRemark + "','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "'," + sFileLength + ",'" + sUser + "','" + sContract_NO + "','" + sFileType + "')");
                if (b)
                {
                    X.Msg.Show(new MessageBoxConfig
                    {
                        Buttons = MessageBox.Button.OK,
                        Icon = MessageBox.Icon.INFO,
                        Title = "Success",
                        Message = string.Format(tpl, this.FileUploadField1.PostedFile.FileName, data.FormatBytes(this.FileUploadField1.PostedFile.ContentLength))
                    });
                    storeWinFile.Reload();
                }
                else
                {
                    X.Msg.Show(new MessageBoxConfig
                    {
                        Buttons = MessageBox.Button.OK,
                        Icon = MessageBox.Icon.ERROR,
                        Title = "Error",
                        Message = "Error in uploading " + this.FileUploadField1.PostedFile.FileName+"<br/>Please contact jig for more information!"
                    });
                }
            }
            else
            {
                X.Msg.Show(new MessageBoxConfig
                {
                    Buttons = MessageBox.Button.OK,
                    Icon = MessageBox.Icon.ERROR,
                    Title = "Fail",
                    Message = "No file uploaded"
                });
            }
        }
        private void fileTemplateManage(string sUser)
        {
            if (this.fufFile.HasFile)
            {
                //string sPath = Path.Combine(MapPath("Files"), fufFile.PostedFile.FileName);
                //FileUploadField1.PostedFile.SaveAs(sPath);
                IWorkbook hssfworkbook = NPOI.SS.UserModel.WorkbookFactory.Create(fufFile.PostedFile.InputStream);
                int iCount = hssfworkbook.NumberOfSheets;
                if (iCount==1)
                {
                    ISheet sheet = hssfworkbook.GetSheetAt(0);
                    int iRowCount = sheet.LastRowNum;                    
                    //List<clsFCEvaluation> data = new List<clsFCEvaluation>();
                    IRow rowTitle = sheet.GetRow(0);
                    for (int i = 1; i < iRowCount+1; i++)
                    {
                        try
                        {
                            IRow row = sheet.GetRow(i);
                            string sSQLNames = "",sSQLValues="";
                            string sContract_No = row.GetCell(0).StringCellValue;
                            int iZeroCount = 0;
                            if (!string.IsNullOrEmpty(sContract_No)&&sContract_No.Length>6&&row.Cells.Count>6)
                            {
                                for (int j = 6; j < row.Cells.Count; j++)
                                {
                                    string sTitle = rowTitle.Cells[row.Cells[j].ColumnIndex - rowTitle.FirstCellNum].ToString();
                                    try
                                    {
                                        if (!string.IsNullOrEmpty(row.Cells[j].ToString()))
                                        {
                                            double dNowValue = double.Parse(row.Cells[j].ToString());
                                            if (dNowValue<=0||dNowValue>10)
                                            {
                                                iZeroCount += 1;
                                                continue;
                                            }
                                            if (sTitle.Contains("CTS(<=10)"))
                                            {
                                                sSQLNames += "Score1,";
                                                sSQLValues += (dNowValue/2.0).ToString("0.00") + ",";
                                            }
                                            if (sTitle.Contains("CHA(<=10)"))
                                            {
                                                sSQLNames += "Score2,";
                                                sSQLValues += (dNowValue / 2.0).ToString("0.00") + ",";
                                            }
                                            if (sTitle.Contains("Main Coordinator(<=10)"))
                                            {
                                                sSQLNames += "Score3,";
                                                sSQLValues += (dNowValue / 2.0).ToString("0.00") + ",";
                                            }
                                            if (sTitle.Contains("User Representative(<=10)"))
                                            {
                                                sSQLNames += "Score4,";
                                                sSQLValues += (dNowValue / 2.0).ToString("0.00") + ",";
                                            }
                                            if (sTitle.Contains("CTM/T(<=10)"))
                                            {
                                                sSQLNames += "Score5,";
                                                sSQLValues += (dNowValue / 2.0).ToString("0.00") + ",";
                                            }
                                            if (sTitle.Contains("CTE/D(<=10)"))
                                            {
                                                sSQLNames += "Score6,";
                                                sSQLValues += (dNowValue / 2.0).ToString("0.00") + ",";
                                            } 
                                        }
                                    }
                                    catch (Exception)
                                    {
                                        //throw;
                                    }
                                }
                            }

                            if (!string.IsNullOrEmpty(sSQLNames)&&iZeroCount<6)
                            {
                                sSQLNames += "Score7,";
                                sSQLValues += "0,";
                                bool b = data.msSQL.ExecuteQuery("INSERT INTO FC_Score(Contract_No,"+sSQLNames+"[By],Role,DateIn) VALUES ('" + sContract_No + "'," + sSQLValues+ "'" + sUser + "','Dep','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')"); 
                            }

                        }
                        catch (Exception)
                        {
                            //throw;
                        }

                    }
                    //storeEvaluationContract.DataSource = data;
                    //storeEvaluationContract.DataBind();
                    storeEvaluationContract.Reload();
                    X.Msg.Show(new MessageBoxConfig
                    {
                        Buttons = MessageBox.Button.OK,
                        Icon = MessageBox.Icon.INFO,
                        Title = "Success",
                        Message ="OK"
                    });
                }
                return;
                
            }
            else
            {
                X.Msg.Show(new MessageBoxConfig
                {
                    Buttons = MessageBox.Button.OK,
                    Icon = MessageBox.Icon.ERROR,
                    Title = "Fail",
                    Message = "No file uploaded"
                });

            }
        }

        private void AutoScore()
        {
            if (DateTime.Now.Day >= data.iStartDay && DateTime.Now.Day <= 20)
            {
                //now first of all,check if it necessary to update the records
                string sHostPath = MapPath(".");
                string sDateRecordPath = Path.Combine(sHostPath, "autoScore.txt");
                DateTime dt = DateTime.Now.AddDays(-30);
                FileInfo fi = new FileInfo(sDateRecordPath);
                if (fi.Exists)
                {
                    dt = fi.LastWriteTime;
                }
                else
                {
                    File.WriteAllText(sDateRecordPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    fi.LastWriteTime = DateTime.Now;
                }
                if ((DateTime.Now - dt).TotalDays > 26)
                {
                    //now marking it
                    File.WriteAllText(sDateRecordPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    fi.LastWriteTime = DateTime.Now;

                    Thread nt = new Thread(new ThreadStart(AutoScoreThread));
                    nt.IsBackground = true;
                    nt.Start();
                }
            }
            
        }
        private void AutoScoreThread()
        {
            string sStartDate = DateTime.Now.ToString("yyyy-MM-") + data.iStartDay.ToString("00");
            if (DateTime.Now.Day<data.iStartDay)
            {
                sStartDate = DateTime.Now.AddMonths(-1).ToString("yyyy-MM-") + data.iStartDay.ToString("00");
            }

            //now for User Part
            data.sSQLRecent = " AND (Deleted is NULL and Blocked is NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'" + DateTime.Now.AddMonths(1).ToString("yyyy-MM-01") + "' )";
            string sSQL = "SELECT DISTINCT(FO) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + data.sSQLRecent;
            DataTable dt = data.msSQL.GetDataTable(sSQL);//("SELECT DISTINCT(FO) FROM FC_SESReport A WHERE Requisitioner IS NOT NULL AND ( TECO_Format IS NOT NULL OR CS_REC_Format IS NOT NULL ) ");
            if (dt!=null&&dt.Rows.Count>0)
            {
                //now auto scoring with 60%
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sContract_No = dt.Rows[i][0].ToString();

                    //for User Part
                    DataTable dtUser = data.msSQL.GetDataTable("SELECT DISTINCT(ISNULL(Requisitioner, 'Null')) FROM FC_SESReport A WHERE  (TECO_Format is not NULL or SES_CONF_Format is not NULL) AND  FO='" + sContract_No + "' " + data.sSQLRecent);
                    if (dtUser!=null&&dtUser.Rows.Count>0)
                    {
                        for (int n = 0; n < dtUser.Rows.Count; n++)
                        {
                            string sUser = dtUser.Rows[n][0].ToString();
                            updateInsertEvaluation(sContract_No, sUser, "User", "Auto", "6", "6", "6", "6", "6", "6");
                        }
                    }

                }
            }
            //now for Dep parts
            data.sSQLRecent = " AND (Deleted is NULL and Blocked is NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'" + DateTime.Now.ToString("yyyy-MM-01") + "' )";
            sSQL = "SELECT DISTINCT(FO) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + data.sSQLRecent + "";
            dt = data.msSQL.GetDataTable(sSQL);
            if (dt != null && dt.Rows.Count > 0)
            {
                //for Dep Part
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sSQLNames = "", sSQLValues = "";
                    sSQLNames = "Score1,Score2,Score3,Score4,Score5,Score6,";
                    sSQLValues = "3,3,3,3,3,3,";
                    //SELECT * from FC_Score WHERE Contract_No='A546058163' AND Role='Dep' and Score4>0
                    DataTable dtTmp = data.msSQL.GetDataTable("SELECT Score1,Score2,Score3,Score4,Score5,Score6 from FC_Score WHERE Contract_No='" + data.object2Str(dt.Rows[i][0]) + "' AND Role='Dep' and DateIn>='" + DateTime.Now.ToString("yyyy-MM-01") + "'");
                    if (dtTmp != null && dtTmp.Rows.Count > 0)
                    {

                        for (int n = 0; n < 6; n++)
                        {
                            for (int j = 0; j < dtTmp.Rows.Count; j++)
                            {
                                if (double.Parse(dtTmp.Rows[j][n].ToString()) > 0)
                                {
                                    sSQLNames = sSQLNames.Replace("Score" + (n + 1).ToString() + ",", "");
                                    sSQLValues = sSQLValues.Remove(0, 2);
                                    break;
                                }
                            }
                        }
                    }
                    else if (dtTmp != null && dtTmp.Rows.Count == 0)
                    {
                        sSQLNames = "Score1,Score2,Score3,Score4,Score5,Score6,";
                        sSQLValues = "3,3,3,3,3,3,";
                    }
                    if (!string.IsNullOrEmpty(sSQLNames) && sSQLNames.Length > 5)
                    {
                        sSQLNames += "Score7,";
                        sSQLValues += "0,";
                        bool b = data.msSQL.ExecuteQuery("INSERT INTO FC_Score(Contract_No," + sSQLNames + "[By],Role,Remark,DateIn) VALUES ('" + data.object2Str(dt.Rows[i][0]) + "'," + sSQLValues + "'AutoDep','Dep','Auto','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
                    } 
                }
            }


            //now at last for Score7 user part
            dt = data.msSQL.GetDataTable("SELECT DISTINCT(Contract_No) FROM FC_Score WHERE Role='User' AND DateIn>='" + DateTime.Now.ToString("yyyy-MM-01") + "' AND DateIn<='" + DateTime.Now.ToString("yyyy-MM-")+(data.iStartDay+1).ToString() + "'");
            if (dt!=null&&dt.Rows.Count>0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string sContract_NO = dt.Rows[i][0].ToString();
                    bool bNow = data.msSQL.ExecuteQuery("UPDATE FC_Score SET Score7=" + getScore7(sContract_NO).ToString("0.00") + "  WHERE Role='User' AND Contract_No='" + sContract_NO + "'  AND DateIn>='" + DateTime.Now.ToString("yyyy-MM-01") + "' AND DateIn<'" + DateTime.Now.ToString("yyyy-MM-") + (data.iStartDay + 1).ToString() + "'");
                }
            }
        }
        private double getScore7(string sContract_NO)
        {
            try
            {                
                DataTable dt = data.msSQL.GetDataTable("select CAST(AVG(Timely) as NUMERIC(10,2)),CAST(AVG(Honesty) as NUMERIC(10,2)) from contractScore7Real where Contract_No='" + sContract_NO + "' AND DateIn>='" + DateTime.Now.AddMonths(-1).ToString("yyyy-MM-" + data.iStartDay.ToString()) + "' AND DateIn<'" + DateTime.Now.ToString("yyyy-MM-" + (data.iStartDay + 1).ToString()) + "'");
                if (dt!=null&&dt.Rows.Count>0)
                {
                    //double dPercentage = 100*double.Parse(dt.Rows[0][1].ToString());
                    //double dConDays = double.Parse(dt.Rows[0][2].ToString());
                    double d1=3.0, d2=3.0;
                    //if (dPercentage <= 1) { d1 = 5.0; }
                    //else if (dPercentage <= 5) { d1 = 4.0; }
                    //else if (dPercentage <= 10) { d1 = 3.0; }
                    //else if (dPercentage <= 15) { d1 = 2.0; }
                    //else { d1 = 0.0; }

                    //if (dConDays <= 5) { d2 = 5.0; }
                    //else if (dConDays <= 10) { d2 = 4.0; }
                    //else if (dConDays <= 15) { d2 = 3.0; }
                    //else if (dConDays <= 20) { d2 = 2.0; }
                    //else if (dConDays <= 30) { d2 = 1.0; }
                    //else { d2 = 0.0; }

                    if (dt.Rows[0][0] != null && !string.IsNullOrEmpty(dt.Rows[0][0].ToString()))
                    {
                        d1 = double.Parse(dt.Rows[0][0].ToString());
                    }
                    if (dt.Rows[0][1] != null && !string.IsNullOrEmpty(dt.Rows[0][1].ToString()))
                    {
                        d2 = double.Parse(dt.Rows[0][1].ToString());
                    }
                                        
                    return (d1 + d2);
                }
            }
            catch (Exception)
            {                
                //throw;
            }
            return 6;
        }

        private void batchEmail()
        {
                //if (DateTime.Now.Day == 1 || DateTime.Now.Day == 4 || DateTime.Now.Day == 7 || DateTime.Now.Day == 10)

                //if (DateTime.Now.Day == 1 || DateTime.Now.Day == 4 || DateTime.Now.Day == 7 || DateTime.Now.Day == 10 || DateTime.Now.Day==14)

                if (DateTime.Now.Day == 1 || DateTime.Now.Day == 4 || DateTime.Now.Day == 7 || DateTime.Now.Day == 10 )

                {
                //now first of all,check if it necessary to update the records
                string sHostPath = MapPath(".");
                string sDateRecordPath = Path.Combine(sHostPath, "autoEmail.txt");
                DateTime dt = DateTime.Now.AddDays(-30);
                FileInfo fi = new FileInfo(sDateRecordPath);
                if (fi.Exists)
                {
                    dt = fi.LastWriteTime;
                }
                else
                {
                    File.WriteAllText(sDateRecordPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    fi.LastWriteTime = DateTime.Now;
                }
                if ((DateTime.Now - dt).TotalDays >= 2)
                {
                    //now marking it
                    File.WriteAllText(sDateRecordPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    fi.LastWriteTime = DateTime.Now;

                    Thread nt = new Thread(new ThreadStart(batchSendEmailThread));
                    nt.IsBackground = true;
                    nt.Start();
                }
            }
        }
        private void batchSendEmailThread()
        {
            try
            {
                int iNum = 0;
                string sInfo = "";
                data.sSQLRecent = " AND (Deleted is NULL and Blocked is NULL)  AND not EXISTS (SELECT * from FC_SESRecord B WHERE A.SES_No=B.SES_No AND B.DateIn<'" + DateTime.Now.AddMonths(1).ToString("yyyy-MM-01") + "' )";
                string sSQL = "SELECT DISTINCT(Requisitioner) from FC_SESReport A where (A.SES_CONF_Format is not NULL or A.TECO_Format is not NULL) " + data.sSQLRecent;//SELECT DISTINCT(Requisitioner) FROM FC_SESReport A WHERE Requisitioner IS NOT NULL AND ( TECO_Format IS NOT NULL OR CS_REC_Format IS NOT NULL ) AND NOT EXISTS ( SELECT * FROM FC_Score B WHERE A.FO = B.Contract_No AND UPPER (B.[By]) = UPPER (A.Requisitioner))

                DataTable dt = data.msSQL.GetDataTable(sSQL);
                if (dt!=null&&dt.Rows.Count>0)
                {
                    
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        if (dt.Rows[i][0]!=null)
                        {
                            string sUserEmail = data.getUserDetails(dt.Rows[i][0].ToString().Trim(),1);
                            string sUserName = data.getUserDetails(dt.Rows[i][0].ToString().Trim(),0);
                            //sending emails
                            if (!string.IsNullOrEmpty(sUserEmail)&&!string.IsNullOrEmpty(sUserName))
                            {
                                sendEmail(sUserName,dt.Rows[i][0].ToString(),sUserEmail,true,"");
                                sInfo += sUserName + "/" + sUserEmail+" As User<br/>";
                                iNum += 1;
                            }
                        }
                    }
                    
                }
                //for departments
                for (int i = 0; i < 6; i++)
                {
                    string[] sResults = data.getEmails(i);
                    if (sResults != null)
                    {
                        for (int j = 0; j < sResults.Length; j++)
                        {
                            string sDepName = "";
                            if (i==0)
                            {
                                sDepName = "CTE/D";
                            }
                            if (i == 1)
                            {
                                sDepName = "Main Coordinator";
                            }
                            if (i == 2)
                            {
                                sDepName = "User Representative";
                            }
                            if (i == 3)
                            {
                                sDepName = "CTS";
                            }
                            if (i == 4)
                            {
                                sDepName = "CHA";
                            }
                            if (i == 5)
                            {
                                sDepName = "CTM/T";
                            }
                            string[] ss = sResults[j].Split(new string[]{"[@]"},StringSplitOptions.RemoveEmptyEntries);
                            if (ss.Length==2)
                            {
                                if (!sInfo.Contains(ss[1]))
                                {
                                    string sUserName = data.getUserDetails(ss[0], 0);
                                    //for resending purposes 2013-09-05
                                    DataTable dtTmp = data.msSQL.GetDataTable("SELECT ID from FC_Score where LOWER([By])='" + ss[0].ToLower() + "' AND DATEDIFF(month, DateIn, GETDATE())<1");
                                    if (dtTmp!=null&&dtTmp.Rows.Count>0)
                                    {
                                        continue;
                                    }
                                    //ignore CTS's other people 2014-03-24
                                    if (sUserName.ToLower() == "ma shanhe" || sUserName.ToLower() == "liu benwen")
                                    {
                                        continue;
                                    }

                                    sendEmail(sUserName, ss[0], ss[1], false, sDepName);
                                    sInfo += sUserName + "/" + ss[1] + " As Dep.<br/>";
                                    iNum += 1; 
                                }
                            }
                        }
                    }
                    
                }
                // data.sendEmail("Gang.Ji@basf-ypc.com.cn", "Email Summary", "<P>Sent:" + iNum.ToString() + "</P>" + sInfo,sPath);
                data.sendEmail("fcl@basf-ypc.com.cn", "Email Summary", "<P>Sent:" + iNum.ToString() + "</P>" + sInfo, sPath);
            }
            catch (Exception)
            {
                //throw;
            }
        }
        private void sendEmail(string sUserName,string sUserID,string sEmailTo,bool isUser,string sDepName)
        {            
            //get user role
            string sRoleEn = " <STRONG>As User</STRONG>", sRoleCn = "<STRONG>以用户身份</STRONG>";
            bool isDep = false;
            if (isUser)
            {
                for (int i = 0; i < 6; i++)
                {
                    string[] sResults = data.getEmails(i);
                    if (sResults != null)
                    {
                        for (int j = 0; j < sResults.Length; j++)
                        {
                            if (sResults[j].ToLower().Contains(sUserID.ToLower() + "[@]"))
                            {
                                isDep = true;
                                if (i == 0)
                                {
                                    sDepName = "CTE/D";
                                }
                                if (i == 1)
                                {
                                    sDepName = "Main Coordinator";
                                }
                                if (i == 2)
                                {
                                    sDepName = "User Representative";
                                }
                                if (i == 3)
                                {
                                    sDepName = "CTS";
                                }
                                if (i == 4)
                                {
                                    sDepName = "CHA";
                                }
                                if (i== 5)
                                {
                                    sDepName = "CTM/T";
                                }
                                break;
                            }
                        }
                    }
                    if (isDep)
                    {
                        break;
                    }
                }

                if (isDep)
                {
                    sRoleEn = " <STRONG>As User And Dep.(" + sDepName + ")</STRONG>";
                    sRoleCn = "<STRONG>以用户以及职能部门身份(" + sDepName + ")</STRONG>";
                }
            }
            else
            {
                sRoleEn = " <STRONG>As Dep.("+sDepName+")</STRONG>";
                sRoleCn = "<STRONG>以职能部门身份("+sDepName+")</STRONG>";
            }

            DateTime dt=DateTime.Now;
            if (dt.Day>data.iStartDay)
            {
                dt = dt.AddMonths(1);
            }
            string MonthEn = dt.ToString("MMMM", System.Globalization.CultureInfo.GetCultureInfo("en-US").DateTimeFormat);

            //string sCon = "<P>Dear " + sUserName + ":</P><P>Please provide evaluation to the frame contractors and your valuable input is most important!<BR>请对框架承包商的绩效进行评估，您的评价非常重要！</P><P>Your evaluation will be closed on <STRONG>10th of " + MonthEn + "</STRONG>, and you should give the evaluation before.<BR>评价将在<STRONG>"+dt.Month.ToString()+"月"+(data.iStartDay-1).ToString()+"日</STRONG>前截止，请注意时间安排。</P><P>Please click the following link to start evaluation"+sRoleEn+". Thanks!<BR>请点击下面链接"+sRoleCn+"开始评价，谢谢！</P><P><A title=\"\" href=\"http://10.137.12.32/fcl/\" target=_blank><U><STRONG>http://10.137.12.32/fcl/</STRONG></U></A></P><P>You can find the instruction to guide the evaluation at the top left corner.<BR>您可以在左上角找到指导书引导您进行评估。</P>";

            string sCon = "<P>Dear " + sUserName + ":</P><P>Please provide evaluation to the frame contractors and your valuable input is most important!<BR>请对框架承包商的绩效进行评估，您的评价非常重要！</P><P>Your evaluation will be closed on <STRONG> 10th of " + MonthEn + "</STRONG>, and you should give the evaluation before.<BR>评价将在<STRONG>" + dt.Month.ToString() + "月" + (data.iStartDay - 1).ToString() + "日</STRONG>前截止，请注意时间安排。</P><P>Please click the following link to start evaluation" + sRoleEn + ". Thanks!<BR>请点击下面链接" + sRoleCn + "开始评价，谢谢！</P><P><A title=\"\" href=\"https://fcl.basf-ypc.com.cn:8443/fcl/\" target=_blank><U><STRONG>https://fcl.basf-ypc.com.cn:8443/fcl/</STRONG></U></A></P><P>You can find the instruction to guide the evaluation at the top left corner.<BR>您可以在左上角找到指导书引导您进行评估。</P>";

            data.sendEmail(sEmailTo, "FC Contractor Performance Evaluation", sCon, sPath);
            //data.sendEmail(sEmailTo, "请尽快完成承包商评估，谢谢!Please complete your FC Contractor Performance Evaluation,Thanks!", sCon, sPath);
        }
        
        protected void RefreshProgress(object sender, DirectEventArgs e)
        {
            try
            {
                string sProgressPath = Path.Combine(System.Web.Hosting.HostingEnvironment.MapPath("/"), "progress.txt");
                int iProgress = 0;
                if (File.Exists(sProgressPath))
                {
                    iProgress = int.Parse(File.ReadAllText(sProgressPath));
                    X.Msg.UpdateProgress((iProgress) / 100f, string.Format("Processing {0} of {1}...", iProgress.ToString(), 100));
                }
                else
                {
                    //iProgress = 100;
                }
                if (Session["progress"]!=null&&((int)Session["progress"])==100)
                {
                    Session["progress"] = null;
                    this.TaskManager1.StopTask("Task1");
                    X.MessageBox.Hide();
                    this.ResourceManager1.AddScript("Ext.Msg.notify('Done', 'Excel has been exported!');");
                }
            }
            catch (Exception)
            {
                //throw;
            }
        }

        [DirectMethod]
        public void setKPIMOM(string sContractNO,string sContent,string sDateMonth)
        {
            sDateMonth = sDateMonth.Substring(0, 4) + "-"+sDateMonth.Substring(4,2);
            DataTable dt = data.msSQL.GetDataTable("select * from KPI where ContractNO='"+sContractNO+ "' and DateMonth='"+sDateMonth+"'");
            bool b = false;
            if (dt!=null&&dt.Rows.Count>0)
            {
                b = data.msSQL.ExecuteQuery("update KPI set MOM='" + sContent.Trim() + "',DateIn='"+ DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where ContractNO='" + sContractNO + "' and DateMonth='"+sDateMonth+"'");
            }
            else
            {
                b = data.msSQL.ExecuteQuery("insert into KPI(ContractNO,MOM,DateMonth,DateIn) values('" + sContractNO+"','"+sContent.Trim()+"','"+sDateMonth+"','" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')");
            }
            if (b)
            {
                Notification.Show(new NotificationConfig
                {
                    Title = "Information",
                    Icon = Icon.Information,
                    Html = "MOM of " + sContractNO + " Updated!"
                });
                storePie1.Reload();
            }
            else
            {
                Notification.Show(new NotificationConfig
                {
                    Title = "Error",
                    Icon = Icon.Error,
                    Html = "Failed in updating MOM of " + sContractNO + "!"
                });
            }
        }
    }
}