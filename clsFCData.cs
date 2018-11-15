using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FCPortal
{
    public class clsFCData
    {
        private string sFcNo, sworkCenter, sfcDescription, scontractor, scontractAdmin, sbuyer,
            smainCoordinator, suserRepresentative, scontact, stelephone, sPricing_Scheme,sRemark;
        private DateTime? dvalidateDate, dexpireDate;
        private float fScore, fbugdet, fnetAmount, funusedBudget, funusedPercentage, fremainingPercentage;
        public clsFCData(string FCNO)
        {
            sFcNo = FCNO;
        }
        public string NO
        {
            get { return sFcNo; }
            set { sFcNo = value; }
        }
        public string fcDescription
        {
            get { return sfcDescription; }
            set { sfcDescription = value; }
        }
        public string contractor
        {
            get { return scontractor; }
            set { scontractor = value; }
        }
        public float bugdet
        {
            get { return fbugdet; }
            set { fbugdet = value; }
        }
        public float netAmount
        {
            get { return fnetAmount; }
            set { fnetAmount = value; }
        }
        public float unusedBudget
        {
            get { return funusedBudget; }
            set { funusedBudget = value; }
        }
        public float unusedPercentage
        {
            get { return funusedPercentage; }
            set { funusedPercentage = value; }
        }
        public float remainingPercentage
        {
            get { return fremainingPercentage; }
            set { fremainingPercentage = value; }
        }
        public string workCenter
        {
            get { return sworkCenter; }
            set { sworkCenter = value; }
        }
        public string contractAdmin
        {
            get { return scontractAdmin; }
            set { scontractAdmin = value; }
        }
        public string buyer
        {
            get { return sbuyer; }
            set { sbuyer = value; }
        }
        public string mainCoordinator
        {
            get { return smainCoordinator; }
            set { smainCoordinator = value; }
        }
        public string userRepresentative
        {
            get { return suserRepresentative; }
            set { suserRepresentative = value; }
        }
        public string contact
        {
            get { return scontact; }
            set { scontact = value; }
        }
        public string telephone
        {
            get { return stelephone; }
            set { stelephone = value; }
        }  
        public float score
        {
            get { return fScore; }
            set { fScore = value; }
        }
        public string pricingScheme
        {
            get { return sPricing_Scheme; }
            set { sPricing_Scheme = value; }
        }
        public string remark
        {
            get { return sRemark; }
            set { sRemark = value; }
        }  
        
        public DateTime? validateDate
        {
            get { return dvalidateDate; }
            set { dvalidateDate = value; }
        }
        public DateTime? expireDate
        {
            get { return dexpireDate; }
            set { dexpireDate = value; }
        }
       
    }
    public class clsFCDataPie
    {
        private string sPricing_Type="", sCT_Director="",sMOM="";
        double dNum = 0,dNum1=0;
        public string Type
        {
            get { return sPricing_Type; }
            set { sPricing_Type = value; }
        }
        public string Name
        {
            get { return sCT_Director; }
            set { sCT_Director = getDepartment(value); }
        }
        public string MOM
        {
            get { return sMOM; }
            set { sMOM = value; }
        }
        public double Data1
        {
            get { return dNum; }
            set {
                    if (value > 1000)
                    {
                        dNum = (int)Math.Floor(value);
                    }
                    else
                    {
                        dNum = Math.Round(value, 2);
                    }
                 }
        }
        public double Data2
        {
            get { return dNum1; }
            set { dNum1 = Math.Round(value,2); }
        }
        public static string getDepartment(string sDirector)
        {
            if (sDirector.Contains("Alexander Wiesel") || sDirector.Contains("Bram Jansen"))
            {
                return "CTE";
            }
            if (sDirector.Contains("Xu Zhixian") || sDirector.Contains("Zhu Ruisong"))
            {
                return "CTM";
            }
            if (sDirector.Contains("Wang Changjun"))
            {
                return "CTA";
            }
            if (sDirector.Contains("Yang Huafei"))
            {
                return "CTS";
            }
            return sDirector;
        }
    }
    public class clsFCDataPie2
    {
        private string sContract_NO, sMonth;
        int iuserScore, idepScore;

        public string Contract_NO
        {
            get { return sContract_NO; }
            set { sContract_NO = value; }
        }
        public string Month
        {
            get { return sMonth; }
            set { sMonth = value; }
        }
        public int userScore
        {
            get { return iuserScore; }
            set { iuserScore = value; }
        }
        public int depScore
        {
            get { return idepScore; }
            set { idepScore = value; }
        }
        public int allScore
        {
            get { return idepScore+iuserScore; }
        }
    }

    public class clsFCSESReport
    {
        private string sFO,sSES_No = "", sShort_Descrption = "", sStart_Date, sEnd_Date, sTECO_Date, sRequisitioner, sClaim_sheets_receive;
        DateTime? dTECO_Format, dCS_REC_Format;
        public string FO
        {
            get { return sFO; }
            set { sFO = value; }
        }
        public string SES_No
        {
            get { return sSES_No; }
            set { sSES_No = value; }
        }
        public string Short_Descrption
        {
            get { return sShort_Descrption; }
            set { sShort_Descrption = value; }
        }
        public string Start_Date
        {
            get { return sStart_Date; }
            set { sStart_Date = value; }
        }
        public string End_Date
        {
            get { return sEnd_Date; }
            set { sEnd_Date = value; }
        }
        public string TECO_Date
        {
            get { return sTECO_Date; }
            set { sTECO_Date = value; }
        }
        public DateTime? TECO_Format
        {
            get { return dTECO_Format; }
            set { dTECO_Format = value; }
        }
        public string Requisitioner
        {
            get { return sRequisitioner; }
            set { sRequisitioner = value; }
        }
        public string SES_Confirmed_on
        {
            get { return sClaim_sheets_receive; }
            set { sClaim_sheets_receive = value; }
        }
        public DateTime? SES_CONF_Format
        {
            get { return dCS_REC_Format; }
            set { dCS_REC_Format = value; }
        }
    }

    public class clsFCEvaluation
    {
        private string sUser = "", sContract_No = "", sFileID;
        double iScore1, iScore2, iScore3, iScore4, iScore5, iScore6, iScore7;
        bool bEvaluated;
        public string User
        {
            get { return sUser; }
            set { sUser = value; }
        }
        public string Contract_No
        {
            get { return sContract_No; }
            set { sContract_No = value; }
        }
        public double Score1
        {
            get { return iScore1; }
            set { iScore1 = value; }
        }
        public double Score2
        {
            get { return iScore2; }
            set { iScore2 = value; }
        }
        public double Score3
        {
            get { return iScore3; }
            set { iScore3 = value; }
        }
        public double Score4
        {
            get { return iScore4; }
            set { iScore4 = value; }
        }
        public double Score5
        {
            get { return iScore5; }
            set { iScore5 = value; }
        }
        public double Score6
        {
            get { return iScore6; }
            set { iScore6 = value; }
        }
        public double Score7
        {
            get { return iScore7; }
            set { iScore7 = value; }
        }
        public bool Evaluated
        {
            get { return bEvaluated; }
            set { bEvaluated = value; }
        }
        public string FileID
        {
            get { return sFileID; }
            set { sFileID = value; }
        }
    }


    public class clsContractFile
    {
        private string sFileID = "", sFileName = "", sUploadUser, sRemark, sContract_NO, sFileLengthStr, sType;
        int iFileLength;
        DateTime? dDateIn;
        public string FileID
        {
            get { return sFileID; }
            set { sFileID = value; }
        }
        public string FileName
        {
            get { return sFileName; }
            set { sFileName = value; }
        }
        public int FileLength
        {
            get { return iFileLength; }
            set { iFileLength = value; }
        }
        public string FileLengthStr
        {
            get { sFileLengthStr = data.FormatBytes(iFileLength); return sFileLengthStr; }
            set { sFileLengthStr = value; }
        }
        public string UploadUser
        {
            get { return sUploadUser; }
            set { sUploadUser = value; }
        }
        public string Contract_NO
        {
            get { return sContract_NO; }
            set { sContract_NO = value; }
        }
        public string Remark
        {
            get { return sRemark; }
            set { sRemark = value; }
        }
        public string Type
        {
            get { return sType; }
            set { sType = value; }
        }
        public DateTime? DateIn
        {
            get { return dDateIn; }
            set { dDateIn = value; }
        }
    }

}