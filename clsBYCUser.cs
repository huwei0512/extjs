using System;
using System.Collections.Generic;
using System.Text;
using System.DirectoryServices;
using System.DirectoryServices.Protocols;

namespace BYC
{
    class clsBYCUser
    {
        string sUserName,sPassword,sError;
        DirectoryEntry de = null;
        public clsBYCUser(string sNovelID,string sPass)
        {
            sUserName = sNovelID;
            sPassword = sPass;              
        }
        public string UserID
        {
            get 
            {                
                return getPropertyValue("cn");//uid 
            }
            //set { sUserName = value; }
        }
        public bool checkUser()
        {
            if (string.IsNullOrEmpty(getPropertyValue("cn"))&&!string.IsNullOrEmpty(sError))
            {
                return false;
            }
            return true;
        }
        public string Orgnization
        {
            get { return getPropertyValue("o");}            
        }
        public string Department
        {
            get { return getPropertyValue("ou"); }
        }
        public string EMail
        {
            get { return getPropertyValue("mail"); }
        }
        public string FullName
        {
            get { return getPropertyValue("sn")+" "+getPropertyValue("givenName"); }
        }
        public string TelephoneNumber
        {
            get { return getPropertyValue("telephoneNumber"); }
        }
        public string PostalCode
        {
            get { return getPropertyValue("postalCode"); }
        }
        public string Street
        {
            get { return getPropertyValue("street"); }
        }
        public string LocationCity
        {
            get { return getPropertyValue("l"); }
        }
        public string LastError
        {
            get { return sError; }
        }

        public void close()
        {
            try
            {
                de.Close();
                de.Dispose();
            }
            catch (Exception)
            {

                //throw;
            }
        }

        private string getPropertyValue(string sPropertyName)
        {
            try
            {
                //string targetOU = "cn=" + sUserName + ",ou=BASF-YPC COMPANY LIMITED,ou=EMPLOYEES,o=auth";

                string targetOU = "cn=" + sUserName + ",ou=Users,ou=7430BASF,DC=byc,DC=com,DC=cn";

                if (de == null)
                {
                    //de = new DirectoryEntry("LDAP://10.4.21.214:389", targetOU, sPassword, AuthenticationTypes.None);

                    de = new DirectoryEntry("LDAP://10.137.1.30:389", targetOU, sPassword, AuthenticationTypes.None);
                }

                DirectorySearcher ds = new DirectorySearcher();
                ds.SearchRoot = de;
                //ds.Filter = ("(objectClass=group)");
                ds.Filter = "cn=" + sUserName;
                //ds.Filter = ("(objectCategory=YBS)(objectClass=user)") ;
                //. Find("ybs", "Group")) 


                foreach (SearchResult result in ds.FindAll())
                {
                    ResultPropertyValueCollection pvc = result.Properties[sPropertyName];
                    if (pvc.Count > 0)
                    {
                        return pvc[0].ToString();
                    }
                    //string name = result.GetDirectoryEntry().Name.ToString();                    
                    //DirectoryEntry deGroup = new DirectoryEntry(result.Path, targetOU, sPassword, AuthenticationTypes.None);
                    //System.DirectoryServices.PropertyCollection pcoll = deGroup.Properties;
                    //int n = pcoll["member"].Count;

                }
            }
            catch (Exception ex)
            {
                sError = ex.Message;
                //throw;
            }            
            return "";
        }
    }
}
