using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.IO;

namespace swxProxyAddress
{   class Program
    {
        static void Main(string[] args)
        {
            FileInfo objFileIN = new FileInfo("external_addresses.txt");
            FileInfo objFileOUT = new FileInfo("groups");
            StreamWriter objSW = objFileOUT.CreateText();

            using (StreamReader objSR = objFileIN.OpenText())
            {
                string strAddress = "";
                while ((strAddress = objSR.ReadLine()) != null)
                {
                    DirectoryEntry objDE = new DirectoryEntry("LDAP://ldap.main.cobaltgroup.com:389/dc=main,dc=cobaltgroup,dc=com","MAIN\\ldap","@DAccess",AuthenticationTypes.Secure);
                    DirectorySearcher objSearcher = new DirectorySearcher();
                    objSearcher.SearchRoot = objDE;
                    objSearcher.Filter = "(proxyAddresses=SMTP:" + strAddress + ")";
                    SearchResult objResult = objSearcher.FindOne();
                    if (objResult != null)
                    {
                        foreach (object objClass in objResult.Properties["objectclass"])
                        {
                            if (objClass.ToString() == "group")
                            {

                                try
                                {
									Console.WriteLine(strAddress + "|" + objResult.Properties["name"][0].ToString());
									objSW.WriteLine(strAddress + "|" + objResult.Properties["name"][0].ToString());
                                }
                                catch
                                {
									Console.WriteLine(strAddress + "|missingname");
                                    objSW.WriteLine(strAddress + "|missingname");
                                }
                            }
                        }
                    }
                }
            }
            objSW.Close();
            Console.ReadKey();
        }
    }
}