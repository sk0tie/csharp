using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.IO;

namespace swxProxyAddress
{
	class Program
    {
        static void Main(string[] args)
        {
			FileInfo objFile = new FileInfo("swxProxyAddresses.txt");
			StreamWriter objSW = objFile.CreateText();
			Console.Write("Processing.. ");
			DirectoryEntry objDE = new DirectoryEntry("LDAP://ldap.main.cobaltgroup.com:389/cn=RESOURCES,cn=Users,dc=main,dc=cobaltgroup,dc=com","MAIN\\ldap","@DAccess",AuthenticationTypes.Secure);
            foreach (DirectoryEntry objDEChild in objDE.Children)
            {
				foreach (string objClass in objDEChild.Properties["objectclass"])
				{
					if ((objClass == "organizationalPerson") && (objDEChild.Properties["proxyaddresses"].Count > 0))
					{
						objSW.Write(objDEChild.Name);
						
						List<string> arrProxyAddresses = new List<string>();
												
						foreach (object objProxyAddress in objDEChild.Properties["proxyaddresses"])
						{
							string strProxyType = objProxyAddress.ToString().ToLower();
							
							if ((strProxyType.StartsWith("smtp")) || (strProxyType.StartsWith("x")))
							{
								arrProxyAddresses.Add(objProxyAddress.ToString());
							}
						}
						
						arrProxyAddresses.Sort(StringComparer.Ordinal);
						
						foreach (string strProxyAddress in arrProxyAddresses)
						{
							objSW.Write("|" + strProxyAddress);
						}
						objSW.WriteLine();
					}
				}
			}
			Console.WriteLine("Done.");
			objSW.Close();
        }
    }
}