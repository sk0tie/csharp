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
			FileInfo objFile = new FileInfo("swxGroups.txt");
			StreamWriter objSW = objFile.CreateText();
			Console.Write("Processing.. ");
			DirectoryEntry objDE = new DirectoryEntry("LDAP://ldap.main.cobaltgroup.com:389/ou=Test,dc=main,dc=cobaltgroup,dc=com","MAIN\\ldap","@DAccess",AuthenticationTypes.Secure);
            foreach (DirectoryEntry objDEChild in objDE.Children)
            {
				foreach (string objClass in objDEChild.Properties["objectclass"])
				{
					if (objClass == "group")
					{
						foreach (object objProxyAddress in objDEChild.Properties["proxyaddresses"])
						{
							if (objProxyAddress.ToString().ToLower().StartsWith("smtp:") == true)
							{
								Console.WriteLine(objDEChild.Path.ToString() + "|" + objProxyAddress.ToString() + "|" + objDEChild.Name.ToString());
								objSW.WriteLine(objDEChild.Path.ToString() + "|" + objProxyAddress.ToString() + "|" + objDEChild.Name.ToString());
							}
						}
					}
				}
			}
			Console.WriteLine("Done.");
			objSW.Close();
        }
    }
}