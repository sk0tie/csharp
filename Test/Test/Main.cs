using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;
using System.IO;

namespace Test
{
	class MainClass
	{
		public static void Main (string[] args)
		{
			DirectoryEntry objDE = new DirectoryEntry("LDAP://ldap.main.cobaltgroup.com:389/ou=Groups,dc=main,dc=cobaltgroup,dc=com","MAIN\\ldap","@DAccess",AuthenticationTypes.Secure);
            foreach (DirectoryEntry objDEChild in objDE.Children)
            {
				Console.WriteLine(objDEChild.Name);
			}
		}
	}
}

