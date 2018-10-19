using System;
using System.Security.Principal;

namespace swxSIDInfo
{
	class MainClass
	{
		public static void Main (string[] args)
		{
			string strUsername = String.Format(@"{0}\{1}", Environment.UserDomainName, Environment.UserName);
			WindowsIdentity objWI = new WindowsIdentity(strUsername);
			Console.WriteLine(objWI.User.AccountDomainSid.ToString());
			Console.ReadLine();
		}
	}
}

