/* Password Expiry Notification (swxPEN.exe)
 * Written by: Scott Wilcox
 * The Cobalt Group, Inc.
 * Date: 07/16/2010
 * Platform: x86
 * Environment: Visual Studio C# 2010 Express
 */

// Description

/* C# Application utilizing the System.DirectoryServices.AccountManagement libraries
 * to access Active Directory Groups and Members and retrive their password expiration
 * date. If a user is within the 14 day notification period, and e-mail will be sent
 * to them, informing them of their impending doom if they do not change it.
 */

// Namespaces

/* Declare the namespaces we'll be using in this application, specifically 'System' and
 * 'System.DirectoryServices.AccountManagement'.  The bulk of this applications logic
 * will utilize these libraries.
 */

using System;
using System.Text;
using System.IO;
using System.DirectoryServices.AccountManagement;
using System.Net.Mail;

namespace swxPEN
{
    class swxPEN
    {
        // Main Entry Point

        static void Main(string[] args)
        {
            // Global Variables for the Mail Subsystem.

            string strDomain = null;
            string strGroup = null;
            string strSmartHost = null;
            string strMailFrom = null;
            string strMailFromDisplay = null;
            string strMailSubject = null;
            string strMailBody = null;

            /* Since this will likely be run daily through some sort of automation, it would
             * be a good idea to create a log file mostly to track who's been receiving
             * notifications but more importantly to track exception errors due to mis-
             * configuration.
             */

            // Create our Log File Object if it doesn't exsist and bind to it.

            StreamWriter swLog;
            if (!File.Exists("swxPEN.log")) { swLog = new StreamWriter("swxPEN.log"); }
            else { swLog = File.AppendText("swxPEN.log"); }

            /* This application will use a configuration file to store the domain, group and
             * mail envolope. If the file doesn't exists, we'll create it with defaults.
             * We'll also use a second plain text file for the message body to easily
             * pull in a formatted message. Once the files are created we'll read back from
             * them it to populate our global variables.
             */

            // Wrap the bulk of our code in a try/catch block so we can cast errors to the log.

            try
            {
                // If the configuration file doesn't exist..
                if (!File.Exists("swxPEN.ini"))
                {
                    // Cast to the log that it will be created.
                    swLog.WriteLine(DateTime.Now + ": WARNING | swxPEN.ini file missing; default configuration will be created.");

                    // Build a new File Object and append to the file with defaults.
                    StreamWriter swIni = new StreamWriter("swxPEN.ini");
                    swIni.WriteLine("# swxPEN.ini Configuration File");
                    swIni.WriteLine("domain=mydomain.com");
                    swIni.WriteLine("group=domain users");
                    swIni.WriteLine("smartHost=mail.mydomain.com");
                    swIni.WriteLine("mailFrom=helpdesk@mydomain.com");
                    swIni.WriteLine("mailFromDisplay=Helpdesk");
                    swIni.WriteLine("mailSubject=REMINDER: Your domain password will in %expireDays days.");

                    // Close the file to free up resources.
                    swIni.Close();
                }

                // If the message body file doesn't exist...
                if (!File.Exists("swxPEN.msg"))
                {
                    StreamWriter swMessage = new StreamWriter("swxPEN.msg");
                    swMessage.WriteLine("%fullName: Your domain password will expire in %expireDays days on %expireDate");
                    swMessage.Close();
                }

                // If the configuration file does exist...
                if (File.Exists("swxPEN.ini"))
                {
                    // Create our File Object and read in each line until the end of the file.
                    StreamReader srIni = new StreamReader("swxPEN.ini");
                    while (srIni.Peek() > -1)
                    {
                        string strLineIn = srIni.ReadLine();
                        if (!strLineIn.StartsWith("#"))
                        {
                            /* If the line doesn't start with a comment, split it with the '='
                             * and populate their respective global variables.
                             */
                            string[] strSplit = strLineIn.Split(new Char[] { '=' });
                            if (strSplit[0] == "domain") { strDomain = strSplit[1]; }
                            if (strSplit[0] == "group") { strGroup = strSplit[1]; }
                            if (strSplit[0] == "smartHost") { strSmartHost = strSplit[1]; }
                            if (strSplit[0] == "mailFrom") { strMailFrom = strSplit[1]; }
                            if (strSplit[0] == "mailFromDisplay") { strMailFromDisplay = strSplit[1]; }
                            if (strSplit[0] == "mailSubject") { strMailSubject = strSplit[1]; }
                        }
                    }

                    // Close the configuration file.
                    srIni.Close();
                }

                // If the message body file exists...
                if (File.Exists("swxPEN.msg"))
                {
                    // Create our File Object and read in the entire message body.
                    StreamReader srMessage = new StreamReader("swxPEN.msg");
                    strMailBody = srMessage.ReadToEnd();

                    // Close the message body file.
                    srMessage.Close();
                }

                // Cast to our log the domain we'll be using.
                swLog.WriteLine(DateTime.Now + ": Contacting " + strDomain + "...");

                // Build our entry object (This is like DirectoryEntry, just simplified).
                PrincipalContext pcContext = new PrincipalContext(ContextType.Domain, strDomain);

                // Find our group by it's identity.
                GroupPrincipal gpGroup = GroupPrincipal.FindByIdentity(pcContext, IdentityType.Name, strGroup);

                // Cast to the log that we've begun scanning our group members and where mail will be sent from.
                swLog.WriteLine(DateTime.Now + ": Scanning " + strGroup + " for domain passwords set to expire within 14 days...");
                swLog.WriteLine(DateTime.Now + ": Notifications will be sent from " + strMailFrom);

                /* Iterate through each member in our group (this will also return members of all
                 * nested groups within). Grab the date their password was set and determine when
                 * their password will expire.  Notify them if their password will expire within
                 * the next 14 days. Exlude those users that have disabled accounts or their
                 * password has already expired.
                 */

                foreach (UserPrincipal upUser in gpGroup.GetMembers(true))
                {
                    DateTime dtLastPasswordSet = Convert.ToDateTime(upUser.LastPasswordSet);
                    TimeSpan tsPasswordExpire = dtLastPasswordSet.AddDays(45) - DateTime.Now;

                    // If enabled and password is set to expire within 14 days..
                    if ((upUser.Enabled == true) && (tsPasswordExpire.TotalDays < 15) && (tsPasswordExpire.TotalDays > 0))
                    {
                        /* Since we want the mail subject and body to utilize variables such as
                         * the expiration date and how long until their password expires, we can
                         * replace virtual string variables (text that starts with %) with
                         * information we've collected during the scan.  We'll use a String
                         * Builder Object instead of using the explicit replace method of
                         * string variables. This will speed up the replacement process and
                         * utilize less memory.
                         */

                        StringBuilder sbMailSubject = new StringBuilder(strMailSubject);
                        StringBuilder sbMailBody = new StringBuilder(strMailBody);

                        // Replace our virtual string variables with real data we've collected.
                        sbMailSubject.Replace("%expireDate", dtLastPasswordSet.AddDays(45).ToString());
                        sbMailSubject.Replace("%expireDays", tsPasswordExpire.Days.ToString());
                        sbMailBody.Replace("%expireDate", dtLastPasswordSet.AddDays(45).ToString());
                        sbMailBody.Replace("%expireDays", tsPasswordExpire.Days.ToString());
                        sbMailBody.Replace("%fullName", upUser.DisplayName.ToString());

                        // Create a mail object for each person we'll be sending mail to.
                        MailMessage message = new MailMessage();

                        // Populate the envelope with who the mail is from, and a proper display name.
                        message.From = new MailAddress(strMailFrom, strMailFromDisplay);

                        // Populate the envelope with who the mail is to, and a proper display name.
                        message.To.Add(new MailAddress(upUser.EmailAddress, upUser.Name));
                        message.Priority = MailPriority.High;

                        // Populate the envelope with the subject and body with replaced variables.
                        message.Subject = sbMailSubject.ToString();
                        message.Body = sbMailBody.ToString();

                        // Bind to the Mail Subsystem, use the configured smarthost and send the message.
                        SmtpClient client = new SmtpClient();
                        client.Host = strSmartHost;
                        client.Send(message);

                        // Cast to the log that the user has been notified.
                        swLog.WriteLine(DateTime.Now + ": " + upUser.Name + " | Password expires in " + tsPasswordExpire.Days + " Days. Notification sent to " + upUser.EmailAddress);
                    }
                }

                // Cast to the log that the process is completed.
                swLog.WriteLine(DateTime.Now + ": Completed scan.");
            }

            // In the event that any problems occured in our try/catch block, write the error to the log.
            catch (Exception e) { swLog.WriteLine(DateTime.Now + ": ERROR | " + e); }

            // Close the log file.
            swLog.Close();
        }
    }
}