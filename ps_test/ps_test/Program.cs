using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Collections.ObjectModel;
using System.Management.Automation.Runspaces;
using System.Management.Automation.Remoting;

namespace ps_test
{
    class Program
    {
        static void Main(string[] args)
        {
            #region vars

            bool useRunspace = false;
            bool useRunspace2 = true;
            bool usePipeline = false;
            Runspace rs = null;
            string kerberosUri = "http://CHAADSYNC.CORR.ROUND-ROCK.TX.US/PowerShell";
            string chaadUri = "http://CHAADSYNC.CORR.ROUND-ROCK.TX.US";
            string schemaUri = "http://schemas.microsoft.com/powershell/Microsoft.Exchange";
            string schemaUriCHAAD = "http://schemas.microsoft.com/powershell";
            string server = "CHAADSYNC.CORR.ROUND-ROCK.TX.US";
            Collection<PSObject> results = null;

            #endregion vars

            #region setupcreds

            Console.Write("Enter password: ");
            string pass = Console.ReadLine();
            System.Security.SecureString secPass = new System.Security.SecureString();
            foreach(char c in pass)
            {
                secPass.AppendChar(c);
            }
            PSCredential credentials = new PSCredential(@"corr\jmcarthur", secPass);

            //WSManConnectionInfo connectionInfo = new WSManConnectionInfo(
            //    new Uri(kerberosUri),
            //    schemaUriCHAAD, credentials);
            //connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;

            #endregion setupcreds

            #region runspace

            if (useRunspace)
            {
                rs = RunspaceFactory.CreateRunspace();
                rs.Open();
                PowerShell ps = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("New-PSSession");
                command.AddParameter("ConfigurationName", "Microsoft.Exchange");
                command.AddParameter("ConnectionUri", kerberosUri);
                command.AddParameter("Credential", credentials);
                command.AddParameter("Authentication", "kerberos");
                ps.Commands = command;
                ps.Runspace = rs;
                Collection<PSSession> result = ps.Invoke<PSSession>();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "ra");
                command.AddParameter("Value", result[0]);
                ps.Commands = command;
                ps.Runspace = rs;
                ps.Invoke();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Import-PSSession -Session $ra");
                ps.Commands = command;
                ps.Runspace = rs;
                ps.Invoke();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Get-Recipient");
                command.AddParameter("ResultSize", "10");
                ps.Commands = command;
                ps.Runspace = rs;
                results = ps.Invoke();
            }
            else if (useRunspace2)
            {
                rs = RunspaceFactory.CreateRunspace();
                rs.Open();
                PowerShell ps = PowerShell.Create();
                PSCommand command = new PSCommand();
                command.AddCommand("New-PSSession");
                command.AddParameter("ComputerName", server);
                command.AddParameter("Credential", credentials);
                command.AddParameter("Authentication", "Kerberos");
                ps.Commands = command;
                ps.Runspace = rs;
                Collection<PSSession> result = ps.Invoke<PSSession>();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddCommand("Set-Variable");
                command.AddParameter("Name", "ra");
                command.AddParameter("Value", result[0]);
                ps.Commands = command;
                ps.Runspace = rs;
                ps.Invoke();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Import-PSSession -Session $ra");
                ps.Commands = command;
                ps.Runspace = rs;
                ps.Invoke();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddScript("Import-Module -Session $ra");
                //command.AddCommand("Import-Module");
                command.AddArgument(@"C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1");
                ps.Commands = command;
                ps.Runspace = rs;
                //results = ps.Invoke();

                ps = PowerShell.Create();
                command = new PSCommand();
                command.AddScript(@"Invoke-Command -Session $ra -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}");
                ps.Commands = command;
                ps.Runspace = rs;
                try
                {
                    results = ps.Invoke();
                }
                catch (Exception e)
                {
                    Console.WriteLine(String.Format("Caught exception running sync: {0}", e.Message));
                }
            }
            //else if (usePipeline)
            //{
            //    rs = RunspaceFactory.CreateRunspace(connectionInfo);
            //    rs.Open();
            //    Pipeline pipe = rs.CreatePipeline();
            //    pipe.Commands.Add("Import-Module");
            //    var command = pipe.Commands[0];
            //    command.Parameters.Add("Name", @"C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync\ADSync.psd1");
            //    results = pipe.Invoke();
            //}

            #endregion runspace

            #region wrapup

            foreach (Object res in results)
            {
                Console.WriteLine(res);
            }
            Console.WriteLine("That's all, folks");
            Console.ReadKey();

            #endregion wrapup
        }

            #region LDAP stuff

            public static Collection<PSObject> GetUsersUsingKerberos(
                        string kerberosUri, string schemaUri, PSCredential credentials, int count)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo(
                new Uri(kerberosUri),
                schemaUri, credentials);
            connectionInfo.AuthenticationMechanism = AuthenticationMechanism.Kerberos;

            using (Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo))
            {
                return GetUserInformation(count, runspace);
            }
        }

        public static Collection<PSObject> GetUserInformation(int count, Runspace runspace)
        {
            using (PowerShell powershell = PowerShell.Create())
            {
                //powershell.AddCommand("Get-User");
                //powershell.AddParameter("ResultSize", count);
                powershell.AddCommand("Enable-RemoteMailbox");
                powershell.AddParameter("Identity", "btables");
                powershell.AddParameter("Alias", "btables");
                powershell.AddParameter("RemoteRoutingAddress", "btables@roundrocktexas.mail.onmicrosoft.com");

                runspace.Open();

                powershell.Runspace = runspace;

                return powershell.Invoke();
            }
        }
    }
            # endregion
}
