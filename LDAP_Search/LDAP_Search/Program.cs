using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using Newtonsoft.Json;
using System.IO;

namespace LDAP_Search
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("It is currently {0}", System.DateTime.Now);
            //Console.WriteLine("One hour ago it was {0}", System.DateTime.Now.AddHours(-1));
            Console.Write("Enter user (login): ");
            String username = Console.ReadLine();

            try
            {
                // create LDAP connection object  

                DirectoryEntry myLdapConnection = createDirectoryEntry();

                //Console.WriteLine("DirectoryEntry.Path: {0}", myLdapConnection.Path);
                //Console.WriteLine("DirectoryEntry.Parent.Path: {0}", myLdapConnection.Parent.Path);
                //Console.ReadKey();

                // create search object which operates on LDAP connection object  
                // and set search object to only find the user specified  

                //Console.WriteLine("Using base DN: {0}", myLdapConnection.Path);

                //string baseDN = "OU=roundrocktexas.gov,DC=corr,DC=round-rock,DC=tx,DC=us";
                string[] PropertiesToLoad = { "department", "division", "title", "displayname", "mail", "physicaldeliveryofficename" };
                string ldapFilter = "(&(objectClass=user)(samaccountname=" + username + ")(mail=*)(objectCategory=person)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))";

                DirectorySearcher search = new DirectorySearcher(myLdapConnection, ldapFilter, PropertiesToLoad);
                search.Filter = "(samaccountname=" + username + ")";

                //search.Filter = "(objectClass=organizationalUnit)";
                
                // create results objects from search object

                SearchResult result = search.FindOne();

                if (result != null)
                {
                    // user exists, cycle through LDAP fields (cn, telephonenumber etc.)  

                    ResultPropertyCollection fields = result.Properties;

                    if (fields.Contains("title"))
                    {
                        Console.WriteLine("Found {0} with value {1}.", "title", fields["title"][0]);
                    }

                    if (fields.Contains("objectsid"))
                    {
                        printBytes((byte[])fields["objectsid"][0], "objectsid");
                    }
                    if (fields.Contains("objectguid"))
                    {
                        printBytes((byte[])fields["objectguid"][0], "objectguid");
                    }

                    foreach (String ldapField in fields.PropertyNames)
                    {
                        // cycle through objects in each field e.g. group membership  
                        // (for many fields there will only be one object such as name)  
                        foreach (Object myCollection in fields[ldapField])
                            Console.WriteLine(String.Format("{0,-20} : {1}",
                                          ldapField, myCollection.ToString()));
                    }
                }
                else
                {
                    // user does not exist  
                    Console.WriteLine("User not found!");
                }


                //This block is for testing a search that returns a
                //collection of SearchResult objects (FindAll())

                //search.Filter = "(&(objectClass=user)(cn=" + username + "*))";
                //SearchResultCollection results = search.FindAll();
                ////var json = JsonConvert.SerializeObject(results);

                //int count = 0;

                //string[] resArray = new String[results.Count];

                //foreach (SearchResult res in results)
                //{

                //    ResultPropertyCollection fields = res.Properties;
                //    //foreach (String ldapField in fields.PropertyNames)
                //    //{
                //    //    foreach (Object myCollection in fields[ldapField])
                //    //        Console.WriteLine(String.Format("{0,-20} : {1}",
                //    //                      ldapField, myCollection.ToString()));
                //    //}
                //    Console.WriteLine(String.Format("{0,-20} : {1}",
                //                          "Name", fields["name"][0]));
                //    resArray[count] = (String)fields["name"][0];
                //    //Console.WriteLine(String.Format("    {0,-20} : *** {1}",
                //    //                      "Name", fields["distinguishedname"][0]));
                //    count += 1;
                //    if (count > 20)
                //    {
                //        break;
                //    }
                //}
                //using (StreamWriter file = File.CreateText(@"file.json"))
                //{
                //    JsonSerializer serializer = new JsonSerializer();
                //    serializer.Serialize(file, resArray);
                //}

                Console.Write("Press any key to exit");
                Console.ReadKey();
            }

            catch (Exception e)
            {
                Console.WriteLine("Exception caught:\n\n" + e.ToString());
            }
        }

        static DirectoryEntry createDirectoryEntry()
        {
            // create and return new LDAP connection with desired settings  

            DirectoryEntry ldapConnection = new DirectoryEntry("corrdc2");
            //ldapConnection.Path = "LDAP://OU=Test,OU=Generic Accounts,OU=Information Technology,OU=roundrocktexas.gov,DC=corr,DC=round-rock,DC=tx,DC=us";
            ldapConnection.Path = "LDAP://OU=roundrocktexas.gov,DC=corr,DC=round-rock,DC=tx,DC=us";
            ldapConnection.AuthenticationType = AuthenticationTypes.Secure;

            return ldapConnection;
        }

        static void printBytes(byte[] bytes, string s)
        {
            var sb = new StringBuilder(s + " [] { ");
            string a;
            string bs;
            foreach (var b in bytes)
            {
                bs = Convert.ToString(b);
                a = int.Parse(bs).ToString("X");
                sb.Append(a + " ");
            }
            sb.Append("}");
            Console.WriteLine(sb.ToString());
            Console.ReadKey();
        }
    }
}
