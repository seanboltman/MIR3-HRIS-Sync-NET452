using System;
//using System.Collections.Generic;
using System.Collections;
using System.Data.SqlClient;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Net;
using System.Security.Authentication;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;


namespace MIR3_AdSync
{
    class Program
    {
        static string officeNames =
            "Redmond,Portland,Tacoma,Spokane,Bellingham,Seattle,Boise,Springfield,Salt Lake City,Richland,Bend,Baton Rouge,Pendleton,Kennewick,Cary";

        public class HRData
        {
            public string UUID;
            public string employeeId;
            public string userName;
            public string firstName;
            public string lastName;
            public string jobTitle;
            public string mobilePhone;
            public string workPhone;
            public string homePhone;
            public string office;
            public string email;
        }
        public class Office
        {
            public string address;
            public string city;
            public string state;
            public string zip;
            public string timeZone;
        }

        static Hashtable MIR3Users = null;
        static Hashtable HRActiveUsersByEmployeeId = null;
        static Hashtable HRActiveUsersByMobilePhone = null;
        static Hashtable HRActiveUsersByWorkPhone = null;
        static Hashtable HRActiveUsersByHomePhone = null;
        static Hashtable HRActiveUsersByFirstLast = null;
        static Hashtable Offices = null;

        static void Main(string[] args)
        {
            EventLog.WriteEntry("MIR3 Active Directory Synch", "Starting...", EventLogEntryType.Information);

            // Setup ability to use TLS 1.2, which is NOW REQUIRED by Onvolve MIR3 (as of October 2018)
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

            // collect local information
            GetOffices();
            GetHRUsers();

            // process to figure out who to add, remove, managers, and rebuild office groups in MIR3 as well
            processMIR3Users();

            EventLog.WriteEntry("MIR3 Active Directory Synch", "Finished", EventLogEntryType.Information);
        }

        static void GetHRUsers()
        {
            HRActiveUsersByEmployeeId = new Hashtable();
            HRActiveUsersByMobilePhone = new Hashtable();
            HRActiveUsersByHomePhone = new Hashtable();
            HRActiveUsersByWorkPhone = new Hashtable();
            HRActiveUsersByFirstLast = new Hashtable();

            HRData data = null;

            string connectionString =
                "Application Name=HR System;Data Source=SQLClust;Initial Catalog=Staff;Integrated Security=True";
            string queryString =
                "SELECT Emp.EmployeeID, Emp.FirstName, Emp.LastName, ResidentOffice = ResidentOffices.Description, Emp.MobilePhone, Emp.OfficePhone, Emp.HomePhone " +
                "FROM Staff.dbo.Employee Emp " +
                "INNER JOIN Staff.dbo.Office AS ResidentOffices ON Emp.ResidentOfficeID = ResidentOffices.OfficeID " +
                "INNER JOIN Staff.dbo.Status AS EmpStatus ON Emp.StatusID = EmpStatus.StatusID " +
                "WHERE EmployeeID > 0 AND (Emp.StatusID = 1 OR Emp.StatusID = 2 OR Emp.StatusID = 5)";

            // Collect all Active employees
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandText = queryString;

            connection.Open();

            using (connection)
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    while (sdr.Read())
                    {
                        // retrieve the AD entry based on employee id
                        DirectoryEntry ADRecord = GetUser(Convert.ToInt16(sdr[0]));
                        if (ADRecord == null) continue;

                        data = new HRData();

                        data.employeeId = sdr[0].ToString();
                        data.userName = ADRecord.Properties["sAMAccountName"].Value.ToString();
                        data.firstName = sdr[1].ToString();
                        data.lastName = sdr[2].ToString();
                        data.jobTitle = ADRecord.Properties["title"].Value.ToString();
                        data.workPhone = sdr[5].ToString().Trim();
                        data.homePhone = sdr[6].ToString().Trim();
                        data.office = sdr[3].ToString();
                        data.email = ADRecord.Properties["Mail"].Value.ToString();
                        data.mobilePhone = sdr[4].ToString().Trim();

                        if (data.mobilePhone.Length == 0 && data.homePhone.Length == 0 && data.workPhone.Length == 0)
                        {
                            Console.WriteLine("Missing phone information in HRIS: " + data.firstName + " " + data.lastName);
                            continue;
                        }

                        // Determine a default phone number
                        string defaultNumber = data.mobilePhone;
                        if (defaultNumber.Length == 0)
                        {
                            defaultNumber = data.homePhone;
                            if (defaultNumber.Length == 0)
                            {
                                defaultNumber = data.workPhone;
                            }
                        }
                        if (data.mobilePhone.Length == 0)
                        {
                            data.mobilePhone = defaultNumber;
                        }
                        if (data.homePhone.Length == 0)
                        {
                            data.homePhone = defaultNumber;
                        }
                        if (data.workPhone.Length == 0)
                        {
                            data.workPhone = defaultNumber;
                        }

                        if (data.homePhone.Length > 0)
                        {
                            if (!HRActiveUsersByHomePhone.ContainsKey(data.homePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "")))
                            {
                                HRActiveUsersByHomePhone.Add(data.homePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""), data);
                            }
                        }

                        if (data.mobilePhone.Length > 0)
                        {
                            HRActiveUsersByMobilePhone.Add(data.mobilePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""), data);
                        }

                        if (data.workPhone.Length > 0)
                        {
                            // We may get duplicates since some employees share the same work line
                            if (!HRActiveUsersByWorkPhone.ContainsKey(data.workPhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "")))
                            {
                                HRActiveUsersByWorkPhone.Add(data.workPhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""), data);
                            }
                        }

                        HRActiveUsersByFirstLast.Add(data.firstName.Trim() + data.lastName.Trim(), data);
                        HRActiveUsersByEmployeeId.Add(data.employeeId, data);
                    }
                }
                connection.Close();
            }
        }

        static void GetOffices()
        {
            Offices = new Hashtable();

            Office data = null;

            string connectionString =
                "Application Name=HR System;Data Source=SQLClust;Initial Catalog=Staff;Integrated Security=True";
            string queryString =
                "SELECT dbo.Office.Description, dbo.Office.Address1, dbo.Office.City, dbo.State.Description AS StateName, dbo.Office.PostalCode " +
                "FROM dbo.Office  " +
                "LEFT OUTER JOIN dbo.State ON dbo.Office.StateID = dbo.State.StateID " +
                "WHERE Address1 is not null and Address1 != '' and Office.OfficeID > 0 and " +
                "Office.Description in (";

            string[] officeName = officeNames.Split(',');
            int x = 0;
            foreach (string str in officeName)
            {
                queryString += "'" + str + "'";
                if (++x != officeName.Length)
                {
                    queryString += ",";
                }
            }
            queryString += ")";

            // Collect all Active employees
            SqlConnection connection = new SqlConnection(connectionString);
            SqlCommand command = new SqlCommand();
            command.Connection = connection;
            command.CommandText = queryString;

            connection.Open();

            using (connection)
            {
                using (SqlDataReader sdr = command.ExecuteReader())
                {
                    while (sdr.Read())
                    {
                        // retrieve the AD entry based on employee id
                        data = new Office();

                        data.address = sdr[1].ToString();
                        data.city = sdr[0].ToString();   // force City value to come from the Description column in SQL
                        data.state = sdr[3].ToString();
                        data.zip = sdr[4].ToString();

                        if (data.city.Equals("Redmond") || data.city.Equals("Seattle") || data.city.Equals("Portland") ||
                            data.city.Equals("Tacoma") || data.city.Equals("Spokane") || data.city.Equals("Bellingham") ||
                            data.city.Equals("Bend") || data.city.Equals("Richland") || data.city.Equals("Pendleton") ||
                            data.city.Equals("Kennewick"))
                        {
                            data.timeZone = "PACIFIC_USA";
                        }
                        else if (data.city.Equals("Boise") || data.city.Equals("Salt Lake City") || data.city.Equals("American Fork"))
                        {
                            data.timeZone = "MOUNTAIN_USA";
                        }

                        else if (data.city.Equals("Springfield") || data.city.Equals("Baton Rouge") || data.city.Equals("Lake Charles"))
                        {
                            data.timeZone = "CENTRAL_USA";
                        }

                        Offices.Add(sdr[0].ToString(), data);
                    }
                }
                connection.Close();

            }
        }

        static void processMIR3Users()
        {
            // First clear out all Staff Groups in MIR3
            clearAllStaffGroups();

            HttpWebRequest request = CreateWebRequest();
            XmlDocument soapEnvelopeXml = new XmlDocument();

            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                <soap:Body>
                    <searchRecipients xmlns=""http://www.mir3.com/ws"">
                    <apiVersion>3.7</apiVersion>
                    <authorization>
                        <username>jsheldon.geoengineers.com</username>
                        <password>geohomesafe123</password>
                    </authorization>
                    <maxResults>2000</maxResults>
                    <includeDetail>1</includeDetail>
                    <query>
                    </query>
                    </searchRecipients>
                </soap:Body>
                </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                try
                {
                    HRData data = null;

                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(response.GetResponseStream());

                    XmlNodeList items = xDoc.GetElementsByTagName("recipientDetail");
                    foreach (XmlNode xItem in items)
                    {
                        XmlElement elem = ((XmlElement)xItem);

                        string UUID = elem.GetElementsByTagName("userUUID")[0].InnerXml;
                        string telephonyId = elem.GetElementsByTagName("telephonyId")[0].InnerXml;
                        string employeeId = elem.GetElementsByTagName("employeeId").Count > 0 ? elem.GetElementsByTagName("employeeId")[0].InnerXml : "0";
                        string firstName = elem.GetElementsByTagName("firstName")[0].InnerXml;
                        string lastName = elem.GetElementsByTagName("lastName")[0].InnerXml;
                        string jobTitle = elem.GetElementsByTagName("jobTitle").Count > 0 ? elem.GetElementsByTagName("jobTitle")[0].InnerXml : "";

                        // We have external non-employees defined, like our Building Manager.  Skip these folks.
                        if (employeeId.Equals("0"))
                        {
                            Console.WriteLine(employeeId);
                            continue;
                        }

                        //if (!employeeId.Equals("2608"))  //("845"))  // ("1176"))
                        //{
                        //    Console.WriteLine(employeeId);
                        //    continue;
                        //}

                        Console.WriteLine("Processing: " + employeeId);


                        data = null;
                        if (HRActiveUsersByMobilePhone.ContainsKey(telephonyId))
                        {
                            data = (HRData)HRActiveUsersByMobilePhone[telephonyId];
                            HRActiveUsersByEmployeeId.Remove(data.employeeId);
                        }
                        else if (HRActiveUsersByEmployeeId.ContainsKey(employeeId))
                        {
                            data = (HRData)HRActiveUsersByEmployeeId[employeeId];
                            HRActiveUsersByEmployeeId.Remove(data.employeeId);
                        }
                        else if (HRActiveUsersByWorkPhone.ContainsKey(telephonyId))
                        {
                            data = (HRData)HRActiveUsersByWorkPhone[telephonyId];
                            HRActiveUsersByEmployeeId.Remove(data.employeeId);
                        }
                        else if (HRActiveUsersByFirstLast.ContainsKey(firstName + lastName))
                        {
                            data = (HRData)HRActiveUsersByFirstLast[firstName + lastName];
                            HRActiveUsersByEmployeeId.Remove(data.employeeId);
                        }
                        else if (HRActiveUsersByHomePhone.ContainsKey(telephonyId))
                        {
                            data = (HRData)HRActiveUsersByHomePhone[telephonyId];
                            HRActiveUsersByEmployeeId.Remove(data.employeeId);
                        }
                        // SPECIAL: Jodie Sheldon:  She has a special TelephonyId which we do not know if we can update or not.
                        else if (telephonyId.Equals("102263918"))
                        {
                            data = (HRData)HRActiveUsersByEmployeeId[employeeId];
                            HRActiveUsersByEmployeeId.Remove(data.employeeId);
                            continue;
                        }
                        else
                        {
                            EventLog.WriteEntry("MIR3 Active Directory Synch", "Deleting this user: " + xItem.ChildNodes[3].InnerXml + " " + xItem.ChildNodes[4].InnerXml, EventLogEntryType.Information);
                            Console.WriteLine("Deleting this user: " + xItem.ChildNodes[3].InnerXml + " " + xItem.ChildNodes[4].InnerXml);
                            removeMIR3User(UUID);
                            continue;
                        }

                        // LOOK FOR CHANGES in data
                        bool updateNeeded = false;
                        //bool updateOfficeNeeded = false;

                        // See if the telephonyID needs to change, Cell, then Work, then Home
                        if (data.mobilePhone.Length > 0 && !telephonyId.Equals(data.mobilePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "")))
                        {
                            updateNeeded = true;
                            elem.GetElementsByTagName("telephonyId")[0].InnerXml = data.mobilePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", "");
                        }

                        // compare other values for change
                        if (!data.firstName.Equals(firstName))
                        {
                            updateNeeded = true;
                            elem.GetElementsByTagName("firstName")[0].InnerXml = data.firstName.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
                        }
                        if (!data.lastName.Equals(lastName))
                        {
                            updateNeeded = true;
                            elem.GetElementsByTagName("lastName")[0].InnerXml = data.lastName.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
                        }
                        if (!data.employeeId.Equals(employeeId))
                        {
                            updateNeeded = true;
                            elem.GetElementsByTagName("employeeId")[0].InnerXml = data.employeeId;
                        }
                        if (jobTitle.Length > 0 && !data.jobTitle.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;").Equals(jobTitle))
                        {
                            updateNeeded = true;
                            elem.GetElementsByTagName("jobTitle")[0].InnerXml = data.jobTitle.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
                        }

                        // Locate Devices node
                        XmlNodeList nodes = xItem.ChildNodes;
                        foreach (XmlNode node in nodes)
                        {
                            if (!node.Name.Equals("devices")) continue;

                            XmlNodeList devices = node.ChildNodes;
                            foreach (XmlNode device in devices)
                            {
                                if (device.ChildNodes.Count == 0) continue;

                                if (device.ChildNodes[0].InnerXml.ToString().Equals("Mobile Phone") ||
                                    device.ChildNodes[0].InnerXml.ToString().Equals("SMS"))
                                {
                                    if (!data.mobilePhone.Equals(device.ChildNodes[1].InnerXml))
                                    {
                                        updateNeeded = true;

                                        device.ChildNodes[1].InnerXml = data.mobilePhone;
                                    }
                                }
                                else if (device.ChildNodes[0].InnerXml.ToString().Equals("Work Phone"))
                                {
                                    if (!data.workPhone.Equals(device.ChildNodes[1].InnerXml))
                                    {
                                        updateNeeded = true;
                                        device.ChildNodes[1].InnerXml = data.workPhone;
                                    }
                                }
                                else if (device.ChildNodes[0].InnerXml.ToString().Equals("Home Phone"))
                                {
                                    if (!data.homePhone.Equals(device.ChildNodes[1].InnerXml))
                                    {
                                        updateNeeded = true;
                                        device.ChildNodes[1].InnerXml = data.homePhone;
                                    }
                                }
                                else if (device.ChildNodes[0].InnerXml.ToString().Equals("Work Email"))
                                {
                                    if (!data.email.Equals(device.ChildNodes[1].InnerXml))
                                    {
                                        Console.WriteLine("Bad Email: " + data.email + " vs. " + device.ChildNodes[1].InnerXml);
                                        updateNeeded = true;
                                        device.ChildNodes[1].InnerXml = data.email;
                                    }
                                }
                            }
                        }

                        if (updateNeeded)
                        {
                            EventLog.WriteEntry("MIR3 Active Directory Synch", "Updating user: " + data.firstName + " " + data.lastName, EventLogEntryType.Information);
                            Console.WriteLine("Updating user: " + data.firstName + " " + data.lastName);
                            updateMIR3User(UUID, xItem.InnerXml, data.employeeId, data.office);
                        }

                        addUserToMIR3Groups(employeeId, data.office);
                    }




                    // Any remaining HRIS users need to be added as new user into MIR3
                    IDictionaryEnumerator denum = HRActiveUsersByEmployeeId.GetEnumerator();
                    DictionaryEntry dentry;
                    while (denum.MoveNext())
                    {
                        dentry = (DictionaryEntry)denum.Current;
                        data = (HRData)dentry.Value;

                        // Create a new MIR3 user if we have a phone number
                        if (data.mobilePhone.Length > 0 || data.homePhone.Length > 0 || data.workPhone.Length > 0)
                        {
                            Console.WriteLine("Adding new user: " + data.firstName + " " + data.lastName);
                            EventLog.WriteEntry("MIR3 Active Directory Synch", "Adding new user: " + data.firstName + " " + data.lastName, EventLogEntryType.Information);

                            addMIR3User(data);
                            addUserToMIR3Groups(data.employeeId, data.office);
                        }
                    }
                }
                catch (Exception ex)
                {
                    EventLog.WriteEntry("MIR3 Active Directory Synch", ex.Message, EventLogEntryType.Error);
                }
            }
        }


        public static void clearAllStaffGroups()
        {
            HttpWebRequest request;
            XmlDocument soapEnvelopeXml;

            // Build user list
            string userListXml = "";
            foreach (string empId in HRActiveUsersByEmployeeId.Keys)
            {
                userListXml += "<recipient><employeeId>" + empId + "</employeeId></recipient>";
            }

            string[] officeName = officeNames.Split(',');
            foreach (string str in officeName)
            {
                request = CreateWebRequest();
                soapEnvelopeXml = new XmlDocument();

                //---------------------------------------------------------------------------------
                // Clear the users in this Office All Staff
                //---------------------------------------------------------------------------------

                // Add new user record
                soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                    <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                    <soap:Body>
                        <removeRecipientsFromGroup xmlns=""http://www.mir3.com/ws"">
                        <apiVersion>4.7</apiVersion>
                        <authorization> 
                        <username>jsheldon.geoengineers.com</username>
                        <password>geohomesafe123</password>
                        </authorization>
                        <recipientGroup>" + str + @" All Staff</recipientGroup>
                          <removeMembers>" + userListXml +
                          @"</removeMembers>
                        </removeRecipientsFromGroup>
                    </soap:Body>
                    </soap:Envelope>");

                using (Stream stream = request.GetRequestStream())
                {
                    soapEnvelopeXml.Save(stream);
                }

                using (WebResponse response = request.GetResponse())
                {
                    using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                    {
                        string soapResult = rd.ReadToEnd();
                        Console.WriteLine("Removed all users from " + str + " All Staff");
                    }
                }
            }


            //---------------------------------------------------------------------------------
            // Clear the general All Staff
            //---------------------------------------------------------------------------------

            request = CreateWebRequest();
            soapEnvelopeXml = new XmlDocument();

            // Add new user record
            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                    <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                    <soap:Body>
                        <removeRecipientsFromGroup xmlns=""http://www.mir3.com/ws"">
                        <apiVersion>4.7</apiVersion>
                        <authorization> 
                        <username>jsheldon.geoengineers.com</username>
                        <password>geohomesafe123</password>
                        </authorization>
                        <recipientGroup>All Staff</recipientGroup>
                          <removeMembers>" + userListXml +
                      @"</removeMembers>
                        </removeRecipientsFromGroup>
                    </soap:Body>
                    </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    Console.WriteLine("Removed all users from All Staff");
                }
            }

        }







        public static void addMIR3User(HRData data)
        {
            HttpWebRequest request = CreateWebRequest();
            XmlDocument soapEnvelopeXml = new XmlDocument();

            //---------------------------------------------------------------------------------
            // Add User Record
            //---------------------------------------------------------------------------------

            // Add new user record
            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                <soap:Body>
                    <addNewRecipients xmlns=""http://www.mir3.com/ws"">
                    <apiVersion>3.7</apiVersion>
                    <authorization> <!-- Login credentials of the authorized user -->
                    <username>jsheldon.geoengineers.com</username>
                    <password>geohomesafe123</password>
                    </authorization>" + buildNewUserXml(data) +
                    @"</addNewRecipients>            
                </soap:Body>
                </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    Console.WriteLine(soapResult);
                }
            }
        }

        public static void updateMIR3User(string UUID, string recipientDetails, string employeeId, string office)
        {
            HttpWebRequest request = CreateWebRequest();
            XmlDocument soapEnvelopeXml = new XmlDocument();

            //---------------------------------------------------------------------------------
            // Update User Record
            //---------------------------------------------------------------------------------

            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                <soap:Body>
                <updateRecipients xmlns=""http://www.mir3.com/ws"">
                  <apiVersion>3.7</apiVersion>
                  <authorization> <!-- Login credentials of the authorized user -->
                    <username>jsheldon.geoengineers.com</username>
                    <password>geohomesafe123</password>
                  </authorization>
                  <updateOneRecipient>
                    <recipient>
                    <userUUID>" + UUID + @"</userUUID>
                    </recipient>
                      <recipientDetail>" + recipientDetails +
                      @"</recipientDetail>
                  </updateOneRecipient>
                </updateRecipients>
                </soap:Body>
                </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    Console.WriteLine(soapResult);
                }
            }
        }

        public static void addUserToMIR3Groups(string employeeId, string office)
        {
            HttpWebRequest request = CreateWebRequest();
            XmlDocument soapEnvelopeXml = new XmlDocument();

            //---------------------------------------------------------------------------------
            // Update Office All Staff group
            //---------------------------------------------------------------------------------

            request = CreateWebRequest();
            soapEnvelopeXml = new XmlDocument();

            // Add new user to office group
            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                            <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                            <soap:Body>
                                <addRecipientsToGroup xmlns=""http://www.mir3.com/ws"">
                                <apiVersion>3.7</apiVersion>
                                <authorization> <!-- Login credentials of the authorized user -->
                                <username>jsheldon.geoengineers.com</username>
                                <password>geohomesafe123</password>
                                </authorization>
                                <recipientGroup>" + office + @" All Staff</recipientGroup>
                                  <addMembers>
                                    <recipient>
                                      <employeeId>" + employeeId + @"</employeeId>
                                    </recipient>
                                  </addMembers>
                                </addRecipientsToGroup>
                            </soap:Body>
                            </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    Console.WriteLine("*** added user: " + employeeId + "  to " + office + " All Staff");
                }
            }

            //---------------------------------------------------------------------------------
            // Update All Staff group
            //---------------------------------------------------------------------------------

            request = CreateWebRequest();
            soapEnvelopeXml = new XmlDocument();

            // Add new user to office group
            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                            <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                            <soap:Body>
                                <addRecipientsToGroup xmlns=""http://www.mir3.com/ws"">
                                <apiVersion>3.7</apiVersion>
                                <authorization> <!-- Login credentials of the authorized user -->
                                <username>jsheldon.geoengineers.com</username>
                                <password>geohomesafe123</password>
                                </authorization>
                                <recipientGroup>All Staff</recipientGroup>
                                  <addMembers>
                                    <recipient>
                                      <employeeId>" + employeeId + @"</employeeId>
                                    </recipient>
                                  </addMembers>
                                </addRecipientsToGroup>
                            </soap:Body>
                            </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    Console.WriteLine("*** added user: " + employeeId + "  to All Staff");
                }
            }
        }

        public static void removeMIR3User(string UUID)
        {
            HttpWebRequest request = CreateWebRequest();
            XmlDocument soapEnvelopeXml = new XmlDocument();

            // Remove user (which also removes them from any groups)
            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                <soap:Body>
                    <deleteRecipients xmlns =""http://www.mir3.com/ws"">
                    <apiVersion>3.7</apiVersion>
                    <authorization> <!-- Login credentials of the authorized user -->
                    <username>jsheldon.geoengineers.com</username>
                    <password>geohomesafe123</password>
                    </authorization>
                    <recipient>
                    <userUUID>" + UUID + @"</userUUID>
                    </recipient>
                    </deleteRecipients>
                </soap:Body>
                </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
                {
                    string soapResult = rd.ReadToEnd();
                    Console.WriteLine(soapResult);
                }
            }
        }

        public static string buildNewUserXml(HRData data)
        {
            string Xml = @"<recipientDetail><pin>9999</pin><telephonyId>{telephonyId}</telephonyId><firstName>{firstName}</firstName><lastName>{lastName}</lastName><locale>en_US</locale><employeeId>{employeeId}</employeeId><jobTitle>{jobTitle}</jobTitle><company>GeoEngineers</company><division>/</division><addresses><address><addressTypeName>work</addressTypeName><address1>{officeAddress}</address1><city>{officeCity}</city><state>{officeState}</state><zip>{officeZip}</zip></address></addresses><timeZone>{timeZone}</timeZone><role>Recipient</role><devices><device><deviceType>Mobile Phone</deviceType><address>{mobilePhone}</address><description>Mobile Phone</description><private>false</private><disabled>false</disabled><defaultPriority>1</defaultPriority><source>MIR3</source></device><device><deviceType>Home Phone</deviceType><address>{homePhone}</address><description>Home Phone</description><private>false</private><disabled>false</disabled><defaultPriority>1</defaultPriority><source>MIR3</source></device><device><deviceType>SMS</deviceType><address>{smsPhone}</address><description>SMS</description><private>false</private><disabled>false</disabled><defaultPriority>1</defaultPriority><source>MIR3</source></device><device><deviceType>Work Phone</deviceType><address>{workPhone}</address><description>Work Phone</description><private>false</private><disabled>false</disabled><defaultPriority>1</defaultPriority><source>MIR3</source></device><device><deviceType>Work Email</deviceType><address>{workEmail}</address><description>Work Email</description> <private>false</private><disabled>false</disabled><sendReports2>TEXT</sendReports2><defaultPriority>1</defaultPriority><source>MIR3</source></device></devices><activeLocationStatus>DEFAULT_LOCATION</activeLocationStatus><locationStatuses><defaultStatus/></locationStatuses><preferences><localePreference><locale>en_US</locale><visible>true</visible></localePreference></preferences></recipientDetail>";

            if (data.mobilePhone.Length > 0)
            {
                Xml = Xml.Replace("{telephonyId}", data.mobilePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""));
            }
            else if (data.homePhone.Length > 0)
            {
                Xml = Xml.Replace("{telephonyId}", data.homePhone.Replace(".", "").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""));
            }

            Xml = Xml.Replace("{firstName}", data.firstName.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));
            Xml = Xml.Replace("{lastName}", data.lastName.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));
            Xml = Xml.Replace("{employeeId}", data.employeeId);
            Xml = Xml.Replace("{jobTitle}", data.jobTitle.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));

            Office o = (Office)Offices[data.office];
            Xml = Xml.Replace("{officeAddress}", o.address.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));
            Xml = Xml.Replace("{officeCity}", o.city.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));
            Xml = Xml.Replace("{officeState}", o.state);
            Xml = Xml.Replace("{officeZip}", o.zip);
            Xml = Xml.Replace("{timeZone}", o.timeZone);

            Xml = Xml.Replace("{mobilePhone}", data.mobilePhone);
            Xml = Xml.Replace("{smsPhone}", data.mobilePhone);
            Xml = Xml.Replace("{workPhone}", data.workPhone);
            Xml = Xml.Replace("{homePhone}", data.homePhone);
            Xml = Xml.Replace("{workEmail}", data.email.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;"));

            return Xml;

        }

        public static DirectoryEntry GetUser(int employeeId)
        {
            return GetDirectoryEntry(string.Format("employeeID={0}", employeeId));
        }

        public static DirectoryEntry GetDirectoryEntry(string filter)
        {
            using (Domain domain = Domain.GetCurrentDomain())
            using (DirectoryEntry rootEntry = domain.GetDirectoryEntry())
            using (DirectorySearcher searcher = new DirectorySearcher(rootEntry))
            {
                searcher.ReferralChasing = ReferralChasingOption.All;
                searcher.Filter = filter;
                searcher.SearchScope = SearchScope.Subtree;

                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    ResultPropertyValueCollection propertyValues = result.Properties["objectsid"];
                    byte[] objectsid = (byte[])propertyValues[0];

                    SecurityIdentifier sid = new SecurityIdentifier(objectsid, 0);

                    NTAccount account = (NTAccount)sid.Translate(typeof(NTAccount));
                    account.ToString(); // This give the DOMAIN\User format for the account

                    return result.GetDirectoryEntry();
                }
                else
                {
                    return null;
                }
            }
        }

        static public HttpWebRequest CreateWebRequest()
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(@"https://ws.mir3.com/services/v1.2/mir3");
            webRequest.Headers.Add(@"SOAP:Action");
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            return webRequest;
        }



        static void GetMIR3Users()
        {
            MIR3Users = new Hashtable();
            HRData data = null;

            HttpWebRequest request = CreateWebRequest();
            XmlDocument soapEnvelopeXml = new XmlDocument();

            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
                <soap:Body>
                    <searchRecipients xmlns=""http://www.mir3.com/ws"">
                    <apiVersion>3.7</apiVersion>
                    <authorization>
                        <username>jsheldon.geoengineers.com</username>
                        <password>geohomesafe123</password>
                    </authorization>
                    <maxResults>1000</maxResults>
                    <query>
                    </query>
                    </searchRecipients>
                </soap:Body>
                </soap:Envelope>");

            using (Stream stream = request.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }

            using (WebResponse response = request.GetResponse())
            {
                try
                {
                    XmlDocument xDoc = new XmlDocument();
                    xDoc.Load(response.GetResponseStream());

                    XmlNodeList items = xDoc.GetElementsByTagName("recipient");
                    foreach (XmlNode xItem in items)
                    {
                        data = new HRData();

                        data.UUID = xItem.ChildNodes[0].InnerXml;
                        data.firstName = xItem.ChildNodes[1].InnerXml;
                        data.lastName = xItem.ChildNodes[2].InnerXml;
                        data.mobilePhone = xItem.ChildNodes[3].InnerXml;

                        MIR3Users.Add(data.mobilePhone, data);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(ex.Message);
                }
            }
        }

    }
}





//            HttpWebRequest request = CreateWebRequest();
//            XmlDocument soapEnvelopeXml = new XmlDocument();
//            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
//                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
//                <soap:Body>
//                    <getVerbiage xmlns=""http://www.mir3.com/ws"">
//                    <apiVersion>3.7</apiVersion>
//                    <authorization> <!-- Login credentials of the authorized user -->
//                    <username>jsheldon.geoengineers.com</username>
//                    <password>geohomesafe123</password>
//                    </authorization>
//                    <key>phone.thankYou</key>
//                    <division>/</division>
//                    </getVerbiage>
//                </soap:Body>
//                </soap:Envelope>");

//            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
//                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
//                <soap:Body>
//                    <getRecipientRoles xmlns=""http://www.mir3.com/ws"">
//                    <apiVersion>3.7</apiVersion>
//                    <authorization> <!-- Login credentials of the authorized user -->
//                    <username>jsheldon.geoengineers.com</username>
//                    <password>geohomesafe123</password>
//                    </authorization>
//                    </getRecipientRoles>
//                </soap:Body>
//                </soap:Envelope>");

//            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
//                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
//                <soap:Body>
//                    <searchRecipients xmlns=""http://www.mir3.com/ws"">
//                    <apiVersion>3.7</apiVersion>
//                    <authorization> <!-- Login credentials of the authorized user -->
//                    <username>jsheldon.geoengineers.com</username>
//                    <password>geohomesafe123</password>
//                    </authorization>
//                    <query>
//                    </query>
//                    </searchRecipients>
//                </soap:Body>
//                </soap:Envelope>");


//            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
//                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
//                <soap:Body>
//                    <searchRecipients xmlns=""http://www.mir3.com/ws"">
//                    <apiVersion>3.7</apiVersion>
//                    <authorization> <!-- Login credentials of the authorized user -->
//                    <username>jsheldon.geoengineers.com</username>
//                    <password>geohomesafe123</password>
//                    </authorization>
//                    <includeDetail>1</includeDetail>
//                    <query>
//                    <and>
//                    <lastName>Thom</lastName>
//                    </and>
//                    </query>
//                    </searchRecipients>
//                </soap:Body>
//                </soap:Envelope>");

//////            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
//////                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
//////                <soap:Body>
//////                    <searchRecipients xmlns=""http://www.mir3.com/ws"">
//////                    <apiVersion>3.7</apiVersion>
//////                    <authorization> <!-- Login credentials of the authorized user -->
//////                    <username>jsheldon.geoengineers.com</username>
//////                    <password>geohomesafe123</password>
//////                    </authorization>
//////                    <includeDetail>1</includeDetail>
//////                    <query>
//////                    </query>
//////                    </searchRecipients>
//////                </soap:Body>
//////                </soap:Envelope>");

////            // update sboltman record
////            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
////                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
////                <soap:Body>
////                <updateRecipients xmlns=""http://www.mir3.com/ws"">
////                <apiVersion>3.7</apiVersion>
////                <authorization> <!-- Login credentials of the authorized user -->
////                <username>jsheldon.geoengineers.com</username>
////                <password>geohomesafe123</password>
////                </authorization>
////                <updateOneRecipient>
////                <recipient>
////                <userUUID>15d53200-0001-3000-80c0-fceb55463ffe</userUUID>
////                </recipient>
////                  <recipientDetail>
////                    <pin>9999</pin>
////                    <telephonyId>2064271232</telephonyId>
////                    <firstName>Sean</firstName>
////                    <lastName>Boltman</lastName>
////                    <locale>en_US</locale>
////                    <employeeId>2064271232</employeeId>
////                    <jobTitle>Senior Engineer</jobTitle>
////                    <company>GeoEngineers</company>
////                    <division>/</division>
////                    <addresses>
////                      <address>
////                        <addressTypeName>work</addressTypeName>
////                        <city>Redmond</city>
////                        <state>WA</state>
////                        <zip>98101</zip>
////                      </address>
////                    </addresses>
////                    <timeZone>PACIFIC_USA</timeZone>
////                    <role>Recipient</role>
////                    <devices>
////                      <device>
////                        <deviceType>Mobile Phone</deviceType>
////                        <address>(206) 427-1232</address>
////                        <description>Mobile Phone</description>
////                        <private>false</private>
////                        <disabled>false</disabled>
////                        <defaultPriority>1</defaultPriority>
////                        <source>MIR3</source>
////                      </device>
////                      <device>
////                        <deviceType>SMS</deviceType>
////                        <address>(206) 427-1232</address>
////                        <description>SMS</description>
////                        <private>false</private>
////                        <disabled>false</disabled>
////                        <defaultPriority>1</defaultPriority>
////                        <source>MIR3</source>
////                      </device>
////                      <device>
////                        <deviceType>Work Phone</deviceType>
////                        <address>(425) 861-6097</address>
////                        <description>Work Phone</description>
////                        <private>false</private>
////                        <disabled>false</disabled>
////                        <defaultPriority>1</defaultPriority>
////                        <source>MIR3</source>
////                      </device>
////                      <device>
////                        <deviceType>Home Phone</deviceType>
////                        <address>(206) 427-1232</address>
////                        <description>Home Phone</description>
////                        <private>false</private>
////                        <disabled>false</disabled>
////                        <defaultPriority>1</defaultPriority>
////                        <source>MIR3</source>
////                      </device>
////                      <device>
////                        <deviceType>Work Email</deviceType>
////                        <address>sboltman@geoengineers.com</address>
////                        <description>Work Email</description>
////                        <private>false</private>
////                        <disabled>false</disabled>
////                        <sendReports2>TEXT</sendReports2>
////                        <defaultPriority>1</defaultPriority>
////                        <source>MIR3</source>
////                      </device>
////                    </devices>
////                    <activeLocationStatus>Seattle</activeLocationStatus>
////                    <locationStatuses>
////                      <defaultStatus/>
////                      <locationStatus>
////                        <name>Seattle</name>
////                      </locationStatus>
////                    </locationStatuses>
////                    <preferences>
////                      <localePreference>
////                        <locale>en_US</locale>
////                        <visible>true</visible>
////                      </localePreference>
////                    </preferences>
////                  </recipientDetail>
////            </updateOneRecipient>
////            </updateRecipients>
////            </soap:Body>
////            </soap:Envelope>");

//            // Lookup SBOLTMAN record
//            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
//                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
//                <soap:Body>
//                    <searchRecipients xmlns=""http://www.mir3.com/ws"">
//                    <apiVersion>3.7</apiVersion>
//                    <authorization> <!-- Login credentials of the authorized user -->
//                    <username>jsheldon.geoengineers.com</username>
//                    <password>geohomesafe123</password>
//                    </authorization>
//                    <includeDetail>1</includeDetail>
//                    <query>
//                    <and>
//                    <lastName>Adams</lastName>
//                    </and>
//                    </query>
//                    </searchRecipients>
//                </soap:Body>
//                </soap:Envelope>");

////            // Add new user record
////            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
////                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
////                <soap:Body>
////                    <addNewRecipients xmlns=""http://www.mir3.com/ws"">
////                    <apiVersion>3.7</apiVersion>
////                    <authorization> <!-- Login credentials of the authorized user -->
////                    <username>jsheldon.geoengineers.com</username>
////                    <password>geohomesafe123</password>
////                    </authorization>
////                    <recipientDetail>
////                      <pin>9999</pin>
////                      <telephonyId>2064271233</telephonyId>
////                      <firstName>Jessica</firstName>
////                      <lastName>Firsow</lastName>
////                      <employeeId>2507</employeeId>
////                      <locale>en_US</locale>
////                      <jobTitle>Senior Engineer</jobTitle>
////                      <company>GeoEngineers</company>
////                      <division>/</division>
////                      <addresses>
////                        <address>
////                          <addressTypeName>work</addressTypeName>
////                          <address1></address1>
////                          <city>Redmond</city>
////                          <state>WA</state>
////                          <zip>98101</zip>
////                        </address>
////                      </addresses>
////                      <timeZone>PACIFIC_USA</timeZone>
////                      <role>Recipient</role>
////                      <devices>
////                        <device>
////                          <deviceType>Mobile Phone</deviceType>
////                          <address>(206) 427-1232</address>
////                          <description>Mobile Phone</description>
////                          <private>false</private>
////                          <disabled>false</disabled>
////                          <defaultPriority>1</defaultPriority>
////                          <source>MIR3</source>
////                        </device>
////                        <device>
////                          <deviceType>SMS</deviceType>
////                          <address>(206) 427-1232</address>
////                          <description>SMS</description>
////                          <private>false</private>
////                          <disabled>false</disabled>
////                          <defaultPriority>1</defaultPriority>
////                          <source>MIR3</source>
////                        </device>
////                        <device>
////                          <deviceType>Work Phone</deviceType>
////                          <address>(425) 861-6097</address>
////                          <description>Work Phone</description>
////                          <private>false</private>
////                          <disabled>false</disabled>
////                          <defaultPriority>1</defaultPriority>
////                          <source>MIR3</source>
////                        </device>
////                        <device>
////                          <deviceType>Home Phone</deviceType>
////                          <address>(206) 427-1232</address>
////                          <description>Home Phone</description>
////                          <private>false</private>
////                          <disabled>false</disabled>
////                          <defaultPriority>1</defaultPriority>
////                          <source>MIR3</source>
////                        </device>
////                        <device>
////                          <deviceType>Work Email</deviceType>
////                          <address>sboltman@geoengineers.com</address>
////                          <description>Work Email</description>
////                          <private>false</private>
////                          <disabled>false</disabled>
////                          <sendReports2>TEXT</sendReports2>
////                          <defaultPriority>1</defaultPriority>
////                          <source>MIR3</source>
////                        </device>
////                      </devices>
////                      <activeLocationStatus>Seattle</activeLocationStatus>
////                      <locationStatuses>
////                        <defaultStatus/>
////                        <locationStatus>
////                          <name>Seattle</name>
////                        </locationStatus>
////                      </locationStatuses>
////                      <preferences>
////                        <localePreference>
////                          <locale>en_US</locale>
////                          <visible>true</visible>
////                        </localePreference>
////                      </preferences>
////                    </recipientDetail>
////                    </addNewRecipients>            
////                </soap:Body>
////                </soap:Envelope>");

////            // Add new user to office group
////            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
////                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
////                <soap:Body>
////                    <addRecipientsToGroup xmlns=""http://www.mir3.com/ws"">
////                    <apiVersion>3.7</apiVersion>
////                    <authorization> <!-- Login credentials of the authorized user -->
////                    <username>jsheldon.geoengineers.com</username>
////                    <password>geohomesafe123</password>
////                    </authorization>
////                    <recipientGroup>Redmond All Staff</recipientGroup>
////                    <addMembers>
////                    <recipient>
////                    <employeeId>2507</employeeId>
////                    </recipient>
////                    <recipientGroup>Redmond All Staff</recipientGroup>
////                    </addMembers>
////                    </addRecipientsToGroup>
////                </soap:Body>
////                </soap:Envelope>");

////            // Remove user (which also removes them from any groups)
////            soapEnvelopeXml.LoadXml(@"<?xml version=""1.0"" encoding=""utf-8""?>
////                <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
////                <soap:Body>
////                    <deleteRecipients xmlns =""http://www.mir3.com/ws"">
////                    <apiVersion>3.7</apiVersion>
////                    <authorization> <!-- Login credentials of the authorized user -->
////                    <username>jsheldon.geoengineers.com</username>
////                    <password>geohomesafe123</password>
////                    </authorization>
////                    <recipient>
////                    <employeeId>2507</employeeId>
////                    </recipient>
////                    </deleteRecipients>
////                </soap:Body>
////                </soap:Envelope>");

//            using (Stream stream = request.GetRequestStream())
//            {
//                soapEnvelopeXml.Save(stream);
//            }

//            using (WebResponse response = request.GetResponse())
//            {
//                using (StreamReader rd = new StreamReader(response.GetResponseStream()))
//                {
//                    string soapResult = rd.ReadToEnd();
//                    Console.WriteLine(soapResult);

//                    try
//                    {
//                        XmlDocument xDoc = new XmlDocument();
//                        xDoc.Load(response.GetResponseStream());

//                        XmlNodeList items = xDoc.GetElementsByTagName("recipientDetail");
//                        foreach (XmlNode xItem in items)
//                        {
//                            string userUUID = xItem.Attributes["userUUID"].Value;
//                            string firstName = xItem.Attributes["firstName"].Value;
//                            string lastName = xItem.Attributes["lastName"].Value;
//                            XmlNodeList devices = xItem.Attributes["devices"].ChildNodes;
//                            foreach (XmlNode device in devices)
//                            {
//                                string deviceType = device.Attributes["deviceType"].Value;
//                                string address = device.Attributes["address"].Value;
//                                string description = device.Attributes["description"].Value;
//                            }

//                        }
//                    }
//                    catch (Exception ex)
//                    {
//                        System.Diagnostics.Debug.WriteLine(ex.Message);
//                    }

//                }
//    }
//}

