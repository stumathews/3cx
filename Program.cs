using System;
using System.Collections.Generic;
using System.Text;
using TCX.Configuration;
using TCX.PBXAPI;
using System.Threading;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.IO;
using NDesk.Options;


/*
 TODO:
 * Add option to print routes
 * FAdd Extension group matching
 *
 
 */

namespace TestCSharp
{
    class Program
    {
        
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileStringA")]
        static extern int GetKeyValueA(string strSection, string strKeyName, string strNull, StringBuilder RetVal, int nSize, string strFileName);
         [DllImport("kernel32.dll")]
        public static extern bool SetConsoleTextAttribute(IntPtr hConsoleOutput,      int wAttributes);
        [DllImport("kernel32.dll")]
        public static extern IntPtr GetStdHandle(uint nStdHandle);
        
        /* Global program options */
        static string csv_file_path = null;
        static bool show_help = false;
        static bool use_csv = false;
        static bool print_rules = false;
        static bool isDebug = false;
        static bool print_gateways = false;
        static bool ignore_warnings = false;
        static bool be_silent = false;
        static bool item_protection = false;

        class MyListener
        {
            static int receivedNotifications=0;
            public class MyEvent
            {
                public MyEvent(int operation, NotificationEventArgs e)
                {
                    operation_ = operation;
                    e_ = e;
                }
                public int operation_;
                public NotificationEventArgs e_;
            };

            String filter_ = null;
            public MyListener(String filter)
            {
                filter_ = filter;
            }

            void eventHandler(Object data)
            {
                MyEvent ev = data as MyEvent;
                if (ev != null)
                {
                    NotificationEventArgs e = ev.e_;
                    switch (ev.operation_)
                    {
                        case 0: //insert
                            {
                                System.Console.WriteLine("Inserted: " + e.EntityName + "(" + e.RecID.ToString() + "):" + (e.ConfObject == null ? "no object" : e.ConfObject.ToString()));
                                ActiveConnection ac = e.ConfObject as ActiveConnection;
                                if (ac != null)
                                {
                                    System.Console.WriteLine("AttachedData(" + ac.AttachedData.Count.ToString() + "):");
                                    foreach (KeyValuePair<String, String> kvp in ac.AttachedData)
                                    {
                                        System.Console.WriteLine(kvp.Key + "=" + kvp.Value);
                                    }
                                }
                                if (e.EntityName == "REGISTRATION")
                                {
                                    DN dn = e.ConfObject as DN;
                                    foreach (RegistrarRecord rr in dn.GetRegistrarContactsEx())
                                    {
                                        System.Console.WriteLine(rr.ToString()+"\nExpites="+rr.Expires.ToString());
                                    }
                                }
                            }
                            break;
                        case 1:
                            {
                                System.Console.WriteLine("Updated: " + e.EntityName + "(" + e.RecID.ToString() + "):" + (e.ConfObject == null ? "no object" : e.ConfObject.ToString()));
                                ActiveConnection ac = e.ConfObject as ActiveConnection;
                                if (ac != null)
                                {
                                    System.Console.WriteLine("AttachedData(" + ac.AttachedData.Count.ToString() + "):");
                                    foreach (KeyValuePair<String, String> kvp in ac.AttachedData)
                                    {
                                        System.Console.WriteLine(kvp.Key + "=" + kvp.Value);
                                    }
                                }
                                if (e.EntityName == "REGISTRATION")
                                {
                                    DN dn = e.ConfObject as DN;
                                    foreach (RegistrarRecord rr in dn.GetRegistrarContactsEx())
                                    {
                                        System.Console.WriteLine(rr.ToString() + "\nExpites=" + rr.Expires.ToString());
                                    }
                                }
                            }
                            break;
                        case 2:
                            {
                                System.Console.WriteLine("Deleted: " + e.EntityName + "(" + e.RecID.ToString() + "):" + (e.ConfObject == null ? "no object" : e.ConfObject.ToString()));

                                ActiveConnection ac = e.ConfObject as ActiveConnection;
                                if (ac != null)
                                {
                                    System.Console.WriteLine("AttachedData(" + ac.AttachedData.Count.ToString() + "):");
                                    foreach (KeyValuePair<String, String> kvp in ac.AttachedData)
                                    {
                                        System.Console.WriteLine(kvp.Key + "=" + kvp.Value);
                                    }
                                }
                                if (e.EntityName == "REGISTRATION")
                                {
                                    DN dn = e.ConfObject as DN;
                                    foreach (RegistrarRecord rr in dn.GetRegistrarContactsEx())
                                    {
                                        System.Console.WriteLine(rr.ToString() + "\nExpites=" + rr.Expires.ToString());
                                    }
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }

            public void ps_Inserted(object sender, NotificationEventArgs e)
            {
                if (filter_ != null && filter_ != e.EntityName)
                    return;
                System.Console.WriteLine((receivedNotifications++).ToString());
                //ThreadPool.QueueUserWorkItem(this.eventHandler, new MyEvent(0, e));
                this.eventHandler(new MyEvent(0, e));
            }
            public void ps_Updated(object sender, NotificationEventArgs e)
            {
                if (filter_ != null && filter_ != e.EntityName)
                    return;
                System.Console.WriteLine((receivedNotifications++).ToString());
                //ThreadPool.QueueUserWorkItem(this.eventHandler, new MyEvent(1, e));
                this.eventHandler(new MyEvent(1, e));
            }
            public void ps_Deleted(object sender, NotificationEventArgs e)
            {
                if (filter_ != null && filter_ != e.EntityName)
                    return;
                System.Console.WriteLine((receivedNotifications++).ToString());
                //ThreadPool.QueueUserWorkItem(this.eventHandler, new MyEvent(2, e));
                this.eventHandler(new MyEvent(2, e));
            }
        }
        struct OutBoundRouteDetails
        {
            public string name;
            public byte strip_count;
            public string prepend;
        }
        struct OutBoundRuleDetails
        {
            public String Calls_to_numbers_with_a_length_of;
            public String Name;
            public String Prefix;
            public ExternalLine[] External_lines; // not used
            public String ExtensionRange;  //Represents extension
            public Group[] DNGroups;    //not used
            public OutBoundRouteDetails[] Outbound_route;
                //r.OutboundRoutes[0].Gateway = ps.GetGateways()[0];
            public void  Initialise()
            {
                this.Calls_to_numbers_with_a_length_of = "";
                this.Name = "";
                this.Prefix = "";
                this.External_lines = null;
                this.ExtensionRange = "";
                this.DNGroups = null;

                /* Its possible to have 3 routes but we're only going to support having one per outbound rule - we initilaise 3 through for future use */
                this.Outbound_route = new OutBoundRouteDetails[3];

                for (int i = 0; i < 3; i++)
                {
                    OutBoundRouteDetails route = new OutBoundRouteDetails();
                    route.name = "";
                    route.prepend = "";
                    route.strip_count = 0;
                    this.Outbound_route[i] = route;
                }

                
                
            }
           

            
        }
        static void Main(string[] args)
        {
            PhoneSystem.ApplicationName = "3CX Outbound Rule importer";
            OutBoundRuleDetails single_use_outbound_rule = new OutBoundRuleDetails();
            single_use_outbound_rule.Initialise();
            
            /* Deal with the command line */
                      

            OptionSet options = new OptionSet()
                .Add("h|?|help", "Show this output", delegate(string v) { show_help = v != null; })
                .Add("f|csv-file=", "Specify path to csv file (eg. c:\\import_rules.csv)", delegate(string v) { csv_file_path = v; use_csv = true; })
                .Add("p|prefix=", "prefix to use (eg. 1232)", delegate(string v) { single_use_outbound_rule.Prefix = v; })
                .Add("n|numlen=", "VALUE number length (eg. 5)", delegate(string v) { single_use_outbound_rule.Calls_to_numbers_with_a_length_of = v; })
               // .Add("po|outrules", "Print outbound rules for the default tenant ", delegate(string v) { print_rules = true; })
               // .Add("pgw|gateways", "Print gateways defined in default tenant ", delegate(string v) { print_gateways = true; })
                .Add("x|name=", "name of rule to use eg. (CstInternalFwdRule) ", delegate(string v) { single_use_outbound_rule.Name = v; })
                .Add("d|debug", "print debug information", delegate(string v) { isDebug = true; })
                .Add("sr|strip=", "strip VALUE digits while routing (eg. 2)", delegate(string v) { single_use_outbound_rule.Outbound_route[0].strip_count = byte.Parse(v); })
                .Add("pr|prepend=", "prepend while routing (eg. 2222)", delegate(string v) { single_use_outbound_rule.Outbound_route[0].prepend = v; })
                .Add("iw|ignore-warnings", "Dont show warnings", delegate(string v) { ignore_warnings = true; })
                .Add("bs|be-silent", "Don't show any output", delegate(string v) { be_silent = true; })
                .Add("e|extension-range=", "Extension range", delegate(string v) { single_use_outbound_rule.ExtensionRange = v; })
                .Add("rn|route-name=", "Outbound root name (eg. DOL)", delegate(string v) { single_use_outbound_rule.Outbound_route[0].name = v; }); // we only support creating one route, hence [0]

                
            List<string> extra = options.Parse(args);
            if (args.Length < 2)
                show_help = true;
            if( extra.Count > 0 )
            {
                foreach (string option in extra)
                {
                    System.Console.WriteLine(option + " not recognised.");
                }
            }

            if( show_help == true )
            {
                Console.WriteLine(PhoneSystem.ApplicationName);
                Console.WriteLine("Usage: orule_importer.exe [OPTIONS]");                
                Console.WriteLine();
                Console.WriteLine("NB:This needs to be run 'physically' on a capable 3CX server.");
                Console.WriteLine();
                Console.WriteLine("Options:"); 
                options.WriteOptionDescriptions(Console.Out);
                Console.WriteLine();
                Console.WriteLine("Eg: orule_importer.exe -e 11115-11118,12345 -p MyPrefix -n 5 -x MyRule --sr 2 --pr 2222 --rn DOL");
                Console.WriteLine("Eg: orule_importer.exe -f c:\\import_rules.csv --bs");
                
                return;
            }


                
            
            #region phone system initialization(init db server)
            RegistryKey regKeyAppRoot;
            if (IntPtr.Size == 4)
            {
                regKeyAppRoot = Registry.LocalMachine.OpenSubKey("SOFTWARE\\3CX\\PhoneSystem");
            }
            else
            {
                regKeyAppRoot = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\3CX\\PhoneSystem");
            }
            String _appPath = (string)regKeyAppRoot.GetValue("AppPath");
            String filePath = Path.Combine(_appPath, @"Bin\3CXPhoneSystem.ini");
            String value = GetKeyValue("ConfService", "ConfPort", filePath);
            Int32 port = 0;
            if (!String.IsNullOrEmpty(value))
            {
                Int32.TryParse(value.Trim(), out port);
                PhoneSystem.CfgServerPort = port;
            }
            value = GetKeyValue("ConfService", "confUser", filePath);
            if (!String.IsNullOrEmpty(value))
                PhoneSystem.CfgServerUser = value;
            value = GetKeyValue("ConfService", "confPass", filePath);
            if (!String.IsNullOrEmpty(value))
                PhoneSystem.CfgServerPassword = value;
            #endregion
           
            PhoneSystem ps = null;
                        
            while (ps == null)
            {
                try
                {
                    /* Get the system off its knees */
                    ps = PhoneSystem.Root;
                    Tenant tenant = ps.GetTenants()[0];
                    List<OutBoundRuleDetails> outbound_rules_to_create = new List<OutBoundRuleDetails>();

                    if (use_csv == true)
                    {
                        StreamReader sr = null;
                        

                        try
                        {
                            sr = File.OpenText(csv_file_path);
                            string line = sr.ReadLine();

                            while (String.IsNullOrEmpty(line) == false)
                            {
                                if (line.StartsWith("//"))
                                {
                                    line = sr.ReadLine();
                                    continue;
                                }

                                if (item_protection == false)
                                {
                                    if (line.StartsWith("[") == true)
                                        item_protection = true;
                                }

                                string[] values = line.Split(',');
                                string[] result = removeItemProtection(item_protection, values);
                                

                                if (result.Length != 7)
                                    throw new Exception("Delimited item length not expected: CSV is in invalid format");

                                OutBoundRuleDetails rule = new OutBoundRuleDetails();
                                rule.Initialise();
                                try
                                {
                                    rule.Name = result[0];
                                    rule.Prefix = result[1];
                                    rule.Calls_to_numbers_with_a_length_of = result[2];
                                    rule.Outbound_route[0].name = result[3];
                                    rule.Outbound_route[0].prepend = result[4];
                                    rule.Outbound_route[0].strip_count = byte.Parse(result[5]);
                                    rule.ExtensionRange = result[6];
                                }
                                catch (Exception ex)
                                {
                                    throw new Exception(ex.Message + ": CSV is in invalid format");
                                }
                                outbound_rules_to_create.Add(rule);
                                line = sr.ReadLine();
                            }
                        }
                        catch (Exception ex)
                        {
                            LogMessage(ex.Message, Severity.Critical);
                            return;                                
                        }
                        finally
                        {
                            if( sr != null )
                                sr.Close(); // Be nice ;-)
                        }
                        /* Loop through datasource and extract details for outbound rules to be created */
                                               

                        
                    }
                    else
                    {
                        // single cmd line mode
                        outbound_rules_to_create.Add(single_use_outbound_rule);
                    }

                    /* With the data loaded, lets create the outbound rules that the data corresponds to*/
                    
                    for (int i = 0; i < outbound_rules_to_create.Count; i++)
                    {
                        CreateAndSaveOutboundRule(ps, tenant, outbound_rules_to_create[i]);                        
                    }
                    
                }
                catch (TCX.Configuration.Exceptions.ObjectSavingException)
                {
                    LogMessage("Error creating and saving outbound rule. ", Severity.Critical);                    
                }
            }
            
        }

        private static string[] removeItemProtection(bool item_protection, string[] values)
        {
            //if (item_protection == true)
            //{
                List<String> continuationString = new List<string>();
                bool currently_processing_continuation = false;
                List<String> result = new List<string>();

                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = values[i].Trim();

                    if (values[i].StartsWith("[") && values[i].EndsWith("]"))
                    {

                        values[i] = values[i].TrimStart('[');
                        values[i] = values[i].TrimEnd(']');
                        result.Add(values[i]);
                        continue;
                    }

                    if (values[i].StartsWith("[") && !values[i].EndsWith("]"))
                    {
                        //continuation of string
                        currently_processing_continuation = true;

                        values[i] = values[i].TrimStart('[');
                        continuationString.Add(values[i]);
                        continue;
                    }
                    if (!values[i].StartsWith("[") && !values[i].EndsWith("]") && !currently_processing_continuation)
                    {
                        result.Add(values[i]);
                        continue;
                    }
                    if (!values[i].StartsWith("[") && !values[i].EndsWith("]") && currently_processing_continuation)
                    {
                        continuationString.Add(values[i]);
                        continue;
                    }
                    if (!values[i].StartsWith("[") && values[i].EndsWith("]") && currently_processing_continuation)
                    {
                        values[i] = values[i].TrimEnd(']');
                        continuationString.Add(values[i]);
                        result.Add(String.Join(",", continuationString.ToArray()));
                        //end of continueation
                        continuationString.Clear();
                        currently_processing_continuation = false;
                        continue;
                    }
                }
                return result.ToArray();
            //}           
            //return values;
        }

       

        private enum Severity { Info, Warn, Debug, Critical }
        
        private static void CreateAndSaveOutboundRule(PhoneSystem ps,Tenant tenant, OutBoundRuleDetails outbound_rule_to_create)
        {            
            OutboundRule[] rules = tenant.GetOutboundRules();
            int last_maximum_priority = 0;

            

            /* Determine the highest priority outbound rule in the system at the moment.
             * While we're here, might as well see if this new rule exists or not...there are allowed to be duplicates though */
            foreach (OutboundRule rule in rules)
            {

                if (rule.Priority > last_maximum_priority)
                    last_maximum_priority = rule.Priority;

                if (rule.Name.Equals(outbound_rule_to_create.Name) == true)
                {                    
                    if (isDebug)
                        LogMessage("Duplicate rule found, " + rule.Name, Severity.Warn);
                }                

                if (print_rules == true)
                {
                    LogMessage(rule.Name, Severity.Info);
                }
            }

            if (tenant != null)
            {
                OutboundRule outbound_rule = null;
                outbound_rule = tenant.CreateOutboundRule();

                if (outbound_rule != null)
                {
                    /* Populate the outbound rule, using defaults where non is explicity set */

                    outbound_rule.Name = String.IsNullOrEmpty(outbound_rule_to_create.Name) ? outbound_rule.Name : outbound_rule_to_create.Name;
                    outbound_rule.Prefix = String.IsNullOrEmpty(outbound_rule_to_create.Prefix) ? outbound_rule.Prefix : outbound_rule_to_create.Prefix;
                    outbound_rule.NumberLengthRanges = String.IsNullOrEmpty(outbound_rule_to_create.Calls_to_numbers_with_a_length_of) ? outbound_rule.NumberLengthRanges : outbound_rule_to_create.Calls_to_numbers_with_a_length_of;
                    
                    /* Get a list of gateways the system is aware of */

                    Gateway[] gateways = ps.GetGateways();
                    bool found_user_specified_gateway_in_system = false;
                    if (gateways.Length > 0)
                    {
                        foreach (Gateway gw in gateways)
                        {
                            if (print_gateways)
                                LogMessage(String.Format("GATEWAY {0}", gw.Name), Severity.Info);
                            
                            /* GetAppPath the gateways that the user wants to add to this outbound_rule_to_create rule */
                             
                            if (gw.Name.Equals(outbound_rule_to_create.Outbound_route[0].name) == true)
                            {
                                outbound_rule.OutboundRoutes[0].Gateway = gw;                                                                
                                found_user_specified_gateway_in_system = true;
                            }

                        }
                        outbound_rule.OutboundRoutes[0].Prepend = outbound_rule_to_create.Outbound_route[0].prepend;
                        outbound_rule.OutboundRoutes[0].StripDigits = outbound_rule_to_create.Outbound_route[0].strip_count;
                    }
                    else
                    {
                        LogMessage("No gateways are defined in the system, no specifying one in the outbound rule", Severity.Warn);
                    }

                    /*Use default gateway if one is available and we cant find the use specified gateway name*/
                    if (!found_user_specified_gateway_in_system && gateways.Length > 0)
                    {
                        outbound_rule.OutboundRoutes[0].Gateway = gateways[0];
                        LogMessage("Using default gateway as route for outbound rule - route not found or not provided.",Severity.Warn);
                    }

                    
                    

                    if (outbound_rule_to_create.External_lines != null)
                    {
                        if (outbound_rule_to_create.External_lines.Length > 0)
                            outbound_rule.ExternalLines = outbound_rule_to_create.External_lines;
                    }

                    if ( String.IsNullOrEmpty(outbound_rule_to_create.ExtensionRange) == false)
                    {
                        string[] ranges = outbound_rule_to_create.ExtensionRange.Split(',');
                        //allocate space for extensio nranges
                        DNRange[] made_ranges = new DNRange[ranges.Length];
                        
                        for (int i = 0; i < ranges.Length; i++)
                        {
                            made_ranges[i] = outbound_rule.CreateDNRange();
                            if (ranges[i].Contains("-") == true)
                            {
                                string[] toAndFrom = ranges[i].Split('-');
                                made_ranges[i].From = toAndFrom[0];
                                made_ranges[i].To = toAndFrom[1];

                            }
                            else
                            {
                                // no range just single extension
                                made_ranges[i].From = ranges[i];
                                made_ranges[i].To = ranges[i];
                                
                            }                            
                        }
                        outbound_rule.DNRanges = made_ranges;
                    }

                    if (outbound_rule_to_create.DNGroups != null)
                    {
                        if (outbound_rule_to_create.DNGroups.Length > 0)
                            outbound_rule.DNGroups = outbound_rule_to_create.DNGroups;
                    }

                    /* NB - ensure that the priority of the successive Outbound rules are at consecutive multiples of 5 larger than the highest priority in the system */
                    outbound_rule.Priority = last_maximum_priority + 5; //this is a static variable because we cannot determine which order the last outbound rule was saved in when calling GetOutBoundRules
                    outbound_rule.Save(); // Save it to the system
                    tenant.Refresh(); // This is so that the next call to the Tenant Object's GetOutboundRules() contains reference to the rule we just created.
                    LogMessage("Saved outbound rule, " + outbound_rule.Name, Severity.Info);
                }
                else
                {
                    LogMessage("Could not create/allocate an Outbound rule template to use to create a new rule.", Severity.Critical);
                    
                }
            }
            else
            {                
                LogMessage("Tenant " + tenant.Name + " is not found but is required.", Severity.Critical);
            }
        }

        private static void LogMessage(string message, Severity severity)
        {
            
            ConsoleColorUtility TextChange = new ConsoleColorUtility();
            switch (severity)
            {
                case Severity.Debug:
                    if (!be_silent)
                    {
                        TextChange.TextColor((int)ConsoleColorUtility.Foreground.Grey);
                        Console.WriteLine("DEBUG: " + message);
                    }
                    break;
                case Severity.Critical:
                     TextChange.TextColor((int)ConsoleColorUtility.Foreground.Red);
                     Console.WriteLine("CRITICAL: " + message);
                    break;
                case Severity.Info:
                    if (!be_silent)
                    {
                        TextChange.TextColor((int)ConsoleColorUtility.Foreground.Green);
                        Console.WriteLine("INFO: " + message);
                    }
                    break;
                case Severity.Warn:
                    if (ignore_warnings == false || !be_silent)
                    {
                        //TextChange.TextColor((int)ConsoleColorUtility.Foreground.Blue +
                        //            (int)ConsoleColorUtility.Foreground.Intensity);
                        TextChange.TextColor((int)ConsoleColorUtility.Foreground.Yellow);
                        Console.WriteLine("WARN: " + message);
                    }
                    break;
            }
            TextChange.ResetColor();
        }

        static public string GetKeyValue(string Section, string KeyName, string FileName)
        {
            //Reading The KeyValue Method
            try
            {
                StringBuilder JStr = new StringBuilder(255);
                int i = GetKeyValueA(Section, KeyName, String.Empty, JStr, 255, FileName);
                return JStr.ToString();
            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        public static TCX.PBXAPI.PBXConnection CreatePbxConn()
        {
            String filePath = Path.Combine(GetAppPath(), @"Bin\3CXPhoneSystem.ini");
            String pbxUser = String.Empty,
                   pbxPass = String.Empty,
                   pbxHost = "127.0.0.1";
            Int32 pbxPort = 5482;//default value according to stepan
            String value = GetKeyValue("General", "PBXUser", filePath);
            if (!String.IsNullOrEmpty(value))
                pbxUser = value;
            value = GetKeyValue("General", "PBXPass", filePath);
            if (!String.IsNullOrEmpty(value))
                pbxPass = value;
            value = GetKeyValue("General", "CMHost", filePath);
            if (!String.IsNullOrEmpty(value))
                pbxHost = value;
            value = GetKeyValue("General", "CMPort", filePath);
            if (!String.IsNullOrEmpty(value))
            {
                Int32.TryParse(value, out pbxPort);
            }
            //obsolete. must not be used ?????????????????????????/
            //PBXConnection pbx = new PBXConnection();
            return new TCX.PBXAPI.PBXConnection(pbxHost, pbxPort, pbxUser, pbxPass);
        }

        public static String GetAppPath()
        {
            RegistryKey regKeyAppRoot;
            if (IntPtr.Size == 4)
            {
                regKeyAppRoot = Registry.LocalMachine.OpenSubKey("SOFTWARE\\3CX\\PhoneSystem");
            }
            else
            {
                regKeyAppRoot = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Wow6432Node\\3CX\\PhoneSystem");
            }
            return (string)regKeyAppRoot.GetValue("AppPath");
        }
    }
}
