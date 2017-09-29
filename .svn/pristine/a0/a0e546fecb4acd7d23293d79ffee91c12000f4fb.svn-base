using System;
using System.Collections.Generic;
using System.Text;
using TCX.Configuration;
using TCX.PBXAPI;
using System.Threading;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.IO;
namespace TestCSharp
{
    class Program
    {
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileStringA")]
        static extern int GetKeyValueA(string strSection, string strKeyName, string strNull, StringBuilder RetVal, int nSize, string strFileName);

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
        static void Main(string[] args)
        {
            String action = "display";
            if (args.Length > 0)
                action = args[0].ToLower();
            {
                Random a = new Random(Environment.TickCount);
                //unique name PhoneSystem.ApplicationName = "TestApi";//any name
                PhoneSystem.ApplicationName = PhoneSystem.ApplicationName + action + a.Next().ToString();
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

            //PhoneSystem.CfgServerHost = "127.0.0.1";
            //PhoneSystem.CfgServerPort = 5485;
            PhoneSystem ps = null;
            while (ps == null)
            {
                try
                {
                    ps = PhoneSystem.Root;
                }
                catch (Exception ex)
                {
                    System.Console.WriteLine(ex.Message);
                }
            }
            {
                DN[] precache = ps.GetDN();
                foreach(DN dn__ in precache)
                {
                    //System.Console.WriteLine(dn__.ToString());
                    Extension ext = dn__ as Extension;
                    if(ext!=null)
                    {
			            {
				            ExtensionRule[] or = ext.ForwardingRules;
				            //System.Console.WriteLine("All:");
				            for(int i=0; i<or.Length; i++)
				            {
					            //System.Console.WriteLine("\t"+or[i].ToString());
				            }
			            }
			            if(ext.OverrideProfile!=null)
			            {
				            ExtensionRule[] or = ext.OverrideProfile.ForwardingRules;
				            //System.Console.WriteLine("Override:");
				            for(int i=0; i<or.Length; i++)
				            {
					            //System.Console.WriteLine("\t"+or[i].ToString());
				            }
			            }
			            if(ext.CurrentProfile!=null)
			            {
				            ExtensionRule[] or = ext.CurrentProfile.ForwardingRules;
				            //System.Console.WriteLine("Current:");
				            for(int i=0; i<or.Length; i++)
				            {
					            //System.Console.WriteLine("\t"+or[i].ToString());
				            }
			            }
                        if (ext.FallbackProfile != null)
			            {
				            ExtensionRule[] or = ext.FallbackProfile.ForwardingRules;
				            //System.Console.WriteLine("Fallback:");
				            for(int i=0; i<or.Length; i++)
				            {
					           // System.Console.WriteLine("\t"+or[i].ToString());
				            }
			            }
                    }
                }
            }
            GC.Collect();
            if (action != "display")
            {
                if (!ps.Connected)
                {
                    System.Console.WriteLine("Action \"" + action + "\" Requires established connection");
                    System.Console.Write("Waiting for ConfigurationServer");
                    while (!ps.Connected)
                    {
                        Thread.Sleep(1000);
                        System.Console.WriteLine(".");
                    }
                    System.Console.WriteLine("Connected");
                }
            }
            switch (action)
            {
                case "disconnect":
                    {
                        //while (true)
                        //{
                        //    DN e = ps.GetDNByNumber("100");
                        //}
                    }
                    break;
                case "nothing":
                    while (true)
                    {
                        Thread.Sleep(5000);
                        GC.Collect();
                    }
                    break;
                case "createpromptset":
                    {
                        try
                        {
                            PromptSet myPromptSet = PhoneSystem.Root.CreatePromptSet();

                            myPromptSet.CultureCode = "dn";
                            myPromptSet.LanguageCode = "dn";
                            myPromptSet.PromptSetName = "My Custom PromptSet";
                            myPromptSet.UseAlternateNumberPronunciation = true;
                            myPromptSet.Version = "0.0.1";
                            myPromptSet.Description = "my TestPrompt";
                            myPromptSet.Folder = "folder in IVR prompts";
                            myPromptSet.PromptSetType = PromptSetType.Custom;
                            System.Collections.Generic.List<Prompt> myListOfPrompts = new System.Collections.Generic.List<Prompt>();
                            for (int i = 0; i < 100; i++)
                            {
                                Prompt newPrompt = myPromptSet.CreatePrompt();
                                newPrompt.ID = "PRMPT" + i.ToString();//set string ID of prompt
                                newPrompt.Filename = "nameoffileinpromptsfolder" + i.ToString() + ".wav";//the name of file in prompts folder
                                newPrompt.Transcription = "text of prompt " + i.ToString();//description of prompt you can copy it from 
                                myListOfPrompts.Add(newPrompt);
                            }
                            myPromptSet.Prompts = myListOfPrompts.ToArray();
                            myPromptSet.Save();
                            System.Console.WriteLine("Prompt is Saved Successfully");
                            Parameter p = PhoneSystem.Root.GetParameterByName("ACPRMSET");
                            p.Value = myPromptSet.Folder;
                            p.Save();
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine("Failed to create prompt"+e.Message);
                        }
                    }
                    break;
                case "display":
                    {
                        while (true)
                        {
                            System.Console.WriteLine("Parameters:");
                            foreach (Parameter p in ps.GetParameters())
                            {
                                System.Console.WriteLine("\t" + p.ToString());
                            }
                            System.Console.WriteLine("Codecs:");
                            foreach (Codec c in ps.GetCodecs())
                            {
                                System.Console.WriteLine("\t" + c.ToString());
                            }
                            System.Console.WriteLine("Conditions:");
                            foreach (RuleCondition rc in ps.GetRuleConditions())
                            {
                                System.Console.WriteLine("\t" + rc.ToString());
                            }
                            System.Console.WriteLine("RuleHours:");
                            foreach (RuleHours rh in ps.GetRuleHourTypes())
                            {
                                System.Console.WriteLine("\t" + rh.ToString());
                            }
                            System.Console.WriteLine("RuleCalltype:");
                            foreach (RuleCallType rct in ps.GetRuleCallTypes())
                            {
                                System.Console.WriteLine("\t" + rct.ToString());
                            }
                            foreach (Tenant t in ps.GetTenants())
                            {
                                System.Console.WriteLine(t.ToString());
                                System.Console.WriteLine("\tOfficeHours:");
                                foreach (HoursRange hr in t.OfficeHoursRanges)
                                {
                                    System.Console.WriteLine("\t\t" + hr);
                                }
                                System.Console.WriteLine("\tExtensions:");
                                foreach (Extension e in t.GetExtensions())
                                {
                                    System.Console.WriteLine("\t\t" + e);
                                    System.Console.WriteLine("\t\t\tProperties:");
                                    foreach (DNProperty d in e.GetProperties())
                                    {
                                        System.Console.WriteLine("\t\t\t\t" + d);
                                    }
                                    System.Console.WriteLine("\t\t\tForwarding:");
                                    foreach (ExtensionRule er in e.ForwardingRules)
                                    {
                                        System.Console.WriteLine("\t\t\t\t" + er);
                                    }
                                }
                                foreach (IVR ivr in t.GetIVRs())
                                {
                                    System.Console.WriteLine("\t\t" + ivr.Name);
                                    System.Console.WriteLine("\t\t\t" + ivr.Number);
                                    System.Console.WriteLine("\t\t\t" + ivr.PromptFilename);
                                    System.Console.WriteLine("\t\t\t" + ivr.Timeout);
                                    System.Console.WriteLine("\t\t\t" + ivr.TimeoutForwardType.ToString());
                                    DN dn = ivr.TimeoutForwardDN;
                                    if (dn != null)
                                        System.Console.WriteLine("\t\t\t" + dn.ToString());
                                    else
                                        System.Console.WriteLine("\t\t\t" + "not specified");
                                }
                            }
                            System.GC.Collect();
                            Thread.Sleep(3000);
                        }
                    }
                    break;
                case "sethours":
                    {
                        Tenant tadd = null;
                        if (args.Length > 1)
                        {
                            foreach (Tenant t in ps.GetTenants())
                            {
                                if (t.Name == args[1])
                                {
                                    tadd = t;
                                    break;
                                }
                            }
                            if (tadd != null)
                            {
                                System.Collections.Generic.List<HoursRange> arr = new System.Collections.Generic.List<HoursRange>();
                                HoursRange r = tadd.CreateHoursRange();
                                r.DayOfWeek = System.DayOfWeek.Monday;

                                r.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);
                                r.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
                                arr.Add(r);
                                r = tadd.CreateHoursRange();
                                r.DayOfWeek = System.DayOfWeek.Tuesday;
                                r.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);
                                r.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
                                arr.Add(r);
                                tadd.OfficeHoursRanges = arr.ToArray();
                                tadd.Save();
                            }
                            else
                            {
                                System.Console.WriteLine("Tenant \"" + args[1] + "\" is not found");
                            }
                        }
                        else
                        {
                            System.Console.WriteLine("Tenant is not specified");
                        }
                    }
                    break;
                case "add":
                    {
                        Tenant tadd = null;
                        if (args.Length > 1)
                        {
                            foreach (Tenant t in ps.GetTenants())
                            {
                                if (t.Name == args[1])
                                {
                                    tadd = t;
                                    break;
                                }
                            }
                            if (tadd != null)
                            {
                                Extension eadd = null;
                                if (args.Length == 2)
                                {
                                    eadd = tadd.CreateExtension();
                                }
                                else
                                {
                                    eadd = tadd.CreateExtension(args[2]);
                                }

                                if (eadd != null)
                                {
                                    eadd.Save();
                                    DNProperty prp = eadd.CreateProperty();
                                    prp.Name = "TestProperty";
                                    prp.Type = PropertyType.String;
                                    prp.Value = "justFortest";
                                    //prp.Description = "Why description is not NULL?";
                                    prp.Save();
                                    System.Console.WriteLine(tadd);
                                    System.Console.WriteLine("\tAdded: " + eadd);
                                }
                                else
                                {
                                    System.Console.WriteLine("Extension \"" + args[2] + "\" already exists or number is malformed");
                                }
                            }
                            else
                            {
                                System.Console.WriteLine("Tenant \"" + args[1] + "\" is not found");
                            }
                        }
                        else
                        {
                            System.Console.WriteLine("Tenant is not specified");
                        }
                    }
                    break;
                case "remove":
                    {
                        if (args.Length == 1)
                        {
                            System.Console.WriteLine("Tenant is not specified");
                        }
                        else if (args.Length == 2)
                        {
                            System.Console.WriteLine("Extension is not specified");
                        }
                        else
                        {
                            bool bfound = false;
                            foreach (Tenant t in ps.GetTenants())
                            {
                                if (t.Name == args[1])
                                {
                                    bfound = true;
                                    bool bremoved = false;
                                    foreach (Extension e in t.GetExtensions())
                                    {
                                        if (e.Number == args[2])
                                        {
                                            System.Console.WriteLine(t);
                                            System.Console.WriteLine("\tRemoved: " + e);
                                            e.Delete();
                                            bremoved = true;
                                            break;
                                        }
                                    }
                                    if (!bremoved)
                                    {
                                        System.Console.WriteLine("Cannot find the extension \"" + args[2] + "\" in tenant \"" + t.Name + "\"");
                                    }
                                    break;
                                }
                            }
                            if (!bfound)
                            {
                                System.Console.WriteLine("Cannot find the tenant \"" + args[1] + "\"");
                            }
                        }
                    }
                    break;
                case "fwd":
                    {
                        Tenant[] tarr = ps.GetTenants();
                        if (tarr.Length > 0)
                        {
                            Extension[] ex = tarr[0].GetExtensions();
                            if (ex.Length > 1)
                            {
                                ExtensionRule er = ex[0].CreateForwardingRule();
                                er.Conditions.Condition = ps.GetRuleConditions()[0];
                                er.Conditions.Hours = ps.GetRuleHourTypes()[0];
                                er.Conditions.CallType = ps.GetRuleCallTypes()[0];
                                er.Data = "1234567";
                                er.Forward.To = DestinationType.External;
                                er.Forward.External = "987654321";

                                HoursRange hr = er.CreateHoursRange();
                                System.Collections.Generic.List<HoursRange> arr = new System.Collections.Generic.List<HoursRange>();
                                HoursRange r = er.CreateHoursRange();

                                r.DayOfWeek = System.DayOfWeek.Wednesday;
                                r.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);
                                r.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
                                arr.Add(r);
                                r = er.CreateHoursRange();
                                r.DayOfWeek = System.DayOfWeek.Thursday;
                                r.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);
                                r.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
                                arr.Add(r);
                                er.HoursRanges = arr.ToArray();
                                System.Collections.Generic.List<ExtensionRule> arrR = new System.Collections.Generic.List<ExtensionRule>();
                                arrR.Add(er);
                                ex[0].ForwardingRules = arrR.ToArray();
                                ex[0].Save();
                            }
                        }
                    }
                    break;
                case "ivrdelete":
                    {
                        Tenant tadd = null;
                        if (args.Length > 1)
                        {
                            foreach (Tenant t in ps.GetTenants())
                            {
                                if (t.Name == args[1])
                                {
                                    tadd = t;
                                    break;
                                }
                            }
                            if (tadd != null)
                            {
                                foreach (IVR ivr in tadd.GetIVRs())
                                {
                                    ivr.Delete();
                                }
                            }
                            else
                            {
                                System.Console.WriteLine("Tenant \"" + args[1] + "\" is not found");
                            }
                        }
                        else
                        {
                            System.Console.WriteLine("Tenant is not specified");
                        }
                        Thread.Sleep(3000);
                    }
                    break;
                case "ivr":
                    {
                        Tenant tadd = null;
                        if (args.Length > 1)
                        {
                            foreach (Tenant t in ps.GetTenants())
                            {
                                if (t.Name == args[1])
                                {
                                    tadd = t;
                                    break;
                                }
                            }
                            if (tadd != null)
                            {
                                IVR ivradd = tadd.CreateIVR();
                                if (ivradd != null)
                                {
                                    ivradd.Name = "TestIVR";
                                    ivradd.PromptFilename = "prompts\\onhold.wav";
                                    ivradd.Timeout = 356;
                                    ivradd.TimeoutForwardType = IVRForwardType.VoiceMail;
                                    ivradd.TimeoutForwardDN = tadd.GetExtensions()[0];
                                    ivradd.Save();
                                }
                            }
                            else
                            {
                                System.Console.WriteLine("Tenant \"" + args[1] + "\" is not found");
                            }
                        }
                        else
                        {
                            System.Console.WriteLine("Tenant is not specified");
                        }
                        Thread.Sleep(3000);
                    }
                    break;
                case "createperf":
                    {
                        Tenant t = ps.GetTenants()[0];
                        DateTime start = DateTime.Now;
                        for (int ei = 200; ei < 700; ei++)
                        {
                            String data = ei.ToString();
                            Extension e = t.CreateExtension(data);
                            e.AuthID = data;
                            e.AuthPassword = data;
                            e.BusyDetection = BusyDetectionType.UsePBXStatus;
                            e.DeliverAudio = false;
                            e.EmailAddress = "stepan@3cx.com";
                            e.Enabled = true;
                            e.FirstName = "F_" + data;
                            e.HidePresence = false;
                            e.Internal = false;
                            e.LastName = "L_" + data;
                            e.NoAnswerTimeout = 60;
                            e.Number = data;
                            e.OutboundCallerID = "OUTCID_" + data;
                            e.QueueStatus = QueueStatusType.LoggedOut;
                            e.RecordCalls = false;
                            e.SIPID = data;
                            e.SupportReinvite = true;
                            e.SupportReplaces = true;
                            e.UserStatus = UserStatusType.Available;
                            e.VMEmailOptions = VMEmailOptionsType.None;
                            e.VMEnabled = true;
                            e.VMPIN = data;
                            e.VMPlayCallerID = true;
                            e.VMPlayMsgDateTime = VMPlayMsgDateTimeType.Play24Hr;
                            e.Save();
                            System.Console.WriteLine(e.ToString());
                        }
                        DateTime end = DateTime.Now;
                        System.Console.WriteLine((end - start).ToString());
                        Thread.Sleep(5000);
                    }
                    break;
                case "deletetenant":
                    {
                        Tenant t = ps.GetTenants()[0];
                        t.Delete();
                    }
                    break;
                case "deleteperf":
                    {
                        Tenant t = ps.GetTenants()[0];
                        Extension[] tmp = t.GetExtensions();
                        String num;
                        DateTime start = DateTime.Now;
                        foreach (Extension e in tmp)
                        {
                            num = e.Number;
                            if (num.Length == 3 && num[0] > '1' && num[0] < '7')
                            {
                                System.Console.WriteLine(e.ToString());
                                e.Delete();
                            }
                        }
                        DateTime end = DateTime.Now;
                        System.Console.WriteLine((end - start).ToString());
                        Thread.Sleep(5000);
                    }
                    break;
                case "makecall":
                    {
                        PBXConnection pbx = CreatePbxConn();
                        for (; ; )
                        {
                            try
                            {
                                pbx.MakeCall(args[1], args[2]);
                            }
                            catch (Exception e)
                            {
                                System.Console.WriteLine(e.ToString());
                            }
                            ConsoleKeyInfo k = System.Console.ReadKey(false);
                            if (k.KeyChar == 'e')
                            {
                                break;
                            }
                        }

                    }
                    break;
                case "makecall2":
                    {
                        PBXConnection pbx = CreatePbxConn();
                        Dictionary<String, String> d = new Dictionary<String, String>();
                        for (int i = 2; i < args.Length; i++)
                        {
                            String[] a = args[i].Split(new char[]{'='});
                            d.Add(a[0], a[1]);
                        }
                        try
                        {
                            pbx.MakeCall(args[1], d);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                    }
                    break;
                case "drop":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.DropCall(System.Convert.ToInt32(args[1]), args[2]);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "divert":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.DivertCall(System.Convert.ToInt32(args[1]), args[2], args[3], (args.Length > 4 && System.Convert.ToInt32(args[4]) == 1) ? true : false);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "refresh":
                    {
                        PBXConnection pbxConnn = CreatePbxConn();
                        Gateway gateway = ps.GetGateways()[0];
                        ExternalLine[] lineArray = gateway.GetExternalLines();
                        foreach (ExternalLine extLine in lineArray)
                        {
                            pbxConnn.RefreshRegistration(extLine.Number);
                        }
                    }
                    Thread.Sleep(5000);
                    break;
                case "transfer":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.TransferCall(System.Convert.ToInt32(args[1]), args[2], args[3]);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "monitor":
                    {

                        {
                            MyListener aa = new MyListener(null);
                            ps.Updated += new NotificationEventHandler(aa.ps_Updated);
                            ps.Inserted += new NotificationEventHandler(aa.ps_Inserted);
                            ps.Deleted += new NotificationEventHandler(aa.ps_Deleted);
                        }
                        while (true)
                        {
                            //ActiveConnection[] prev=null;
                            foreach (DN dn in ps.GetDN())
                            {
                                System.Console.WriteLine(dn + " " + (dn.IsRegistered ? "REGISTERED" : "NOT REGISTERED") + "VMB(" + dn.VoiceMailBox.ToString() + ")");
                                ActiveConnection[] a = dn.GetActiveConnections();
                                ActiveConnection[] ab = dn.GetActiveConnections();
                                //for(int i=0; i<a.Length; i++)
                                //{
                                //    if(a[i]==ab[i])
                                //        System.Console.WriteLine("equal");
                                //    if(a[i]!=ab[i])
                                //        System.Console.WriteLine("notequal");
                                //}
                                if (dn.IsRegistered || a.Length > 0)
                                {
                                    //prev = a;
                                    foreach (ActiveConnection ac in a)
                                    {
                                        //if (prev != null && prev.Length > 0)
                                        //{
                                        //    if(ac!=prev[0])
                                        //        System.Console.WriteLine("notequal");
                                        //    if(ac==prev[0])
                                        //        System.Console.WriteLine("equal");
                                        //}
                                        System.Console.WriteLine(ac.CallID);
                                        System.Console.WriteLine(ac.InternalParty);
                                        System.Console.WriteLine(ac.IsInbound);
                                        System.Console.WriteLine(ac.IsOutbound);
                                        System.Console.WriteLine(ac.Status);
                                        System.Console.WriteLine(DateTime.Now - ac.LastChangeStatus);
                                        System.Console.WriteLine(ac.ExternalParty);
                                        System.Console.WriteLine("AttachedData(" + ac.AttachedData.Count.ToString() + "):");
                                        foreach (KeyValuePair<String, String> kvp in ac.AttachedData)
                                        {
                                            System.Console.WriteLine(kvp.Key + "=" + kvp.Value);
                                        }
                                    }
                                }
                            }
                            Thread.Sleep(1000);
                        }
                    }
                    break;
                case "gwpars":
                    {
                        for (; ; )
                        {
                            foreach (GatewayParameter p in ps.GetGatewayParameters())
                            {
                                //System.Console.WriteLine(p.Name);
                                //System.Console.WriteLine("\tSourceID");
                                foreach (GatewayParameterValue v in p.PossibleValuesAsSourceID)
                                {
                                    //System.Console.WriteLine("\t\t" + v.Name);
                                }
                                //System.Console.WriteLine("\tInbound");
                                foreach (GatewayParameterValue v in p.PossibleValuesAsInbound)
                                {
                                    //System.Console.WriteLine("\t\t" + v.Name);
                                }
                                //System.Console.WriteLine("\tOutbound");
                                foreach (GatewayParameterValue v in p.PossibleValuesAsOutbound)
                                {
                                    //System.Console.WriteLine("\t\t" + v.Name);
                                }
                                System.Console.Write(".");
                            }
                            Thread.Sleep(0);
                        }
                    }
                    break;
                case "ruleupdatetest":
                    {
                        ExternalLine _extLine = ps.GetVoipProviders()[0].GetExternalLines()[0];
                        bool bEndCall = false;
                        for (; ; )
                        {
                            ExternalLineRule lineRuleOff = null;
                            ExternalLineRule lineRuleOutOff = null;
                            foreach (ExternalLineRule extLineRule in _extLine.RoutingRules)
                            {
                                if (extLineRule.Conditions.Condition.Type == RuleConditionType.ForwardAll //ask stepan?????? 
                                    && extLineRule.Conditions.Hours.Type == RuleHoursType.OfficeHours
                                    && extLineRule.Conditions.CallType.Type == RuleCallTypeType.AllCalls)
                                {
                                    lineRuleOff = extLineRule;
                                }
                                if (extLineRule.Conditions.Condition.Type == RuleConditionType.ForwardAll //ask stepan?????? 
                                    && extLineRule.Conditions.Hours.Type == RuleHoursType.OutOfOfficeHours
                                    && extLineRule.Conditions.CallType.Type == RuleCallTypeType.AllCalls)
                                {
                                    lineRuleOutOff = extLineRule;
                                }
                            }
                            if (lineRuleOff != null)
                            {
                                if (!bEndCall)
                                {
                                    lineRuleOff.Forward.To = DestinationType.Extension;
                                    lineRuleOff.Forward.Internal = ps.GetTenants()[0].GetExtensions()[0];
                                }
                                else
                                {
                                    lineRuleOff.Forward.To = DestinationType.None;
                                }
                            }
                            if (lineRuleOutOff != null)
                            {
                                if (!bEndCall)
                                {
                                    lineRuleOutOff.Forward.To = DestinationType.Extension;
                                    lineRuleOutOff.Forward.Internal = ps.GetTenants()[0].GetExtensions()[0];
                                }
                                else
                                {
                                    lineRuleOutOff.Forward.To = DestinationType.None;
                                }
                            }
                            _extLine.Save();
                            bEndCall = !bEndCall;
                            Thread.Sleep(5000);
                        }
                    }
                    break;
                case "setmwi":
                    {
                        DN dn = ps.GetDNByNumber(args[1]);
                        VMBInformation vbm = new VMBInformation(2, 1);
                        dn.VoiceMailBox = vbm;
                        Thread.Sleep(5000);
                    }
                    break;
                case "outboundrule":
                    {
                        Tenant t = ps.GetTenants()[0];
                        OutboundRule r = t.CreateOutboundRule();
                        r.OutboundRoutes[0].Gateway = ps.GetGateways()[0];
                        r.OutboundRoutes[0].Gateway = null;
                    }
                    break;
                case "parkcodes":
                    try
                    {
                        String newParkCode = args[1];
                        String newUnparkCode = args[2];

                        Parameter parkCode = ps.GetParameterByName("PARK");
                        Parameter unparkCode = ps.GetParameterByName("UNPARK");
                        DN oldPark = null;
                        DN oldUnpark = null;
                        //check park code
                        if (parkCode == null || newParkCode != parkCode.Value)
                        {
                            DN d = ps.GetDNByNumber(newParkCode);
                            if (d != null && !(d is ParkExtension))
                            {
                                throw new Exception("park code is allocated to enother entity");
                            }
                            if (parkCode != null)
                            {
                                oldPark = (ps.GetDNByNumber(parkCode.Value) as ParkExtension);
                            }
                        }
                        //check unpark code
                        if (unparkCode == null || newParkCode != unparkCode.Value)
                        {
                            DN d = ps.GetDNByNumber(newUnparkCode);
                            if (d != null && !(d is ParkExtension))
                            {
                                throw new Exception("unpark code is allocated to enother entity");
                            }
                            if (unparkCode != null)
                            {
                                oldUnpark = (ps.GetDNByNumber(unparkCode.Value) as ParkExtension);
                            }
                        }
                        if (parkCode == null)
                        {
                            parkCode = ps.CreateParameter();
                            parkCode.Type = ParameterType.String;
                            parkCode.Name = "PARK";
                        }

                        if (unparkCode == null)
                        {
                            unparkCode = ps.CreateParameter();
                            unparkCode.Type = ParameterType.String;
                            unparkCode.Name = "UNPARK";
                        }

                        ParkExtension[] pe = ps.GetTenants()[0].GetParkExtensions();
                        ParkExtension parkUpdate = null;
                        ParkExtension unparkUpdate = null;

                        foreach (DN p in pe)
                        {
                            if (oldPark == p)
                            {
                                parkUpdate = oldPark as ParkExtension;
                            }
                            if (oldUnpark == p)
                            {
                                unparkUpdate = oldUnpark as ParkExtension;
                            }
                        }

                        if (parkUpdate == null || unparkUpdate == null)
                        {
                            foreach (ParkExtension p in pe)
                            {
                                p.Delete();
                            }
                            Thread.Sleep(2000); //just timeout for update
                            parkUpdate = ps.GetTenants()[0].CreateParkExtension();
                            unparkUpdate = ps.GetTenants()[0].CreateParkExtension();
                        }

                        parkUpdate.Number = newParkCode;
                        parkUpdate.Save();
                        parkCode.Value = newParkCode;
                        parkCode.Save();
                        unparkUpdate.Number = newUnparkCode;
                        unparkCode.Value = newUnparkCode;
                        unparkCode.Save();
                        unparkUpdate.Save();


                    }
                    catch (Exception ex)
                    {
                        System.Console.WriteLine(ex.ToString());
                    }
                    break;
                case "parkmonitor":
                    for (; ; )
                    {
                        /*                        ParkExtension orbit = ps.GetDNByNumber(ps.GetParameterByName("PARK").Value) as ParkExtension; //this method will be deprecated and replaced with Tenant.GetDNByNumber() method in future releases*/
                        ParkExtension orbit = null;
                        ParkExtension unorbit = null;
                        foreach (DN dn in ps.GetTenants()[0].GetDN())
                        {
                            if (dn.Number == ps.GetParameterByName("PARK").Value)
                            {
                                orbit = dn as ParkExtension;
                            }
                            if (dn.Number == ps.GetParameterByName("UNPARK").Value)
                            {
                                unorbit = dn as ParkExtension;
                            }
                        }
                        ActiveConnection[] parkedCalls = orbit.GetActiveConnections();

                        foreach (ActiveConnection ac in parkedCalls)
                        {
                            if (ac.InternalParty != null)
                            {
                                System.Console.WriteLine("Call(" + ac.CallID + "): parked on " + orbit.Number + "remoteParty is " + ac.InternalParty + "(" + ac.ExternalParty + ")");
                            }
                        }
                        ActiveConnection[] unparkedCalls = unorbit.GetActiveConnections();
                        foreach (ActiveConnection ac in unparkedCalls)
                        {
                            if (ac.InternalParty != null)
                            {
                                System.Console.WriteLine("Call(" + ac.CallID + "): parked on " + orbit.Number + "remoteParty is " + ac.InternalParty + "(" + ac.ExternalParty + ")");
                            }
                        }
                        Thread.Sleep(5000);
                    }
                    break;
                case "notifymonitor":
                    {
                        String filter = null;
                        if (args.Length > 1)
                            filter = args[1];
                        MyListener a = new MyListener(filter);
                        ps.Updated += new NotificationEventHandler(a.ps_Updated);
                        ps.Inserted += new NotificationEventHandler(a.ps_Inserted);
                        ps.Deleted += new NotificationEventHandler(a.ps_Deleted);
                        while (true)
                        {
                            Thread.Sleep(5000);
                        }
                    }
                    break;
                case "record":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.RecordCall(System.Convert.ToInt32(args[1]), args[2], System.Convert.ToInt32(args[3])!=0);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "bargein":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.BargeinCall(System.Convert.ToInt32(args[1]), args[2], false);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "listen":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.Listen(System.Convert.ToInt32(args[1]), args[2]);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "whisper":
                    {
                        try
                        {
                            PBXConnection pbx = CreatePbxConn();
                            pbx.WhisperTo(System.Convert.ToInt32(args[1]), args[2], args[3]);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(10000);
                    }
                    break;
                case "refreshtest":
                    {
                        PhoneSystem.InternalStat stat=new PhoneSystem.InternalStat();
                        Extension extensionForTest = ps.GetTenants()[0].GetExtensions()[0];
                        for (int i=0;;i++)
                        {
                            extensionForTest.Refresh();
                            Thread.Sleep(10);
                            if ((i++ % 100) == 0)
                            {
                                GC.Collect();
                                ps.GetInternalStat(stat);
                                System.Console.WriteLine(stat.ToString());
                            }
                        }
                    }
                    break;
                case "updatetest":
                    {
                        MyListener the = new MyListener(null);
                        ps.Updated += new NotificationEventHandler(the.ps_Updated);
                        ps.Inserted += new NotificationEventHandler(the.ps_Inserted);
                        ps.Deleted += new NotificationEventHandler(the.ps_Deleted);
                        bool audioopt = true;
                        while (true)
                        {
                            audioopt = !audioopt;
                            try
                            {
                                foreach (Extension a in ps.GetTenants()[0].GetExtensions())
                                {
                                    a.DeliverAudio = audioopt;
                                    a.Save();
                                }
                            }
                            catch (Exception)
                            {
                                System.Console.WriteLine("Exception in updatetest");
                            }
                            GC.Collect();
                            Thread.Sleep(10000);
                        }
                    }
                    break;
                case "setmulticast":
                    {
                        RingGroup dn = ps.GetDNByNumber(args[1]) as RingGroup;
                        if (dn != null && dn.RingStrategy == RingGroup.StrategyType.Paging)
                        {
                            if (args.Length > 2)
                            {
                                DNProperty prop = null;
                                prop = dn.GetPropertyByName("MULTICASTADDR");
                                if (prop == null)
                                {
                                    prop = dn.CreateProperty();
                                    prop.Name = "MULTICASTADDR";
                                }
                                prop.Value = args[2];
                                prop.Save();
                                prop = dn.GetPropertyByName("MULTICASTPORT");
                                if (prop == null)
                                {
                                    prop = dn.CreateProperty();
                                    prop.Name = "MULTICASTPORT";
                                }
                                prop.Value = args[3];
                                prop.Save();
                                prop = dn.GetPropertyByName("MULTICASTCODEC");
                                if (prop == null)
                                {
                                    prop = dn.CreateProperty();
                                    prop.Name = "MULTICASTCODEC";
                                }
                                prop.Value = args[4];
                                prop.Save();
                                prop = dn.GetPropertyByName("MULTICASTPTIME");
                                if (prop == null)
                                {
                                    prop = dn.CreateProperty();
                                    prop.Name = "MULTICASTPTIME";
                                }
                                prop.Value = args[5];
                                prop.Save();
                            }
                            else
                            {
                                DNProperty prop = null;
                                prop = dn.GetPropertyByName("MULTICASTADDR");
                                if (prop != null)
                                {
                                    prop.Delete();
                                }
                                prop = dn.GetPropertyByName("MULTICASTPORT");
                                if (prop != null)
                                {
                                    prop.Delete();
                                }
                                prop = dn.GetPropertyByName("MULTICASTCODEC");
                                if (prop != null)
                                {
                                    prop.Delete();
                                }
                                prop = dn.GetPropertyByName("MULTICASTPTIME");
                                if (prop != null)
                                {
                                    prop.Delete();
                                }
                            }
                        }
                    }
                    break;
                case "extsave":
                    {
                        Extension testStorage = ps.GetDNByNumber("699") as Extension;
                        testStorage.EnableSRTP = !testStorage.EnableSRTP;
                        testStorage.Save();
                        Thread.Sleep(5000);
                        break;
                    }
                case "phonebook":
                    {
                        Tenant t = ps.GetTenants()[0];
                        DN e=null;
                        foreach (DN dn in t.GetDN())
                        {
                            if (dn.Number == "108")
                            {
                                e = dn;
                            }
                        }
                        PhoneBookEntry a = t.CreatePhoneBookEntry();
                        a.FirstName = "TenantFN";
                        a.LastName = "TenantLN";
                        a.PhoneNumber = "54321";
                        a.Save();
                        a = e.CreatePhoneBookEntry();
                        a.FirstName = "ExtFN";
                        a.LastName = "ExtLN";
                        a.PhoneNumber = "7890";
                        a.Save();
                        Thread.Sleep(2000);
                        t.Refresh();
                        e.Refresh();
                        foreach (PhoneBookEntry pbe in t.GetPhoneBookEntries())
                        {
                            System.Console.WriteLine(pbe.ToString());
                        }
                        foreach (PhoneBookEntry pbe in e.GetPhoneBookEntries())
                        {
                            System.Console.WriteLine(pbe.ToString());
                        }
                        Thread.Sleep(2000);
                        break;
                    }
                case "createfax":
                    {
                        FaxExtension fe = ps.GetTenants()[0].CreateFaxExtension();
                        fe.Number = args[1];
                        fe.AuthID = args[2];
                        fe.AuthPassword = args[3];
                        fe.Save();
                        Thread.Sleep(2000);
                        break;
                    }
                    break;
                case "createfwdprofile":
                    {
                        Tenant[] tarr = ps.GetTenants();
                        if (tarr.Length > 0)
                        {
                            Extension[] ex = tarr[0].GetExtensions();
                            FwdProfile[] fps = ex[0].FwdProfiles;
                            if (fps.Length > 0)
                            {
                                ex[0].CurrentProfile = fps[0];
                                ex[0].Save();
                            }
                            if (ex.Length > 1)
                            {
                                ExtensionRule er = ex[0].CreateForwardingRule();
                                er.Conditions.Condition = ps.GetRuleConditions()[0];
                                er.Conditions.Hours = ps.GetRuleHourTypes()[0];
                                er.Conditions.CallType = ps.GetRuleCallTypes()[0];
                                er.Data = "1234567";
                                er.Forward.To = DestinationType.External;
                                er.Forward.External = "987654321";

                                HoursRange hr = er.CreateHoursRange();
                                System.Collections.Generic.List<HoursRange> arr = new System.Collections.Generic.List<HoursRange>();
                                HoursRange r = er.CreateHoursRange();

                                r.DayOfWeek = System.DayOfWeek.Wednesday;
                                r.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);
                                r.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
                                arr.Add(r);
                                r = er.CreateHoursRange();
                                r.DayOfWeek = System.DayOfWeek.Thursday;
                                r.StartTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 0, 0);
                                r.EndTime = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 0, 0);
                                arr.Add(r);
                                er.HoursRanges = arr.ToArray();
                                System.Collections.Generic.List<ExtensionRule> arrR = new System.Collections.Generic.List<ExtensionRule>();
                                arrR.Add(er);
                                ExtensionRule er2 = ex[0].CreateForwardingRule();
                                er2.Conditions.Condition = ps.GetRuleConditions()[1];
                                er2.Conditions.Hours = ps.GetRuleHourTypes()[1];
                                er2.Conditions.CallType = ps.GetRuleCallTypes()[1];
                                er2.Data = "1234567";
                                er2.Forward.To = DestinationType.External;
                                er2.Forward.External = "1234567890";
                                arrR.Add(er2);
                                FwdProfile fp = ex[0].CreateFwdProfile("Tratata");
                                fp.ForwardingRules = arrR.ToArray();
                                fp.Save();
                                FwdProfile fp1 = ex[0].CreateFwdProfile("Tratata1");
                                arrR.Clear();
                                arrR.Add(fp.ForwardingRules[0]);
                                fp1.ForwardingRules = arrR.ToArray();
                                fp1.Save();
                                ex[0].Save();
                            }
                        }
                    }
                    break;
                case "invoke":
                    {
                        String command_name = args[1];
                        Dictionary<String, String> parameters=new Dictionary<String, String>();
                        for (int i = 2; i < args.Length; i++)
                        {
                            String[] a = args[i].Split(new char[]{'='});
                            if (a.Length >= 2)
                            {
                                parameters.Add(a[0], String.Join("=", a, 1, a.Length - 1));
                            }
                            else
                            {
                                System.Console.WriteLine(args[i] + " ignored");
                            }
                        }

                        PBXConnection pbx = CreatePbxConn();
                        try
                        {
                            pbx.InvokeCommand(command_name, parameters);
                        }
                        catch (Exception e)
                        {
                            System.Console.WriteLine(e.ToString());
                        }
                        Thread.Sleep(2000);
                    }
                    break;
                case "newphonemodel":
                    {
                        PhoneModel a = ps.CreatePhoneModel();
                        a.CanBlankSDP = true;
                        a.CanReceiveOnly = true;
                        a.CanReinvite = true;
                        a.CanReplaces = true;
                        a.Manufacturer = "MyPhone";
                        a.ModelName = "MyModel";
                        a.Revision = "123";
                        a.UserAgentIdentifier = "my user agent";
                        a.Save();
                    }
                    break;
                case "statnotify":
                    {
                        String filter = null;
                        if (args.Length > 1)
                            filter = args[1];
                        MyListener a = new MyListener(filter);
                        ps.Updated += new NotificationEventHandler(a.ps_Updated);
                        ps.Inserted += new NotificationEventHandler(a.ps_Inserted);
                        ps.Deleted += new NotificationEventHandler(a.ps_Deleted);
                        Statistics[] myStat;
                        myStat = ps.InitializeStatistics();
                        foreach(Statistics s in myStat)
                        {
                            System.Console.WriteLine(s.ToString());
                        }

                        while (true)
                        {
                            Thread.Sleep(5000);
                        }
                    }
                    break;
                case "statdelete":
                    {
                        Statistics myStat = ps.CreateStatistics("MYSTAT");
                        Dictionary<String, String> original = myStat.Content;
                        System.Console.WriteLine("Current:");
                        foreach (KeyValuePair<String, String> kv in original)
                        {
                            System.Console.WriteLine(kv.Key + "=" + kv.Value);
                        }
                        System.Console.WriteLine(myStat.ToString());
                        myStat.Delete();
                        System.Console.WriteLine("After deleted:");
                        Dictionary<String, String> a = myStat.Content;
                        foreach (KeyValuePair<String, String> kv in a)
                        {
                            System.Console.WriteLine(kv.Key + "=" + kv.Value);
                        }
                        System.Console.WriteLine(myStat.ToString());
                        Thread.Sleep(5000);
                        myStat.Content = original;
                        System.Console.WriteLine("After assigning content:");
                        a = myStat.Content;
                        foreach (KeyValuePair<String, String> kv in a)
                        {
                            System.Console.WriteLine(kv.Key + "=" + kv.Value);
                        }
                        System.Console.WriteLine(myStat.ToString());
                        Thread.Sleep(5000);
                        myStat.update();
                        System.Console.WriteLine("After update:");
                        a = myStat.Content;
                        foreach (KeyValuePair<String, String> kv in a)
                        {
                            System.Console.WriteLine(kv.Key + "=" + kv.Value);
                        }
                        System.Console.WriteLine(myStat.ToString());
                        Thread.Sleep(10000);
                    }
                    break;
                case "statcheck":
                    {
                        String[] strs ={ "tratata", "blatata" };
                        Statistics myStat;
                        myStat = ps.CreateStatistics("MYSTAT");
                        bool swap = false;
                        String filter = null;
                        if (args.Length > 1)
                            filter = args[1];
                        MyListener a = new MyListener(filter);
                        ps.Updated += new NotificationEventHandler(a.ps_Updated);
                        ps.Inserted += new NotificationEventHandler(a.ps_Inserted);
                        ps.Deleted += new NotificationEventHandler(a.ps_Deleted);
                        int i = 0;
                        while (true)
                        {
                            if((++i%5)==0)
                                myStat.clearall();
                            else
                            {
                                myStat["s1"] = strs[swap ? 1 : 0];
                                myStat["s2"] = strs[swap ? 0 : 1];
                            }
                            swap = !swap;
                            try
                            {
                                myStat.update();
                                System.Console.WriteLine("("+i.ToString()+")NewStat="+myStat.ToString()+"\n------------");
                            }
                            catch (Exception e)
                            {
                                System.Console.WriteLine("Exception catched" + e.ToString());
                            }
                            Thread.Sleep(1);
                        }
                    }
                    break;
                case "delfromprofile":
                    {
                        Extension ext = ps.GetDNByNumber("102") as Extension;

                        List<ExtensionRule> extRuleList = new List<ExtensionRule>(ext.ForwardingRules);
                        ExtensionRule ruleToDelete = ext.ForwardingRules[2];
                        if (extRuleList.Contains(ruleToDelete))
                            extRuleList.Remove(ruleToDelete);
                        extRuleList.Clear();
                        ext.ForwardingRules = extRuleList.ToArray();
                        ext.Save();
                        Thread.Sleep(5000);
                        break;
                    }
                    break;
                case "delcurrentprofile":
                    {
                        Extension ext = ps.GetDNByNumber("112") as Extension;
                        FwdProfile fp = ext.CurrentProfile;
                        if (fp != null)
                        {
                            fp.Delete();
                        }
                    }
                    break;
                case "setofficehours":
                    {
                        Tenant a = ps.GetTenants()[0];
                        List<HoursRange> myHours = new List<HoursRange>();
                        HoursRange hr = a.CreateHoursRange();
                        hr.DayOfWeek = System.DayOfWeek.Monday;
                        hr.StartTime = new DateTime(1970, 1, 1, 8, 0, 0);
                        hr.EndTime = new DateTime(1970, 1, 1, 9, 0, 0);
                        myHours.Add(hr);
                        a.OfficeHoursRanges = myHours.ToArray();
                        a.Save();
                        Thread.Sleep(5000);
                        //hr = a.CreateHoursRange();
                        //hr.DayOfWeek = System.DayOfWeek.Tuesday;
                        //t.Hour = 9;
                        //t.Minute = 0;
                        //hr.StartTime = t;
                        //t.Hour = 10;
                        //t.Minute = 0;
                        //hr.EndTime = t;
                        //myHours.Add(hr);
                        //hr = a.CreateHoursRange();
                        //hr.DayOfWeek = System.DayOfWeek.Wednesday;
                        //t.Hour = 11;
                        //t.Minute = 0;
                        //hr.StartTime = t;
                        //t.Hour = 12;
                        //t.Minute = 0;
                        //hr.EndTime = t;
                        //myHours.Add(hr);
                        //hr = a.CreateHoursRange();
                        //hr.DayOfWeek = System.DayOfWeek.Thursday;
                        //t.Hour = 13;
                        //t.Minute = 0;
                        //hr.StartTime = t;
                        //t.Hour = 14;
                        //t.Minute = 0;
                        //hr.EndTime = t;
                        //myHours.Add(hr);
                        //hr = a.CreateHoursRange();
                        //hr.DayOfWeek = System.DayOfWeek.Friday;
                        //t.Hour = 15;
                        //t.Minute = 0;
                        //hr.StartTime = t;
                        //t.Hour = 16;
                        //t.Minute = 0;
                        //hr.EndTime = t;
                        //myHours.Add(hr);

                    }
                    break;

                case "propertysavedelete":
                    {
                        Extension a = PhoneSystem.Root.GetDNByNumber("100") as Extension;
                        bool yesno=false;
                        DNProperty c = a.GetPropertyByName("TESTPROP");
                        if (c != null)
                            c.Delete();
                        for (; ; )
                        {
                            try
                            {
                                a.Save();
                                a.Refresh();
                                c = a.GetPropertyByName("TESTPROP");
                                if (yesno && c == null
                                    || (!yesno && c != null))
                                {
                                    System.Console.WriteLine("Failed: expected - " + (yesno ? "doesn't exisit" : "exists"));
                                    continue;
                                }
                                if (yesno)
                                {
                                    c.Delete();
                                    yesno = false;
                                }
                                else
                                {
                                    c = a.CreateProperty();
                                    c.Description = "Blablabla";
                                    c.Name = "TESTPROP";
                                    c.Value = "MyVal";
                                    c.Save();
                                    yesno = true;
                                }
                            }
                            catch
                            {
                                System.Console.WriteLine("Failed: expected - " + (yesno ? "doesn't exisit" : "exists"));
                                a.Refresh();
                                c = a.GetPropertyByName("TESTPROP");
                                yesno = (c != null);
                            }
                        }
                    }
                    break;
                case "createsp":
                    {

                        Tenant a = ps.GetTenants()[0];
                        foreach(ParkExtension parks in a.GetParkExtensions())
                        {
                            if (parks.Number == args[1])
                            {
                                System.Console.WriteLine("Already exists");
                                return;
                            }   
                        }
                        ParkExtension sp = a.CreateParkExtension();
                        sp.Number = args[1];
                        sp.Save();
                        Thread.Sleep(5000);
                        break;
                    }
                default:
                    {
                        System.Console.WriteLine("Available operations are: add, remove, display(default)");
                    }
                    break;
            }
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
