using System;
using System.IO;
using System.Management;
using System.Text;

namespace InventoryAuditService
{
    class Audit
    {
        public void Execute()
        {
            setMonitorData(setPCData());
        }

        //gets the PC operating system name
        private static string GetOSFriendlyName()
        {
            string result = string.Empty;
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
            foreach (ManagementObject os in searcher.Get())
            {
                result = os["Caption"].ToString();
                break;
            }
            return result;
        }
        //Converts UInt16 values to a readable string
        private static string getString(UInt16[] Val)
        {
            StringBuilder sb = new StringBuilder();
            try
            {
                foreach (UInt16 u in Val)
                    sb.Append(char.ConvertFromUtf32(u));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return sb.ToString();
        }

        private static MonitorData setMonitorData(string serial)
        {

            MonitorData md = new MonitorData();
            //add to database
            Database d = new Database();

            try
            {
                //this queries the information of all available monitors
                ManagementObjectSearcher searcher = new ManagementObjectSearcher(@"root\WMI", "SELECT * FROM WmiMonitorID");
                foreach (ManagementObject obj in searcher.Get())
                {
                    foreach (PropertyData p in obj.Properties)
                    {
                        if (p.Value != null)
                        {
                            if (p.Value.GetType().ToString().Equals("System.UInt16[]"))
                            {

                                switch (p.Name)
                                {
                                    case "ManufacturerName":
                                        {
                                            md.vendorM = getString((UInt16[])p.Value);
                                            break;
                                        }
                                    case "SerialNumberID":
                                        {
                                            md.serialNumberM = getString((UInt16[])p.Value);
                                            break;
                                        }
                                    case "UserFriendlyName":
                                        {
                                            md.modelM = getString((UInt16[])p.Value);
                                            break;
                                        }
                                }

                            }
                        } else if(p.Value == null)
                        {
                            if(!d.verifyMonitor(md.serialNumberM))
                            {
                                md.attachedPC = serial;
                                md.attachedPC = " " + serial;
                                md.modelM = "Check";
                                md.serialNumberM = "The";
                                md.vendorM = "Monitor";
                                d.InsertMonitor(md);

                            }
                           
                        }
                    }

                    /*if (md.modelM != null)
                    {
                        if (!d.verifyMonitor(md.serialNumberM))
                        {
                            md.attachedPC = serial;
                            d.insertMonitor(md);
                        }
                    }*/
                }

            }
            catch (Exception e)
            {
                WriteToFile("Exception thrown in setMonitor " + e.Message);               
                d.InsertMonitor(md);
            }

            return md;

        }

        private static string setPCData()
        {
            PCData pd = new PCData();
            Database d = new Database();

            try
            {

                System.Management.SelectQuery query = new System.Management.SelectQuery(@"Select * from Win32_ComputerSystem");

                //initialize the searcher with the query it is supposed to execute
                using (System.Management.ManagementObjectSearcher searcher1 = new System.Management.ManagementObjectSearcher(query))
                {
                    //execute the query
                    foreach (System.Management.ManagementObject process in searcher1.Get())
                    {
                        //print system info
                        process.Get();

                        pd.vendorPC = "" + process["Manufacturer"];
                        pd.modelPC = "" + process["Model"];
                    }
                }
                //to start searching at Windows BIOS table for the device serial number
                //shows the serial number of the PC
                ManagementObjectSearcher MOS = new ManagementObjectSearcher("Select * From Win32_BIOS");

                foreach (ManagementObject getserial in MOS.Get())
                {
                    pd.serialNumberPC = getserial["SerialNumber"].ToString();
                }

                pd.systemName = System.Environment.MachineName;
                pd.version = GetOSFriendlyName();
                pd.domain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
                pd.assetName = pd.vendorPC + " " + pd.modelPC;


                if (!d.verifyPC(pd.serialNumberPC))
                {
                    d.InsertPC(pd);
                }
                else
                    d.UpdatePC(pd);
            }
            catch (Exception t)
            {
                WriteToFile("Exception thrown in setPC " + t.Message);
                pd.systemName = "Generic system name";
                pd.vendorPC = "this";
                pd.serialNumberPC = "Check";
                pd.version = "Generic PC Windows Version";
                pd.domain = "Generic domain";
                pd.modelPC = "PC";
                pd.assetName = pd.vendorPC + pd.modelPC + " should be checked";
                d.InsertPC(pd);
            }

            //pd.show();
            return pd.serialNumberPC;

        }

        private static void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Audit Tool Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filepath))
            {
                // Create a file to write to.   
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                    sw.Close();
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                    sw.Close();
                }
            }
        }

    }

}

