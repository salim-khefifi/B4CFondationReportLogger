using log4net.Config;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Runtime.InteropServices;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace WindowsService
{
    public partial class Service1 : ServiceBase
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);
        public Service1()
        {
            InitializeComponent();
            var assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            XmlConfigurator.Configure(new FileInfo(Path.Combine(assemblyFolder, "Logger.config")));
        }

        System.Timers.Timer timer = null;

        public string ReadXmlConfig(string NodePath, string xmlTag)
        {
                string result = string.Empty;

                XmlDocument doc = new XmlDocument();
                //doc.Load(Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + @"C:\SALIM\B4CES_Work\B4CFondationReportLogger\ActivityReports.xml");
                doc.Load(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\ActivityReports.xml");
                //XmlNodeList node =  doc.SelectNodes("/Cruise/DatabaseConfiguration");        
                XmlNodeList node = doc.SelectNodes("/" + NodePath);
                //string SeparateUnitsCode = node[0]["DataBaseServer"].InnerText;
                result = node[0][xmlTag].InnerText;
                //log.Debug(result);
                return result;
        }

        protected override void OnStart(string[] args)
        {
            // Starting Service
            string path = @"C:\SALIM\B4CES_Work\B4CFondationReportLogger\sample.txt";

            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format("Windows Service is called on " + DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt") + ""));
                writer.Close();
            }

            //Create a timer with a 0.5 second interval
            timer = new System.Timers.Timer();
            timer.Interval = 500; // 0.5 second
            timer.Elapsed += new System.Timers.ElapsedEventHandler(OnTimer);
            timer.Enabled = true;

            ActivityReportsToDo();
            ActivityReports();
            ActivityReportsWaitingForSupport();
            ActivityReportsWaitingForCustomer();
            ActivityReportsWaitingForApproval();
            ActivityReportsPending();
            ActivityReportsCanceled();
            ActivityReportsDone();
        }

        public void OnTimer(object sender, System.Timers.ElapsedEventArgs args)
        {
            //Working();
            //timer.Stop();
        }

        public void SendMail(string body) 
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                //string subject = "Activity Report [B4CES Team Dev]";
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var message = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [In PROGRESS] :" + "\n" + body 
                })
                {
                    try
                    {
                        message.CC.Add(ccAddress);
                        smtp.Send(message);
                        //MessageBox.Show("Successfully send email");
                        log.Debug("Successfully send email");

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailToDo(string bodyToDo)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [To Do] :" + "\n" + bodyToDo 
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        smtp.Send(messageToDo);
                        //MessageBox.Show("Successfully send email");
                        log.Debug("Successfully send email");

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailWaitSupp(string bodyWaitSupp)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [WAITING FOR SUPPORT] :" + "\n" + bodyWaitSupp
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        smtp.Send(messageToDo);
                        //MessageBox.Show("Successfully send email");
                        log.Debug("Successfully send email");

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailWaitCust(string bodyWaitCust)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [WAITING FOR CUSTOMER] :" + "\n" + bodyWaitCust
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        smtp.Send(messageToDo);
                        //MessageBox.Show("Successfully send email");
                        log.Debug("Successfully send email");

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailWaitApp(string bodyWaitApp)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [WAITING FOR APPROVAL] :" + "\n" + bodyWaitApp
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        smtp.Send(messageToDo);
                        //MessageBox.Show("Successfully send email");
                        log.Debug("Successfully send email");

                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailPending(string bodyPending)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [PENDING] :" + "\n" + bodyPending 
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        if (bodyPending.ToString() != "\n" + "\n" + "\n" + "\n" + "\n" + "\n" )
                        {
                            smtp.Send(messageToDo);
                            log.Debug("Successfully send email");
                        }
                        else
                        {
                            log.Debug("Body Must Be Not Empty");
                        }
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailCanceled(string bodyCanceled)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [CANCELED] :" + "\n" + bodyCanceled
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        if ( bodyCanceled.ToString() != "\n" + "\n" + "\n" + "\n" + "\n" + "\n")
                        {
                            smtp.Send(messageToDo);
                            log.Debug("Successfully send email");
                        }
                        else
                        {
                            log.Debug("Body Must Be Not NULL");
                        }
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void SendMailDone(string bodyDone)
        {
            try
            {
                string FromAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "FromAddress");
                string ToAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "ToAddress");
                string CcAddress = ReadXmlConfig("ActivityReports/MailConfiguration", "CcAdress");
                string FromPassword = ReadXmlConfig("ActivityReports/MailConfiguration", "FromPassword");
                var fromAddress = new MailAddress(FromAddress, "B4Creation Software Solution Team");
                var toAddress = new MailAddress(ToAddress, "B4Creation Software Solution Team");
                var ccAddress = new MailAddress(CcAddress, "B4Creation Software Solution Team");
                string fromPassword = FromPassword;
                string subject = "Activity Report [Dev Team]";
                var smtp = new SmtpClient
                {
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Credentials = new NetworkCredential(fromAddress.Address, fromPassword)

                };

                using (var messageToDo = new MailMessage(fromAddress, toAddress)
                {
                    Subject = subject,
                    Body = "Tasks [DONE] :" + "\n" + bodyDone
                })
                {
                    try
                    {
                        messageToDo.CC.Add(ccAddress);
                        smtp.Send(messageToDo);
                        log.Debug("Successfully send email");
                    }
                    catch (Exception ex)
                    {
                        //MessageBox.Show(ex.Message);
                        log.Debug(ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {

                log.Error(ex.Message);
            }
        }
        public void ActivityReports() //fonction principale : Consume JiraApi & Send Mail
        {

            //Consume JIRA API: status = IN PROGRESS
            string targetRequest = "'DCPM' AND status = 'IN PROGRESS'";
            string basicUrl = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(basicUrl + targetRequest);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequest.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequest.Headers.Add("Authorization", "Basic" + " " + authHeader);

            // IN PROGRESS
            string SalimBody = "";
            string HayetBody = "";
            string MarwaBody = "";
            string KhaledBody = "";
            string SadokBody = "";
            string SamiBody = "";
            string RamziBody = "";
          

            try
            {
                var httpWebResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_INPROG.txt", result);

                    JObject joResponse = JObject.Parse(result);
                    JToken issues = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_INPROG.txt", issues.ToString());
                    foreach (JToken item in issues)
                    {
                        JToken assignee = item["fields"]["assignee"]["displayName"];
                        JToken summary = item["fields"]["summary"];
                        switch (assignee.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summary.ToString()!= "")
                                {
                                    SalimBody = SalimBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                
                                break;
                            case "Hayet Manai":
                                if (summary.ToString() != "")
                                {
                                    HayetBody = HayetBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summary.ToString() != "")
                                {
                                   MarwaBody = MarwaBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summary.ToString() != "")
                                {
                                   KhaledBody = KhaledBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summary.ToString() != "")
                                {
                                    SadokBody = SadokBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summary.ToString() != "")
                                {
                                    SamiBody = SamiBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summary.ToString() != "")
                                {
                                    RamziBody = RamziBody + "\n" + assignee + ":" + summary.ToString();
                                }
                                break;
                                
                        }
                    }
                    string body = SamiBody + "\n" + SadokBody + "\n" + KhaledBody + "\n" + MarwaBody + "\n" + SalimBody + "\n" + HayetBody + "\n" + RamziBody;
                    SendMail(body);
                }

            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);
            }
        }
        public void ActivityReportsToDo()
        {
            // Consume JIRA API: Status = TO DO
            string targetRequestToDo = "'DCPM' AND status = 'TO DO'";
            string basicUrlToDo = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestToDo = (HttpWebRequest)WebRequest.Create(basicUrlToDo + targetRequestToDo);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestToDo.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestToDo.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //TO DO
            string SalimBodyToDo = "";
            string HayetBodyToDo = "";
            string MarwaBodyToDo = "";
            string KhaledBodyToDo = "";
            string SadokBodyToDo = "";
            string SamiBodyToDo = "";
            string RamziBodyToDo = "";
            try
            {
                var httpWebResponseToDo = (HttpWebResponse)httpWebRequestToDo.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponseToDo.GetResponseStream()))
                {
                    var resultToDo = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_TODO.txt", resultToDo);
                    JObject joResponse = JObject.Parse(resultToDo);
                    JToken issuesToDo = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_TODO.txt", issuesToDo.ToString());
                    foreach (JToken item in issuesToDo)
                    {
                        JToken assigneeToDo = item["fields"]["assignee"]["displayName"];
                        JToken summaryToDo = item["fields"]["summary"];
                        switch (assigneeToDo.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryToDo.ToString() != "")
                                {
                                    SalimBodyToDo = SalimBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryToDo.ToString() != "")
                                {
                                    HayetBodyToDo = HayetBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryToDo.ToString() != "")
                                {
                                    MarwaBodyToDo = MarwaBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryToDo.ToString() != "")
                                {
                                    KhaledBodyToDo = KhaledBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryToDo.ToString() != "")
                                {
                                    SadokBodyToDo = SadokBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryToDo.ToString() != "")
                                {
                                    SamiBodyToDo = SamiBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryToDo.ToString() != "")
                                {
                                    RamziBodyToDo = RamziBodyToDo + "\n" + assigneeToDo + ":" + summaryToDo.ToString();
                                }
                                break;
                        }
                    }
                    string bodyToDo = SamiBodyToDo + "\n" + SadokBodyToDo + "\n" + KhaledBodyToDo + "\n" + MarwaBodyToDo + "\n" + SalimBodyToDo + "\n" + HayetBodyToDo + "\n" + RamziBodyToDo;
                    SendMailToDo(bodyToDo);
                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        public void ActivityReportsWaitingForSupport()
        {
            // Consume JIRA API: Status = WAITING FOR SUPPORT
            string targetRequestWaitingForSupport = "'DCPM' AND status = 'WAITING FOR SUPPORT'";
            string basicUrlWaitingForSupport = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestWaitingForSupport = (HttpWebRequest)WebRequest.Create(basicUrlWaitingForSupport + targetRequestWaitingForSupport);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestWaitingForSupport.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestWaitingForSupport.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //WAITING FOR SUPPORT
            string SalimBodyWaitSupp = "";
            string HayetBodyWaitSupp = "";
            string MarwaBodyWaitSupp = "";
            string KhaledBodyWaitSupp = "";
            string SadokBodyWaitSupp = "";
            string SamiBodyWaitSupp = "";
            string RamziBodyWaitSupp = "";

            try
            {
                var httpWebResponseWaitingForSupport = (HttpWebResponse)httpWebRequestWaitingForSupport.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponseWaitingForSupport.GetResponseStream()))
                {
                    var resultWaitSupp = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_WAITSUPP.txt", resultWaitSupp);
                    JObject joResponse = JObject.Parse(resultWaitSupp);
                    JToken issuesWaitSupp = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_TODO.txt", issuesWaitSupp.ToString());
                    foreach (JToken item in issuesWaitSupp)
                    {
                        JToken assigneeWaitSupp = item["fields"]["assignee"]["displayName"];
                        JToken summaryWaitSupp = item["fields"]["summary"];
                        switch (assigneeWaitSupp.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    SalimBodyWaitSupp = SalimBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    HayetBodyWaitSupp = HayetBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    MarwaBodyWaitSupp = MarwaBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    KhaledBodyWaitSupp = KhaledBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    SadokBodyWaitSupp = SadokBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    SamiBodyWaitSupp = SamiBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryWaitSupp.ToString() != "")
                                {
                                    RamziBodyWaitSupp = RamziBodyWaitSupp + "\n" + assigneeWaitSupp + ":" + summaryWaitSupp.ToString();
                                }
                                break;
                        }
                    }
                    string bodyWaitSupp = SamiBodyWaitSupp + "\n" + SadokBodyWaitSupp + "\n" + KhaledBodyWaitSupp + "\n" + MarwaBodyWaitSupp + "\n" + SalimBodyWaitSupp + "\n" + HayetBodyWaitSupp + "\n" + RamziBodyWaitSupp;
                    SendMailWaitSupp(bodyWaitSupp);
                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        public void ActivityReportsWaitingForCustomer()
        {
            // Consume JIRA API: Status = WAITING FOR CUSTOMER
            string targetRequestWaitingForCustomer = "'DCPM' AND status = 'WAITING FOR CUSTOMER'";
            string basicUrlWaitingForCustomer = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestWaitingForCustomer = (HttpWebRequest)WebRequest.Create(basicUrlWaitingForCustomer + targetRequestWaitingForCustomer);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestWaitingForCustomer.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestWaitingForCustomer.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //WAITING FOR CUSTOMER
            string SalimBodyWaitCust = "";
            string HayetBodyWaitCust = "";
            string MarwaBodyWaitCust = "";
            string KhaledBodyWaitCust = "";
            string SadokBodyWaitCust = "";
            string SamiBodyWaitCust = "";
            string RamziBodyWaitCust = "";

            try
            {
                var httpWebResponseWaitingForCustomer = (HttpWebResponse)httpWebRequestWaitingForCustomer.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponseWaitingForCustomer.GetResponseStream()))
                {
                    var resultWaitCust = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_WAITCust.txt", resultWaitCust);
                    JObject joResponse = JObject.Parse(resultWaitCust);
                    JToken issuesWaitCust = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_TODO.txt", issuesWaitCust.ToString());
                    foreach (JToken item in issuesWaitCust)
                    {
                        JToken assigneeWaitCust = item["fields"]["assignee"]["displayName"];
                        JToken summaryWaitCust = item["fields"]["summary"];
                        switch (assigneeWaitCust.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    SalimBodyWaitCust = SalimBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    HayetBodyWaitCust = HayetBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    MarwaBodyWaitCust = MarwaBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    KhaledBodyWaitCust = KhaledBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    SadokBodyWaitCust = SadokBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    SamiBodyWaitCust = SamiBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryWaitCust.ToString() != "")
                                {
                                    RamziBodyWaitCust = RamziBodyWaitCust + "\n" + assigneeWaitCust + ":" + summaryWaitCust.ToString();
                                }
                                break;
                        }
                    }
                    string bodyWaitCust = SamiBodyWaitCust + "\n" + SadokBodyWaitCust + "\n" + KhaledBodyWaitCust + "\n" + MarwaBodyWaitCust + "\n" + SalimBodyWaitCust + "\n" + HayetBodyWaitCust + "\n" + RamziBodyWaitCust;
                    SendMailWaitCust(bodyWaitCust);
                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        public void ActivityReportsWaitingForApproval()
        {
            // Consume JIRA API: Status = WAITING FOR APPROVAL
            string targetRequestWaitingForApproval = "'DCPM' AND status = 'WAITING FOR APPROVAL'";
            string basicUrlWaitingForApproval = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestWaitingForApproval = (HttpWebRequest)WebRequest.Create(basicUrlWaitingForApproval + targetRequestWaitingForApproval);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestWaitingForApproval.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestWaitingForApproval.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //WAITING FOR CUSTOMER
            string SalimBodyWaitApp = "";
            string HayetBodyWaitApp = "";
            string MarwaBodyWaitApp = "";
            string KhaledBodyWaitApp = "";
            string SadokBodyWaitApp = "";
            string SamiBodyWaitApp = "";
            string RamziBodyWaitApp = "";

            try
            {
                var httpWebResponseWaitingForApproval = (HttpWebResponse)httpWebRequestWaitingForApproval.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponseWaitingForApproval.GetResponseStream()))
                {
                    var resultWaitApp = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_WAITApp.txt", resultWaitApp);
                    JObject joResponse = JObject.Parse(resultWaitApp);
                    JToken issuesWaitApp = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_WAITApp.txt", issuesWaitApp.ToString());
                    foreach (JToken item in issuesWaitApp)
                    {
                        JToken assigneeWaitApp = item["fields"]["assignee"]["displayName"];
                        JToken summaryWaitApp = item["fields"]["summary"];
                        switch (assigneeWaitApp.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    SalimBodyWaitApp = SalimBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    HayetBodyWaitApp = HayetBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    MarwaBodyWaitApp = MarwaBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    KhaledBodyWaitApp = KhaledBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    SadokBodyWaitApp = SadokBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    SamiBodyWaitApp = SamiBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryWaitApp.ToString() != "")
                                {
                                    RamziBodyWaitApp = RamziBodyWaitApp + "\n" + assigneeWaitApp + ":" + summaryWaitApp.ToString();
                                }
                                break;
                        }
                    }
                    string bodyWaitApp = SamiBodyWaitApp + "\n" + SadokBodyWaitApp + "\n" + KhaledBodyWaitApp + "\n" + MarwaBodyWaitApp + "\n" + SalimBodyWaitApp + "\n" + HayetBodyWaitApp + "\n" + RamziBodyWaitApp;
                    SendMailWaitApp(bodyWaitApp);
                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        public void ActivityReportsPending()
        {
            // Consume JIRA API: Status = PENDING
            string targetRequestPending = "'DCPM' AND status = 'PENDING'";
            string basicUrlPending = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestPending = (HttpWebRequest)WebRequest.Create(basicUrlPending + targetRequestPending);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestPending.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestPending.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //PENDING
            string SalimBodyPending = "";
            string HayetBodyPending = "";
            string MarwaBodyPending = "";
            string KhaledBodyPending = "";
            string SadokBodyPending = "";
            string SamiBodyPending = "";
            string RamziBodyPending = "";

            try
            {
                var httpWebResponsePending = (HttpWebResponse)httpWebRequestPending.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponsePending.GetResponseStream()))
                {
                    var resultPending = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_Pending.txt", resultPending);
                    JObject joResponse = JObject.Parse(resultPending);
                    JToken issuesPending = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_Pending.txt", issuesPending.ToString());
                    foreach (JToken item in issuesPending)
                    {
                        JToken assigneePending = item["fields"]["assignee"]["displayName"];
                        JToken summaryPending = item["fields"]["summary"];
                        switch (assigneePending.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryPending.ToString() != "")
                                {
                                    SalimBodyPending = SalimBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryPending.ToString() != "")
                                {
                                    HayetBodyPending = HayetBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryPending.ToString() != "")
                                {
                                    MarwaBodyPending = MarwaBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryPending.ToString() != "")
                                {
                                    KhaledBodyPending = KhaledBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryPending.ToString() != "")
                                {
                                    SadokBodyPending = SadokBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryPending.ToString() != "")
                                {
                                    SamiBodyPending = SamiBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryPending.ToString() != "")
                                {
                                    RamziBodyPending = RamziBodyPending + "\n" + assigneePending + ":" + summaryPending.ToString();
                                }
                                break;
                        }
                    }
                    string bodyPending = SamiBodyPending + "\n" + SadokBodyPending + "\n" + KhaledBodyPending + "\n" + MarwaBodyPending + "\n" + SalimBodyPending + "\n" + HayetBodyPending + "\n" + RamziBodyPending;
                    SendMailPending(bodyPending);
                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        public void ActivityReportsCanceled()
        {
            // Consume JIRA API: Status = CANCELED
            string targetRequestCanceled = "'DCPM' AND status = 'CANCELED'";
            string basicUrlCanceled = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestCanceled = (HttpWebRequest)WebRequest.Create(basicUrlCanceled + targetRequestCanceled);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestCanceled.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestCanceled.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //CANCELED
            string SalimBodyCanceled = "";
            string HayetBodyCanceled = "";
            string MarwaBodyCanceled = "";
            string KhaledBodyCanceled = "";
            string SadokBodyCanceled = "";
            string SamiBodyCanceled = "";
            string RamziBodyCanceled = "";

            try
            {
                var httpWebResponseCanceled = (HttpWebResponse)httpWebRequestCanceled.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponseCanceled.GetResponseStream()))
                {
                    var resultCanceled = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_Canceled.txt", resultCanceled);
                    JObject joResponse = JObject.Parse(resultCanceled);
                    JToken issuesCanceled = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_Pending.txt", issuesCanceled.ToString());
                    foreach (JToken item in issuesCanceled)
                    {
                        JToken assigneeCanceled = item["fields"]["assignee"]["displayName"];
                        JToken summaryCanceled = item["fields"]["summary"];
                        switch (assigneeCanceled.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryCanceled.ToString() != "")
                                {
                                    SalimBodyCanceled = SalimBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryCanceled.ToString() != "")
                                {
                                    HayetBodyCanceled = HayetBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryCanceled.ToString() != "")
                                {
                                    MarwaBodyCanceled = MarwaBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryCanceled.ToString() != "")
                                {
                                    KhaledBodyCanceled = KhaledBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryCanceled.ToString() != "")
                                {
                                    SadokBodyCanceled = SadokBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryCanceled.ToString() != "")
                                {
                                    SamiBodyCanceled = SamiBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryCanceled.ToString() != "")
                                {
                                    RamziBodyCanceled = RamziBodyCanceled + "\n" + assigneeCanceled + ":" + summaryCanceled.ToString();
                                }
                                break;
                        }
                    }
                    string bodyCanceled = SamiBodyCanceled + "\n" + SadokBodyCanceled + "\n" + KhaledBodyCanceled + "\n" + MarwaBodyCanceled + "\n" + SalimBodyCanceled + "\n" + HayetBodyCanceled + "\n" + RamziBodyCanceled;
                    SendMailCanceled(bodyCanceled);

                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        public void ActivityReportsDone()
        {
            // Consume JIRA API: Status = DONE
            string targetRequestDone = "'DCPM' AND status = 'DONE'";
            string basicUrlDone = "https://b4csoftwaresolutions.atlassian.net/rest/api/2/search?jql=project=";
            var httpWebRequestDone = (HttpWebRequest)WebRequest.Create(basicUrlDone + targetRequestDone);
            // Auth JIRA Token:
            string userName = "salim.khefifi@b4creation.org";
            string userPassword = "XbYtWD42WXInIeNMXdvID37E";
            httpWebRequestDone.Method = "GET";
            string authHeader = System.Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(userName + ":" + userPassword));
            httpWebRequestDone.Headers.Add("Authorization", "Basic" + " " + authHeader);

            //DONE
            string SalimBodyDone = "";
            string HayetBodyDone = "";
            string MarwaBodyDone = "";
            string KhaledBodyDone = "";
            string SadokBodyDone = "";
            string SamiBodyDone = "";
            string RamziBodyDone = "";
            string MohamedBodyDone = "";

            try
            {
                var httpWebResponseDone = (HttpWebResponse)httpWebRequestDone.GetResponse();
                using (var streamReader = new StreamReader(httpWebResponseDone.GetResponseStream()))
                {
                    var resultDone = streamReader.ReadToEnd();
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\test_DCPM_Done.txt", resultDone);
                    JObject joResponse = JObject.Parse(resultDone);
                    JToken issuesDone = joResponse["issues"];
                    File.WriteAllText(@"C:\SALIM\B4CES_Work\B4CFondationReportLogger\testissues_DCPM_Pending.txt", issuesDone.ToString());
                    foreach (JToken item in issuesDone)
                    {
                        JToken assigneeDone =string.Empty;
                        try
                        {
                             assigneeDone = item["fields"]["assignee"]["displayName"];
                        }
                        catch (Exception ex)
                        {

                            log.Error("Unassigned Tasks " + ex.Message);
                        }
                        JToken summaryDone = item["fields"]["summary"];
                        switch (assigneeDone.ToString())
                        {
                            case "salim KHEFIFI":
                                if (summaryDone.ToString() != "")
                                {
                                    SalimBodyDone = SalimBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "Hayet Manai":
                                if (summaryDone.ToString() != "")
                                {
                                    HayetBodyDone = HayetBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "Marwa Ayari":
                                if (summaryDone.ToString() != "")
                                {
                                    MarwaBodyDone = MarwaBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "khaled boudagga":
                                if (summaryDone.ToString() != "")
                                {
                                    KhaledBodyDone = KhaledBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "Sadok AGILI":
                                if (summaryDone.ToString() != "")
                                {
                                    SadokBodyDone = SadokBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "Sami KACHAI":
                                if (summaryDone.ToString() != "")
                                {
                                    SamiBodyDone = SamiBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "Ramzi Kehili":
                                if (summaryDone.ToString() != "")
                                {
                                    RamziBodyDone = RamziBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                            case "Mohamed Amine Souissi":
                                if(summaryDone.ToString() != "")
                                {
                                    MohamedBodyDone = MohamedBodyDone + "\n" + assigneeDone + ":" + summaryDone.ToString();
                                }
                                break;
                        }
                    }
                    string bodyDone = SamiBodyDone + "\n" + SadokBodyDone + "\n" + KhaledBodyDone + "\n" + MarwaBodyDone + "\n" + SalimBodyDone + "\n" + HayetBodyDone + "\n" + RamziBodyDone + "\n" + MohamedBodyDone;
                    SendMailDone(bodyDone);

                }
            }
            catch (Exception ex)
            {
                log.Debug("Failed to fetch data:" + ex.Message);

            }
        }
        protected override void OnStop()
        {
            log.Info("service stopped");
        }
    }
}
