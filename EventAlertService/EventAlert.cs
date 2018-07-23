using System;
using System.Collections;
using System.Data;
using System.ServiceProcess;
using System.Timers;//For Timers
using System.Net.Mail;//For Email
using System.Configuration;
using System.Data.SqlClient;

namespace EventAlertService
{
    partial class EventAlert : ServiceBase
    {   
        //Timer
        Timer timer1 = new Timer();
        //Interval
        String interval = ConfigurationManager.AppSettings["Interval"];
        //Remind Before
        int reminderTime = Convert.ToInt32(ConfigurationManager.AppSettings["Reminder"]);
        //Connection String
        String connectionString = ConfigurationManager.ConnectionStrings["EventAlertService.Properties.Settings.IntellimediaConnectionString"].ConnectionString;
        public EventAlert()
        {
            InitializeComponent();

            
            //To check if source Exist
            if (!System.Diagnostics.EventLog.SourceExists("EmailServiceCheck"))
            {
                //Create a EventLog
                System.Diagnostics.EventLog.CreateEventSource("EmailServiceCheck", "Chaitanya");
            }
            EmailServiceCheck.Source = "EmailServiceCheck";
            EmailServiceCheck.Log = "Chaitanya";

            timer1.Interval = Convert.ToInt32(interval);
            timer1.Elapsed += timer1_Elapsed;
        }

       void timer1_Elapsed(object sender, ElapsedEventArgs e)
        {
            SearchForEvents();
        }

        private void SearchForEvents()
        {
            using(SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                String SQLQuery = "SELECT EventID FROM Events WHERE DATEDIFF(mi,GETDATE(),EventTime)<=" + reminderTime + "AND DATEDIFF(mi,GETDATE(),EventTime)>0";
                using(SqlCommand command = new SqlCommand(SQLQuery,connection))
                {
                    using(SqlDataReader Reader1 = command.ExecuteReader())
                    {
                        while(Reader1.Read())
                        {
                            String EventID = Reader1["EventID"].ToString();
                            EmailServiceCheck.WriteEntry(EventID);
                            GetTheSubscribers(Convert.ToInt32(EventID));
                        }
                        connection.Close();
                        Reader1.Close();
                    }
                }
            }
        }

        private void GetTheSubscribers(int EventID)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                String SQLQuery = "SELECT * FROM EventSubscribers WHERE EventID = " + EventID;
                using (SqlCommand command = new SqlCommand(SQLQuery, connection)) 
                {
                    using (SqlDataReader Reader = command.ExecuteReader())
                    {
                        while (Reader.Read())
                        {
                            String EmailID = Reader["EmailID"].ToString();
                            EmailServiceCheck.WriteEntry(EmailID);
                            SendEmail(EmailID,EventID);
                        }
                        connection.Close();
                        Reader.Close();
                    }
                }
            }
        }

        private void SendEmail(string EmailID,int EventID)
        {
            SmtpClient smtpClient = new SmtpClient();
            String EmailFromString = ConfigurationManager.AppSettings["EmailFrom"];
            String Password = ConfigurationManager.AppSettings["Password"];
            String smtpHost = ConfigurationManager.AppSettings["Host"];
            String smtpPort = ConfigurationManager.AppSettings["Port"];
            String EnableSs1 = ConfigurationManager.AppSettings["EnableSs1"];
            String UseDefaultCredentials = ConfigurationManager.AppSettings["UseDefaultCredentials"];
            MailAddress EmailFrom = new MailAddress(EmailFromString);
            MailMessage mailMsg = new MailMessage();

            //Email Body
            String EventName = null;
            String EventTime = null;
            String SubscriberName = null;

            try
            {
                using(SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    String SQLQuery1 = "SELECT * FROM Events WHERE EventID = " + EventID;
                    using(SqlCommand command = new SqlCommand(SQLQuery1,connection))
                    {
                        using (SqlDataReader Reader1 = command.ExecuteReader())
                        {
                            while (Reader1.Read())
                            {
                                EventName = Reader1["EventName"].ToString();
                                EventTime = Reader1["EventTime"].ToString();
                            }
                        }
                    }

                    String SQLQuery2 = "SELECT * FROM EventSubscribers WHERE EmailID = '" + EmailID + "'";
                    using(SqlCommand command = new SqlCommand(SQLQuery2,connection))
                    {
                        using(SqlDataReader Reader2 = command.ExecuteReader())
                        {
                            while(Reader2.Read())
                            {
                                SubscriberName = Reader2["Name"].ToString();
                            }
                        }
                    }
                    connection.Close();
                }
                //EmailServiceCheck.WriteEntry(EventName + " " + EventTime + " " + SubscriberName);

                mailMsg.To.Add(EmailID);
                mailMsg.From = EmailFrom;
                mailMsg.Subject = EventName+ " Reminder";

                //MessageBody
                mailMsg.IsBodyHtml = true;
                mailMsg.Body = "Hi " + SubscriberName + ",<br/>Your event " + EventName + " is going to start at " + EventTime + ".<br/> Please have a look at the following URL.<br/>https://www.google.com<br/>Regards<br/>" + EventName + " Support Team";

                smtpClient.Host = smtpHost;
                smtpClient.Port = Convert.ToInt32(smtpPort);
                smtpClient.EnableSsl = Convert.ToBoolean(EnableSs1);
                smtpClient.UseDefaultCredentials = Convert.ToBoolean(UseDefaultCredentials);
                smtpClient.Credentials = new System.Net.NetworkCredential(EmailFromString, Password);

                smtpClient.Send(mailMsg);

                EmailServiceCheck.WriteEntry("Mail Send Successfully to "+ EmailID);
            }
            
            finally
            {
                mailMsg = null;
                EmailFrom = null;
            }
        }

        protected override void OnStart(string[] args)
        {
            EmailServiceCheck.WriteEntry("Service is Started");
            timer1.Start();
        }

        protected override void OnStop()
        {
            EmailServiceCheck.WriteEntry("Service is Stopped");
        }
    }
}