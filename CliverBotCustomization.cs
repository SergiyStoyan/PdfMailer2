//********************************************************************************************
//Author: Sergey Stoyan, CliverSoft.com
//        http://cliversoft.com
//        stoyan@cliversoft.com
//        sergey.stoyan@gmail.com
//        27 February 2007
//Copyright: (C) 2007, Sergey Stoyan
//********************************************************************************************

using System;
using System.Linq;
using System.Net;
using System.Text;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using System.Web;
using System.Data;
using System.Web.Script.Serialization;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Net.Mail;
using Cliver;
using System.Configuration;
using System.Windows.Forms;
using Cliver.Bot;
using Cliver.BotGui;
using Microsoft.Win32;
using System.Reflection;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Cliver.PdfMailer2
{
    public class Program
    {
        [STAThread]
        static void Main()
        {
            //Cliver.Bot.Program.Run();//It is the entry when the app runs as a console app.
            Cliver.BotGui.Program.Run();//It is the entry when the app uses the default GUI.

            SettingsForm sf = new SettingsForm();
            //Application.Run(sf);
        }
        internal static readonly SettingsClass Settings = Cliver.Serializable.Load<SettingsClass>("Settings.txt");

        public class SettingsClass : Serializable
        {
            public Dictionary<string, PartyProfile> PartyProfileNames2PartyProfile = new Dictionary<string, PartyProfile>();
            public Dictionary<string, BuyerProfile> BuyerProfileNames2BuyerProfile = new Dictionary<string, BuyerProfile>();
            public Dictionary<string, BrokerProfile> BrokerProfileNames2BrokerProfile = new Dictionary<string, BrokerProfile>();
            public Dictionary<string, AgentProfile> AgentProfileNames2AgentProfile = new Dictionary<string, AgentProfile>();
            public Dictionary<string, EscrowProfile> EscrowProfileNames2EscrowProfile = new Dictionary<string, EscrowProfile>();
            public Dictionary<string, EmailTemplateProfile> EmailTemplateProfileNames2EmailTemplateProfileProfile = new Dictionary<string, EmailTemplateProfile>();
            public Dictionary<string, EmailServerProfile> EmailServerProfileNames2EmailServerProfile = new Dictionary<string, EmailServerProfile>();
            public string[] AttachmentFiles = new string[0];

            public string PartyProfileName;
            public string BuyerProfileName;
            public string BrokerProfileName;
            public string AgentProfileName;
            public string EscrowProfileName;
            public string EmailTemplateProfileName;
            public string EmailServerProfileName;
            public int[] SelectedAttachmentIds = new int[0];

            public DateTime CloseOfEscrow = DateTime.Now;
            public string Emd;
            public bool ShortSaleAddendum;
            public bool OtherAddendum1;
            public bool OtherAddendum2;
            public bool UseRandomDelay;
            public int MinRandomDelayMss;
            public int MaxRandomDelayMss;
        }

        public abstract class Profile
        {
            public string _ProfileName;
        }

        public class EmailTemplateProfile : Profile
        {
            public string Subject;
            public string Body;
        }

        public class EmailServerProfile : Profile
        {
            public string SmtpHost;
            public int SmtpPort;
            public string SmtpPassword;
            public string SenderEmail;
        }

        public class PartyProfile : Profile
        {
            public string BuyerProfileName;
            public string BrokerProfileName;
            public string AgentProfileName;
            public string EscrowProfileName;
        }

        public class BuyerProfile : Profile
        {
            public string Name;
            public string InitialFile;
            public string SignatureFile;
            public bool UseCoBuyer;
            public string CoBuyerName;
            public string CoBuyerInitialFile;
            public string CoBuyerSignatureFile;
        }

        public class AgentProfile : Profile
        {
            public string Name;
            public string LicenseNo;
            public string Email;
            public string InitialFile;
            public string SignatureFile;
        }

        public class BrokerProfile : Profile
        {
            public string Name;
            public string LicenseNo;
            public string Company;
            public string Address;
            public string City;
            public string State;
            public string Zip;
            public string Phone;
        }

        public class EscrowProfile : Profile
        {
            public string TitleCompany;
            public string Officer;
        }
    }

    public class CustomBotGui : Cliver.BotGui.BotGui
    {
        override public string[] GetConfigControlNames()
        {
            return new string[] { "General", "Input", "Output", /*"Web", "Browser", "Spider", "Proxy",*/ "Log" };
        }

        override public Cliver.BaseForm GetToolsForm()
        {
            return new SettingsForm();
        }

        //override public Type GetBotThreadControlType()
        //{
        //    return typeof(IeRoutineBotThreadControl);
        //    //return typeof(WebRoutineBotThreadControl);
        //}
    }

    public class CustomBot : Cliver.Bot.Bot
    {
        new static public string GetAbout()
        {
            return @"PDF MAILER
Created: " + Cliver.Bot.Program.GetCustomizationCompiledTime().ToString() + @"
Developed by: www.cliversoft.com";
        }

        new static public void SessionCreating()
        {
            InternetDateTime.CHECK_TEST_PERIOD_VALIDITY(2016, 11, 25);
        }

        new static public void SessionClosing()
        {
        }

        override public void CycleStarting()
        {
        }

        public class DataItem : InputItem
        {
            readonly public string Status;
            readonly public string ML_Id;
            readonly public string Address;
            readonly public string UnitNumber;
            readonly public string City;
            readonly public string ZipCode;
            readonly public string ParcelNumber;
            readonly public string BldgDes;
            readonly public string SubdivisionName;
            readonly public string Sub;
            readonly public string ShortSale;
            readonly public string ApproxLivArea;
            readonly public string BedsTotal;
            readonly public string BathsTotal;
            readonly public string Garage;
            readonly public string PvPool;
            readonly public string Spa;
            readonly public string YearBuilt;
            readonly public string LotSqft;
            readonly public string CurrentPrice;
            readonly public string OfferAmt;
            readonly public string AdditionalTerms;
            readonly public string Agent2AgentRemarks;
            readonly public string ListAgentFullName;
            readonly public string ListAgentEmail;
            readonly public string ListOfficeName;
            readonly public string ActualCloseDate;
            readonly public string ListAgentDirectWorkPhone;
            readonly public string DOM;
            readonly public string OfferSentDate;

            override public void PROCESSOR(BotCycle bc)
            {
                CustomBot cb = (CustomBot)bc.Bot;

                string d = Log.OutputDir + "\\" + Log.This.Id + "_" + DateTime.Now.GetSecondsSinceUnixEpoch();
                Directory.CreateDirectory(d);
                string pdf = d + "\\" + PathRoutines.GetFileNameFromPath(template_pdf);
                lock (template_pdf)
                {
                    File.Copy(template_pdf, pdf);
                }
                
                //PdfReader.unethicalreading = true;
                //PdfReader pr;
                //pr = new PdfReader(pdf);
                
                //MemoryStream ms = new MemoryStream();
                //pr.RemoveUsageRights();
                ////pr.SelectPages("7,8");
                //PdfStamper ps = new PdfStamper(pr, ms);

                //ps.FormFlattening = true;

                //var pcb = ps.GetOverContent(1);
                //add_image(pcb, employee_signature, new System.Drawing.Point(140, 213));
                //add_image(pcb, preparer_signature, new System.Drawing.Point(180, 120));
                //pcb = ps.GetOverContent(2);
                //add_image(pcb, employer_signature, new System.Drawing.Point(60, 256));
                //add_image(pcb, employer_signature, new System.Drawing.Point(65, 30));
                //ps.Close();
                //pr.Close();

                bc.Add(new PdfItem(pdf));
            }
            static readonly string template_pdf = Log.GetAppCommonDataDir() + "\\RPA template.pdf";
        }

        public class PdfItem : InputItem
        {
            readonly public DataItem Data;
            readonly public string File;

            public PdfItem(string pdf_file)
            {
                File = pdf_file;
            }

            override public void PROCESSOR(BotCycle bc)
            {
                CustomBot cb = (CustomBot)bc.Bot;

                MailMessage mm = new MailMessage(
                    Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail,
                    Data.ListAgentEmail
                    )
                {
                    Subject = Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile[Program.Settings.EmailTemplateProfileName].Subject,
                    Body = Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile[Program.Settings.EmailTemplateProfileName].Body,
                    From = new MailAddress(Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail)
                };
                foreach (int i in Program.Settings.SelectedAttachmentIds)
                    mm.Attachments.Add(new Attachment(Program.Settings.AttachmentFiles[i]));
                Log.Write("Emailing to " + mm.To + ": " + mm.Subject);
                cb.email(mm);
            }
        }

        void email(MailMessage mm)
        {
            try
            {
                Log.Write("Emailing to " + mm.To + ": " + mm.Subject);
                smtp_client.Send(mm);
            }
            catch (Exception e)
            {
                Log.Error(e);
                Log.Write("Host: " + smtp_client.Host
                    + "\r\nPort: " + smtp_client.Port
                    + "\r\nEnableSsl: " + smtp_client.EnableSsl
                    + "\r\nDeliveryMethod: " + smtp_client.DeliveryMethod
                    + "\r\nUserName: " + Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail
                    + "\r\nSmtpPassword: " + Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SmtpPassword
                    + "\r\nFrom: " + mm.From.Address
                    + "\r\nTo: " + mm.To
                    );
            }
        }
        SmtpClient smtp_client = new SmtpClient
        {
            Host = Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SmtpHost,
            Port = Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SmtpPort,
            EnableSsl = true,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(
                Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail,
                Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SmtpPassword
                )
        };
    }
}