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

            public PartyProfile PartyProfile
            {
                get
                {
                    return PartyProfileNames2PartyProfile[PartyProfileName];
                }
            }
            public BuyerProfile BuyerProfile
            {
                get
                {
                    return BuyerProfileNames2BuyerProfile[BuyerProfileName];
                }
            }
            public BrokerProfile BrokerProfile
            {
                get
                {
                    return BrokerProfileNames2BrokerProfile[BrokerProfileName];
                }
            }
            public AgentProfile AgentProfile
            {
                get
                {
                    return AgentProfileNames2AgentProfile[AgentProfileName];
                }
            }
            public EscrowProfile EscrowProfile
            {
                get
                {
                    return EscrowProfileNames2EscrowProfile[EscrowProfileName];
                }
            }
            public EmailTemplateProfile EmailTemplateProfile
            {
                get
                {
                    return EmailTemplateProfileNames2EmailTemplateProfileProfile[EmailTemplateProfileName];
                }
            }

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

            Cliver.BotGui.Program.BindProgressBar2InputItemQueue<PdfItem>();
            Session.GetInputItemQueue<PdfItem>().PickNext = pick_next_PdfItem;
        }

       static InputItem pick_next_PdfItem()
        {
            lock (emails2sent_time)
            {
                int delay = Program.Settings.MinRandomDelayMss + (int)((float)(Program.Settings.MaxRandomDelayMss - Program.Settings.MinRandomDelayMss) * random.NextDouble());
                TimeSpan min_wait_period = TimeSpan.MaxValue;
                PdfItem min_wait_period_pi = null;
                System.Collections.IEnumerator ie = Session.GetInputItemQueue<PdfItem>().GetEnnumerator();
                for (ie.Reset(); ie.MoveNext(); )
                {
                    PdfItem pi = ((PdfItem)ie.Current);
                    string to_email = pi.__ParentItem.ListAgentEmail;
                    DateTime sent_time;
                    if (!emails2sent_time.TryGetValue(to_email, out sent_time))
                    {
                        emails2sent_time[to_email] = DateTime.Now;
                        return pi;
                    }
                    TimeSpan wait_period = sent_time.AddMilliseconds(delay) - DateTime.Now;
                    if (wait_period < TimeSpan.FromSeconds(0))
                    {
                        emails2sent_time[to_email] = DateTime.Now;
                        return pi;
                    }
                    if (wait_period > min_wait_period)
                    {
                        min_wait_period = wait_period;
                        min_wait_period_pi = pi;
                    }
                }
                if (min_wait_period_pi != null && Session.GetInputItemQueue<DataItem>().CountOfNew < 1)
                {
                    Thread.Sleep(min_wait_period);
                    emails2sent_time[min_wait_period_pi.__ParentItem.ListAgentEmail] = DateTime.Now;
                    return min_wait_period_pi;
                }
                return null;
            }
        }

        new static public void SessionClosing()
        {
            //while (send_pdf()) ;
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
                string output_pdf = d + "\\" + PathRoutines.GetFileNameFromPath(template_pdf);
                //lock (template_pdf)
                //{
                //    File.Copy(template_pdf, pdf);
                //}

                PdfReader.unethicalreading = true;
                PdfReader pr = new PdfReader(template_pdf);
                //pr.RemoveUsageRights();
                //pr.SelectPages("7,8");
                PdfStamper ps = new PdfStamper(pr, new FileStream(output_pdf, FileMode.Create, FileAccess.Write, FileShare.None));
                
                //string fs = "";
                //foreach (KeyValuePair<string, AcroFields.Item> kvp in ps.AcroFields.Fields)
                //    fs += "\n{\"" + kvp.Key + "\", \"\"},";
                
                set_field(ps.AcroFields, "Todays Date", DateTime.Today.ToShortDateString() );
                set_field(ps.AcroFields, "Buyer Name", Program.Settings.BuyerProfile.Name );
                set_field(ps.AcroFields, "Address and Unit Number", Address + " " + UnitNumber);
                set_field(ps.AcroFields, "City/Town", City);
                //set_field(ps.AcroFields, "CLARK", );
                set_field(ps.AcroFields, "Zip", ZipCode);
                set_field(ps.AcroFields, "PARCEL NUMBER", ParcelNumber);
                set_field(ps.AcroFields, "OfferAmt", OfferAmt);
                //set_field(ps.AcroFields, "OfferAmt in words", );
                set_field(ps.AcroFields, "EMD", Program.Settings.Emd);
                //set_field(ps.AcroFields, "Check Box1", );
                //set_field(ps.AcroFields, "Balance", );
                set_field(ps.AcroFields, "Co Buyer Name", Program.Settings.BuyerProfile.CoBuyerName);
                //set_field(ps.AcroFields, "<address> <UnitNumber> <City/town> NV <ZIP Code>", );
                set_field(ps.AcroFields, "ML#", ML_Id);
                set_field(ps.AcroFields, "Title Company", Program.Settings.EscrowProfile.TitleCompany);
                set_field(ps.AcroFields, "Escrow Officer", Program.Settings.EscrowProfile.Officer);
                set_field(ps.AcroFields, "Close of Escrow", Program.Settings.CloseOfEscrow.ToShortDateString());
                set_field(ps.AcroFields, "ADDITIONAL TERMS", AdditionalTerms);
                set_field(ps.AcroFields, "Additional  terms", AdditionalTerms);
                set_field(ps.AcroFields, "Buyer Broker", Program.Settings.BrokerProfile.Name);
                set_field(ps.AcroFields, "Agent Name", Program.Settings.AgentProfile.Name);
                set_field(ps.AcroFields, "Company Name", Program.Settings.BrokerProfile.Company);
                set_field(ps.AcroFields, "Agents License", Program.Settings.AgentProfile.LicenseNo);
                set_field(ps.AcroFields, "Brokers License", Program.Settings.BrokerProfile.LicenseNo);
                set_field(ps.AcroFields, "Office Address", Program.Settings.BrokerProfile.Address);
                set_field(ps.AcroFields, "Office Phone", Program.Settings.BrokerProfile.Phone);
                set_field(ps.AcroFields, "City State Zip", Program.Settings.BrokerProfile.Zip);
                set_field(ps.AcroFields, "Agent Email", Program.Settings.AgentProfile.Email);
                //set_field(ps.AcroFields, "Response Month", );
                //set_field(ps.AcroFields, "Response day", );
                //set_field(ps.AcroFields, "year", );
                //set_field(ps.AcroFields,  "Date_2", );
                //set_field(ps.AcroFields, "Date_3", );
                //set_field(ps.AcroFields, "Sellers Broker", );
                set_field(ps.AcroFields, "Agents Name_2", Program.Settings.AgentProfile.Name);
                set_field(ps.AcroFields, "Company Name_2", Program.Settings.BrokerProfile.Company);
                set_field(ps.AcroFields, "Agents License Number_2", Program.Settings.AgentProfile.LicenseNo);
                set_field(ps.AcroFields, "Brokers License Number_2", Program.Settings.BrokerProfile.LicenseNo);
                set_field(ps.AcroFields, "Office Address_2", Program.Settings.BrokerProfile.Address);
                set_field(ps.AcroFields, "Phone_2", Program.Settings.BrokerProfile.Phone);
                set_field(ps.AcroFields, "City State Zip_2", Program.Settings.BrokerProfile.Zip);
                set_field(ps.AcroFields, "Email_2", Program.Settings.AgentProfile.Email);
                
                ps.FormFlattening = true;

                //var pcb = ps.GetOverContent(1);
                //add_image(pcb, employee_signature, new System.Drawing.Point(140, 213));
                //add_image(pcb, preparer_signature, new System.Drawing.Point(180, 120));
                //pcb = ps.GetOverContent(2);
                //add_image(pcb, employer_signature, new System.Drawing.Point(60, 256));
                //add_image(pcb, employer_signature, new System.Drawing.Point(65, 30));

                ps.Close();
                pr.Close();
                               
                bc.Add(new PdfItem(output_pdf));
                //CustomBot.send_pdf(output_pdf, ListAgentEmail);
            }
            static readonly string template_pdf = Log.GetAppCommonDataDir() + "\\RPA template - Acrobat forms.pdf";

        }

        static void set_field(AcroFields form, string field_key, string value)
        {
            switch (form.GetFieldType(field_key))
            {
                case AcroFields.FIELD_TYPE_CHECKBOX:
                case AcroFields.FIELD_TYPE_RADIOBUTTON:
                //bool v;
                //if (bool.TryParse(value, out v))
                //    value = !v ? "false" : "true";
                //else
                //{
                //    int i;
                //    if (int.TryParse(value, out i))
                //        value = i == 0 ? "false" : "true";
                //    else
                //        value = string.IsNullOrEmpty(value) ? "false" : "true";
                //}
                //form.SetField(field_key, value);
                //break;
                case AcroFields.FIELD_TYPE_COMBO:
                case AcroFields.FIELD_TYPE_LIST:
                case AcroFields.FIELD_TYPE_NONE:
                case AcroFields.FIELD_TYPE_PUSHBUTTON:
                case AcroFields.FIELD_TYPE_SIGNATURE:
                case AcroFields.FIELD_TYPE_TEXT:
                    form.SetField(field_key, value);
                    break;
                default:
                    throw new Exception("Unknown option: " + form.GetFieldType(field_key));
            }
        }

        static void add_image(PdfContentByte pcb, System.Drawing.Image image, System.Drawing.Point point)
        {
            image = ImageRoutines.GetCroppedByColor(image, System.Drawing.Color.Transparent);
            Image i = Image.GetInstance(image, (BaseColor)null);
            var ratio = Math.Min((float)image_max_size.Width / image.Width, (float)image_max_size.Height / image.Height);
            i.ScalePercent(ratio * 100);
            i.SetAbsolutePosition(point.X, point.Y);
            pcb.AddImage(i);
        }
        static System.Drawing.Size image_max_size = new System.Drawing.Size(200, 50);
        
        class PdfItem : InputItem
        {
            public readonly DataItem __ParentItem;
            public readonly string Pdf;

            internal PdfItem(string pdf)
            {
                Pdf = pdf;
            }

            override public void PROCESSOR(BotCycle bc)
            {
                CustomBot cb = (CustomBot)bc.Bot;

                cb.send(Pdf, __ParentItem.ListAgentEmail);
            }
        }

        //static bool send_pdf(string pdf = null, string email = null)
        //{
        //    lock (pdfs2email)
        //    {
        //        if (pdf != null)
        //            pdfs2email[pdf] = email;
        //        pdf = get_next_pdf(pdf == null);
        //        if (pdf != null)
        //            send(pdf, pdfs2email[pdf]);
        //        return pdf != null;
        //    }
        //}

        //static string get_next_pdf(bool blocking_wait)
        //{
        //    lock (pdfs2email)
        //    {
        //        TimeSpan min_wait_period = TimeSpan.MaxValue;
        //        string min_wait_period_pdf = null;
        //        foreach (string pdf in pdfs2email.Keys)
        //        {
        //            if (pdf == null)
        //                continue;
        //            int delay = Program.Settings.MinRandomDelayMss + (int)((float)(Program.Settings.MaxRandomDelayMss - Program.Settings.MinRandomDelayMss) * random.NextDouble());
        //            DateTime sent_time;
        //            if (!emails2sent_time.TryGetValue(pdfs2email[pdf], out sent_time))
        //            {
        //                emails2sent_time[pdfs2email[pdf]] = DateTime.Now;
        //                pdfs2email[pdf] = null;
        //                return pdf;
        //            }
        //            TimeSpan wait_period = sent_time.AddMilliseconds(delay) - DateTime.Now;
        //            if (wait_period < TimeSpan.FromSeconds(0))
        //            {
        //                emails2sent_time[pdfs2email[pdf]] = DateTime.Now;
        //                pdfs2email[pdf] = null;
        //                return pdf;
        //            }
        //            else if (blocking_wait && wait_period > min_wait_period)
        //            {
        //                min_wait_period = wait_period;
        //                min_wait_period_pdf = pdf;
        //            }
        //        }
        //        if (blocking_wait && min_wait_period_pdf != null)
        //        {
        //            Thread.Sleep(min_wait_period);
        //            emails2sent_time[pdfs2email[min_wait_period_pdf]] = DateTime.Now;
        //            pdfs2email[min_wait_period_pdf] = null;
        //            return min_wait_period_pdf;
        //        }
        //        return null;
        //    }
        //}
        static Random random = new Random();
        static Dictionary<string, DateTime> emails2sent_time = new Dictionary<string, DateTime>();
        //static SingleValueWorkItemDictionary<SingleValueWorkItem<string>, string> pdfs2email = Session.GetSingleValueWorkItemDictionary<SingleValueWorkItem<string>, string>();

        void send(string pdf, string to_email)
        {
            to_email = "sergey.stoyan@gmail.com";
            MailMessage mm = new MailMessage(
                Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail,
                to_email
                )
            {
                Subject = Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile[Program.Settings.EmailTemplateProfileName].Subject,
                Body = Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile[Program.Settings.EmailTemplateProfileName].Body,
                From = new MailAddress(Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail)
            };
            foreach (int i in Program.Settings.SelectedAttachmentIds)
            {
                mm.Attachments.Add(new Attachment(pdf));
                mm.Attachments.Add(new Attachment(Program.Settings.AttachmentFiles[i]));
            }
            Log.Write("Emailing to " + mm.To + ": " + mm.Subject);
            try
            {
                smtp_client.Send(mm);
            }
            catch (Exception e)
            {
                Log.Write("Host: " + smtp_client.Host
                    + "\r\nPort: " + smtp_client.Port
                    + "\r\nEnableSsl: " + smtp_client.EnableSsl
                    + "\r\nDeliveryMethod: " + smtp_client.DeliveryMethod
                    + "\r\nUserName: " + Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SenderEmail
                    + "\r\nSmtpPassword: " + Program.Settings.EmailServerProfileNames2EmailServerProfile[Program.Settings.EmailServerProfileName].SmtpPassword
                    + "\r\nFrom: " + mm.From.Address
                    + "\r\nTo: " + mm.To
                    );
                throw new Session.FatalException("SMTP error.", e);
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