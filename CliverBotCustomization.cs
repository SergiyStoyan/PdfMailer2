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
            public EmailServerProfile EmailServerProfile
            {
                get
                {
                    return EmailServerProfileNames2EmailServerProfile[EmailServerProfileName];
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
            return new string[] { "General", "Input", /*"Output", "Web", "Browser", "Spider", "Proxy",*/ "Log" };
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

        new static public void FatalError(string message)
        {
        }

        new static public void SessionCreating()
        {
            InternetDateTime.CHECK_TEST_PERIOD_VALIDITY(2016, 11, 25);

            Cliver.BotGui.Program.BindProgressBar2InputItemQueue<EmailItem>();
            BotCycle.TreatExceptionAsFatal = true;
            Session.GetInputItemQueue<EmailItem>().PickNext = pick_next_PdfItem;
        }

        static InputItem pick_next_PdfItem(System.Collections.IEnumerator items_ennumerator)
        {
            lock (emails2sent_time)
            {
                int delay = Program.Settings.MinRandomDelayMss + (int)((float)(Program.Settings.MaxRandomDelayMss - Program.Settings.MinRandomDelayMss) * random.NextDouble());
                TimeSpan min_wait_period = TimeSpan.MaxValue;
                EmailItem min_wait_period_ei = null;
                for (items_ennumerator.Reset(); items_ennumerator.MoveNext();)
                {
                    EmailItem ei = ((EmailItem)items_ennumerator.Current);
                    string to_email = ei.Parent.ListAgentEmail;
                    DateTime sent_time;
                    if (!emails2sent_time.TryGetValue(to_email, out sent_time))
                    {
                        emails2sent_time[to_email] = DateTime.Now;
                        return ei;
                    }
                    TimeSpan wait_period = sent_time.AddMilliseconds(delay) - DateTime.Now;
                    if (wait_period < TimeSpan.FromSeconds(0))
                    {
                        emails2sent_time[to_email] = DateTime.Now;
                        return ei;
                    }
                    if (wait_period > min_wait_period)
                    {
                        min_wait_period = wait_period;
                        min_wait_period_ei = ei;
                    }
                }
                if (min_wait_period_ei != null && Session.GetInputItemQueue<DataItem>().CountOfNew < 1)
                {
                    Thread.Sleep(min_wait_period);
                    emails2sent_time[min_wait_period_ei.Parent.ListAgentEmail] = DateTime.Now;
                    return min_wait_period_ei;
                }
                return null;
            }
        }
        static Random random = new Random();
        static Dictionary<string, DateTime> emails2sent_time = new Dictionary<string, DateTime>();

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
               // throw new Session.FatalException("fdsfsfs");
                CustomBot cb = (CustomBot)bc.Bot;

                string d = Log.OutputDir + "\\" + Log.This.Id + "_" + DateTime.Now.GetSecondsSinceUnixEpoch();
                Directory.CreateDirectory(d);

                string output_pdf = d + "\\" + PathRoutines.GetFileNameFromPath(template_pdf);
                {
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

                    set_field(ps.AcroFields, "Todays Date", DateTime.Today.ToShortDateString());
                    set_field(ps.AcroFields, "Buyer Name", Program.Settings.BuyerProfile.Name);
                    set_field(ps.AcroFields, "Address and Unit Number", Address + " " + UnitNumber);
                    set_field(ps.AcroFields, "City/Town", City);
                    //set_field(ps.AcroFields, "CLARK", );
                    set_field(ps.AcroFields, "Zip", ZipCode);
                    set_field(ps.AcroFields, "PARCEL NUMBER", ParcelNumber);
                    set_field(ps.AcroFields, "OfferAmt", OfferAmt);
                    string oa = Regex.Replace(OfferAmt, @"[^\d]", "");
                    if (oa.Length > 0)
                        set_field(ps.AcroFields, "OfferAmt in words", ConvertionRoutines.NumberToWords(int.Parse(oa)));
                    set_field(ps.AcroFields, "EMD", Program.Settings.Emd);
                    //set_field(ps.AcroFields, "Check Box1", );
                    //set_field(ps.AcroFields, "Balance", );
                    set_field(ps.AcroFields, "Co Buyer Name", Program.Settings.BuyerProfile.CoBuyerName);
                    set_field(ps.AcroFields, "<address> <UnitNumber> <City/town> NV <ZIP Code>", Address + " " + UnitNumber + ", " + City + " NV " + ZipCode);
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
                    set_field(ps.AcroFields, "City State Zip", Program.Settings.BrokerProfile.City + " " + Program.Settings.BrokerProfile.State + " " + Program.Settings.BrokerProfile.Zip);
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

                    for (int i = 1; i <= pr.NumberOfPages; i++)
                    {
                        var pcb = ps.GetOverContent(i);
                        add_image(pcb, System.Drawing.Image.FromFile(Program.Settings.BuyerProfile.InitialFile), new System.Drawing.Point(497, 67));
                        if(Program.Settings.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Program.Settings.BuyerProfile.CoBuyerInitialFile), new System.Drawing.Point(536, 67));
                    }
                    {
                        var pcb = ps.GetOverContent(9);
                        add_image(pcb, System.Drawing.Image.FromFile(Program.Settings.BuyerProfile.SignatureFile), new System.Drawing.Point(60, 190));
                        if (Program.Settings.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Program.Settings.BuyerProfile.CoBuyerSignatureFile), new System.Drawing.Point(60, 155));
                    }

                    ps.Close();
                    pr.Close();
                }

                string output_addendum_pdf = null;
                if (Program.Settings.ShortSaleAddendum)
                {
                    output_addendum_pdf = d + "\\" + PathRoutines.GetFileNameFromPath(template_addendum_pdf);

                    PdfReader.unethicalreading = true;
                    PdfReader pr = new PdfReader(template_addendum_pdf);
                    PdfStamper ps = new PdfStamper(pr, new FileStream(output_addendum_pdf, FileMode.Create, FileAccess.Write, FileShare.None));

                    //string fs = "";
                    //foreach (KeyValuePair<string, AcroFields.Item> kvp in ps.AcroFields.Fields)
                    //    fs += "\n{\"" + kvp.Key + "\", \"\"},";

                    if (Program.Settings.BuyerProfile.UseCoBuyer)
                        set_field(ps.AcroFields, "<Buyer Name> and <Co Buyer Name>", Program.Settings.BuyerProfile.Name + " and " + Program.Settings.BuyerProfile.CoBuyerName);
                    else
                        set_field(ps.AcroFields, "<Buyer Name> and <Co Buyer Name>", Program.Settings.BuyerProfile.Name);
                    set_field(ps.AcroFields, "Todays Date", DateTime.Today.ToShortDateString());
                    set_field(ps.AcroFields, "<Address> <UnitNumber> <City/Town> NV <Zip Code>", Address + " " + UnitNumber + ", " + City + " NV " + ZipCode);
                    set_field(ps.AcroFields, "Agent Name", Program.Settings.AgentProfile.Name);
                    //set_field(ps.AcroFields, "Agent Phone", );

                    ps.FormFlattening = true;

                    ps.Close();
                    pr.Close();
                }

                bc.Add(new EmailItem(output_pdf, output_addendum_pdf));
            }
            static readonly string template_pdf = Log.AppDir + "\\RPA.pdf";
            static readonly string template_addendum_pdf = Log.AppDir + "\\addendum.pdf";
        }

        static void set_field(AcroFields form, string field_key, string value)
        {
            switch (form.GetFieldType(field_key))
            {
                case AcroFields.FIELD_TYPE_CHECKBOX:
                case AcroFields.FIELD_TYPE_RADIOBUTTON:
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
        static System.Drawing.Size image_max_size = new System.Drawing.Size(105, 22);

        class EmailItem : InputItem
        {
            public DataItem Parent { get { return (DataItem)__ParentItem; } }
            public readonly string Pdf;
            public readonly string Addendum;

            internal EmailItem(string pdf, string addendum)
            {
                Pdf = pdf;
                Addendum = addendum;
            }

            override public void PROCESSOR(BotCycle bc)
            {
                CustomBot cb = (CustomBot)bc.Bot;

                cb.send(Parent.ListAgentEmail, Pdf, Addendum);
            }
        }

        void send(string to_email, params string[] attachments)
        {
            MailMessage mm = new MailMessage(
                Program.Settings.EmailServerProfile.SenderEmail,
                to_email
                )
            {
                Subject = Program.Settings.EmailTemplateProfile.Subject,
                Body = Program.Settings.EmailTemplateProfile.Body,
                From = new MailAddress(Program.Settings.EmailServerProfile.SenderEmail)
            };
            foreach (string a in attachments)
            {
                mm.Attachments.Add(new Attachment(a));
            }
            foreach (int i in Program.Settings.SelectedAttachmentIds)
            {
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
                    + "\r\nUserName: " + Program.Settings.EmailServerProfile.SenderEmail
                    + "\r\nSmtpPassword: " + Program.Settings.EmailServerProfile.SmtpPassword
                    + "\r\nFrom: " + mm.From.Address
                    + "\r\nTo: " + mm.To
                    );
                throw new Session.FatalException("SMTP error.", e);
            }
        }

        SmtpClient smtp_client = new SmtpClient
        {
            Host = Program.Settings.EmailServerProfile.SmtpHost,
            Port = Program.Settings.EmailServerProfile.SmtpPort,
            EnableSsl = true,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(
                 Program.Settings.EmailServerProfile.SenderEmail,
                 Program.Settings.EmailServerProfile.SmtpPassword
                 )
        };
    }
}