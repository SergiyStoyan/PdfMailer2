using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.Configuration;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;

namespace Cliver.PdfMailer2
{
    public partial class Settings
    {
        public static readonly EmailClass Email;

        public class EmailClass : Cliver.Settings
        {
            public Dictionary<string, EmailServerProfile> EmailServerProfileNames2EmailServerProfile = new Dictionary<string, EmailServerProfile>();
            public string EmailServerProfileName;
            public bool UseRandomDelay;
            public int MinRandomDelayMss;
            public int MaxRandomDelayMss;
            [ScriptIgnore]
            public EmailServerProfile EmailServerProfile
            {
                get
                {
                    return EmailServerProfileNames2EmailServerProfile[EmailServerProfileName];
                }
            }
        }

        public class EmailServerProfile : Profile
        {
            public string SmtpHost;
            public int SmtpPort;
            public string SmtpPassword;
            public string SenderEmail;
        }
    }
}