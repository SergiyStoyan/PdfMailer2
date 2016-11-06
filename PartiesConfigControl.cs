using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace Cliver.PdfMailer2
{
    public partial class PartiesConfigControl : Cliver.BotGui.ConfigControl
    {
        new public const string NAME = "Parties";

        public PartiesConfigControl()
        {
            InitializeComponent();
            Init(NAME);
        }

        override protected void set_tool_tip()
        {
            //toolTip1.SetToolTip(this.DbConnectionString, "Database connection string");
        }

        private void bDelete_Click(object sender, EventArgs e)
        {

        }

        private void bSave_Click(object sender, EventArgs e)
        {

        }
    }
}

