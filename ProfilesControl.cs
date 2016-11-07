using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cliver.PdfMailer2
{
    public partial class ProfilesControl : UserControl
    {
        public ProfilesControl()
        {
            InitializeComponent();
        }

        private void bSave_Click(object sender, EventArgs e)
        {
            if (Add == null || !Add())
                return;
            Names.Items.Remove(Names.Text);
            Names.Items.Insert(0, Names.Text);
            Names.SelectedIndex = 0;
        }
        public delegate bool OnAdd();
        public OnAdd Add = null;

        private void bDelete_Click(object sender, EventArgs e)
        {
            if (!Message.YesNo("Are you sure?"))
                return;
            if (Delete == null)
                return;
            Delete();
            Names.Items.Remove(Names.Text);
            Names.Text = "";
        }
        public delegate void OnDelete();
        public OnDelete Delete = null;

        private void Names_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Select == null)
                return;
            Select();
        }
        public delegate void OnSelect();
        public OnSelect Select = null;
    }
}
