using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VstsProjectDocumenter
{
    public partial class AddProjectUrlForm : Form
    {
        public AddProjectUrlForm()
        {
            InitializeComponent();
        }

        public DialogResult ShowDialog(out string projectUrl, out string projectName)
        {
            var dr = this.ShowDialog();
            projectUrl = AddProjectUrlTextBox.Text;
            projectName = projectNameTextBox.Text;
            return dr;
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
        }
    }
}
