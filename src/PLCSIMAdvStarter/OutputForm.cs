using System.Drawing;
using System.Windows.Forms;

namespace PLCSIMAdvStarter
{
    public sealed partial class OutputForm : Form
    {
        public OutputForm(string windowTitle,string messageTitle, string message)
        {
            InitializeComponent();

            Text = windowTitle.Trim();
            labelTitle.Text = messageTitle.Trim();
            textBoxMessage.Text = message.Trim();
            textBoxMessage.SelectionStart = 0;
            textBoxMessage.SelectionLength = 0;
        }
    }
}
