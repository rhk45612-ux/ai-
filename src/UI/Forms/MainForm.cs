using System.Windows.Forms;

namespace MyApp
{
    public class MainForm : Form
    {
        public MainForm()
        {
            Text = "AI 프로그램";
            Width = 1200;
            Height = 800;

            var view = new Views.ExcelUnitSizeView
            {
                Dock = DockStyle.Fill
            };

            Controls.Add(view);
        }
    }
}
