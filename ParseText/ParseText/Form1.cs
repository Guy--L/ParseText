using System;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace ParseText
{
    public partial class Form1 : Form
    {
        public bool doCharts { get { return graphsOn.Checked; } }
        public bool notoutset { get { return string.IsNullOrWhiteSpace(output.Text); } }
        public string outdir { get { return output.Text; } }

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {
            output.Text = "";
        }

        public void WriteLine(string s)
        {
            progress.WriteLine(s);
        }

        private static string[] nLabels = new string[] { "n^2", "n0", "nInf", "tC", "t95" };

        private void button3_Click(object sender, EventArgs e)
        {
            var oldcolor = button3.BackColor;
            button3.BackColor = Color.Green;
            button3.Enabled = false;
            progress.Clear();

            foreach (var s in inputs.Items)
            {
                var file = s as string;
                WriteLine(file);
                Program.ControlXLInDir(file);
                WriteLine("input " + s + " processed");
            }

            WriteLine("Done");
            button3.BackColor = oldcolor;
            button3.Enabled = true;
        }

        private bool heading = false;

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = folderBrowser.ShowDialog();
            if (dr == DialogResult.OK)
            {
                inputs.Items.Add(folderBrowser.SelectedPath);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = folderBrowser.ShowDialog();
            if (dr == DialogResult.OK)
            {
                output.Text = folderBrowser.SelectedPath;

                bool exists = Directory.Exists(output.Text);
                if (!exists) Directory.CreateDirectory(output.Text);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Properties.Settings.Default["indirectories"].ToString().Split(',').Select(s => inputs.Items.Add(s)).ToList();
            output.Text = Properties.Settings.Default["outdirectory"].ToString();

            bool exists = Directory.Exists(output.Text);
            if (!exists) Directory.CreateDirectory(output.Text);

            button3.Enabled = inputs.Items.Count > 0;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var ins = string.Join(",", inputs.Items.Cast<String>().Select(s => s).ToArray());
            Properties.Settings.Default["indirectories"] = ins;
            Properties.Settings.Default["outdirectory"] = output.Text;
            Properties.Settings.Default.Save();
            Program.Release();
        }

        private void inputs_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = inputs.IndexFromPoint(e.Location);
            if (index != ListBox.NoMatches)
            {
                inputs.Items.RemoveAt(index);
            }
        }

        private void output_Leave(object sender, EventArgs e)
        {
            bool exists = Directory.Exists(output.Text);
            if (!exists) Directory.CreateDirectory(output.Text);
        }
    }

    public static class WinFormsExtensions
    {
        public static void WriteLine(this TextBox source, string value)
        {
            if (source.Text.Length == 0)
                source.Text = value;
            else
                source.AppendText("\r\n" + value);
        }
    }
}
