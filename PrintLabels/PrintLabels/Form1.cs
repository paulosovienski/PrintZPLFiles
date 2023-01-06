using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PrintLabels
{
    public partial class Form1 : Form
    {

        List<ProcessStartInfo> lPFI = new List<ProcessStartInfo>();
        List<string> lPaths = new List<string>();
        private StreamReader streamToPrint;
        List<ConfigPrinter> lConfigPrinter = new List<ConfigPrinter>();
        List<ItemToPrint> lItemToPrin = new List<ItemToPrint>();

        bool error = false;

        public Form1()
        {
            InitializeComponent();

            var builder = new ConfigurationBuilder().AddJsonFile($"appsettings.json", true, true);
            var config = builder.Build();            

            for (int i = 1; i <= 5; i++)
            {
                var p = new ConfigPrinter()
                {
                    Printer = config.GetSection("ConfigPrinter" + i + ":Printer").Value,
                    Server = config.GetSection("ConfigPrinter" + i + ":Server").Value
                };
                if (!string.IsNullOrEmpty(p.Printer))
                {
                    lConfigPrinter.Add(p);
                }
            }           

            comboBox1.DisplayMember = "Printer";
            comboBox1.DataSource = lConfigPrinter;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string s = "";

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    string[] files = Directory.GetFiles(fbd.SelectedPath);

                    foreach (var item in files)
                    {

                        try
                        {
                            streamToPrint = new StreamReader(item);
                            string label = streamToPrint.ReadToEnd();
                            lPaths.Add(item);
                            s += string.Format(item.Substring((item.Length - 17), 17));
                            s += Environment.NewLine;

                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                    textBox1.Text = s;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string printer = comboBox1.Text;
            string server = lConfigPrinter.Where(x => x.Printer == printer).FirstOrDefault().Server;

            foreach (var item in lPaths)
            {
                var args = string.Format(@"LPR -S {0} -P {1} -ol {2}", server, printer, item);

                try
                {
                    using (var process = new Process())
                    {
                        process.StartInfo = new ProcessStartInfo
                        {
                            FileName = "cmd.exe",
                            Arguments = $"/c {args}",
                            RedirectStandardOutput = true,
                            RedirectStandardError = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };

                        process.Start();
                        var err = process.StandardError.ReadToEnd();

                        if (!string.IsNullOrEmpty(err))
                        {
                            MessageBox.Show("Print failed: " + err);
                            break;
                        }
                        var msg = process.StandardOutput.ReadToEnd();
                        process.WaitForExit();

                    }
                }
                catch (Exception ex)
                {
                    var err = ex.Message;
                    MessageBox.Show("Fail to print:" + ex.Message);
                }
            }
            MessageBox.Show("Files printed with success");
        }

    }
}
