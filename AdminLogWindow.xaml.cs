using System;
using System.IO;
using System.Windows;

namespace SchBom_Convert
{
    public partial class AdminLogWindow : Window
    {
        public AdminLogWindow()
        {
            InitializeComponent();
            string logDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".sys");
            string logPath = System.IO.Path.Combine(logDir, "sysdata.dat");
            if (File.Exists(logPath))
                LogTextBox.Text = File.ReadAllText(logPath);
            else
                LogTextBox.Text = "尚無紀錄";
        }
    }
} 