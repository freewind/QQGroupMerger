using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using System.Diagnostics;


namespace QQGroupMerger {

    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window {

        private List<string> selectedMhtFiles = new List<string>();

        public MainWindow() {
            InitializeComponent();
            listBox1.ItemsSource = selectedMhtFiles;
        }

        private void button1_Click(object sender, RoutedEventArgs e) {
            var dialog = new OpenFileDialog();
            dialog.Filter = "QQ聊天记录导出文件(*.mht)|*.mht";
            dialog.Multiselect = true;
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                var filenames = dialog.FileNames;
                foreach (var filename in filenames) {
                    selectedMhtFiles.Add(filename);
                }
                listBox1.Items.Refresh();
            }
        }

        private string readMhtFile(string filename) {
            var reader = new StreamReader(filename);
            var sb = new StringBuilder();
            try {
                while (reader.Peek() != -1) {
                    sb.AppendLine(reader.ReadLine());
                }
                return sb.ToString();
            } finally {
                reader.Close();
            }
        }

        private void button2_Click(object sender, RoutedEventArgs e) {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) {
                var targetDir = dialog.SelectedPath;
                var readers = new List<MhtReader>();
                foreach (var file in selectedMhtFiles) {
                    var reader = new MhtReader(file);
                    readers.Add(reader);
                    reader.Parse();
                }

                var merger = new MhtMerger(readers);
                merger.Merge();
                merger.WriteToMultiHtml(targetDir);

                // delete tmp dirs
                foreach (var reader in readers) {
                    reader.DeleteTmpDir();
                }

                System.Windows.Forms.MessageBox.Show("合并成功");
                Process.Start("explorer.exe", targetDir);
            }
        }
    }
}


