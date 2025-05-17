using Microsoft.Win32;
using System.IO;
using System.IO.Packaging;
using System.IO.Ports;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PMBusInterface.Data;

namespace PMBusInterface
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string AvalailablePorts { get; private set; }

        public string ScanString { get; private set; }

        public string PSUString { get; private set; }

        public string FRUString { get; private set; }

        private readonly SerialPort _serialService;
       

        public MainWindow()
        {
            InitializeComponent();

            AvalailablePorts = "";
            ScanString = "";
            PSUString = "";
            FRUString = "";

            _serialService = new SerialPort();

            

            FindAvailablePorts();
        }

        private void FindAvailablePorts()
        {
            // here you should find all available ports
            available_ports.ItemsSource = new COM_Port[]
            {
                new COM_Port {Name = "Port 1"},
                new COM_Port {Name = "Port 2"},
            };
        }

        public void ScanPort(object sender, RoutedEventArgs e)
        {
            //_serialService.Write("SCAN");

            // change scan string (ScanString)
            if(available_ports.SelectedItem is COM_Port port)
            {
                addresses.Text = port.Name;
            }
        }

        public void CheckFRU(object sender, RoutedEventArgs e)
        {
            // read uint from input field (read addres from text box)
            // change scan string (FRUString)
            FRUoutput.Text = fru_input.Text;
        }

        public void CheckPSU(object sender, RoutedEventArgs e)
        {
            // read uint from input field (read addres from text box)
            // change scan string (PSUString)
            PSUoutput.Text = psu_input.Text;
        }

        public void ExportToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                Report_Handler.CreateReport(output_path.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error - " + ex);
            }
        }

        public void ChooseOutputFolder(object sender, RoutedEventArgs e)
        {
            var folderDialog = new OpenFolderDialog
            {
                // todo 
            };

            if (folderDialog.ShowDialog() == true)
            {
                output_path.Text = folderDialog.FolderName;
                // Do something with the result
            }
        }
    }
}