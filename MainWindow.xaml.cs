using Microsoft.Win32;
using PMBusInterface.Data;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.IO.Ports;
using System.Linq;
using System.Linq;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace PMBusInterface
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private enum ReportType { None, PSU, FRU, Scan }
        private ReportType currentReportType = ReportType.None;
        private readonly StringBuilder currentReportBuffer = new StringBuilder();
        private DispatcherTimer _scanTimer;
        private readonly StringBuilder _receiveBuffer = new StringBuilder();
        private readonly object _bufferLock = new object();
        private readonly ManualResetEventSlim _dataReceivedEvent = new ManualResetEventSlim(false);

        private CancellationTokenSource _readCancellationTokenSource;
        private SerialPort _serialPort;

        public MainWindow()
        {
            InitializeComponent();
            InitializeSerialPort();
            LoadAvailablePorts();
            SetupScanTimer();
        }

        private void InitializeSerialPort()
        {
            _serialPort = new SerialPort
            {
                BaudRate = 115200,
                Parity = Parity.None,
                StopBits = StopBits.One,
                DataBits = 8,
                Handshake = Handshake.None,
                ReadTimeout = 4000,
                WriteTimeout = 4000,

            };

            _serialPort.DataReceived += SerialPort_DataReceived;
        }

        private void LoadAvailablePorts()
        {
            available_ports.ItemsSource = SerialPort.GetPortNames()
                .Select(p => new COM_Port { Name = p, Description = GetPortDescription(p) })
                .ToList();

            if (available_ports.Items.Count > 0)
                available_ports.SelectedIndex = 0;
        }

        private async void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            try
            {
                // Читаем все доступные данные
                var bytesToRead = _serialPort.BytesToRead;
                if (bytesToRead <= 0) return;

                var buffer = new byte[bytesToRead];
                var bytesRead = _serialPort.Read(buffer, 0, bytesToRead);
                var data = Encoding.ASCII.GetString(buffer, 0, bytesRead);

                lock (_bufferLock)
                {
                    _receiveBuffer.Append(data);
                }

                _dataReceivedEvent.Set();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка обработки данных: {ex}");
                Dispatcher.Invoke(() => addresses.Text += $"Ошибка: {ex.Message}\n");
            }
        }
        
        private void SetupScanTimer()
        {
            _scanTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };

            _scanTimer.Tick += async (s, e) =>
            {
                if (!_serialPort.IsOpen) return;

                try
                {
                    string dataToProcess = null;

                    lock (_bufferLock)
                    {
                        if (_receiveBuffer.Length > 0)
                        {
                            dataToProcess = _receiveBuffer.ToString();
                            _receiveBuffer.Clear();
                        }
                    }

                    if (dataToProcess != null)
                    {
                        await ProcessReceivedData(dataToProcess);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка таймера: {ex}");
                }
            };
        }

        private async Task ProcessReceivedData(string data)
        {
            await Dispatcher.InvokeAsync(() =>
            {
                foreach (var line in data.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (line.Contains("[PSU_START]"))
                    {
                        currentReportType = ReportType.PSU;
                        currentReportBuffer.Clear();
                        
                        continue;
                    }
                    else if (line.Contains("[FRU_START]"))
                    {
                        currentReportType = ReportType.FRU;
                        currentReportBuffer.Clear();
                        
                        continue;
                    }
                    else if (line.Contains("[PSU_END]"))
                    {
                        PSUoutput.Text = currentReportBuffer.ToString();
                        currentReportType = ReportType.None;
                        continue;
                    }
                    else if (line.Contains("[FRU_END]"))
                    {
                        FRUoutput.Text = currentReportBuffer.ToString();
                        currentReportType = ReportType.None;
                        continue;
                    }

                    switch (currentReportType)
                    {
                        case ReportType.PSU:
                            currentReportBuffer.AppendLine(line);
                            break;
                        case ReportType.FRU:
                            currentReportBuffer.AppendLine(line);
                            break;
                        default:
                            addresses.Text += line + "\n";
                            break;
                    }
                }
            });
         }


        private static string GetPortDescription(string portName)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher(
                    $"SELECT * FROM Win32_SerialPort WHERE DeviceID = '{portName}'"))
                {
                    var port = searcher.Get().OfType<ManagementObject>().FirstOrDefault();
                    return port?["Description"]?.ToString() ?? "Unknown device";
                }
            }
            catch
            {
                return "Unknown device";
            }
        }

        private void ScanPort(object sender, RoutedEventArgs e)
        {
            if (available_ports.SelectedItem == null)
            {
                MessageBox.Show("Выберите COM-порт!");
                return;
            }

            try
            {
                var port = (COM_Port)available_ports.SelectedItem;

                if (_serialPort.IsOpen)
                    _serialPort.Close();

                _serialPort.PortName = port.Name;
                _serialPort.Open();

                addresses.Text = "Scanning...\n";
                _serialPort.DiscardInBuffer();
                _serialPort.DiscardOutBuffer();

                // Отправка команды
                _serialPort.Write("SCAN\r\n");

                _scanTimer.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
                Debug.WriteLine(ex.ToString());
            }
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            LoadAvailablePorts(); // Перезагружаем список портов
        }

       
        private void CheckFRU(object sender, RoutedEventArgs e)
        {
            if (!_serialPort.IsOpen)
            {
                MessageBox.Show("Сначала откройте COM-порт!");
                return;
            }

            FRUoutput.Text = "Ожидание данных FRU...\n";
            _serialPort.DiscardInBuffer();
            _serialPort.Write($"CHECK_FRU {fru_input.Text}\r\n");
        }


        private void CheckPSU(object sender, RoutedEventArgs e)
        {
            if (!_serialPort.IsOpen)
            {
                MessageBox.Show("Сначала откройте COM-порт!");
                return;
            }

            PSUoutput.Text = "Ожидание данных PSU...\n";
            _serialPort.DiscardInBuffer();
            _serialPort.Write($"CHECK_PSU {psu_input.Text}\r\n");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _scanTimer.Stop();

            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
            }

            _readCancellationTokenSource?.Cancel();
        }

        private void ExportToExcel(object sender, RoutedEventArgs e)
        {
            try
            {
                var psuData = PSUoutput.Text;
                var fruData = FRUoutput.Text;

                Report_Handler.CreateReport(
                    string.IsNullOrEmpty(output_path.Text) ?
                        Report_Handler.GetDefaultFilePath() :
                        output_path.Text,
                    psuData,
                    fruData
                );


            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}");
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
        private void ClearOutput_Click(object sender, RoutedEventArgs e)
        {
            addresses.Text = string.Empty;
        }

    }
}