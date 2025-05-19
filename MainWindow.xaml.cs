using Microsoft.Win32;
using PMBusInterface.Data;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.IO.Ports;
using System.Linq;
using System.Management;
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
using System.Windows.Threading;

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

        private SerialPort _serialPort;
        private readonly DispatcherTimer _scanTimer = new DispatcherTimer();
        /*AvalailablePorts = "";
            ScanString = "";
            PSUString = "";
            FRUString = "";

            _serialPort = new SerialPort();*/



        public MainWindow()
        {
            InitializeComponent();
            InitializeSerialPort();
            LoadAvailablePorts();
            SetupScanTimer();
            

            //FindAvailablePorts();
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
                ReadTimeout = 500,
                WriteTimeout = 500
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

        private void SetupScanTimer()
        {
            _scanTimer.Interval = TimeSpan.FromMilliseconds(100); // Чаще проверяем
            _scanTimer.Tick += (s, e) =>
            {
                if (!_serialPort.IsOpen) return;

                try
                {
                    // Чтение всех доступных данных
                    if (_serialPort.BytesToRead > 0)
                    {
                        string data = _serialPort.ReadExisting();
                        Dispatcher.Invoke(() => {
                            addresses.Text += data;
                            // Автопрокрутка
                            var scrollViewer = FindVisualChild<ScrollViewer>(addresses);
                            scrollViewer?.ScrollToEnd();
                        });
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Ошибка чтения: {ex}");
                    _scanTimer.Stop();
                }
            };
        }

        // Вспомогательный метод для поиска ScrollViewer
        private static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                var child = VisualTreeHelper.GetChild(obj, i);
                if (child is T result)
                    return result;
                var childResult = FindVisualChild<T>(child);
                if (childResult != null)
                    return childResult;
            }
            return null;
        }

        private string GetPortDescription(string portName)
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

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
           /* var data = _serialPort.ReadExisting();
            Dispatcher.Invoke(() =>
            {
                // Здесь будет обработка входящих данных
                // Например:
                if (data.Contains("Found device"))
                {
                    addresses.Text += data + "\n";
                }
            });*/
        }
        private void HexInput_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Разрешаем только 0-9, A-F, a-f и x
            e.Handled = !System.Text.RegularExpressions.Regex.IsMatch(
                e.Text, @"^[0-9a-fA-Fx]+$");

            // Автоматически добавляем 0x если нужно
            if (fru_input.Text.Length == 0 && e.Text.ToLower() != "0")
            {
                fru_input.Text = "0x";
                fru_input.CaretIndex = 2;
            }
        }
        private void ScanPort(object sender, RoutedEventArgs e)
        {
            if (available_ports.SelectedItem == null)
            {
                MessageBox.Show("Выберите COM-порт!");
                return;
            }

            var port = (COM_Port)available_ports.SelectedItem;

            try
            {
                // Закрыть порт, если открыт
                if (_serialPort.IsOpen)
                    _serialPort.Close();

                // Настройка и открытие
                _serialPort.PortName = port.Name;
                _serialPort.Open();

                // Диагностика
                Debug.WriteLine($"Порт {port.Name} открыт: {_serialPort.IsOpen}");
                addresses.Text = "Scanning...\n";

                // Важно: очистить буферы перед отправкой!
                _serialPort.DiscardInBuffer();
                _serialPort.DiscardOutBuffer();

                // Отправка команды с переносом строки
                _serialPort.Write("SCAN\r\n");  // \n или \r\n в зависимости от МК

                // Запуск таймера для чтения
                _scanTimer.Start();
            }
            catch (Exception ex)
            {
                addresses.Text += $"Ошибка: {ex.Message}\n";
                Debug.WriteLine(ex.ToString());
            }
        }

        private void ComboBox_DropDownOpened(object sender, EventArgs e)
        {
            LoadAvailablePorts(); // Перезагружаем список портов
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


        private async void CheckFRU(object sender, RoutedEventArgs e)
        {
            if (!_serialPort.IsOpen)
            {
                MessageBox.Show("Сначала откройте COM-порт!");
                return;
            }

            if (string.IsNullOrWhiteSpace(fru_input.Text) ||
                !fru_input.Text.StartsWith("0x") ||
                fru_input.Text.Length < 3)
            {
                MessageBox.Show("Введите корректный адрес (например: 0x50)");
                return;
            }

            try
            {
                // Очистка предыдущих данных
                FRUoutput.Text = "Запрос данных FRU...\n";

                // Отправка команды
                string command = $"CHECK_FRU {fru_input.Text}\n";
                _serialPort.Write(command);

                // Ожидание ответа с таймаутом
                await WaitForResponseAsync("FRU_DATA", TimeSpan.FromSeconds(3));
            }
            catch (Exception ex)
            {
                FRUoutput.Text += $"Ошибка: {ex.Message}\n";
            }
        }

        private async Task WaitForResponseAsync(string expectedStart, TimeSpan timeout)
        {
            DateTime startTime = DateTime.Now;
            StringBuilder response = new StringBuilder();

            while (DateTime.Now - startTime < timeout)
            {
                if (_serialPort.BytesToRead > 0)
                {
                    string data = _serialPort.ReadExisting();
                    response.Append(data);

                    if (response.ToString().Contains(expectedStart))
                    {
                        Dispatcher.Invoke(() => {
                            FRUoutput.Text += data;
                        });
                        return;
                    }
                }
                await Task.Delay(100);
            }
            throw new TimeoutException("Устройство не ответило");
        }

        private void CheckPSU(object sender, RoutedEventArgs e)
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
        private void ClearOutput_Click(object sender, RoutedEventArgs e)
        {
            addresses.Text = string.Empty;
        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            _scanTimer.Stop();
            if (_serialPort.IsOpen)
                _serialPort.Close();
        }
    }
}