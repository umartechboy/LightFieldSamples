using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using PrincetonInstruments.LightField.AddIns;

namespace LightFieldAddInSamples.BAP_Lab_SimpleMultipointSpectroscope
{
    public partial class SimpleMultipointSpectroscopeControl : UserControl
    {
        // ── LightField interfaces ──────────────────────────────────────────────
        private readonly ILightFieldApplication _app;
        private readonly IExperiment            _experiment;

        // ── Hardware & State ───────────────────────────────────────────────────
        private MarlinDriver _marlin;
        private IImageDataSet _lastDataSet;
        private bool _controlsReady = false;

        // ── Scan State ─────────────────────────────────────────────────────────
        private CancellationTokenSource _scanCts;
        private readonly List<ScanPointRecord> _scanRecords = new List<ScanPointRecord>();

        public class ScanPointRecord
        {
            public double X;
            public double Y;
            public double Z;
            public string Timestamp;
            public string Filename;
        }

        public SimpleMultipointSpectroscopeControl(ILightFieldApplication app)
        {
            _app = app;
            _experiment = app.Experiment;

            InitializeComponent();

            _marlin = new MarlinDriver();
            _marlin.LineReceived += OnMarlinLineReceived;

            RefreshPorts();
            
            // Set default output folder to My Documents
            OutputFolderText.Text = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BAP_Lab_Export");

            _controlsReady = true;
            InitBlackPreview();
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TAB 1: SETUP
        // ═══════════════════════════════════════════════════════════════════════

        private void RefreshPorts_Click(object sender, RoutedEventArgs e) => RefreshPorts();

        private void RefreshPorts()
        {
            string current = PortCombo.SelectedItem as string;
            PortCombo.Items.Clear();
            string[] ports = SerialPort.GetPortNames();

            if (ports.Length == 0)
            {
                SerialStatusText.Text = "Status: no COM ports detected";
                SerialStatusText.Foreground = Brushes.Gray;
                return;
            }

            foreach (string port in ports) PortCombo.Items.Add(port);
            
            if (current != null && Array.IndexOf(ports, current) >= 0)
                PortCombo.SelectedItem = current;
            else
                PortCombo.SelectedIndex = 0;

            if (!_marlin.IsConnected)
            {
                SerialStatusText.Text = $"Status: {ports.Length} port(s) found – select one";
                SerialStatusText.Foreground = Brushes.DarkGoldenrod;
            }
        }

        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            if (_marlin.IsConnected)
            {
                _marlin.Disconnect();
                ConnectButton.Content = "Connect";
                SerialStatusText.Text = "Status: disconnected";
                SerialStatusText.Foreground = Brushes.Gray;
            }
            else
            {
                if (PortCombo.SelectedItem == null) return;
                string port = PortCombo.SelectedItem.ToString();
                int baud = int.Parse((BaudCombo.SelectedItem as ComboBoxItem).Content.ToString());
                try
                {
                    _marlin.Connect(port, baud);
                    ConnectButton.Content = "Disconnect";
                    SerialStatusText.Text = $"Status: connected to {port} @ {baud}";
                    SerialStatusText.Foreground = Brushes.DarkGreen;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to connect:\n" + ex.Message);
                }
            }
        }

        private void OnMarlinLineReceived(object sender, string line)
        {
            // Update terminal on UI thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                TerminalText.AppendText(line + "\n");
                TerminalText.ScrollToEnd();
            }));
        }

        private void SendCmdButton_Click(object sender, RoutedEventArgs e)
        {
            string cmd = ManualCommandBox.Text;
            if (string.IsNullOrWhiteSpace(cmd)) return;
            _marlin.SendRaw(cmd);
            ManualCommandBox.Text = "";
        }

        private void ManualCommandBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
                SendCmdButton_Click(sender, e);
        }

        // ── Homing & Coordinates ──────────────────────────────────────────────

        private void UpdatePosUI()
        {
            CurrentPosText.Text = $"X: {_marlin.CurrentX:F2}  Y: {_marlin.CurrentY:F2}  Z: {_marlin.CurrentZ:F2}";
        }

        private void HomeXY_Click(object sender, RoutedEventArgs e) { _marlin.HomeXY(); UpdatePosUI(); }
        private void HomeZ_Click(object sender, RoutedEventArgs e) { _marlin.HomeZ(); UpdatePosUI(); }
        private void SetOrigin_Click(object sender, RoutedEventArgs e) { _marlin.SetOrigin(); UpdatePosUI(); }
        private void SaveEEPROM_Click(object sender, RoutedEventArgs e) { _marlin.SaveToEEPROM(); }

        // ── Manual Movement ───────────────────────────────────────────────────

        private double GetMoveStep()
        {
            if (MoveStepCoarse.IsChecked == true) return 1.0;
            if (MoveStepFine.IsChecked == true) return 0.1;
            return 0.01;
        }

        private int GetFeedRate()
        {
            if (int.TryParse(MoveFeedRate.Text, out int f) && f > 0) return f;
            return 3000;
        }

        private void MoveXMinus_Click(object sender, RoutedEventArgs e) { _marlin.MoveRelativeX(-GetMoveStep(), GetFeedRate()); UpdatePosUI(); }
        private void MoveXPlus_Click(object sender, RoutedEventArgs e) { _marlin.MoveRelativeX(GetMoveStep(), GetFeedRate()); UpdatePosUI(); }
        private void MoveYMinus_Click(object sender, RoutedEventArgs e) { _marlin.MoveRelativeY(-GetMoveStep(), GetFeedRate()); UpdatePosUI(); }
        private void MoveYPlus_Click(object sender, RoutedEventArgs e) { _marlin.MoveRelativeY(GetMoveStep(), GetFeedRate()); UpdatePosUI(); }
        private void MoveZMinus_Click(object sender, RoutedEventArgs e) { _marlin.MoveRelativeZ(-GetMoveStep(), GetFeedRate()); UpdatePosUI(); }
        private void MoveZPlus_Click(object sender, RoutedEventArgs e) { _marlin.MoveRelativeZ(GetMoveStep(), GetFeedRate()); UpdatePosUI(); }

        // ── Pre-Acquire Setup ─────────────────────────────────────────────────

        private void InitBlackPreview()
        {
            if (!_controlsReady) return;

            int w = 1024;
            int h = 256;
            WriteableBitmap bmp = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgr32, null);
            int stride = bmp.BackBufferStride / 4;
            int[] pixels = new int[stride * h];

            bmp.Lock();
            Marshal.Copy(pixels, 0, bmp.BackBuffer, stride * h);
            bmp.AddDirtyRect(new Int32Rect(0, 0, w, h));
            bmp.Unlock();
            SetupImageControl.Source = bmp;
        }

        private void TestAcquireButton_Click(object sender, RoutedEventArgs e)
        {
            // Just acquires and shows in LightField UI. Does not append to scan records.
            if (!ValidateAcquisition()) return;

            try
            {
                TestAcquireButton.IsEnabled = false;

                // Read exactly how many frames LightField is currently configured to take
                int frames = Convert.ToInt32(_experiment.GetValue(ExperimentSettings.AcquisitionFramesToStore) ?? 1);
                
                IImageDataSet dataSet = _experiment.Capture(frames);
                _lastDataSet?.Dispose();
                _lastDataSet = dataSet;

                // Push to live Image preview 
                UpdateLiveImageFrame(dataSet);
                
                // Show acquired image on Setup tab panel too
                SetupImageControl.Source = CreateDisplayBitmap(dataSet.GetFrame(0, 0));

                // Push to LightField built-in display
                IDisplayViewer viewer = _app.DisplayManager?.GetDisplay(DisplayLocation.ExperimentWorkspace, 0);
                viewer?.Display("BAP Test Acquire", dataSet);

                AcqStatusText.Text = $"✓ Test acquired {frames} frame(s)";
            }
            catch (Exception ex)
            {
                AcqStatusText.Text = "✗ Acquire failed: " + ex.Message;
            }
            finally
            {
                TestAcquireButton.IsEnabled = true;
            }
        }

        private bool CameraExists
        {
            get
            {
                foreach (IDevice device in _experiment.ExperimentDevices)
                    if (device.Type == DeviceType.Camera) return true;
                return false;
            }
        }

        private bool ValidateAcquisition()
        {
            if (!CameraExists) { MessageBox.Show("Camera not found!"); return false; }
            if (!_experiment.IsReadyToRun) { MessageBox.Show("Experiment not ready!"); return false; }
            return true;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TAB 2: PROCESS (Scan Loop)
        // ═══════════════════════════════════════════════════════════════════════

        private void GoToButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_marlin.IsConnected) { MessageBox.Show("Connect Marlin first."); return; }
            if (double.TryParse(GoToX.Text, out double x) && double.TryParse(GoToY.Text, out double y))
            {
                _marlin.MoveXY(x, y, GetFeedRate());
                UpdatePosUI();
            }
        }

        private async void StartScanButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_marlin.IsConnected) { MessageBox.Show("Please connect serial port in Setup tab."); return; }
            if (!ValidateAcquisition()) return;

            if (!int.TryParse(GridXPoints.Text, out int xPoints) || xPoints < 1) return;
            if (!int.TryParse(GridYPoints.Text, out int yPoints) || yPoints < 1) return;
            if (!double.TryParse(GridXStep.Text, out double xStep)) return;
            if (!double.TryParse(GridYStep.Text, out double yStep)) return;
            if (!int.TryParse(GridDwellMs.Text, out int dwell) || dwell < 0) return;
            
            // Generate row-by-row traversal
            List<Point> path = new List<Point>();
            for (int y = 0; y < yPoints; y++)
            {
                for (int x = 0; x < xPoints; x++)
                {
                    path.Add(new Point(x * xStep, y * yStep));
                }
            }

            int feedRate = GetFeedRate();

            // Prepare UI
            StartScanButton.IsEnabled = false;
            StopScanButton.IsEnabled = true;
            ExportBothButton.IsEnabled = false;
            ScanProgressBar.Maximum = path.Count;
            ScanProgressBar.Value = 0;
            _scanRecords.Clear();

            _scanCts = new CancellationTokenSource();
            var token = _scanCts.Token;

            // Read exactly how many frames LightField is currently configured to take per point
            int frames = Convert.ToInt32(_experiment.GetValue(ExperimentSettings.AcquisitionFramesToStore) ?? 1);

            // Run Background Task
            try
            {
                await Task.Run(() => ScanLoop(path, feedRate, dwell, frames, token), token);
            }
            catch (OperationCanceledException)
            {
                ScanStatusText.Text = "Scan Cancelled.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Scan error: " + ex.Message);
            }
            finally
            {
                StartScanButton.IsEnabled = true;
                StopScanButton.IsEnabled = false;
                ExportBothButton.IsEnabled = true;
                _scanCts.Dispose();
                _scanCts = null;
            }
        }

        private void StopScanButton_Click(object sender, RoutedEventArgs e)
        {
            _scanCts?.Cancel();
            StopScanButton.IsEnabled = false;
            ScanStatusText.Text = "Cancelling...";
        }

        private void ScanLoop(List<Point> path, int feedRate, int dwellMs, int frames, CancellationToken token)
        {
            for (int i = 0; i < path.Count; i++)
            {
                token.ThrowIfCancellationRequested();

                Point p = path[i];
                
                // 1) Move Stage
                _marlin.MoveXY(p.X, p.Y, feedRate);
                _marlin.WaitForMotion(); // M400
                Thread.Sleep(dwellMs);   // Hardware settling
                
                // 2) Update UI that we arrived
                Dispatcher.Invoke(() => {
                    UpdatePosUI();
                    ScanStatusText.Text = $"Point {i+1}/{path.Count}  X: {p.X:F2}  Y: {p.Y:F2}";
                });

                token.ThrowIfCancellationRequested();

                // 3) Acquire on UI thread
                IImageDataSet dataSet = null;
                Dispatcher.Invoke(() => {
                    dataSet = _experiment.Capture(frames);
                    _lastDataSet?.Dispose();
                    _lastDataSet = dataSet;
                    UpdateLiveImageFrame(dataSet);
                });

                // 4) Save memory immediately
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                string pngName = $"Scan_X{p.X:F1}_Y{p.Y:F1}_{timestamp}.png";
                
                Dispatcher.Invoke(() => {
                    SaveDataSetToPng(dataSet, Path.Combine(OutputFolderText.Text, pngName));
                    ScanProgressBar.Value = i + 1;
                });

                // 5) Record memory
                _scanRecords.Add(new ScanPointRecord
                {
                    X = p.X,
                    Y = p.Y,
                    Z = _marlin.CurrentZ,
                    Timestamp = timestamp,
                    Filename = pngName
                });
            }

            Dispatcher.Invoke(() => {
                ScanStatusText.Text = "Scan Complete.";
            });
        }

        private void UpdateLiveImageFrame(IImageDataSet dataset)
        {
            try
            {
                IImageData frame = dataset.GetFrame(0, 0);
                WriteableBitmap bmp = CreateDisplayBitmap(frame);
                LiveImageControl.Source = bmp;
            }
            catch { /* Ignore draw errors */ }
        }

        private WriteableBitmap CreateDisplayBitmap(IImageData frame)
        {
            Array rawData = frame.GetData();
            int width = frame.Width;
            int height = frame.Height;
            
            WriteableBitmap bmp = new WriteableBitmap(width, height, 96, 96, PixelFormats.Bgr32, null);

            // Calculate min and max for auto-contrast stretching
            ushort minVal = ushort.MaxValue;
            ushort maxVal = ushort.MinValue;
            
            for (int i = 0; i < rawData.Length; i++)
            {
                ushort v = (ushort)rawData.GetValue(i);
                if (v < minVal) minVal = v;
                if (v > maxVal) maxVal = v;
            }

            // Prevent divide by zero if image is completely uniform
            double range = (maxVal - minVal);
            if (range == 0) range = 1;

            int stride = bmp.BackBufferStride / 4;
            int[] pixels = new int[stride * height];

            for (int row = 0; row < height; row++)
            {
                for (int col = 0; col < width; col++)
                {
                    ushort raw = (ushort)rawData.GetValue(row * width + col);
                    // Map [minVal, maxVal] to [0, 255]
                    double normalized = (raw - minVal) / range;
                    byte gray = (byte)(normalized * 255.0);
                    
                    pixels[row * stride + col] = (gray << 16) | (gray << 8) | gray;
                }
            }

            bmp.Lock();
            Marshal.Copy(pixels, 0, bmp.BackBuffer, stride * height);
            bmp.AddDirtyRect(new Int32Rect(0, 0, width, height));
            bmp.Unlock();

            return bmp;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  TAB 3: EXPORT
        // ═══════════════════════════════════════════════════════════════════════

        private void BrowseFolderButton_Click(object sender, RoutedEventArgs e)
        {
            // Simple hook for selecting a folder. Forms FolderBrowserDialog used generally.
            using (var fbd = new System.Windows.Forms.FolderBrowserDialog())
            {
                fbd.SelectedPath = OutputFolderText.Text;
                if (fbd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    OutputFolderText.Text = fbd.SelectedPath;
                }
            }
        }

        private void ExportBothButton_Click(object sender, RoutedEventArgs e)
        {
            if (_scanRecords.Count == 0)
            {
                MessageBox.Show("No scan data available to export. Run a scan first.");
                return;
            }

            string targetDir = OutputFolderText.Text;
            if (!Directory.Exists(targetDir))
            {
                Directory.CreateDirectory(targetDir);
            }

            // Note: PNGs are already saved to OutputFolderText.Text during the scan!
            // We just need to dump the JSON metadata.

            try
            {
                string jsonPath = Path.Combine(targetDir, $"ScanMetadata_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("[");
                for (int i = 0; i < _scanRecords.Count; i++)
                {
                    var r = _scanRecords[i];
                    sb.AppendLine("  {");
                    sb.AppendLine($"    \"X\": {r.X},");
                    sb.AppendLine($"    \"Y\": {r.Y},");
                    sb.AppendLine($"    \"Z\": {r.Z},");
                    sb.AppendLine($"    \"Timestamp\": \"{r.Timestamp}\",");
                    sb.AppendLine($"    \"Filename\": \"{r.Filename}\"");
                    sb.Append("  }");
                    if (i < _scanRecords.Count - 1) sb.Append(",");
                    sb.AppendLine();
                }
                sb.AppendLine("]");

                File.WriteAllText(jsonPath, sb.ToString());
                ExportStatusText.Text = $"Exported metadata to: {jsonPath}";
                ExportStatusText.Foreground = Brushes.DarkGreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Export failed:\n" + ex.Message);
            }
        }

        private void SaveDataSetToPng(IImageDataSet dataset, string filepath)
        {
            try
            {
                IImageData frame = dataset.GetFrame(0, 0);
                WriteableBitmap bmp = CreateDisplayBitmap(frame);

                string dir = Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                using (FileStream fs = new FileStream(filepath, FileMode.Create))
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bmp));
                    encoder.Save(fs);
                }
            }
            catch (Exception ex)
            {
                Dispatcher.Invoke(() => {
                    TerminalText.AppendText($"Error saving {filepath}: {ex.Message}\n");
                });
            }
        }
    }
}
