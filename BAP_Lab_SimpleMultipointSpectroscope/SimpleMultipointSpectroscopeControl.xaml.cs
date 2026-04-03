using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using PrincetonInstruments.LightField.AddIns;

namespace LightFieldAddInSamples.BAP_Lab_SimpleMultipointSpectroscope
{
    public class ScanPointUI : INotifyPropertyChanged
    {
        private int _index;
        private double _x;
        private double _y;
        
        public int Index { get { return _index; } set { _index = value; OnPropertyChanged(nameof(Index)); } }
        public double X { get { return _x; } set { _x = value; OnPropertyChanged(nameof(X)); } }
        public double Y { get { return _y; } set { _y = value; OnPropertyChanged(nameof(Y)); } }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
    public partial class SimpleMultipointSpectroscopeControl : UserControl
    {
        // ── LightField interfaces ──────────────────────────────────────────────
        private readonly ILightFieldApplication _app;
        private readonly IExperiment            _experiment;

        // ── Hardware & State ───────────────────────────────────────────────────
        private MarlinDriver _marlin;
        private IImageDataSet _lastDataSet;
        private bool _controlsReady = false;
        private bool _isLiveVideo = false;

        // ── Keyboard Movement ──────────────────────────────────────────────────
        private readonly System.Windows.Threading.DispatcherTimer _keyMoveTimer = new System.Windows.Threading.DispatcherTimer();
        private readonly HashSet<Key> _pressedKeys = new HashSet<Key>();
        private bool _isHotRegionActive = false;

        // ── Grid Object State ──────────────────────────────────────────────────
        private ObservableCollection<ScanPointUI> _gridPoints;
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
            OutputFolderText.Text = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "BAP_Lab_Export");

            _gridPoints = new ObservableCollection<ScanPointUI>();
            _gridPoints.CollectionChanged += GridPoints_CollectionChanged;
            GridPointsDataGrid.ItemsSource = _gridPoints;

            _controlsReady = true;
            InitBlackPreview();

            // Keyboard Motion Timer
            _keyMoveTimer.Interval = TimeSpan.FromMilliseconds(100);
            _keyMoveTimer.Tick += KeyMoveTimer_Tick;
            _keyMoveTimer.Start();
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
                PortCombo.SelectedIndex = ports.Length - 1;

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
            DrawGridPreview();
        }

        private void HomeXY_Click(object sender, RoutedEventArgs e) { _marlin.HomeXY(); UpdatePosUI(); }
        private void HomeZ_Click(object sender, RoutedEventArgs e) { _marlin.HomeZ(); UpdatePosUI(); }
        private void SetOrigin_Click(object sender, RoutedEventArgs e) { _marlin.SetOrigin(); UpdatePosUI(); }
        private void SaveEEPROM_Click(object sender, RoutedEventArgs e) { _marlin.SaveToEEPROM(); }

        // ── Manual Movement ───────────────────────────────────────────────────

        private int GetFeedRate()
        {
            if (int.TryParse(MoveFeedRate.Text, out int f) && f > 0) return f;
            return 3000;
        }

        private void MoveBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn) || btn.Tag == null) return;
            string tag = btn.Tag.ToString();
            if (tag.Length < 2) return;
            
            char axis = char.ToUpper(tag[0]);
            string rest = tag.Substring(1);
            bool negative = rest.StartsWith("-");
            if (negative) rest = rest.Substring(1);
            
            if (!int.TryParse(rest, out int rank)) return;
            
            // Get value from corresponding textbox
            double step = 0;
            if (axis == 'Z')
            {
                TextBox tb = (rank == 1) ? ZStepSize1 : (rank == 2) ? ZStepSize2 : ZStepSize3;
                double.TryParse(tb.Text, out step);
            }
            else
            {
                TextBox tb = (rank == 1) ? StepSize1 : (rank == 2) ? StepSize2 : StepSize3;
                double.TryParse(tb.Text, out step);
            }
            
            double dist = negative ? -step : step;
            int fr = GetFeedRate();
            
            if (axis == 'X') _marlin.MoveRelativeX(dist, fr);
            else if (axis == 'Y') _marlin.MoveRelativeY(dist, fr);
            else if (axis == 'Z') _marlin.MoveRelativeZ(dist, fr);
            UpdatePosUI();
        }

        // ── Keyboard & HotRegion Logic ─────────────────────────────────────────

        private void HotRegion_MouseEnter(object sender, MouseEventArgs e)
        {
            _isHotRegionActive = true;
            this.Focus(); // Ensure the control receives keyboard events
            if (KeyboardHotRegion != null)
                KeyboardHotRegion.Background = new SolidColorBrush(Color.FromArgb(40, 0, 255, 0)); // Subtle Green
            if (HotRegionStatus != null)
            {
                HotRegionStatus.Text = "CONTROL ACTIVE";
                HotRegionStatus.Foreground = Brushes.LimeGreen;
                HotRegionStatus.FontWeight = FontWeights.Bold;
            }
        }

        private void HotRegion_MouseLeave(object sender, MouseEventArgs e)
        {
            _isHotRegionActive = false;
            if (KeyboardHotRegion != null)
                KeyboardHotRegion.Background = Brushes.Transparent;
            if (HotRegionStatus != null)
            {
                HotRegionStatus.Text = "Hover to activate";
                HotRegionStatus.Foreground = Brushes.Gray;
                HotRegionStatus.FontWeight = FontWeights.Normal;
            }
            _pressedKeys.Clear();
        }

        protected override void OnPreviewKeyDown(KeyEventArgs e)
        {
            base.OnPreviewKeyDown(e);
            if (_isHotRegionActive)
            {
                // Capture image on Space bar
                if (e.Key == Key.Space)
                {
                    DoCaptureOnce();
                    e.Handled = true;
                    return;
                }

                if (!_pressedKeys.Contains(e.Key))
                {
                    _pressedKeys.Add(e.Key);
                    // Immediate Move for better responsiveness
                    ExecuteMoveForKey(e.Key);
                    e.Handled = true;
                }
            }
        }

        protected override void OnPreviewKeyUp(KeyEventArgs e)
        {
            base.OnPreviewKeyUp(e);
            _pressedKeys.Remove(e.Key);
        }

        private void KeyMoveTimer_Tick(object sender, EventArgs e)
        {
            if (!_isHotRegionActive || !_marlin.IsConnected || _pressedKeys.Count == 0) return;

            // Continually move as long as keys are held
            foreach (var key in _pressedKeys)
                ExecuteMoveForKey(key);

            UpdatePosUI();
        }

        private void ExecuteMoveForKey(Key k)
        {
            if (!_marlin.IsConnected) return;

            bool isCtrlDown = Keyboard.Modifiers.HasFlag(ModifierKeys.Control);
            double stepXY = 0;
            double stepZ = 0;

            // XY Logic: No Ctrl = Rank 3 (+++), Ctrl = Rank 2 (++)
            if (isCtrlDown) double.TryParse(StepSize2.Text, out stepXY);
            else            double.TryParse(StepSize3.Text, out stepXY);

            // Z Logic:  No Ctrl = Rank 2 (++),  Ctrl = Rank 1 (+)
            if (isCtrlDown) double.TryParse(ZStepSize1.Text, out stepZ);
            else            double.TryParse(ZStepSize2.Text, out stepZ);

            // Update Label String
            if (HotRegionStatus != null)
            {
                string mode = isCtrlDown ? "FINE" : "COARSE";
                HotRegionStatus.Text = $"{mode}: XY={stepXY} Z={stepZ}";
            }

            int fr = GetFeedRate();

            // Perform movement
            switch (k)
            {
                case Key.Up:    _marlin.MoveRelativeY(stepXY, fr); break;
                case Key.Down:  _marlin.MoveRelativeY(-stepXY, fr); break;
                case Key.Left:  _marlin.MoveRelativeX(-stepXY, fr); break;
                case Key.Right: _marlin.MoveRelativeX(stepXY, fr); break;
                case Key.PageUp:   _marlin.MoveRelativeZ(-stepZ, fr); break;
                case Key.PageDown: _marlin.MoveRelativeZ(stepZ, fr); break;
            }
        }

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

        private async void LiveVideo_Checked(object sender, RoutedEventArgs e)
        {
            if (_isLiveVideo) return;
            _isLiveVideo = true;
            TestAcquireButton.IsEnabled = false;

            try
            {
                while (_isLiveVideo)
                {
                    // Check conditions without showing message boxes repeatedly
                    if (!CameraExists || !_experiment.IsReadyToRun)
                    {
                        LiveVideoCheckBox.IsChecked = false;
                        _isLiveVideo = false;
                        break;
                    }

                    int frames = Convert.ToInt32(_experiment.GetValue(ExperimentSettings.AcquisitionFramesToStore) ?? 1);

                    IImageDataSet dataSet = await Task.Run(() =>
                    {
                        try { return _experiment.Capture(frames); }
                        catch { return null; }
                    });

                    if (dataSet != null)
                    {
                        _lastDataSet?.Dispose();
                        _lastDataSet = dataSet;
                        
                        IImageData frame = dataSet.GetFrame(0, 0);
                        bool showSpectrum = DisplayTypeCombo.SelectedIndex == 1;
                        
                        WriteableBitmap bmp = showSpectrum ? CreateSpectrumBitmap(frame) : CreateDisplayBitmap(frame);
                        SetupImageControl.Source = bmp;
                        
                        // Also update scan tab preview if scanning (though this loop is for Setup tab)
                        UpdateLiveImageFrame(dataSet);
                    }
                    
                    await Task.Delay(10);
                }
            }
            finally
            {
                _isLiveVideo = false;
                TestAcquireButton.IsEnabled = true;
            }
        }

        private void LiveVideo_Unchecked(object sender, RoutedEventArgs e)
        {
            _isLiveVideo = false;
        }

        private void TestAcquireButton_Click(object sender, RoutedEventArgs e)
        {
            // Just acquires and shows in LightField UI. Does not append to scan records.
            if (!ValidateAcquisition()) return;
            DoCaptureOnce();
        }

        private void DoCaptureOnce()
        {
            try
            {
                TestAcquireButton.IsEnabled = false;

                // Read exactly how many frames LightField is currently configured to take
                int frames = Convert.ToInt32(_experiment.GetValue(ExperimentSettings.AcquisitionFramesToStore) ?? 1);
                
                IImageDataSet dataSet = _experiment.Capture(frames);
                _lastDataSet?.Dispose();
                _lastDataSet = dataSet;

                // Push to live Image preview (Scan Tab)
                UpdateLiveImageFrame(dataSet);
                
                // Show on Setup tab based on selection
                IImageData frame = dataSet.GetFrame(0, 0);
                if (DisplayTypeCombo.SelectedIndex == 1)
                    SetupImageControl.Source = CreateSpectrumBitmap(frame);
                else
                    SetupImageControl.Source = CreateDisplayBitmap(frame);

                // Push to LightField built-in display
                IDisplayViewer viewer = _app.DisplayManager?.GetDisplay(DisplayLocation.ExperimentWorkspace, 0);
                viewer?.Display("BAP Test Acquire", dataSet);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Acquire failed: " + ex.Message);
            }
            finally
            {
                if (!_isLiveVideo) TestAcquireButton.IsEnabled = true;
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
        //  TAB 2: PROCESS (Grid & Scan)
        // ═══════════════════════════════════════════════════════════════════════

        private void GridPoints_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            for (int i = 0; i < _gridPoints.Count; i++) _gridPoints[i].Index = i + 1;

            if (e.OldItems != null)
                foreach (INotifyPropertyChanged item in e.OldItems) item.PropertyChanged -= GridPoint_PropertyChanged;
            if (e.NewItems != null)
                foreach (INotifyPropertyChanged item in e.NewItems) item.PropertyChanged += GridPoint_PropertyChanged;

            DrawGridPreview();
        }

        private void GridPoint_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "X" || e.PropertyName == "Y") DrawGridPreview();
        }

        private string ReadJsonString(string content, string key)
        {
            var match = Regex.Match(content, $"\\\"{key}\\\"\\s*:\\s*\\\"(.*?)\\\"");
            return match.Success ? match.Groups[1].Value : null;
        }

        private void LoadGridButton_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                Title = "Load Grid Configuration"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    string content = File.ReadAllText(dlg.FileName);
                    
                    // Parse Configuration
                    string xp = ReadJsonString(content, "XPoints"); if (xp != null) GridXPoints.Text = xp;
                    string yp = ReadJsonString(content, "YPoints"); if (yp != null) GridYPoints.Text = yp;
                    string xs = ReadJsonString(content, "XStep"); if (xs != null) GridXStep.Text = xs;
                    string ys = ReadJsonString(content, "YStep"); if (ys != null) GridYStep.Text = ys;
                    string dwell = ReadJsonString(content, "DwellMs"); if (dwell != null) GridDwellMs.Text = dwell;
                    
                    string pattern = ReadJsonString(content, "Pattern");
                    if (pattern != null)
                    {
                        foreach (ComboBoxItem item in GridPatternCombo.Items)
                        {
                            if (item.Content as string == pattern)
                            {
                                GridPatternCombo.SelectedItem = item;
                                break;
                            }
                        }
                    }
                    
                    string stag = ReadJsonString(content, "Staggered");
                    if (stag != null && bool.TryParse(stag, out bool bStag)) StaggeredCheck.IsChecked = bStag;

                    // Parse Points
                    var matches = Regex.Matches(content, @"\""X\""\s*:\s*([-0-9.]+)\s*,\s*\""Y\""\s*:\s*([-0-9.]+)");
                    
                    if (matches.Count > 0)
                    {
                        _gridPoints.Clear();
                        foreach (Match m in matches)
                        {
                            if (double.TryParse(m.Groups[1].Value, out double x) && double.TryParse(m.Groups[2].Value, out double y))
                            {
                                _gridPoints.Add(new ScanPointUI { X = x, Y = y });
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("No valid coordinate data found in file.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to load JSON: " + ex.Message);
                }
            }
        }

        private void SaveGridButton_Click(object sender, RoutedEventArgs e)
        {
            if (_gridPoints.Count == 0)
            {
                MessageBox.Show("No points in grid to save.");
                return;
            }

            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                Title = "Save Grid Configuration",
                FileName = $"GridConfig_{DateTime.Now:yyyyMMdd_HHmm}.json"
            };

            if (dlg.ShowDialog() == true)
            {
                try
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("{");
                    
                    // Config Block
                    sb.AppendLine("  \"Config\": {");
                    sb.AppendLine($"    \"XPoints\": \"{GridXPoints.Text}\",");
                    sb.AppendLine($"    \"YPoints\": \"{GridYPoints.Text}\",");
                    sb.AppendLine($"    \"XStep\": \"{GridXStep.Text}\",");
                    sb.AppendLine($"    \"YStep\": \"{GridYStep.Text}\",");
                    sb.AppendLine($"    \"DwellMs\": \"{GridDwellMs.Text}\",");
                    sb.AppendLine($"    \"Pattern\": \"{(GridPatternCombo.SelectedItem as ComboBoxItem)?.Content}\",");
                    sb.AppendLine($"    \"Staggered\": \"{StaggeredCheck.IsChecked == true}\"");
                    sb.AppendLine("  },");

                    // Points Block
                    sb.AppendLine("  \"Points\": [");
                    for (int i = 0; i < _gridPoints.Count; i++)
                    {
                        sb.AppendLine("    {");
                        sb.AppendLine($"      \"X\": {_gridPoints[i].X},");
                        sb.AppendLine($"      \"Y\": {_gridPoints[i].Y}");
                        sb.Append("    }");
                        if (i < _gridPoints.Count - 1) sb.Append(",");
                        sb.AppendLine();
                    }
                    sb.AppendLine("  ]");
                    sb.AppendLine("}");

                    File.WriteAllText(dlg.FileName, sb.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to save JSON: " + ex.Message);
                }
            }
        }

        private void GeneratePatternButton_Click(object sender, RoutedEventArgs e)
        {
            if (!int.TryParse(GridXPoints.Text, out int cols) || cols < 1) return;
            if (!int.TryParse(GridYPoints.Text, out int rows) || rows < 1) return;
            if (!double.TryParse(GridXStep.Text, out double xStep) || xStep <= 0) return;
            if (!double.TryParse(GridYStep.Text, out double yStep) || yStep <= 0) return;

            string pattern = (GridPatternCombo.SelectedItem as ComboBoxItem)?.Content as string ?? "Row-by-Row";
            bool staggered = StaggeredCheck.IsChecked == true;

            var points = new List<Point>();

            if (pattern == "Row-by-Row")
            {
                for (int y = 0; y < rows; y++)
                {
                    double xOffset = (staggered && y % 2 != 0) ? xStep / 2.0 : 0;
                    for (int x = 0; x < cols; x++) points.Add(new Point(x * xStep + xOffset, y * yStep));
                }
            }
            else if (pattern == "Column-by-Column")
            {
                for (int x = 0; x < cols; x++)
                {
                    double yOffset = (staggered && x % 2 != 0) ? yStep / 2.0 : 0;
                    for (int y = 0; y < rows; y++) points.Add(new Point(x * xStep, y * yStep + yOffset));
                }
            }
            else if (pattern == "Serpentine Row")
            {
                for (int y = 0; y < rows; y++)
                {
                    double xOffset = (staggered && y % 2 != 0) ? xStep / 2.0 : 0;
                    if (y % 2 == 0)
                    {
                        for (int x = 0; x < cols; x++) points.Add(new Point(x * xStep + xOffset, y * yStep));
                    }
                    else
                    {
                        for (int x = cols - 1; x >= 0; x--) points.Add(new Point(x * xStep + xOffset, y * yStep));
                    }
                }
            }
            else if (pattern == "Serpentine Column")
            {
                for (int x = 0; x < cols; x++)
                {
                    double yOffset = (staggered && x % 2 != 0) ? yStep / 2.0 : 0;
                    if (x % 2 == 0)
                    {
                        for (int y = 0; y < rows; y++) points.Add(new Point(x * xStep, y * yStep + yOffset));
                    }
                    else
                    {
                        for (int y = rows - 1; y >= 0; y--) points.Add(new Point(x * xStep, y * yStep + yOffset));
                    }
                }
            }

            // Center all points around (0,0)
            if (points.Count > 0)
            {
                double minX = points.Min(p => p.X);
                double maxX = points.Max(p => p.X);
                double minY = points.Min(p => p.Y);
                double maxY = points.Max(p => p.Y);
                double cx = (minX + maxX) / 2.0;
                double cy = (minY + maxY) / 2.0;

                _gridPoints.Clear();
                foreach (var p in points)
                {
                    _gridPoints.Add(new ScanPointUI { X = p.X - cx, Y = p.Y - cy });
                }
            }
        }

        private void DrawGridPreview()
        {
            if (GridCanvas == null || !_controlsReady) return;
            GridCanvas.Children.Clear();

            double targetW = 200;
            double targetH = 200;

            double minX = _marlin.CurrentX, maxX = _marlin.CurrentX;
            double minY = _marlin.CurrentY, maxY = _marlin.CurrentY;

            if (_gridPoints.Count > 0)
            {
                minX = Math.Min(minX, _gridPoints.Min(p => p.X));
                maxX = Math.Max(maxX, _gridPoints.Max(p => p.X));
                minY = Math.Min(minY, _gridPoints.Min(p => p.Y));
                maxY = Math.Max(maxY, _gridPoints.Max(p => p.Y));
            }

            double rangeX = maxX - minX;
            double rangeY = maxY - minY;
            double padding = Math.Max(rangeX, rangeY) * 0.15;
            if (padding == 0) padding = 1.0;

            minX -= padding; maxX += padding;
            minY -= padding; maxY += padding;

            double scaleX = targetW / (maxX - minX);
            double scaleY = targetH / (maxY - minY);
            double scale = Math.Min(scaleX, scaleY);

            double offsetX = (targetW - (maxX - minX) * scale) / 2.0 - minX * scale;
            double offsetY = (targetH - (maxY - minY) * scale) / 2.0 - minY * scale;

            Point ToCanvas(double x, double y) => new Point(x * scale + offsetX, targetH - (y * scale + offsetY));

            // Draw Paths
            if (_gridPoints.Count > 1)
            {
                for (int i = 0; i < _gridPoints.Count - 1; i++)
                {
                    Point p1 = ToCanvas(_gridPoints[i].X, _gridPoints[i].Y);
                    Point p2 = ToCanvas(_gridPoints[i+1].X, _gridPoints[i+1].Y);
                    
                    Line line = new Line { X1 = p1.X, Y1 = p1.Y, X2 = p2.X, Y2 = p2.Y, Stroke = Brushes.LightGray, StrokeThickness = 1.5 };
                    GridCanvas.Children.Add(line);
                }
            }

            // Draw Points
            for(int i = 0; i < _gridPoints.Count; i++)
            {
                Point c = ToCanvas(_gridPoints[i].X, _gridPoints[i].Y);
                Ellipse dot = new Ellipse { Width = 7, Height = 7, Fill = Brushes.SteelBlue };
                Canvas.SetLeft(dot, c.X - 3.5);
                Canvas.SetTop(dot, c.Y - 3.5);
                dot.Tag = _gridPoints[i];
                GridCanvas.Children.Add(dot);
            }

            // Draw Stage Pos
            Point stageP = ToCanvas(_marlin.CurrentX, _marlin.CurrentY);
            Ellipse stage = new Ellipse { Width = 9, Height = 9, Fill = Brushes.LimeGreen, Stroke = Brushes.DarkGreen, StrokeThickness = 1 };
            Canvas.SetLeft(stage, stageP.X - 4.5);
            Canvas.SetTop(stage, stageP.Y - 4.5);
            GridCanvas.Children.Add(stage);
        }

        private void GridCanvas_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.OriginalSource is Ellipse dot && dot.Tag is ScanPointUI p)
            {
                GridPointsDataGrid.SelectedItem = p;
                GridPointsDataGrid.ScrollIntoView(p);
                
                if (_marlin.IsConnected)
                {
                    _marlin.MoveXY(p.X, p.Y, GetFeedRate());
                    UpdatePosUI();
                }
                else
                {
                    MessageBox.Show("Please connect serial port in Setup tab to move the stage.");
                }
            }
        }

        private void GridCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.OriginalSource is Ellipse dot && dot.Tag is ScanPointUI p)
            {
                CanvasStatusText.Text = $"Idx: {p.Index} -> X: {p.X:F2}, Y: {p.Y:F2}";
            }
            else
            {
                CanvasStatusText.Text = "Hover over a point...";
            }
        }

        private async void StartScanButton_Click(object sender, RoutedEventArgs e)
        {
            if (!_marlin.IsConnected) { MessageBox.Show("Please connect serial port in Setup tab."); return; }
            if (_gridPoints.Count == 0) { MessageBox.Show("Please generate or add points to the grid first."); return; }
            if (!ValidateAcquisition()) return;

            if (!int.TryParse(GridDwellMs.Text, out int dwell) || dwell < 0) return;

            // Make a snapshot copy of the points for the background thread
            List<ScanPointUI> path = _gridPoints.ToList();

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

        private void ScanLoop(List<ScanPointUI> path, int feedRate, int dwellMs, int frames, CancellationToken token)
        {
            for (int i = 0; i < path.Count; i++)
            {
                token.ThrowIfCancellationRequested();

                ScanPointUI p = path[i];
                
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
                string csvName = $"Scan_X{p.X:F1}_Y{p.Y:F1}_{timestamp}.csv";
                
                Dispatcher.Invoke(() => {
                    SaveDataSetToPng(dataSet, System.IO.Path.Combine(OutputFolderText.Text, pngName));
                    SaveDataSetToCsv(dataSet, System.IO.Path.Combine(OutputFolderText.Text, csvName));
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

            // 6) Auto-Export Metadata
            Dispatcher.Invoke(() => {
                AutoSaveMetadata();
                ScanStatusText.Text = "Scan Complete.";
            });
        }

        private void AutoSaveMetadata()
        {
            try
            {
                string targetDir = OutputFolderText.Text;
                if (!Directory.Exists(targetDir)) Directory.CreateDirectory(targetDir);

                string jsonPath = System.IO.Path.Combine(targetDir, $"ScanMetadata_{DateTime.Now:yyyyMMdd_HHmmss}.json");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("{");
                
                // Camera Settings
                sb.AppendLine("  \"CameraSettings\": {");
                sb.AppendLine($"    \"ExposureTime\": \"{_experiment.GetValue(CameraSettings.ShutterTimingExposureTime)}\",");
                sb.AppendLine($"    \"Frames\": \"{_experiment.GetValue(ExperimentSettings.AcquisitionFramesToStore)}\",");
                sb.AppendLine($"    \"Grating\": \"{_experiment.GetValue(SpectrometerSettings.GratingSelected)}\"");
                sb.AppendLine("  },");

                // Scan Config
                sb.AppendLine("  \"ScanConfig\": {");
                sb.AppendLine($"    \"XPoints\": \"{GridXPoints.Text}\",");
                sb.AppendLine($"    \"YPoints\": \"{GridYPoints.Text}\",");
                sb.AppendLine($"    \"XStep\": \"{GridXStep.Text}\",");
                sb.AppendLine($"    \"YStep\": \"{GridYStep.Text}\",");
                sb.AppendLine($"    \"DwellMs\": \"{GridDwellMs.Text}\",");
                sb.AppendLine($"    \"Pattern\": \"{(GridPatternCombo.SelectedItem as ComboBoxItem)?.Content}\",");
                sb.AppendLine($"    \"Staggered\": \"{StaggeredCheck.IsChecked == true}\"");
                sb.AppendLine("  },");

                // Points History
                sb.AppendLine("  \"ScanRecords\": [");
                for (int i = 0; i < _scanRecords.Count; i++)
                {
                    var r = _scanRecords[i];
                    sb.AppendLine("    {");
                    sb.AppendLine($"      \"X\": {r.X},");
                    sb.AppendLine($"      \"Y\": {r.Y},");
                    sb.AppendLine($"      \"Z\": {r.Z},");
                    sb.AppendLine($"      \"Timestamp\": \"{r.Timestamp}\",");
                    sb.AppendLine($"      \"PNG\": \"{r.Filename}\",");
                    sb.AppendLine($"      \"CSV\": \"{r.Filename.Replace(".png", ".csv")}\"");
                    sb.Append("    }");
                    if (i < _scanRecords.Count - 1) sb.Append(",");
                    sb.AppendLine();
                }
                sb.AppendLine("  ]");
                sb.AppendLine("}");

                File.WriteAllText(jsonPath, sb.ToString());
            }
            catch { /* Ignore auto-export failure */ }
        }

        private void UpdateLiveImageFrame(IImageDataSet dataset)
        {
            try
            {
                IImageData frame = dataset.GetFrame(0, 0);
                bool showSpectrum = (ScanDisplayTypeCombo != null && ScanDisplayTypeCombo.SelectedIndex == 1);
                
                WriteableBitmap bmp = showSpectrum ? CreateSpectrumBitmap(frame) : CreateDisplayBitmap(frame);
                LiveImageControl.Source = bmp;
            }
            catch { /* Ignore draw errors */ }
        }
        
        private WriteableBitmap CreateSpectrumBitmap(IImageData frame)
        {
            Array rawData = frame.GetData();
            int width = frame.Width;
            int height = frame.Height;
            
            // Output bitmap height (match the UI border height roughly or fixed)
            int dispH = 150;
            WriteableBitmap bmp = new WriteableBitmap(width, dispH, 96, 96, PixelFormats.Bgr32, null);
            
            // 1. Calculate Averages
            double[] averages = new double[width];
            double maxAvg = 1;
            for (int x = 0; x < width; x++)
            {
                double sum = 0;
                for (int y = 0; y < height; y++)
                {
                    double val = Convert.ToDouble(rawData.GetValue(y * width + x));
                    sum += val;
                }
                averages[x] = sum / height;
                if (averages[x] > maxAvg) maxAvg = averages[x];
            }

            // 2. Prepare pixel buffer (Black background)
            int stride = bmp.BackBufferStride / 4;
            int[] pixels = new int[stride * dispH];
            
            // 3. Draw Graph (Yellow line)
            int lastY = -1;
            for (int x = 0; x < width; x++)
            {
                // Normalize to height
                int barH = (int)((averages[x] / maxAvg) * (dispH - 10));
                int currentY = dispH - 1 - barH;
                
                if (currentY < 0) currentY = 0;
                if (currentY >= dispH) currentY = dispH - 1;

                if (lastY == -1) // First pixel
                {
                    pixels[currentY * stride + x] = 0xFFFF00; // Red+Green=Yellow
                }
                else
                {
                    // Draw vertical segment from lastY to currentY to make it look continuous
                    int yStart = Math.Min(lastY, currentY);
                    int yEnd = Math.Max(lastY, currentY);
                    for (int y = yStart; y <= yEnd; y++)
                    {
                        pixels[y * stride + x] = 0xFFFF00;
                    }
                }
                lastY = currentY;
            }

            bmp.Lock();
            Marshal.Copy(pixels, 0, bmp.BackBuffer, pixels.Length);
            bmp.AddDirtyRect(new Int32Rect(0, 0, width, dispH));
            bmp.Unlock();
            return bmp;
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
                string jsonPath = System.IO.Path.Combine(targetDir, $"ScanMetadata_{DateTime.Now:yyyyMMdd_HHmmss}.json");
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

                string dir = System.IO.Path.GetDirectoryName(filepath);
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

        private void SaveDataSetToCsv(IImageDataSet dataset, string filepath)
        {
            try
            {
                IImageData frame = dataset.GetFrame(0, 0);
                Array rawData = frame.GetData();
                int width = frame.Width;
                int height = frame.Height;

                string dir = System.IO.Path.GetDirectoryName(filepath);
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                using (StreamWriter sw = new StreamWriter(filepath))
                {
                    for (int r = 0; r < height; r++)
                    {
                        var row = new List<string>();
                        for (int c = 0; c < width; c++)
                        {
                            row.Add(rawData.GetValue(r * width + c).ToString());
                        }
                        sw.WriteLine(string.Join(",", row));
                    }
                }
            }
            catch (Exception ex)
            {
                // Ignore silent failure
            }
        }
    }
}
