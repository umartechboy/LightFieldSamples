using System;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using PrincetonInstruments.LightField.AddIns;

namespace LightFieldAddInSamples.BAP_Lab_SimpleMultipointSpectroscope
{
    /// <summary>
    /// Code-behind for SimpleMultipointSpectroscopeControl.xaml
    ///
    /// PoC features:
    ///   A) Enumerate serial (COM) ports — useful for the Marlin stage.
    ///   B) Configure ROI (Full / Binned / Custom) and acquire an image.
    ///   C) Display the acquired image in the LightField viewer.
    ///   D) Save the last acquired frame as a PNG using a SaveFileDialog.
    ///   E) Maintain a ListBox of acquired sample titles.
    /// </summary>
    public partial class SimpleMultipointSpectroscopeControl : UserControl
    {
        // ── LightField interfaces ──────────────────────────────────────────────
        private readonly ILightFieldApplication _app;
        private readonly IExperiment            _experiment;

        // ── Image cache ───────────────────────────────────────────────────────
        private IImageDataSet _lastDataSet;   // cached after every Acquire

        // ── ROI state ─────────────────────────────────────────────────────────
        private int _binW   = 1, _binH   = 1;      // full-sensor binning
        private int _roiX   = 0, _roiY   = 0;
        private int _roiW   = 100, _roiH  = 100;
        private int _roiXBin = 1, _roiYBin = 1;

        private bool _controlsReady = false;        // guard early TextChanged

        ///////////////////////////////////////////////////////////////////////////
        public SimpleMultipointSpectroscopeControl(ILightFieldApplication app)
        {
            _app        = app;
            _experiment = app.Experiment;

            // InitializeComponent MUST come first – it creates all the named controls.
            // Any method that touches XAML-named elements must be called AFTER this.
            InitializeComponent();

            RefreshPorts();          // populate COM port list on load
            _controlsReady = true;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  SECTION A  –  Serial Port
        // ═══════════════════════════════════════════════════════════════════════

        private void RefreshPorts_Click(object sender, RoutedEventArgs e)
        {
            RefreshPorts();
        }

        private void RefreshPorts()
        {
            PortCombo.Items.Clear();
            string[] ports = SerialPort.GetPortNames();

            if (ports.Length == 0)
            {
                SerialStatusText.Text       = "Status: no COM ports detected";
                SerialStatusText.Foreground = Brushes.Gray;
                return;
            }

            foreach (string port in ports)
                PortCombo.Items.Add(port);

            PortCombo.SelectedIndex         = 0;
            SerialStatusText.Text           = $"Status: {ports.Length} port(s) found – select one (not connected)";
            SerialStatusText.Foreground     = Brushes.DarkGoldenrod;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  SECTION B  –  ROI radio buttons
        // ═══════════════════════════════════════════════════════════════════════

        private void FullChip_Checked(object sender, RoutedEventArgs e)
        {
            if (!_controlsReady) return;
            BinningPanel.Visibility  = Visibility.Collapsed;
            CustomRoiPanel.Visibility = Visibility.Collapsed;

            if (CameraExists)
                _experiment.SetFullSensorRegion();
        }

        private void FullSensorBinned_Checked(object sender, RoutedEventArgs e)
        {
            if (!_controlsReady) return;
            BinningPanel.Visibility  = Visibility.Visible;
            CustomRoiPanel.Visibility = Visibility.Collapsed;

            if (CameraExists)
                _experiment.SetBinnedSensorRegion(_binW, _binH);
        }

        private void CustomRegion_Checked(object sender, RoutedEventArgs e)
        {
            if (!_controlsReady) return;
            BinningPanel.Visibility  = Visibility.Collapsed;
            CustomRoiPanel.Visibility = Visibility.Visible;

            ApplyCustomRoi();
        }

        // ── Binning text-changed ───────────────────────────────────────────────

        private void BinWidth_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(BinWidth.Text)) return;
            if (int.TryParse(BinWidth.Text, out int v) && v > 0)
            {
                _binW = v;
                if (CameraExists && FullSensorBinned.IsChecked == true)
                    _experiment.SetBinnedSensorRegion(_binW, _binH);
            }
        }

        private void BinHeight_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(BinHeight.Text)) return;
            if (int.TryParse(BinHeight.Text, out int v) && v > 0)
            {
                _binH = v;
                if (CameraExists && FullSensorBinned.IsChecked == true)
                    _experiment.SetBinnedSensorRegion(_binW, _binH);
            }
        }

        // ── Custom ROI text-changed ────────────────────────────────────────────

        private void RoiX_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(RoiX.Text)) return;
            if (int.TryParse(RoiX.Text, out int v)) { _roiX = v; ApplyCustomRoi(); }
        }
        private void RoiY_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(RoiY.Text)) return;
            if (int.TryParse(RoiY.Text, out int v)) { _roiY = v; ApplyCustomRoi(); }
        }
        private void RoiWidth_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(RoiWidth.Text)) return;
            if (int.TryParse(RoiWidth.Text, out int v) && v > 0) { _roiW = v; ApplyCustomRoi(); }
        }
        private void RoiHeight_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(RoiHeight.Text)) return;
            if (int.TryParse(RoiHeight.Text, out int v) && v > 0) { _roiH = v; ApplyCustomRoi(); }
        }
        private void RoiXBin_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(RoiXBin.Text)) return;
            if (int.TryParse(RoiXBin.Text, out int v) && v > 0) { _roiXBin = v; ApplyCustomRoi(); }
        }
        private void RoiYBin_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!_controlsReady || string.IsNullOrEmpty(RoiYBin.Text)) return;
            if (int.TryParse(RoiYBin.Text, out int v) && v > 0) { _roiYBin = v; ApplyCustomRoi(); }
        }

        private void ApplyCustomRoi()
        {
            if (!CameraExists || CustomRegion.IsChecked != true) return;
            var roi     = new RegionOfInterest(_roiX, _roiY, _roiW, _roiH, _roiXBin, _roiYBin);
            RegionOfInterest[] rois = { roi };
            _experiment.SetCustomRegions(rois);
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  SECTION B  –  Acquire + Save
        // ═══════════════════════════════════════════════════════════════════════

        private void AcquireButton_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateAcquisition()) return;

            // Parse exposure and frame count
            if (!double.TryParse(ExposureMs.Text, out double expMs) || expMs <= 0)
            {
                MessageBox.Show("Please enter a valid exposure time (ms).");
                return;
            }
            if (!int.TryParse(FramesBox.Text, out int frames) || frames < 1)
            {
                MessageBox.Show("Please enter a valid frame count (≥ 1).");
                return;
            }

            try
            {
                AcquireButton.IsEnabled = false;
                AcqStatusText.Text      = "Acquiring…";
                AcqStatusText.Foreground = Brushes.DarkGoldenrod;

                // Set exposure
                _experiment.SetValue(CameraSettings.ShutterTimingExposureTime, expMs);
                _experiment.SetValue(ExperimentSettings.AcquisitionFramesToStore, frames);

                // Capture synchronously
                IImageDataSet dataSet = _experiment.Capture(frames);

                // Dispose the previous dataset to free memory
                if (_lastDataSet != null)
                {
                    _lastDataSet.Dispose();
                    _lastDataSet = null;
                }
                _lastDataSet = dataSet;

                // Display in the LightField viewer
                IDisplay display = _app.DisplayManager;
                if (display != null)
                {
                    IDisplayViewer viewer = display.GetDisplay(DisplayLocation.ExperimentWorkspace, 0);
                    if (viewer != null)
                        viewer.Display("BAP Acquire", dataSet);
                }

                // Build a title for the list box
                string roiDesc  = GetRoiDescription();
                string title    = $"{DateTime.Now:HH:mm:ss}  |  {frames} frame(s)  |  {roiDesc}";
                SamplesListBox.Items.Insert(0, title);   // newest on top
                SamplesListBox.SelectedIndex = 0;

                SavePngButton.IsEnabled  = true;
                AcqStatusText.Text       = $"✓ Acquired {frames} frame(s) at {expMs} ms";
                AcqStatusText.Foreground  = Brushes.DarkGreen;
            }
            catch (Exception ex)
            {
                AcqStatusText.Text       = "✗ Acquire failed: " + ex.Message;
                AcqStatusText.Foreground  = Brushes.DarkRed;
            }
            finally
            {
                AcquireButton.IsEnabled = true;
            }
        }

        ///////////////////////////////////////////////////////////////////////////
        private void SavePng_Click(object sender, RoutedEventArgs e)
        {
            if (_lastDataSet == null)
            {
                MessageBox.Show("No image has been acquired yet.");
                return;
            }

            // Show save dialog
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title            = "Save Acquired Image as PNG",
                DefaultExt       = ".png",
                Filter           = "PNG Image (*.png)|*.png",
                FileName         = $"BAPAcquire_{DateTime.Now:yyyyMMdd_HHmmss}.png"
            };

            bool? result = dlg.ShowDialog();
            if (result != true) return;

            try
            {
                IImageData frame = _lastDataSet.GetFrame(0, 0);  // region 0, frame 0
                Array rawData    = frame.GetData();

                int width  = frame.Width;
                int height = frame.Height;

                // Build a 32-bpp BGRA WriteableBitmap from the raw 16-bit data
                WriteableBitmap bmp = new WriteableBitmap(
                    width, height, 96, 96,
                    PixelFormats.Bgr32, null);

                // Find max for auto-normalisation (avoids pure-black PNG)
                ushort maxVal = 1;
                for (int i = 0; i < rawData.Length; i++)
                {
                    ushort v = (ushort)rawData.GetValue(i);
                    if (v > maxVal) maxVal = v;
                }

                // Build a managed pixel buffer (stride-aligned) then copy into the bitmap
                int    stride    = bmp.BackBufferStride / 4;   // ints per row (≥ width)
                int[]  pixels    = new int[stride * height];   // must match BackBuffer size

                for (int row = 0; row < height; row++)
                {
                    for (int col = 0; col < width; col++)
                    {
                        ushort raw  = (ushort)rawData.GetValue(row * width + col);
                        byte   gray = (byte)(raw * 255 / maxVal);
                        pixels[row * stride + col] = (gray << 16) | (gray << 8) | gray;
                    }
                }

                bmp.Lock();
                Marshal.Copy(pixels, 0, bmp.BackBuffer, stride * height);
                bmp.AddDirtyRect(new Int32Rect(0, 0, width, height));
                bmp.Unlock();

                // Encode and save
                using (FileStream fs = new FileStream(dlg.FileName, FileMode.Create))
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(bmp));
                    encoder.Save(fs);
                }

                AcqStatusText.Text       = $"✓ Saved → {Path.GetFileName(dlg.FileName)}";
                AcqStatusText.Foreground  = Brushes.DarkGreen;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save PNG:\n" + ex.Message);
            }
        }

        ///////////////////////////////////////////////////////////////////////////
        private void ClearList_Click(object sender, RoutedEventArgs e)
        {
            SamplesListBox.Items.Clear();
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  Helpers
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>Returns true when a camera device is present in the experiment.</summary>
        private bool CameraExists
        {
            get
            {
                foreach (IDevice device in _experiment.ExperimentDevices)
                    if (device.Type == DeviceType.Camera)
                        return true;
                return false;
            }
        }

        /// <summary>Validates camera presence and ready state before acquisition.</summary>
        private bool ValidateAcquisition()
        {
            if (!CameraExists)
            {
                MessageBox.Show("This add-in requires a camera in the experiment!");
                return false;
            }
            if (!_experiment.IsReadyToRun)
            {
                MessageBox.Show("The experiment is not ready to run. Check for errors in LightField.");
                return false;
            }
            return true;
        }

        /// <summary>Returns a short human-readable ROI description for the list box entry.</summary>
        private string GetRoiDescription()
        {
            if (FullChip.IsChecked == true)
                return "Full Chip";
            if (FullSensorBinned.IsChecked == true)
                return $"Binned {_binW}×{_binH}";
            return $"ROI ({_roiX},{_roiY}) {_roiW}×{_roiH} bin={_roiXBin}×{_roiYBin}";
        }
    }
}
