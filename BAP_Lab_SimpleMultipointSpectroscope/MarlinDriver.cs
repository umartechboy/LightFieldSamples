using System;
using System.IO.Ports;
using System.Text;
using System.Threading;

namespace LightFieldAddInSamples.BAP_Lab_SimpleMultipointSpectroscope
{
    /// <summary>
    /// Lightweight Marlin G-code serial driver.
    /// Runs a background reader thread; raises LineReceived for every response line.
    /// All public methods are safe to call from any thread.
    /// </summary>
    public class MarlinDriver : IDisposable
    {
        private SerialPort _port;
        private Thread     _readerThread;
        private volatile bool _running;
        private readonly StringBuilder _lineBuffer = new StringBuilder();

        // ── Events ─────────────────────────────────────────────────────────────
        /// <summary>Raised on the reader thread for every complete line from Marlin.</summary>
        public event EventHandler<string> LineReceived;

        // ── State ──────────────────────────────────────────────────────────────
        public bool   IsConnected => _port != null && _port.IsOpen;
        public double CurrentX   { get; private set; }
        public double CurrentY   { get; private set; }
        public double CurrentZ   { get; private set; }

        // ═══════════════════════════════════════════════════════════════════════
        //  Connection
        // ═══════════════════════════════════════════════════════════════════════

        public void Connect(string portName, int baudRate)
        {
            if (IsConnected) Disconnect();

            _port = new SerialPort(portName, baudRate, Parity.None, 8, StopBits.One)
            {
                ReadTimeout  = 200,
                WriteTimeout = 500,
                NewLine      = "\n"
            };
            _port.Open();

            _running      = true;
            _readerThread = new Thread(ReadLoop)
            {
                IsBackground = true,
                Name         = "MarlinReaderThread"
            };
            _readerThread.Start();

            // Request firmware info so the terminal gets initial output
            SendRaw("M115");
        }

        public void Disconnect()
        {
            _running = false;
            try { _port?.Close(); } catch { /* ignore */ }
            _port?.Dispose();
            _port = null;
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  G-code commands
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>Send any raw G-code string. Appends a newline automatically.</summary>
        public void SendRaw(string gcode)
        {
            if (!IsConnected) return;
            try { _port.WriteLine(gcode.TrimEnd()); }
            catch { /* swallow write errors; reader loop will detect disconnect */ }
        }

        // ── Homing ─────────────────────────────────────────────────────────────

        public void Home()
        {
            SendRaw("G28");
            CurrentX = 0; CurrentY = 0; CurrentZ = 0;
        }

        public void HomeXY()
        {
            SendRaw("G28 X Y");
            CurrentX = 0; CurrentY = 0;
        }

        public void HomeZ()
        {
            SendRaw("G28 Z");
            CurrentZ = 0;
        }

        // ── Absolute movement ─────────────────────────────────────────────────

        public void MoveXY(double x, double y, int feedRate = 3000)
        {
            SendRaw(string.Format("G1 X{0:F3} Y{1:F3} F{2}", x, y, feedRate));
            CurrentX = x; CurrentY = y;
        }

        public void MoveXYZ(double x, double y, double z, int feedRate = 3000)
        {
            SendRaw(string.Format("G1 X{0:F3} Y{1:F3} Z{2:F3} F{3}", x, y, z, feedRate));
            CurrentX = x; CurrentY = y; CurrentZ = z;
        }

        public void MoveZ(double z, int feedRate = 300)
        {
            SendRaw(string.Format("G1 Z{0:F3} F{1}", z, feedRate));
            CurrentZ = z;
        }

        // ── Relative movement ─────────────────────────────────────────────────

        public void MoveRelativeZ(double deltaZ, int feedRate = 300)
        {
            SendRaw("G91");
            SendRaw(string.Format("G1 Z{0:F3} F{1}", deltaZ, feedRate));
            SendRaw("G90");
            CurrentZ += deltaZ;
        }

        public void MoveRelativeX(double deltaX, int feedRate = 3000)
        {
            SendRaw("G91");
            SendRaw(string.Format("G1 X{0:F3} F{1}", deltaX, feedRate));
            SendRaw("G90");
            CurrentX += deltaX;
        }

        public void MoveRelativeY(double deltaY, int feedRate = 3000)
        {
            SendRaw("G91");
            SendRaw(string.Format("G1 Y{0:F3} F{1}", deltaY, feedRate));
            SendRaw("G90");
            CurrentY += deltaY;
        }

        // ── Coordinate management ─────────────────────────────────────────────

        /// <summary>Set current position as origin (G92 X0 Y0 Z0).</summary>
        public void SetOrigin()
        {
            SendRaw("G92 X0 Y0 Z0");
            CurrentX = 0; CurrentY = 0; CurrentZ = 0;
        }

        /// <summary>Save settings to EEPROM (M500).</summary>
        public void SaveToEEPROM()
        {
            SendRaw("M500");
        }

        /// <summary>Wait for all queued moves to finish (M400).</summary>
        public void WaitForMotion()
        {
            SendRaw("M400");
        }

        /// <summary>Dwell / pause in place for given milliseconds (G4).</summary>
        public void Dwell(int milliseconds)
        {
            SendRaw(string.Format("G4 P{0}", milliseconds));
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  Background reader
        // ═══════════════════════════════════════════════════════════════════════

        private void ReadLoop()
        {
            while (_running)
            {
                try
                {
                    int b = _port.ReadByte();
                    if (b < 0) continue;

                    char c = (char)b;
                    if (c == '\n' || c == '\r')
                    {
                        string line = _lineBuffer.ToString().Trim();
                        _lineBuffer.Clear();
                        if (line.Length > 0)
                            LineReceived?.Invoke(this, line);
                    }
                    else
                    {
                        _lineBuffer.Append(c);
                    }
                }
                catch (TimeoutException) { /* normal between characters */ }
                catch { break; /* port closed or error -- exit thread */ }
            }
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  IDisposable
        // ═══════════════════════════════════════════════════════════════════════

        public void Dispose() => Disconnect();
    }
}
