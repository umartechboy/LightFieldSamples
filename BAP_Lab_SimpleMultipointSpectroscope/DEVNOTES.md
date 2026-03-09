# BAP_Lab_SimpleMultipointSpectroscope — Developer Notes

## Project Info
- **Namespace:** `LightFieldAddInSamples.BAP_Lab_SimpleMultipointSpectroscope`
- **Assembly:** `LightFieldCSharpAddInSamples`
- **Target framework:** .NET Framework 4.8, x64

---

## Critical Build Rules

### 1. System.IO.Ports needs an explicit reference
In .NET Framework 4.8 projects, `SerialPort` lives in a **separate** `System.IO.Ports.dll`.
It is NOT included through the default `System` reference.
Add this to the `.csproj` `<ItemGroup>` references:
```xml
<Reference Include="System.IO.Ports" />
```

### 2. InitializeComponent() must come FIRST
Any method that touches XAML-named controls (e.g., `PortCombo`, `SerialStatusText`)
**must be called after** `InitializeComponent()`. Calling them before causes
`NullReferenceException` at load time.

```csharp
// CORRECT
InitializeComponent();
RefreshPorts();  // touches PortCombo -- OK here

// WRONG
RefreshPorts();  // PortCombo is null -- crash!
InitializeComponent();
```

### 3. WriteableBitmap / Marshal.Copy buffer must be stride-aligned
`WriteableBitmap.BackBufferStride` is row width in **bytes**, and may be larger
than `width * bytesPerPixel` due to 4-byte alignment.
Always allocate the pixel buffer as `stride * height`, not `width * height`:

```csharp
int stride = bmp.BackBufferStride / 4;       // stride in int32s
int[] pixels = new int[stride * height];     // NOT width * height
// ... fill pixels ...
bmp.Lock();
Marshal.Copy(pixels, 0, bmp.BackBuffer, stride * height);
bmp.AddDirtyRect(new Int32Rect(0, 0, width, height));
bmp.Unlock();
```

### 4. XAML must be pure ASCII / UTF-8 with no emoji or special Unicode
Emoji characters (`💾`, `▶`) and Unicode box-drawing chars (`═`, `–`) inside XAML
`Content="..."` attributes can cause XAML parser errors or encoding issues in
some VS configurations. Use plain ASCII text for button labels etc.

### 5. QualificationData("IsSample","True") filters the add-in
Princeton's LightField Add-in Manager shows a "Samples" category for add-ins
decorated with `[QualificationData("IsSample","True")]`. Omitting this tag
causes the add-in to appear at the top level of the list — useful for production add-ins.

---

## SDK Key APIs

| Task | Code |
|---|---|
| Get experiment | `IExperiment exp = app.Experiment;` |
| Check camera present | iterate `exp.ExperimentDevices`, check `device.Type == DeviceType.Camera` |
| Check ready | `exp.IsReadyToRun` |
| Set full frame | `exp.SetFullSensorRegion()` |
| Set binned | `exp.SetBinnedSensorRegion(xBin, yBin)` |
| Set custom ROI | `exp.SetCustomRegions(new RegionOfInterest[]{roi})` |
| Set exposure | `exp.SetValue(CameraSettings.ShutterTimingExposureTime, ms)` |
| Set frames | `exp.SetValue(ExperimentSettings.AcquisitionFramesToStore, n)` |
| Capture (sync) | `IImageDataSet ds = exp.Capture(frames)` |
| Get frame data | `IImageData f = ds.GetFrame(regionIdx, frameIdx)` then `Array raw = f.GetData()` |
| Display image | `app.DisplayManager.GetDisplay(DisplayLocation.ExperimentWorkspace, 0).Display("name", ds)` |
| Dispose dataset | `ds.Dispose()` |

## Deployment
Build output goes to:
`$(LIGHTFIELD_ROOT)\AddIns\LightFieldCSharpAddInSamples\`

LightField must be restarted (or add-in reloaded) to pick up new builds.
