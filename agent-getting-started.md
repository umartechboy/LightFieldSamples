Here is a structured Markdown cheat sheet derived from the SDK manual, optimized for an AI coding agent to get up to speed quickly on LightField Add-in and Automation development.

---

# LightField SDK Development Guide

## 1. Project Setup & Architecture

LightField extensions are built on Microsoft's Managed Add-in Framework (MAF).

* 
**Target Framework:** .NET Framework 4.0.


* 
**Required Assembly References:** * `System.AddIn`.


* 
`PrincetonInstruments.LightFieldAddInSupportServices.dll` (Set `Copy Local = true`).


* 
`PrincetonInstruments.LightFieldViewV4.dll` (Set `Copy Local = false`).




* 
**Automation Reference:** If building a standalone executable instead of an in-app Add-in, also reference `PrincetonInstruments.LightField.AutomationV4.dll`.



## 2. Core Add-In Class Structure

Every Add-in must be decorated with the `[AddIn]` attribute and inherit from specific base classes so the LightField Add-in Manager can discover it.

```csharp
using System.AddIn;
using PrincetonInstruments.LightField.AddIns;

// 1. Attribute is strictly required for discovery
[AddIn("MyAddinTitle", Version = "1.0.0", Publisher = "Teledyne Princeton Instruments")]
// 2. Must derive from AddInBase and ILightFieldAddIn
public class MyAddIn : AddInBase, ILightFieldAddIn 
{
    private ILightFieldApplication _app;

    // 3. Required Property: Defines where the UI elements appear
    public UISupport UISupport { get { return UISupport.ApplicationToolBar; } }

    // 4. Required Method: Entry point when activated by the user
    public void Activate(ILightFieldApplication app) 
    {
        _app = app; // Cache the main application interface
    }

    // 5. Required Method: Cleanup when deactivated or application closes
    public void Deactivate() 
    {
        // Dispose of resources, unsubscribe from events, etc.
    }
}

```

## 3. UI Zones (`UISupport` Enum)

The `UISupport` property tells the application where to place the Add-in. It returns enumeration flags:

* 
`None = 0` 


* 
`Menu = 1` 


* 
`ApplicationToolBar = 2` (Supports Checkbox, Button, Button with Bitmap) 


* 
`DataToolBar = 4` 


* 
`ExperimentSetting = 8` (Supports Expander objects only) 


* 
`ExperimentView = 16` (Supports Dedicated Tab objects only) 



*Note: Flags can be combined using a logical OR to place the Add-in in multiple zones*.

## 4. Primary Interfaces (The Application Hierarchy)

The `ILightFieldApplication` interface is passed into the `Activate()` method and serves as the root object to access all subsystem managers.

### A. Hardware & Acquisition (`IExperiment`)

Retrieved via `ILightFieldApplication.Experiment`. This interface controls devices and initiates data acquisition.

* 
**Actions:** `Acquire()`, `Stop()`, `Capture(int frames)`.


* 
**Device Management:** `Add(IDevice device)`, `Remove(IDevice device)`, `Clear()`.


* 
**Settings (Get/Set):** Use `GetValue(string settingName)` and `SetValue(string settingName, object value)`. Setting names are constants defined in the `CameraSettings` and `ExperimentSettings` classes.


* 
**Events:** Subscribe to `ExperimentStarted`, `ExperimentCompleted`, `ImageDataSetReceived`, and `SettingChanged`.



### B. File IO (`IFileManager`)

Retrieved via `ILightFieldApplication.FileManager`. Used to handle LightField's native SPE files.

* 
**Methods:** `OpenFile()`, `CreateFile()`, `SaveFile()`, `CloseFile()`.


* 
**Exporting:** Supports converting data via `CreateExportSettings()` to formats like `Fits, Spc, Tiff, Csv, Avi`, followed by `Export()` or `ExportAsync()`.



### C. Data Construction (`IDataManager`)

Retrieved via `ILightFieldApplication.DataManager`. Used strictly to create standard LightField `IImageDataSet` objects from raw memory arrays.

* 
**Methods:** `CreateImageDataSet(System.Array Data, RegionOfInterest roi, ImageDataFormat Format)`.



### D. Views & Plotting (`IDisplay / IDisplayViewer`)

Retrieved via `ILightFieldApplication.DisplayManager`. Controls the UI data viewers.

* 
**Methods:** `Display(string sourceName, IImageDataSet imageDataSet)` plots datasets directly to the screen. Supports operations like `ZoomIn()`, `ZoomToBestFit()`, and `SetIntensityLevels()`.



## 5. Automation Initialization (Standalone App)

If creating a standalone C# application to drive LightField externally, do not implement `ILightFieldAddIn`. Instead, initialize the host process directly:

```csharp
// Constructor takes (bool isApplicationVisible, List<string> commandLineOptions)
var automation = new PrincetonInstruments.LightField.Automation.Automation(true, null);

// Retrieve the root application interface
ILightFieldApplication application = automation.LightFieldApplication;

```

(Reference )

---

Umar Hasan, this markdown covers the exact namespaces, required inheritances, and core object hierarchies the agent will need. You can feed this directly into Antigravity to skip the framework hallucination phase.