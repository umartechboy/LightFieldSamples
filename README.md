# LightField C# Add-ins — BAP Lab

A collection of Princeton Instruments LightField add-ins, including the
**BAP Lab SimpleMultipointSpectroscope** — a motorised XY-stage scanning add-in
for automated multi-point Raman / fluorescence spectroscopy.

---

## Contents

```
CSharp Add-Ins/
├── BAP_Lab_SimpleMultipointSpectroscope/   ← main add-in (stage + scan + export)
│   ├── SimpleMultipointSpectroscopeAddin.cs
│   ├── SimpleMultipointSpectroscopeControl.xaml / .xaml.cs
│   └── MarlinDriver.cs
├── LightFieldCSharpAddInSamples.sln        ← open this in Visual Studio
├── LightFieldCSharpAddInSamples.csproj
└── <Princeton Instruments sample add-ins>
```

---

## Prerequisites

| Requirement | Version | Notes |
|-------------|---------|-------|
| Windows | 10 / 11 | 64-bit |
| [Visual Studio](https://visualstudio.microsoft.com/) | 2019 or 2022 | Community edition is free |
| .NET Framework | 4.8 | Included with Windows 10+ |
| Princeton Instruments LightField | 6.x | Must be installed; provides the reference DLLs |
| [Antigravity](https://antigravity.dev) | latest | AI coding assistant used for development |

---

## Development Setup

### 1 — Install Visual Studio

1. Download **Visual Studio Community** from <https://visualstudio.microsoft.com/>.
2. During installation select the workload:
   - ✅ **.NET desktop development**
3. Ensure **.NET Framework 4.8 targeting pack** is included (it is by default).

### 2 — Install Antigravity

Antigravity is the AI coding assistant used to edit this project via chat.
Code changes are made through the Antigravity chat interface; Visual Studio
is used in parallel purely to **build and test**.

1. Go to <https://antigravity.dev> and sign in with your Google account.
2. Download and install the Antigravity desktop app for Windows.
3. Open the app and connect it to this repository folder.
4. Open any file through the Antigravity chat to start editing.

> **Workflow:** describe a change in Antigravity chat → Antigravity edits the
> source files → switch to Visual Studio → press **Ctrl+Shift+B** to build →
> test in LightField.

### 3 — Clone the Repository

```bash
git clone https://github.com/umartechboy/LightFieldSamples.git
cd LightFieldSamples
```

Or use **GitHub Desktop** → **File → Clone repository** and paste the URL.

### 4 — Set the `LIGHTFIELD_ROOT` Environment Variable

The project resolves the LightField SDK DLLs through an environment variable.
Set it once and it works for all future builds.

1. Open **Start → Edit the system environment variables**.
2. Under **System variables**, click **New**:
   - **Variable name:** `LIGHTFIELD_ROOT`
   - **Variable value:** path to your LightField installation, e.g.
     ```
     C:\Program Files\Princeton Instruments\LightField
     ```
3. Click **OK** and **restart Visual Studio**.

> If `LIGHTFIELD_ROOT` is not set, the build falls back to a local `bin\`
> folder (the project will still open but DLL references will be unresolved —
> you will see red squiggles).

### 5 — Build in Visual Studio

1. Open `LightFieldCSharpAddInSamples.sln`.
2. Set the build target to **x64** (the platform dropdown in the toolbar).
3. Press **Ctrl+Shift+B** (Build Solution).

On a successful build, the output DLL is written directly into:
```
%LIGHTFIELD_ROOT%\AddIns\LightFieldCSharpAddInSamples\
```

LightField will discover it automatically on the next launch.

---

## Enabling the Add-in in LightField

1. Launch **LightField**.
2. Click **Add-ins** in the top menu.
3. Tick **BAP Lab – SimpleMultipointSpectroscope**.
4. Click **OK** — the panel appears in the workspace.

---

## Usage

See [`BAP_Lab_SimpleMultipointSpectroscope/README.md`](BAP_Lab_SimpleMultipointSpectroscope/README.md)
for the full user guide covering:
Connect → Physical Home → Set Center → Set Grid → Scan → Export Data.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Red squiggles in VS on LightField types | `LIGHTFIELD_ROOT` not set or LightField not installed |
| Add-in not listed in LightField | Build succeeded but targeting **x86** instead of **x64** |
| DLL not found after build | Check Output Path in project properties matches the AddIns folder |
| Marlin not connecting | Wrong COM port or baud rate; try 115200 first |

---

*BAP Lab — University of Engineering and Technology*
