# ResolutionSwitchTool
This tool is a shortcut that will allow to switch between two defined resolution for the main screen in one click.

## Requirements

* Windows with the primary display driven by an NVIDIA GPU (the script uses the
  Windows display API, so it works alongside the NVIDIA Control Panel).
* Python 3.10 or later available on the system path.

## Usage

1. Place `resolution_switch.py` somewhere convenient (for example in
   `C:\Tools\ResolutionSwitchTool`).
2. (Recommended) Run the bundled PowerShell script to create a shortcut
   automatically. The helper looks for Python through the Windows `py`
   launcher first so it works even when only the Microsoft Store aliases are
   on your `PATH`:

   ```powershell
   cd C:\Tools\ResolutionSwitchTool
   powershell -ExecutionPolicy Bypass -File .\create_resolution_switch_shortcut.ps1
   ```

   Run the command from the folder that contains both files (or pass
   `-ScriptPath` to point at the Python script explicitly). The shortcut is
   placed on your desktop as **Resolution Switch.lnk** and targets `pythonw.exe`
   when available so no console window appears. You can re-run the script with
   `-ShortcutPath` or `-PythonExecutable` parameters to customize the shortcut
   location or the interpreter. If PowerShell reports that it cannot create the
   shortcut, make sure you are running on Windows (the script relies on the
   Windows Script Host COM API).
3. Pin the created shortcut to the taskbar for one-click toggling.

The script hides its console window automatically so it can run silently when
triggered from the taskbar. If you need to debug it from a terminal, add the
`--show-console` flag to keep the window visible.

Running the script without arguments toggles between:

* **2560 × 1440** (native)
* **1680 × 1050**

For each resolution, the refresh rate is automatically set to the maximum value
reported by Windows for that mode. To reduce the amount of flicker during a
switch, the tool reuses existing Windows display modes and skips the change if
the display is already set to the requested values.

You can override the refresh rate when needed:

```powershell
python resolution_switch.py --refresh 244
```

You can also explicitly set a resolution instead of toggling by using
`--set WIDTHxHEIGHT`:

```powershell
python resolution_switch.py --set 2560x1440
```

If you want to combine both options, the refresh rate is applied to the
explicit resolution:

```powershell
python resolution_switch.py --set 1680x1050 --refresh 244
```

To troubleshoot or watch the status messages when running from a terminal, use
`--show-console`:

```powershell
python resolution_switch.py --show-console
```
