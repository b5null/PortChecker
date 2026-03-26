# PortChecker
A lightweight Windows utility that scans for TCP ports that are both allowed through the Windows Firewall and not currently in use by any active listener.

---

# What it does
CheckPort queries the Windows Firewall policy via COM and cross-references it against active TCP listeners to find available ports. It checks for ports allowed for specific process identities (e.g. SYSTEM, ANY) and returns the first available one.

---

# Usage
```
PortChecker.exe
```

No arguments required. Output is printed to the console.

## Example output:
```
[*] Looking for available ports..
[*] SYSTEM Is allowed through port 10
If no port is found:
[*] Looking for available ports..
[-] No available ports found
[-] Firewall will block our COM connection. Exiting
```

---

# How it works
1. Retrieves all active TCP listeners using IPGlobalProperties
2. Instantiates the Windows Firewall Manager COM object (HNetCfg.FwMgr) via reflection — no external dependencies needed
3. Checks if the firewall is enabled on the current profile
4. Iterates ports 10–65534, calling IsPortAllowed for each identity name
5. Returns the first port that is both firewall-allowed and not already bound

---

# Building
Requires .NET Framework 4.0+. Compile with:
```
cmdC:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /target:exe /out:PortChecker.exe Program.cs ^
  /reference:C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.dll ^
  /reference:C:\Windows\Microsoft.NET\Framework64\v4.0.30319\System.Core.dll
```

No NuGet packages or interop DLLs required — firewall access is done via late-bound COM reflection.

---

# Requirements

- Windows (any version with .NET Framework 4.0+)
- No admin privileges required to run
- No external dependencies

---

# Notes
- Port scanning starts at 10 and stops at 65534
- Identity names checked: SYSTEM, ANY (in that order)
- If the firewall is disabled entirely, all ports are considered allowed

---

# ⚠️ Disclaimer

For educational and authorized testing only. Use only with explicit permission. The authors assume no liability for misuse.

---

# Author

- 💀 B5null
