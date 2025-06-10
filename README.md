# **Battery Optimizer for Windows 11**  

![Power Management](https://img.shields.io/badge/Power-Management-blue) ![Windows 11](https://img.shields.io/badge/OS-Windows%2011-green) ![WiFi Control](https://img.shields.io/badge/Feature-WiFi%20Control-orange)  

A PowerShell script to **maximize battery life** on Windows 11 with **configurable WiFi control** and **automatic restoration**.  

---

## **📌 Quick Start**  
### **Basic Commands**  
Run in **PowerShell as Administrator**:  

```powershell
# Power-saving mode (WiFi OFF by default)
.\BatteryOptimizer.ps1 -Mode Optimize-Power

# Ultra battery saver (WiFi always OFF)
.\BatteryOptimizer.ps1 -Mode Optimize-Ultra

# Restore original settings
.\BatteryOptimizer.ps1 -Mode Revert
```

### **Keep WiFi ON (Power Mode Only)**  
```powershell
.\BatteryOptimizer.ps1 -Mode Optimize-Power -EnableWiFi
```

### **Auto-Revert After X Minutes**  
```powershell
.\BatteryOptimizer.ps1 -Mode Optimize-Ultra -RevertAfterMinutes 30  # Reverts in 30 mins
```

---

## **⚙️ Modes & Effects**  
| **Mode** | **Brightness** | **WiFi** | **Bluetooth** | **Extra Optimizations** |  
|----------|--------------|--------|------------|----------------------|  
| **Optimize-Power** | 50% | ❌ Disabled (unless `-EnableWiFi`) | ❌ Disabled | Display timeout (5 min), Background apps off |  
| **Optimize-Ultra** | 30% | ❌ Always Disabled | ❌ Disabled | + Search indexing off, SuperFetch disabled, Diagnostics tracking off |  
| **Revert** | Restores original | Restores original | Restores original | Full system restoration |  

---

## **📂 Files & Locations**  
| **File** | **Location** | **Purpose** |  
|---------|------------|-----------|  
| **Logs** | `.\Logs\BatteryOptimizer_<timestamp>.log` | Detailed execution logs |  
| **Backup** | `%APPDATA%\BatterySettingsBackup.json` | Original system settings |  
| **State** | `%APPDATA%\BatteryOptimizerState.txt` | Tracks current mode |  

---

## **🔧 Advanced Options**  
| **Parameter** | **Description** | **Example** |  
|-------------|---------------|------------|  
| `-Mode` | Optimization mode (`Optimize-Power`, `Optimize-Ultra`, `Revert`) | `-Mode Optimize-Ultra` |  
| `-EnableWiFi` | Keeps WiFi ON in **Power mode only** | `-Mode Optimize-Power -EnableWiFi` |  
| `-RevertAfterMinutes` | Auto-revert after X minutes | `-RevertAfterMinutes 60` |  

---

## **🚨 Troubleshooting**  
- **"Cannot switch modes directly"** → Run **`Revert` first** before changing modes.  
- **WiFi not disabling?** → Check adapter name in logs.  
- **Brightness not changing?** → Some monitors don’t support software control.  
- **Permission errors?** → **Always run as Administrator.**  

---

## **❓ FAQ**  
**Q: How do I check the current mode?**  
→ Look in `%APPDATA%\BatteryOptimizerState.txt`.  

**Q: Can I customize brightness levels?**  
→ Edit `$settings.PowerBrightness` and `$settings.UltraBrightness` in the script.  

**Q: Is this reversible?**  
→ **Yes!** `Revert` mode **fully restores** original settings.  

**Q: Why does Ultra mode force WiFi off?**  
→ WiFi consumes significant power—this ensures **maximum battery savings**.  

---

## **🛠️ Uninstall**  
1. Run:  
   ```powershell
   .\BatteryOptimizer.ps1 -Mode Revert
   ```  
2. Delete the script + logs (`.\Logs\`).  
3. (Optional) Remove `%APPDATA%\BatterySettingsBackup.json`.  

---

## **📜 License**  
**Free to use and modify.** For personal & professional use.  

🚀 **Tip:** Combine with Windows’ built-in **Battery Saver** for best results!  

--- 

**Enjoy longer battery life!** 🔋✨
