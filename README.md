# 🏔️ Nordic Nexus: Drawing Collector v3.1

<p align="center">
  <img src="https://img.shields.io/badge/PowerShell-5.1+-5391FE?style=for-the-badge&logo=powershell&logoColor=white" alt="PowerShell 5.1+"/>
  <img src="https://img.shields.io/badge/Epicor-v10/Kinetic-blue?style=for-the-badge" alt="Epicor ERP"/>
  <img src="https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?style=for-the-badge&logo=windows&logoColor=white" alt="Windows 10/11"/>
</p>

---

## ⚡ Overview

**Nordic Nexus** is an advanced engineering automation suite designed to streamline the lifecycle of drawing collection, indexing, and order processing. Built for high-volume engineering environments, it bridges the gap between **Epicor ERP**, **SolidWorks**, and **Outlook** to automate tasks that previously took hours.

### 🚀 Key Capabilities
- **Automated Epicor Integration:** Monitor and extract parts directly from Sales Orders and Job orders.
- **Smart Drawing Collection:** Instantly gather matching PDFs and DXFs from complex network hierarchies.
- **Revision Intelligence:** Sophisticated logic to always identify and prioritize the latest drawing revision.
- **Email Automation:** Watch for PDM notifications and automatically generate transmittal packages.
- **Deep BOM Expansion:** Traverses assembly structures to ensure 100% coverage of child components.

---

## 🛠️ Project Structure

| Component | Description |
|-----------|-------------|
| `EpicorOrderMonitor.ps1` | The heartbeat. Monitors Epicor for new Sales Orders. |
| `EmailOrderMonitor.ps1` | Watches Outlook for PDM and Engineering notifications. |
| `SimpleCollector.ps1` | Core engine for identifying and copying technical drawings. |
| `HubService.ps1` | Centralized orchestration for local and network services. |
| `PDFIndexManager.ps1` | GUI for managing and rebuilding the search index. |
| `Setup-EpicorCredentials.ps1` | Secure credential management for API access. |

---

## 🏁 Quick Start

### 1. Configure the Environment
Initialize your settings in `config.json`:
```json
{
    "indexFolder": "C:\\Data\\Index",
    "outputFolder": "C:\\Data\\Output",
    "epicorUrl": "https://epicor.yourdomain.com/Kinetic"
}
```

### 2. Setup Credentials
Run the setup script to securely store your Epicor API keys:
```powershell
.\Launch-Setup-EpicorCredentials.bat
```

### 3. Initialize the Index
Open the Index Manager to crawl your network drives:
```powershell
.\PDFIndexManager.ps1
```

---

## 📊 System Requirements

- **OS:** Windows 10 / 11
- **Runtime:** PowerShell 5.1+
- **API Access:** Epicor Kinetic REST API v2
- **Integration (Optional):** Outlook 2016+, SolidWorks 2018+

---

<p align="center">
  <strong>Nordic Minesteel Technologies</strong><br/>
  <em>Modernizing Engineering Workflows</em>
</p>
