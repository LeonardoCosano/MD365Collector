# 🛡️ MD365Collector

**MD365Collector** is a **PowerShell** tool designed to collect and analyze security logs from **Microsoft Defender for Office 365**.  
It includes utilities to efficiently collect, search and parse logs, making security event analysis and management easier.

---

## 🎯 Objective

Provide a simple and automated solution to **collect, filter, and prepare logs from Microsoft Defender for Office 365**, focusing on incident analysis.

---

## ⚙️ Installation & Usage

### 1️⃣ Clone the repository
```bash
git clone https://github.com/leonardocosano/MD365Collector.git
cd MD365Collector
```

### 2️⃣ Import the module
```powershell
Import-Module .\MD365Collector.ps1
```

### 3️⃣ Check the environment
```powershell
IsEnvironmentReadyForMD365Collector
```

### 4️⃣ Set up the environment
```powershell
SetEnvironmentReadyForMD365Collector
```

### 5️⃣ Run the cmdlets
If you haven't run the tool before, I suggest start by full automatic collection:
```powershell
StartCollection -user pepito@acme.com -start 2025-11-26T00:00:00Z -end 2025-11-27T00:00:00Z
```

> ⚠️ **Authentication required:**  
> You must authenticate using a valid Microsoft 365 user account.  
> Application-based authentication is currently in the *roadmap*.  
> The authenticated user must have the necessary permissions to access security logs.

---

## 🧩 Features

### ✅ Current features
- Collection of **Audit Logs** from Microsoft Office 365.

### 🔜 Upcoming features
- Advanced Audit Log collection.  
- Audit Log parsing and processing.  
- Collection of **Cloud App Activity Logs**.  
- Parsing of Cloud App Activity logs.  
- Generation of statistics and reports for both log types.

---

## 💡 Usage examples
*(Section under construction — examples of cmdlet execution and common use cases will be added soon.)*

---

## 🧰 Requirements
- PowerShell 7.x or higher.  
- Valid access to a Microsoft 365 environment.  
- Proper permissions to access security logs.

---

## 🤝 Contributing

Contributions are **welcome** 🙌  
To contribute:
1. Fork the project.  
2. Create a new branch (`feature/new-feature`).  
3. Submit a pull request with your changes.  
4. Follow PowerShell best practices and style guidelines.

You can also open *issues* to report bugs or suggest improvements.

---

## 🔒 License
This project is licensed under the **MIT License**. 

---

## 👤 Author
Developed by **Leonardo Cosano**.

---

## 🧭 Project status
> 🚧 Actively under development — initial functional version.

---

## 🌟 Acknowledgements
Thanks to the Microsoft 365 security administrators and analysts community for their continuous knowledge sharing and real-world use cases.

---