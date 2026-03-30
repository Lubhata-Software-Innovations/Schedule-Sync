# 🔄 Lubhata Schedule Sync  
### Professional Revit Data Management Bridge

**Lubhata Schedule Sync** is a high-performance **pyRevit** tool designed to bridge the gap between **Autodesk Revit** and **Microsoft Excel**.

It enables BIM Managers and Architects to:
- Export Revit schedules to CSV  
- Perform bulk edits in Excel  
- Sync updates back into Revit with intelligent validation  

---

## 📺 Demo Video  
See the tool in action and understand the workflow:  

👉 [Watch the Demo on LinkedIn](https://www.linkedin.com/posts/lubhata-software-and-innovations_bim-revit-aec-activity-7419240343545630721-vsur?utm_source=share&utm_medium=member_desktop&rcm=ACoAAFSq0noBBacmECC4aZQkcQaaSbdyZQkD_Ho)

---

## ✨ Key Features  

### 🧠 Smart Export  
- Automatically captures `ElementId` (even if hidden in the schedule)  
- Ensures a reliable and bulletproof data link  

### 🔁 Bi-Directional Sync  
- Edit text, numbers, and parameters in Excel  
- Push updates back into Revit in seconds  

### ✅ Data Validation  
- Handles Revit storage types intelligently:  
  - String  
  - Integer  
  - Double  
- Prevents sync errors  

### 🎨 Modern UI/UX  
- Clean, dark-themed WPF interface  
- Fully integrated into the Revit environment  

### 🔒 Read-Only Protection  
- Skips calculated values  
- Ignores read-only parameters  
- Maintains database integrity  

---

## 🚀 Installation  

### Option 1: Automatic Setup (Recommended)  
1. Go to the **Releases** page  
2. Download `Lubhata_Schedule_Sync.exe`  
3. Run the installer  
4. It will automatically detect your pyRevit installation  
5. Restart Revit or reload pyRevit  

---

### Option 2: Manual Installation  
1. Download the source code  
2. Extract the folder  
3. Copy the `Lubhata.extension` folder  
4. Paste it into your pyRevit extensions directory:  
   `%AppData%\pyRevit\Extensions`

5. Reload pyRevit  

---

## 📖 How to Use  

### 📤 Export  
1. Open **Lubhata Schedule Sync**  
2. Select a schedule  
3. Click **EXPORT TO CSV**  

---

### ✏️ Edit  
1. Open the CSV in Excel  
2. ⚠️ **Important:**  
   - Do NOT modify the `ElementId` column  
   - Do NOT change header names  
3. Edit other parameter values as needed  

---

### 📥 Sync  
1. Switch to the **IMPORT** tab  
2. Select your edited CSV file  
3. Click **SYNC DATA TO REVIT**  

---

## 🛠 Developer Documentation  

### 🧩 Code Architecture & Logic Flow  

Built using **Python (IronPython)** with **pyRevit** and the **Revit API**.

---

### 🖥 UI Layer (WPFWindow)  
- Embedded XAML (zero external dependencies)  
- Handles button events (Export, Import, Links)  

---

### 📊 Data Extraction (`run_smart_export`)  
- Iterates through `ViewSchedule` fields  
- Uses `FilteredElementCollector` scoped to View ID  
- Exports only scheduled elements  
- Supports Revit 2024+ using `get_safe_id` helper  

---

### 🔄 Data Injection (`run_smart_import`)  
- Uses `csv.DictReader` for header-based mapping  

#### ⚙️ Transaction Management  
- All updates wrapped in a single `revit.Transaction`  
- Ensures performance + undo support  

#### 🔢 Type Casting  
- Detects `param.StorageType`  
- Uses:
  - `SetValueString()` for doubles (unit-aware)  
  - `Set()` for strings & integers  

---

## ⚙️ Environment Requirements  

| Component | Version |
|----------|--------|
| pyRevit  | 4.8.10+ |
| Revit    | 2019 – 2025+ |
| OS       | Windows 10 / 11 |

---

## ⚖️ Legal & License  

**Copyright (c) 2026 Lubhata Software & Innovations**

Permission is granted for:
- Personal use  
- Internal organizational use  

🚫 **Restrictions:**  
- Redistribution  
- Modification for redistribution  
- Sublicensing  
- Resale  
- Source code disclosure  

---

## 📬 Commercial Use  

For commercial inquiries or custom BIM development:

🌐 Visit: **Lubhata.com**  
📧 Email: **info@lubhata.com**  

---

## ❤️ Built With Passion  

Built with ❤️ by **Lubhata Software & Innovations**
