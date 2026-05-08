# Syncfusion ASP.NET Core – Patient Report Creation Demo

This repository contains a complete showcase sample demonstrating how to build an **Automated Patient Report Generator** using **Syncfusion DocIO** and **Syncfusion PDF** libraries in an ASP.NET Core MVC application. The sample illustrates how healthcare professionals can automate medical report generation by merging dynamic XML data into Word templates and exporting them as professional PDF documents.

---

## 📁 Project Structure

```
├── Controllers/
│   └── HomeController.cs
├── Models/
│   └── ErrorViewModel.cs
├── Views/
│   ├── Home/
│   │   ├── Index.cshtml
│   │   └── Privacy.cshtml
│   └── Shared/
├── wwwroot/
│   └── Data/
│       ├── Template.docx
│       └── PatientDetails.xml
└── README.md
```

---

## ✨ Features

- Upload a **Word template** (.docx, .doc, .rtf) with mail merge fields as placeholders for patient information.
- Upload an **XML data file** containing patient records, medical history, diagnoses, and treatment details.
- Automatically detect XML structure and apply the appropriate mail merge strategy:
  - **Simple/Flat Merge** – for plain key-value XML data (single patient).
  - **Group Merge** – for XML containing multiple patient records.
  - **Nested Group Merge** – for hierarchical XML with nested medical data (e.g., patient → diagnoses → medications).
- Generate a **single merged PDF** for consolidated patient reports.
- Generate **multiple PDFs** (split by page breaks) bundled as a **ZIP file** for batch patient reports.
- Default template and sample data files included for quick testing.
- Real-time document processing with immediate file cleanup (no data storage).

## 🚀 Getting Started

### Prerequisites

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

---

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Healthcare-Patient-Reports
```

---

### 2. Navigate to the Project Directory

```bash
cd Patient-Report-Creation
```

---

### 3. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```
---

### 4. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```
---

### 5. Run the Application

```bash
dotnet run
```

---

### 6. Open in Browser

Navigate to the localhost URL shown in the terminal output, for example:

```
https://localhost:5001
```

---

## 📋 How to Use

1. Open the application in your browser.
2. **Upload Word Template** (optional):
   - Click the "Report Template" card
   - Select a Word document (.docx, .doc, .rtf, .dotx) containing mail merge fields (e.g., `«PatientName»`, `«Diagnosis»`)
   - Or use the default `Template.docx` provided
3. **Upload Patient Data** (optional):
   - Click the "Patient Data" card
   - Select an XML file with patient information
   - Or use the default `PatientDetails.xml` provided
4. **Choose Output Format**:
   - **Single PDF Document** – Merges all patient records into one consolidated PDF
   - **Multiple PDF Files** – Creates separate PDFs for each patient (downloaded as ZIP)
5. Click **Generate Patient Report** to download the result.
6. View the generated PDF(s) containing professional medical reports.

---

## 📄 XML Data Format Examples

### Simple XML (Single Patient)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Patients>
    <Patient>
        <PatientName>John Smith</PatientName>
        <PatientID>P12345</PatientID>
        <Age>45</Age>
        <Gender>Male</Gender>
        <BloodType>O+</BloodType>
        <Diagnosis>Type 2 Diabetes</Diagnosis>
        <Physician>Dr. Sarah Johnson</Physician>
        <AdmitDate>2026-03-15</AdmitDate>
    </Patient>
</Patients>
```

### Group XML (Multiple Patients)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Patients>
    <Patient>
        <PatientName>John Smith</PatientName>
        <PatientID>P12345</PatientID>
        <Diagnosis>Type 2 Diabetes</Diagnosis>
    </Patient>
    <Patient>
        <PatientName>Emily Davis</PatientName>
        <PatientID>P12346</PatientID>
        <Diagnosis>Hypertension</Diagnosis>
    </Patient>
</Patients>
```

### Nested Group XML (Hierarchical Medical Data)

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Patients>
    <Patient>
        <PatientName>John Smith</PatientName>
        <PatientID>P12345</PatientID>
          <Diagnosis>
            <Condition>Type 2 Diabetes</Condition>
            <Severity>Moderate</Severity>
              <Medication>
                <Name>Metformin</Name>
                <Dosage>500mg</Dosage>
                <Frequency>Twice daily</Frequency>
              </Medication>
          </Diagnosis>
          <VitalSign>
            <Type>Blood Pressure</Type>
            <Value>140/90 mmHg</Value>
          </VitalSign>
    </Patient>
    <Patient>
        <PatientName>Sam</PatientName>
        <PatientID>P12346</PatientID>
          <Diagnosis>
            <Condition>Type 1 Diabetes</Condition>
            <Severity>Moderate</Severity>
              <Medication>
                <Name>Metformin</Name>
                <Dosage>200mg</Dosage>
                <Frequency>Twice daily</Frequency>
              </Medication>
          </Diagnosis>
          <VitalSign>
            <Type>Blood Pressure</Type>
            <Value>130/90 mmHg</Value>
          </VitalSign>
    </Patient>
</Patients>
```

---

## 🏥 Mail Merge Field Examples

In your Word template, use these merge field formats:

### Simple Fields
- `«PatientName»`
- `«PatientID»`
- `«Age»`
- `«Diagnosis»`

### Group Fields (for repeating sections)
```
«TableStart:Patient»
Patient Name: «PatientName»
Patient ID: «PatientID»
«TableEnd:Patient»
```

### Nested Group Fields
```
«TableStart:Patient»
Patient: «PatientName»
  «TableStart:Diagnoses»
  Diagnosis: «Condition»
    «TableStart:Medications»
    - Medication: «Name» (Dosage: «Dosage»)
    «TableEnd:Medications»
  «TableEnd:Diagnoses»
«TableEnd:Patient»
```

---

## 🔧 Technical Implementation

### Mail Merge Detection Logic

The application automatically detects the XML structure and chooses the appropriate merge strategy:

1. **Nested Group Merge** – Used when XML contains hierarchical data with multiple levels of nested arrays
2. **Group Merge** – Used when XML contains multiple records at the same level
3. **Simple Merge** – Used for flat key-value XML data (single record)

### Page Break Splitting

When "Multiple PDF Files" is selected, the application:
1. Inserts bookmarks at each page break location
2. Extracts content between bookmarks
3. Converts each section to a separate PDF
4. Bundles all PDFs into a ZIP archive

### Data Processing Flow

```
Upload Files → Load Streams → Parse XML → 
Detect Structure → Execute Mail Merge → 
Convert to PDF → Split (if multiple) → 
Return File → Cleanup Memory
```

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Simple Mail Merge](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/simple-mail-merge)
- [Mail Merge for Groups](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/mail-merge-for-group)
- [Mail Merge for Nested Groups](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/mail-merge-for-nested-groups)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)

---

## ✅ Benefits

- **Eliminates Manual Work** – Automates repetitive medical report creation tasks
- **Batch Processing** – Handle multiple patient reports simultaneously
- **Complex Data Support** – Handles hierarchical medical data with nested groups
- **Professional Output** – Generates print-ready, standardized PDF reports
- **Template Flexibility** – Easily customize Word templates without code changes
- **Fast Processing** – Real-time document generation and conversion

---

## 📣 Try It Out

Clone the repository, run the sample, and explore how **Syncfusion DocIO** can streamline medical document workflows in healthcare environments.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific business requirements.

You can modify the templates, data sources(XML,Json,mdb,db,etc.,), processing logic, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our Syncfusion Support Team for guidance.

---

## 🎯 Use Cases

- **Patient Discharge Summaries** – Generate standardized discharge reports
- **Medical History Reports** – Compile comprehensive patient history documents
- **Lab Result Reports** – Create formatted laboratory test result documents
- **Prescription Reports** – Generate medication lists and prescription summaries
- **Diagnostic Reports** – Produce radiology, pathology, and other diagnostic reports
- **Batch Report Generation** – Process multiple patient reports for administrative purposes

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads..

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)
