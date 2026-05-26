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

## 🚀 Getting Started

### Prerequisites

- [.NET 6.0 SDK](https://dotnet.microsoft.com/download) or later
- Visual Studio 2022 or VS Code
- A valid **Syncfusion License Key** (or use the free Community License)

### 1. Clone the Repository

```bash
git clone https://github.com/SyncfusionExamples/DocIO-Healthcare-Patient-Reports
cd Patient-Report-Creation
```

### 2. Install Dependencies

Restore all NuGet packages:

```bash
dotnet restore
```

### 3. Add Syncfusion License Key

In your `Program.cs`, register your Syncfusion license:

```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("YOUR_LICENSE_KEY");
```

### 4. Run the Application

```bash
dotnet run
```

Open your browser and navigate to:
```
https://localhost:5001
```

---

## 📋 How to Use

### Basic Workflow

1. **Upload Report Template (Optional)**
   - Upload a Word document (.docx, .doc, .rtf, .html) with mail merge fields
   - Or use the default `Template.docx` provided
   - Drag & drop supported

2. **Upload Patient Data (Optional)**
   - Upload an XML file with patient information
   - Or use the default `PatientDetails.xml` provided

3. **Select Report Type**
   - **All Records** – Generate reports for all patients in XML
   - **Selective Report** – Enter Patient ID to generate report for specific patient

4. **Choose Output Format**
   - **Single PDF Document** – Consolidated report for all records
   - **Multiple PDF Files** – Separate PDFs per patient (downloaded as ZIP)

5. **Generate Patient Report**
   - Click "Generate Patient Report" button
   - Download automatically starts
   - File format: `GeneratedDocument.pdf` or `converted_pdfs.zip`

---

## 🎯 Use Cases

### Healthcare Scenarios

- **Patient Discharge Summaries** – Generate standardized discharge reports with patient demographics and treatment details
- **Medical History Reports** – Compile comprehensive patient history documents with diagnoses and medications
- **Lab Result Reports** – Create formatted laboratory test result documents with multiple test records
- **Batch Report Generation** – Process multiple patient reports simultaneously for administrative purposes

---

## 🔗 Resources

- [Syncfusion DocIO Getting Started](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started)
- [Simple Mail Merge](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/simple-mail-merge)
- [Mail Merge for Groups](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/mail-merge-for-group)
- [Mail Merge for Nested Groups](https://help.syncfusion.com/document-processing/word/word-library/net/mail-merge/mail-merge-for-nested-groups)
- [Word to PDF Conversion](https://help.syncfusion.com/document-processing/word/conversions/word-to-pdf/overview)

---

## 📣 Try It Out

Clone the repository, run the sample, and discover how **Syncfusion DocIO** can streamline medical document workflows in healthcare environments.

### Customization

This sample application is provided as a reference implementation and can be freely customized to suit your specific healthcare requirements.

You can modify the templates, data sources, processing logic, and output formats based on your use case. If you have any questions, need clarification, or require assistance while customizing this sample, please feel free to contact our [Syncfusion Support Team](https://support.syncfusion.com/support/tickets/create) for guidance.

---

## ⚠️ Limitations

- XML element names must match Word template merge field names for successful data binding
- For repeating XML elements, the template group name must match the repeating tag name, not its parent wrapper (e.g., if `<Component>` repeats inside `<Components>`, use `«GroupStart:Component»` in template)
- Highly complex or unusual XML structures may require template and data alignment adjustments

---

## 📄 License and Copyright

> This is a commercial product and requires a paid license for possession or use. Syncfusion® licensed software, including this component, is subject to the terms and conditions of Syncfusion®. To acquire a license, visit https://www.syncfusion.com/account/downloads.

Are you already a Syncfusion user? You can download the product setup [here](https://www.syncfusion.com/account/downloads). If you're not yet a Syncfusion user, you can download a [30-day free trial](https://www.syncfusion.com/downloads).

---

## 📞 Support

For technical support and questions:
- [Syncfusion Support Portal](https://support.syncfusion.com/support/tickets/create)
- [Documentation](https://help.syncfusion.com/)
- [Community Forums](https://www.syncfusion.com/forums)

---