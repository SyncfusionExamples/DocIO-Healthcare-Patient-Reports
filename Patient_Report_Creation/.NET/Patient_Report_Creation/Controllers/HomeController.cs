using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using Patient_Report_Creation.Models;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Pdf;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Dynamic;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Patient_Report_Creation.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _hostingEnvironment;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment hostingEnvironment)
        {
            _logger = logger;
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult GenerateReport(IFormFile file, IFormFile xmlFile, string outputFormat, string reportType, string patientId)
        {          
            try
            {
                // Load Word document stream (uploaded or default)
                Stream wordStream = GetWordDocument(file);

                if (wordStream == null)
                    return View("Index");

                Stream xmlStream = GetXmlDocument(file, xmlFile);
                if (xmlStream == null)
                    return View("Index");

                // If selective report, validate PatientID
                if (reportType == "selective")
                {
                    if (string.IsNullOrWhiteSpace(patientId))
                    {
                        ViewBag.Message = "Please enter a Patient ID for selective report";
                        return View("Index");
                    }

                    patientId = patientId.Trim();
                }

                return CreatePDF(wordStream, xmlStream, outputFormat, reportType, patientId);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating document");
                ViewBag.Message = $"Error: {ex.Message}";
                return View("Index");
            }
        }

        /// <summary>
        /// Retrieves a Word document stream from the uploaded file or a default template.
        /// </summary>

        private Stream GetWordDocument(IFormFile file)
        {
            // Case 1: Uploaded file exists and has content
            if (file != null && file.Length > 0)
            {
                string extension = Path.GetExtension(file.FileName).ToLower();
                string[] supportedExtensions = { ".doc", ".docx", ".dot", ".dotx", ".dotm", ".docm", ".xml", ".rtf" };
                // Validate the file extension
                if (supportedExtensions.Contains(extension))
                {
                    // Copy the uploaded file into an in-memory stream
                    MemoryStream stream = new MemoryStream();
                    file.CopyTo(stream);
                    // Reset stream position to the beginning for downstream reading
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please choose a Word format document to convert to PDF.";
                    return null;
                }
            }
            else
            {
                // Load default file from wwwroot\Data\
                string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Template.docx");
                using (var fileStream = new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read))
                {
                    var memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    memoryStream.Position = 0;
                    return memoryStream;
                }
            }
        }

        /// <summary>
        /// Creates a PDF document by performing a mail merge on a Word template using XML data.
        /// </summary>
        private IActionResult CreatePDF(Stream stream, Stream xmlStream, string outputFormat, string reportType, string patientId)
        {
            if (string.IsNullOrWhiteSpace(outputFormat))
            {
                ViewBag.Message = "Output format is required.";
                return View("Index");
            }
            using (WordDocument document = new WordDocument(stream, FormatType.Automatic))
            {
                if (xmlStream != null && xmlStream.Length > 0)
                {
                    bool mergeSuccess = false;
                    //Determine whether PatientID should be used
                    string effectivePatientId = reportType == "selective" ? patientId?.Trim() : null;
                    // Generate report for all records
                    mergeSuccess = ProcessXmlAndPerformMerge(document, xmlStream,patientId);
                    
                    if (!mergeSuccess)
                        return View("Index");
                    if (!string.IsNullOrWhiteSpace(effectivePatientId))
                    {
                        _logger.LogInformation($"Generated selective report for Patient ID: {effectivePatientId}");
                    }
                    else
                    {
                        _logger.LogInformation("Generated report for all records");
                    }
                }
                // Generate and return the output file based on requested format
                return GenerateOutputFile(document, outputFormat);
            }
        }      
        /// <summary>
        /// SIMPLIFIED: Processes XML data by parsing to ExpandoObject, then analyzes structure 
        /// to detect repeating groups and performs mail merge.
        /// No longer needs separate XML traversal for finding repeating groups!
        /// </summary>
        private bool ProcessXmlAndPerformMerge(WordDocument document, Stream xmlStream, string patientId)
        {
            try
            {
                xmlStream.Position = 0;
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlStream);

                XmlNode documentElement = xmlDocument.DocumentElement;

                // Convert XML to dynamic ExpandoObject structure (single iteration)
                ExpandoObject parsedData = new ExpandoObject();
                GetDataAsExpandoObject(documentElement, ref parsedData);

                // ✅ SIMPLIFIED: Analyze ExpandoObject structure directly to find repeating groups
                var groupInfo = AnalyzeExpandoObjectStructure(parsedData, documentElement.LocalName);

                // Execute appropriate mail merge based on detected structure
                if (groupInfo.HasRepeatingGroup)
                {
                    return ExecuteGroupMailMerge(document, groupInfo.GroupName, 
                                                groupInfo.GroupItems, patientId);
                }
                else
                {
                    return ExecuteSimpleMailMerge(document, documentElement, parsedData);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to process XML and perform mail merge");
                ViewBag.Message = "Error processing XML data.";
                return false;
            }
        }

        /// <summary>
        /// Analyzes ExpandoObject structure to detect repeating groups.
        /// Returns group information including name and items.
        /// </summary>
        private (bool HasRepeatingGroup, string GroupName, List<ExpandoObject> GroupItems) 
            AnalyzeExpandoObjectStructure(ExpandoObject parsedData, string rootName)
        {
            try
            {
                var rootDict = parsedData as IDictionary<string, object>;
                
                // Navigate to the root element content
                if (!rootDict.ContainsKey(rootName))
                {
                    _logger.LogWarning($"Root element '{rootName}' not found in parsed data");
                    return (false, null, null);
                }

                var rootValue = rootDict[rootName];
                if (!(rootValue is List<ExpandoObject> rootList) || rootList.Count == 0)
                {
                    _logger.LogWarning("Invalid root structure in parsed data");
                    return (false, null, null);
                }

                // Analyze the structure starting from root content
                var rootContent = rootList[0] as IDictionary<string, object>;
                var groupInfo = FindRepeatingGroupInExpando(rootContent, new List<string> { rootName });

                if (groupInfo.HasRepeatingGroup)
                {
                    _logger.LogInformation($"Detected repeating group '{groupInfo.GroupName}' with {groupInfo.GroupItems.Count} items at path: {string.Join(" → ", groupInfo.Path)}");

                    return (
                        groupInfo.HasRepeatingGroup,
                        groupInfo.GroupName,
                        groupInfo.GroupItems
                    );

                }

                _logger.LogInformation("No repeating groups detected - will use simple mail merge");
                return (false, null, null);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error analyzing ExpandoObject structure");
                return (false, null, null);
            }
        }

        /// <summary>
        /// Recursively searches ExpandoObject to find repeating groups.
        /// FIXED: Properly detects primary groups (e.g., Patients) before finding nested groups (e.g., Vitals)
        /// </summary>
        private (bool HasRepeatingGroup, string GroupName, List<ExpandoObject> GroupItems, List<string> Path) 
            FindRepeatingGroupInExpando(IDictionary<string, object> currentDict, List<string> currentPath, int depth = 0)
        {
            if (currentDict == null)
                return (false, null, null, null);

            // STEP 1: Check for TRUE repeating groups (multiple items at this level)
            foreach (var kvp in currentDict)
            {
                if (kvp.Value is List<ExpandoObject> list && list.Count > 1)
                {
                    // Found a list with multiple items - this is a repeating group!
                    var path = new List<string>(currentPath) { kvp.Key };
                    return (true, kvp.Key, list, path);
                }
            }

            // STEP 2: Check for PRIMARY group (single item that should be treated as main group)
            // CRITICAL: Detect this BEFORE recursing to avoid finding nested groups first
            // Example: Patients (1 item) with nested Vitals (multiple items) → Patients is the primary group
            foreach (var kvp in currentDict)
            {
                if (kvp.Value is List<ExpandoObject> list && list.Count == 1)
                {
                    var singleItem = list[0] as IDictionary<string, object>;
                    
                    if (singleItem != null && HasOwnScalarFields(singleItem) && IsPrimaryGroup(singleItem, kvp.Key, depth))
                    {
                        // This single-item list IS the primary group
                        var path = new List<string>(currentPath) { kvp.Key };
                        _logger.LogInformation($"Detected PRIMARY group '{kvp.Key}' at depth {depth} (single record with nested data)");
                        return (true, kvp.Key, list, path);
                    }
                }
            }

            // STEP 3: Recurse into nested structures ONLY if no primary group found at this level
            foreach (var kvp in currentDict)
            {
                if (kvp.Value is List<ExpandoObject> list && list.Count > 0)
                {
                    var firstItem = list[0] as IDictionary<string, object>;
                    if (firstItem != null)
                    {
                        var childPath = new List<string>(currentPath) { kvp.Key };
                        var result = FindRepeatingGroupInExpando(firstItem, childPath, depth + 1);
                        
                        if (result.HasRepeatingGroup)
                            return result;
                    }
                }
            }

            return (false, null, null, null);
        }

        /// <summary>
        /// FIXED: Determines if a single-item list represents a PRIMARY group.
        /// Primary group = main data container that may have nested repeating child records.
        /// Example: Patients (primary) contains Vitals/Medications (nested children)
        /// </summary>
        private bool IsPrimaryGroup(IDictionary<string, object> singleItem, string groupName, int depth)
        {
            if (singleItem == null)
                return false;
            // Number of direct scalar fields (e.g., Name, ID, Date, etc.)
            int stringFieldCount = 0;
            // Number of nested lists (e.g., Vitals, Medications, Components)
            int nestedListCount = 0;
            // Number of nested lists that contain MORE THAN ONE item
            // (i.e., true repeating child records)
            int multiItemListCount = 0;
            // Analyze the structure of this single record
            foreach (var kvp in singleItem)
            {
                if (kvp.Value is string)
                {
                    // Found a scalar data field
                    stringFieldCount++;
                }
                else if (kvp.Value is List<ExpandoObject> childList)
                {
                    // Found a nested collection
                    nestedListCount++;
                    // If the nested collection repeats, track it
                    if (childList.Count > 1)
                    {
                        multiItemListCount++;
                    }
                }
            }

            // ✅ RULE 1: TRUE primary record (has fields AND nested data)
            // Example: Patient with Vitals / Medications
            if (stringFieldCount > 0 && nestedListCount > 0)
            {
                return true;
            }

            // ✅ RULE 2: Single wrapper that owns repeating children (root-level)
            // Example: Root → Patients → Patient[]
            if (stringFieldCount == 0 && multiItemListCount > 0 && depth <= 1)
            {
                return true;
            }

            // Flat, scalar-only object (Validator case)
            return false;
        }
        private bool HasOwnScalarFields(IDictionary<string, object> dict)
        {
            // A PRIMARY group must have at least one scalar field
            return dict.Values.Any(v => v is string && !string.IsNullOrWhiteSpace(v.ToString()));
        }

        /// <summary>
        /// ULTRA-SIMPLIFIED: Executes mail merge with pre-analyzed group data.
        /// No XML traversal, no navigation paths - just use the data we already have!
        /// </summary>
        private bool ExecuteGroupMailMerge(WordDocument document, string groupName, 
                                           List<ExpandoObject> groupItems, string patientId)
        {
            try
            {
                // Filter by PatientID if selective report requested
                if (!string.IsNullOrWhiteSpace(patientId))
                {
                    groupItems = groupItems
                        .Where(item =>
                        {
                            var dict = item as IDictionary<string, object>;

                            if (dict == null) return false;

                            // Match PatientID / MRN
                            if (dict.ContainsKey("PatientID"))
                                return dict["PatientID"]?.ToString()
                                    .Equals(patientId, StringComparison.OrdinalIgnoreCase) == true;

                            if (dict.ContainsKey("MRN"))
                                return dict["MRN"]?.ToString()
                                    .Equals(patientId, StringComparison.OrdinalIgnoreCase) == true;

                            return false;
                        })
                        .ToList();

                    // Validate filtered results
                    if (groupItems == null || groupItems.Count == 0)
                    {
                        string message = $"No records found for Patient ID: {patientId}";
                        _logger.LogWarning(message);
                        ViewBag.Message = message;
                        return false;
                    }
                }

                _logger.LogInformation($"Processing {groupItems.Count} items in group '{groupName}'");

                // Determine merge strategy based on data structure
                bool hasNestedGroups = DetectNestedGroups(groupItems);
                bool hasMultipleRecords = groupItems.Count > 1;

                MailMergeDataTable dataTable = new MailMergeDataTable(groupName, groupItems);

                // Execute appropriate mail merge based on data structure
                if (hasNestedGroups)
                {
                    document.MailMerge.StartAtNewPage = true;
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    _logger.LogInformation($"Executed nested group mail merge for '{groupName}'");
                }
                else if (hasMultipleRecords)
                {
                    document.MailMerge.StartAtNewPage = true;
                    document.MailMerge.ExecuteGroup(dataTable);
                    _logger.LogInformation($"Executed group mail merge for {groupItems.Count} records");
                }
                else
                {
                    // Single record - simple merge
                    var firstItem = groupItems[0] as IDictionary<string, object>;
                    string[] fieldNames = firstItem.Keys.ToArray();
                    string[] fieldValues = firstItem.Values
                        .Select(v => v?.ToString() ?? string.Empty)
                        .ToArray();
                    document.MailMerge.Execute(fieldNames, fieldValues);
                    _logger.LogInformation($"Executed simple mail merge for single record");
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to execute group mail merge");
                ViewBag.Message = "Error executing group mail merge.";
                return false;
            }
        }

        /// <summary>
        /// Determines whether a list of group items contains nested or repeating groups.
        /// This is used to identify hierarchical XML structures that require
        /// nested mail merge processing instead of a simple group merge.
        /// </summary>
        private bool DetectNestedGroups(List<ExpandoObject> groupItems)
        {
            // Iterate through each primary group item
            foreach (var item in groupItems)
            {
                // Treat the group item as a dictionary for property access
                var itemDict = item as IDictionary<string, object>;
                // Examine each value inside the group item
                foreach (var value in itemDict.Values)
                {
                    // Identify child collections represented as lists of ExpandoObject
                    if (value is List<ExpandoObject> childList)
                    {
                        // If the child list contains more than one item,
                        // it is a repeating group
                        if (childList.Count > 1)
                            return true;
                        // If the child list contains exactly one item,
                        // check whether that item itself contains nested lists
                        if (childList.Count == 1)
                        {
                            var nestedItem = childList[0] as IDictionary<string, object>;

                            // If the nested item contains any list values,
                            // a deeper nested group exists
                            if (nestedItem?.Values.OfType<List<ExpandoObject>>().Any() == true)
                                return true;
                        }
                    }
                }
            }
            // No nested or repeating groups detected
            return false;
        }

        /// <summary>
        /// Executes simple mail merge for documents without repeating groups.
        /// Flattens root-level fields including single-item nested lists.
        /// </summary>
        private bool ExecuteSimpleMailMerge(WordDocument document, XmlNode documentElement, ExpandoObject parsedData)
        {
            try
            {
                // Convert the parsed ExpandoObject into a dictionary
                var rootDict = parsedData as IDictionary<string, object>;
                // Get the XML document root element name
                string documentRootName = documentElement.LocalName;
                // Validate that the parsed data contains the XML root element
                if (!rootDict.ContainsKey(documentRootName))
                {
                    ViewBag.Message = "XML root element not found.";
                    return false;
                }
                // Expect the root element to be represented as a list
                // (even if it contains only one item)
                var documentRootList = rootDict[documentRootName] as List<ExpandoObject>;
                if (documentRootList == null || documentRootList.Count == 0)
                {
                    ViewBag.Message = "Invalid XML structure.";
                    return false;
                }
                // Flatten nested XML structure into name/value field pairs
                // and capture multi-item lists that are skipped
                var rootItem = documentRootList[0] as IDictionary<string, object>;
                var (fields, skippedLists) = FlattenRootFields(rootItem);

                if (fields.Count == 0)
                {
                    ViewBag.Message = "No merge fields found in XML.";
                    return false;
                }
                // Extract merge field names and values
                string[] fieldNames = fields.Keys.ToArray();
                string[] fieldValues = fields.Values.ToArray();
                // Execute the simple mail merge
                document.MailMerge.Execute(fieldNames, fieldValues);

                // Log results
                if (skippedLists.Count > 0)
                {
                    _logger.LogInformation($"Mail merge completed with {fields.Count} fields. " +
                                           $"Skipped {skippedLists.Count} multi-item lists: {string.Join(", ", skippedLists)}");
                }
                else
                {
                    _logger.LogInformation($"Mail merge completed successfully with {fields.Count} fields.");
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to perform simple mail merge");
                ViewBag.Message = "Error processing merge fields.";
                return false;
            }
        }

        /// <summary>
        /// Flattens root-level fields from ExpandoObject, handling nested single-item lists.
        /// Returns dictionary of fields and list of skipped multi-item lists.
        /// </summary>
        private (Dictionary<string, string>, List<string>) FlattenRootFields(IDictionary<string, object> rootItem)
        {
            var fields = new Dictionary<string, string>();
            var skippedLists = new List<string>();

            foreach (var kvp in rootItem)
            {
                if (kvp.Value is string stringValue)
                {
                    fields[kvp.Key] = stringValue;
                }
                else if (kvp.Value is List<ExpandoObject> list)
                {
                    if (list.Count == 1)
                    {
                        // Flatten single-item nested list
                        var nestedItem = list[0] as IDictionary<string, object>;
                        if (nestedItem != null)
                        {
                            foreach (var nestedKvp in nestedItem)
                            {
                                string fieldKey = nestedKvp.Key;
                                if (fields.ContainsKey(fieldKey))
                                {
                                    fieldKey = $"{kvp.Key}_{nestedKvp.Key}";
                                    _logger.LogWarning($"Key collision detected: '{nestedKvp.Key}' renamed to '{fieldKey}'");
                                }
                                fields[fieldKey] = nestedKvp.Value?.ToString() ?? string.Empty;
                            }
                        }
                    }
                    else if (list.Count > 1)
                    {
                        skippedLists.Add(kvp.Key);
                        _logger.LogWarning($"List '{kvp.Key}' contains {list.Count} items but was not detected as a repeating group. " +
                                           $"This data will be skipped in the merge.");
                    }
                    else
                    {
                        _logger.LogInformation($"Empty list found for key '{kvp.Key}' - no data to merge.");
                    }
                }
                else if (kvp.Value != null)
                {
                    fields[kvp.Key] = kvp.Value.ToString();
                }
            }

            return (fields, skippedLists);
        }

        /// <summary>
        /// Generates output file based on requested type (single PDF or multiple PDFs in ZIP).
        /// </summary>
        private IActionResult GenerateOutputFile(WordDocument document, string type)
        {
            if (string.Equals(type, "single", StringComparison.OrdinalIgnoreCase))
            {
                MemoryStream pdfStream = SaveAsPDF(document);
                return File(pdfStream, "application/pdf", "GeneratedDocument.pdf");
            }
            else if (string.Equals(type, "multiple", StringComparison.OrdinalIgnoreCase))
            {
                byte[] zipBytes = SplitByPageBreak(document);
                if (zipBytes != null && zipBytes.Length > 0)
                    return File(zipBytes, "application/zip", "converted_pdfs.zip");
                else
                {
                    MemoryStream pdfStream = SaveAsPDF(document);
                    return File(pdfStream, "application/pdf", "GeneratedDocument.pdf");
                }
            }

            return View("Index");
        }
        /// <summary>
        /// Splits document by page breaks using bookmarks and returns ZIP bytes
        /// Each section between page breaks becomes a separate PDF
        /// </summary>
        private byte[] SplitByPageBreak(WordDocument document)
        {
            // Find all page breaks in the document
            List<Entity> entities = document.FindAllItemsByProperty(EntityType.Break, "BreakType", "PageBreak");
            if (entities == null || entities.Count == 0)
                return null;

            WSection section = document.Sections[0];
            WTextBody body = section.Body;
            int bookmarkIndex = 1;
            // Step 1: Insert a NEW paragraph at the very beginning with BookmarkStart
            WParagraph firstBookmarkPara = new WParagraph(document);
            firstBookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Insert(0, firstBookmarkPara);

            // Step 2: Iterate page break entities → insert bookmark paragraph directly after each
            foreach (Entity entity in entities)
            {
                WParagraph breakParagraph = entity.Owner as WParagraph;

                if (breakParagraph == null) continue;

                // Get the current index of this paragraph in the body
                int paraIndex = body.ChildEntities.IndexOf(breakParagraph);

                if (paraIndex < 0) continue;

                // Insert new paragraph right after the page break paragraph
                // Close current bookmark and open next bookmark in same paragraph
                WParagraph bookmarkPara = new WParagraph(document);
                bookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
                bookmarkIndex++;
                bookmarkPara.AppendBookmarkStart($"Page_Bookmark_{bookmarkIndex}");
                body.ChildEntities.Insert(paraIndex + 1, bookmarkPara);
            }

            // Step 3: Insert a NEW paragraph at the very end with BookmarkEnd
            WParagraph lastBookmarkPara = new WParagraph(document);
            lastBookmarkPara.AppendBookmarkEnd($"Page_Bookmark_{bookmarkIndex}");
            body.ChildEntities.Add(lastBookmarkPara);
            // Step 4: Create ZIP file and convert each bookmarked section to PDF
            var zipStream = new MemoryStream();
            using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 1; i <= bookmarkIndex; i++)
                {
                    try
                    {
                        // Navigate to each bookmark section
                        BookmarksNavigator navigator = new BookmarksNavigator(document);
                        navigator.MoveToBookmark($"Page_Bookmark_{i}", true, true);
                        WordDocumentPart documentPart = navigator.GetContent();

                        if (documentPart == null) continue;

                        // Extract content as new WordDocument.
                        using (WordDocument extractedDoc = documentPart.GetAsWordDocument())
                        using (DocIORenderer render = new DocIORenderer())
                        // Convert extracted document into PDF.
                        using (PdfDocument pdfDocument = render.ConvertToPDF(extractedDoc))
                        using (MemoryStream pdfStream = new MemoryStream())
                        {
                            pdfDocument.Save(pdfStream);
                            pdfStream.Position = 0;

                            // Write PDF into ZIP entry
                            var entry = zip.CreateEntry($"Document_{i}.pdf", CompressionLevel.Fastest);
                            using (var entryStream = entry.Open())
                            {
                                pdfStream.CopyTo(entryStream);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log and continue to next section if one fails
                        _logger.LogError(ex, $"Error converting bookmark {i} to PDF");
                    }
                }
            }
            return zipStream.ToArray();
        }
        /// <summary>
        /// Convert document into PDF and returns PDF stream
        /// </summary>
        private MemoryStream SaveAsPDF(WordDocument document)
        {
            using (DocIORenderer renderer = new DocIORenderer())
            {
                // Convert the Word document to PDF format
                using (PdfDocument pdfDocument = renderer.ConvertToPDF(document))
                {
                    MemoryStream pdfStream = new MemoryStream();
                    pdfDocument.Save(pdfStream);
                    pdfStream.Position = 0;

                    return pdfStream;
                }
            }
        }

        /// <summary>
        /// Retrieves an XML document stream based on the uploaded files.
        /// </summary>

        private Stream GetXmlDocument(IFormFile file, IFormFile xmlFile)
        {
            // Case 1: Primary file (Word document) is uploaded
            if (file != null && file.Length > 0)
            {
                // Ensure an XML file is also uploaded
                if (xmlFile != null && xmlFile.Length > 0)
                {
                    // Copy uploaded XML file into an in-memory stream.
                    MemoryStream stream = new MemoryStream();
                    xmlFile.CopyTo(stream);
                    // Reset stream position so it can be read from the beginning
                    stream.Position = 0;
                    return stream;
                }
                else
                {
                    ViewBag.Message = "Please upload an XML data file along with the Word document.";
                    return null;
                }
            }
            else
            {
                // Case 2: No uploaded file — load default XML from the application's web root
                string defaultXmlPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "PatientDetails.xml");
                using (var fileStream = new FileStream(defaultXmlPath, FileMode.Open, FileAccess.Read))
                {
                    var memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    memoryStream.Position = 0;
                    return memoryStream;
                }
            }
        }

        /// <summary>
        /// Gets the data as ExpandoObject - exact match to Syncfusion reference
        /// </summary>
        private void GetDataAsExpandoObject(XmlNode node, ref ExpandoObject dynamicObject)
        {
            // Check if node has any child ELEMENT nodes (not just text/entity nodes)
            bool hasChildElements = node.ChildNodes.Cast<XmlNode>()
            .Any(n => n.NodeType == XmlNodeType.Element);
            // If the node has no child elements, treat it as a leaf node
            if (!hasChildElements)
            {
                // Leaf node - simple text value (handles HTML entities correctly)
                dynamicObject.TryAdd(node.LocalName, node.InnerText);
            }
            else
            {
                // Complex node - has child elements
                // This node contains child elements and will be represented
                // as a list of ExpandoObjects (to support repeating XML elements)
                List<ExpandoObject> childObjects;
                // Check if the current ExpandoObject already contains this element name
                // Reuse existing list if the element already exists
                if ((dynamicObject as IDictionary<string, object>).ContainsKey(node.LocalName))
                    childObjects = (dynamicObject as IDictionary<string, object>)[node.LocalName]
                    as List<ExpandoObject>;
                else
                {
                    // Otherwise create a new list for this element
                    childObjects = new List<ExpandoObject>();
                    dynamicObject.TryAdd(node.LocalName, childObjects);
                }
                // Create a new ExpandoObject to represent this instance of the complex node
                ExpandoObject childObject = new ExpandoObject();
                // Recursively process each child ELEMENT node
                foreach (XmlNode childNode in node.ChildNodes)
                {
                    if (childNode.NodeType == XmlNodeType.Element)
                    {
                        GetDataAsExpandoObject(childNode, ref childObject);
                    }
                }
                // Add the fully populated child object to the list
                childObjects.Add(childObject);
            }
        }         
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }       
    }
}
