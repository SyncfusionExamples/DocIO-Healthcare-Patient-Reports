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
                using (FileStream fileStream = new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read))
                {
                    MemoryStream memoryStream = new MemoryStream();
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
                    mergeSuccess = ProcessXmlAndPerformMerge(document, xmlStream, effectivePatientId);
                    
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
        /// Processes XML data and performs a Word mail merge operation.
        /// Uses template group names as source of truth for identifying repeating groups.
        /// </summary>
        private bool ProcessXmlAndPerformMerge(WordDocument document, Stream xmlStream, string patientId)
        {
            try
            {
                // Ensure the stream is read from the beginning
                xmlStream.Position = 0;

                // Load XML document
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlStream);
                XmlElement rootNode = xmlDoc.DocumentElement;
                string rootName = rootNode.Name;

                // Convert XML to ExpandoObject
                ExpandoObject parsedData = new ExpandoObject();
                GetDataAsExpandoObject(rootNode, ref parsedData);

                // Find matching group between XML and template
                (string groupName, List<ExpandoObject> groupItems, bool isNested) =
                    FindMatchingGroup(parsedData, rootName, document);

                // No matching group found - perform simple merge
                if (groupName == null)
                {
                    _logger.LogInformation("No matching group found between XML and template - performing simple mail merge");
                    return ExecuteSimpleMailMerge(document, rootNode, parsedData);
                }

                _logger.LogInformation($"Performing mail merge for group '{groupName}' (IsNested={isNested})");

                // Execute group-based mail merge
                return ExecuteGroupMailMerge(document, groupName, groupItems, patientId, isNested);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ProcessXmlAndPerformMerge: {ex.Message}");
                return false;
            }
        }
        /// <summary>
        /// Finds matching groups between XML structure and Word template merge groups.
        /// Returns the first matching group found with its items and structure type.
        /// </summary>
        private (string GroupName, List<ExpandoObject> Items, bool IsNested)
            FindMatchingGroup(ExpandoObject parsedData, string rootName, WordDocument document)
        {
            try
            {
                string[] templateGroupNames = document.MailMerge.GetMergeGroupNames();

                if (templateGroupNames == null || templateGroupNames.Length == 0)
                {
                    _logger.LogInformation("No merge groups found in template");
                    return (null, null, false);
                }

                _logger.LogInformation($"Template contains {templateGroupNames.Length} group(s): {string.Join(", ", templateGroupNames)}");

                IDictionary<string, object> rootDict = parsedData as IDictionary<string, object>;

                // Get root value directly - validation happens in SearchForTemplateGroup
                object rootValue = rootDict?[rootName];

                return SearchForTemplateGroup(rootValue, templateGroupNames, rootName);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error finding matching group: {ex.Message}");
                return (null, null, false);
            }
        }
        /// <summary>
        /// Recursively searches through XML structure to find tags matching template group names.
        /// All validation (type checking, count checking) happens here.
        /// </summary>
        private (string GroupName, List<ExpandoObject> Items, bool IsNested)
            SearchForTemplateGroup(object itemValue, string[] templateGroupNames, string currentTag)
        {
            // Unified validation: Check if itemValue is valid List<ExpandoObject>
            if (!(itemValue is List<ExpandoObject> itemList) || itemList.Count == 0)
            {
                _logger.LogDebug($"Tag '{currentTag}' does not contain valid data, skipping.");
                return (null, null, false);
            }

            // Check if current tag matches any template group
            bool isMatchingGroup = templateGroupNames.Any(g =>
                g.Equals(currentTag, StringComparison.OrdinalIgnoreCase));

            if (isMatchingGroup)
            {
                _logger.LogInformation($"✓ Found matching group '{currentTag}' with {itemList.Count} item(s)");
                bool hasNested = HasNestedGroups(itemList, templateGroupNames);
                return (currentTag, itemList, hasNested);
            }

            // Current tag doesn't match - search children
            IDictionary<string, object> firstItem = itemList[0] as IDictionary<string, object>;

            if (firstItem != null)
            {
                foreach (KeyValuePair<string, object> kvp in firstItem)
                {
                    // Recursively search - pass raw object value
                    (string GroupName, List<ExpandoObject> Items, bool IsNested) result =
                        SearchForTemplateGroup(kvp.Value, templateGroupNames, kvp.Key);

                    if (result.GroupName != null)
                        return result;
                }
            }

            return (null, null, false);
        }
        /// <summary>
        /// Determines if group items contain nested groups that match template group names.
        /// </summary>
        private bool HasNestedGroups(List<ExpandoObject> groupItems, string[] templateGroupNames)
        {
            if (groupItems == null || groupItems.Count == 0)
                return false;

            // Check first item (structure should be consistent across items)
            IDictionary<string, object> firstItem = groupItems[0] as IDictionary<string, object>;

            if (firstItem == null)
                return false;

            foreach (KeyValuePair<string, object> kvp in firstItem)
            {
                // Check if property is a list (potential nested group)
                if (kvp.Value is List<ExpandoObject> nestedList)
                {
                    // Check if this tag matches any template group name
                    bool isTemplateGroup = templateGroupNames.Any(g =>
                        g.Equals(kvp.Key, StringComparison.OrdinalIgnoreCase));

                    if (isTemplateGroup && nestedList.Count > 0)
                    {
                        _logger.LogInformation($"  → Found nested group '{kvp.Key}' with {nestedList.Count} item(s)");
                        return true;
                    }
                }
            }

            return false;
        }            
        /// <summary>
        /// Executes a Word mail merge for a detected group of data.
        /// Handles multiple scenarios:
        /// - Filters records by Patient ID (if provided)
        /// - Executes nested group mail merge (for hierarchical data)
        /// - Executes group mail merge (for multiple flat records)
        /// - Executes simple mail merge (for a single record)
        /// </summary>
        private bool ExecuteGroupMailMerge(WordDocument document, string groupName,
                                           List<ExpandoObject> groupItems, string patientId,
                                           bool hasNestedGroups)
        {
            try
            {
                // Filter by PatientID if selective report requested
                if (!string.IsNullOrWhiteSpace(patientId))
                {
                    groupItems = groupItems
                        .Where(item =>
                        {
                            // Convert item to dictionary for field access
                            IDictionary<string, object> dict = item as IDictionary<string, object>;
                            if (dict == null) return false;

                            // Match PatientID field (case-insensitive)
                            if (dict.ContainsKey("PatientID"))
                                return dict["PatientID"]?.ToString()
                                    .Equals(patientId, StringComparison.OrdinalIgnoreCase) == true;
                            // Match MRN field (alternative identifier)
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
                // Create a MailMergeDataTable required for grouped mail merge execution
                MailMergeDataTable dataTable = new MailMergeDataTable(groupName, groupItems);
                //Execute nested group mail merge
                if (hasNestedGroups)
                {
                    document.MailMerge.StartAtNewPage = true;
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    _logger.LogInformation($"Executed nested group mail merge for '{groupName}'");
                }
                else
                {
                    // Execute group merge for both single and multiple records
                    document.MailMerge.StartAtNewPage = groupItems.Count > 1;
                    document.MailMerge.ExecuteGroup(dataTable);
                    _logger.LogInformation($"Executed group mail merge for {groupItems.Count} record(s) in '{groupName}'");
                }
                // Create a MailMergeDataTable required for grouped mail merge execution
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
        /// Executes simple mail merge for documents without repeating groups.
        /// Flattens root-level fields including single-item nested lists.
        /// </summary>
        private bool ExecuteSimpleMailMerge(WordDocument document, XmlNode documentElement, ExpandoObject parsedData)
        {
            try
            {
                // Convert the parsed ExpandoObject into a dictionary
                IDictionary<string, object> rootDict = parsedData as IDictionary<string, object>;
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
                List<ExpandoObject> documentRootList = rootDict[documentRootName] as List<ExpandoObject>;
                if (documentRootList == null || documentRootList.Count == 0)
                {
                    ViewBag.Message = "Invalid XML structure.";
                    return false;
                }
                // Flatten nested XML structure into name/value field pairs
                // and capture multi-item lists that are skipped
                IDictionary<string, object> rootItem = documentRootList[0] as IDictionary<string, object>;
                (Dictionary<string, string> fields, List<string> skippedLists) = FlattenRootFields(rootItem);

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
            Dictionary<string, string> fields = new Dictionary<string, string>();
            List<string> skippedLists = new List<string>();

            foreach (KeyValuePair<string, object> kvp in rootItem)
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
                        IDictionary<string, object> nestedItem = list[0] as IDictionary<string, object>;
                        if (nestedItem != null)
                        {
                            foreach (KeyValuePair<string, object> nestedKvp in nestedItem)
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
            MemoryStream zipStream = new MemoryStream();
            using (ZipArchive zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
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
                            ZipArchiveEntry entry = zip.CreateEntry($"Document_{i}.pdf", CompressionLevel.Fastest);
                            using (Stream entryStream = entry.Open())
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
                using (FileStream fileStream = new FileStream(defaultXmlPath, FileMode.Open, FileAccess.Read))
                {
                    MemoryStream memoryStream = new MemoryStream();
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
