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
        /// Converts the XML into a dynamic <see cref="ExpandoObject"/> structure.
        /// Analyzes the structure to detect repeating groups and nested groups.
        /// Executes mail merge based on structure.
        /// </summary>
        private bool ProcessXmlAndPerformMerge(WordDocument document, Stream xmlStream, string patientId)
        {
            try
            {
                // Ensure the stream is read from the beginning
                xmlStream.Position = 0;

                // Load as XmlDocument (not XDocument) to work with your existing method
                var xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlStream);
                // Get the root element of the XML document
                var rootNode = xmlDoc.DocumentElement;
                var rootName = rootNode.Name;

                // Convert XML structure into a dynamic ExpandoObject
                // This allows flexible property access without rigid classes
                var parsedData = new ExpandoObject();
                GetDataAsExpandoObject(rootNode, ref parsedData);

                // Analyze the parsed structure to:
                // - Detect if a repeating group exists
                // - Identify the group name
                // - Extract group items
                // - Check for nested repeating groups
                var (hasGroup, groupName, groupItems, hasNestedGroups) =
                    AnalyzeExpandoObjectStructure(parsedData, rootName,document);
                // If no repeating group is found, perform a simple mail merge
                if (!hasGroup)
                {
                    //simple merge method
                    _logger.LogInformation("No repeating group found - performing simple mail merge");
                    return ExecuteSimpleMailMerge(document, rootNode, parsedData);
                }

                _logger.LogInformation($"Performing mail merge for group '{groupName}' (HasNestedGroups={hasNestedGroups})");
                // Execute group-based mail merge Handles both flat and nested group scenario
                return ExecuteGroupMailMerge(document, groupName, groupItems, patientId, hasNestedGroups);
            }
            catch (Exception ex)
            {
                _logger.LogError($"Error in ProcessXmlAndPerformMerge: {ex.Message}");
                return false;
            }
        }


        /// <summary>
        /// Analyzes the parsed XML (as an <see cref="ExpandoObject"/>) to determine:
        /// - Whether a repeating group exists
        /// - The name of the group
        /// - The items belonging to that group
        /// - Whether the group contains nested repeating groups
        /// Additionally, it validates that the detected group exists in the Word template.
        /// </summary>
        private (bool HasRepeatingGroup, string GroupName, List<ExpandoObject> GroupItems, bool HasNestedGroups)
    AnalyzeExpandoObjectStructure(ExpandoObject parsedData, string rootName, WordDocument document)
        {
            try
            {
                // Convert ExpandoObject into dictionary for key-based access
                var rootDict = parsedData as IDictionary<string, object>;
                // Validate that the root element exists in parsed data
                if (!rootDict.ContainsKey(rootName))
                {
                    _logger.LogWarning($"Root element '{rootName}' not found in parsed XML");
                    return (false, null, null, false);
                }
                // Extract value corresponding to the root element
                var rootValue = rootDict[rootName];
                // Ensure the root value is a non-empty list of ExpandoObjects
                // This represents structured XML data suitable for processing
                if (!(rootValue is List<ExpandoObject> rootList) || rootList.Count == 0)
                {
                    _logger.LogWarning($"Root element '{rootName}' does not contain valid data");
                    return (false, null, null, false);
                }
                // Retrieve merge group names defined in the Word template
                // This is used later to validate detected XML groups
                string[] templateGroupNames = document.MailMerge.GetMergeGroupNames();
                // Identify the main repeating group within the XML structure
                var result = FindMainDataGroup(rootList, rootName);

                if (result.Found)
                {
                    // Validate group exists in template before returning true
                    bool groupExistsInTemplate = templateGroupNames.Any(g =>
                        g.Equals(result.GroupName, StringComparison.OrdinalIgnoreCase));
                    // If group exists in XML but not in template, ignore it
                    if (!groupExistsInTemplate)
                    {
                        _logger.LogInformation($"Group '{result.GroupName}' found in XML but NOT in template - returning HasRepeatingGroup=false");
                        return (false, null, null, false);
                    }

                    // Group exists in both XML and template
                    // Check if the group contains nested repeating structures
                    bool hasNested = CheckForNestedGroups(result.Items);
                    _logger.LogInformation($"Valid group '{result.GroupName}' found in both XML and template with {result.Items.Count} item(s), HasNestedGroups={hasNested}");
                    // Return successful group detection result
                    return (true, result.GroupName, result.Items, hasNested);
                }
                // No repeating group detected in XML
                _logger.LogInformation("No group found in XML structure");
                return (false, null, null, false);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors during analysis
                _logger.LogError($"Error analyzing structure: {ex.Message}");
                return (false, null, null, false);
            }
        }
        /// <summary>
        /// Recursively searches for the primary repeating data group within a list of ExpandoObjects.
        /// Detection rules:
        /// 1. If the current level contains multiple items → it is treated as a true repeating group.
        /// 2. If a single item exists:
        ///    - If it has simple fields and repeating children → hierarchical parent (valid group)
        ///    - If it has only simple fields → simple data record (valid group)
        ///    - If it has no fields but contains nested lists → treated as a wrapper, recurse into children
        /// </summary>
        private GroupSearchResult FindMainDataGroup(List<ExpandoObject> itemList, string groupName)
        {
            if (itemList == null || itemList.Count == 0)
                return GroupSearchResult.NotFound;

            // RULE 1: Multiple items at THIS level = TRUE REPEATING GROUP
            if (itemList.Count > 1)
            {
                _logger.LogInformation($"✓ Found repeating group '{groupName}' with {itemList.Count} items");
                return new GroupSearchResult
                {
                    Found = true,
                    GroupName = groupName,
                    Items = itemList
                };
            }
            // RULE 2: Single item - check if it's a HIERARCHICAL PARENT
            if (itemList.Count == 1)
            {
                // Convert the single item into a dictionary for analysis
                var singleItem = itemList[0] as IDictionary<string, object>;
                if (singleItem != null)
                {

                    // Analyze the structure of the item:
                    // - Does it have simple fields?
                    // - Does it contain repeating child groups?
                    // - Does it contain nested lists?
                    var analysis = AnalyzeItemStructure(singleItem);

                    // CASE 1: HIERARCHICAL PARENT
                    // Has simple fields AND repeating child groups
                    // Example: Patient → Visits[], Medications[]
                    if (analysis.HasSimpleFields && analysis.HasRepeatingGroups)
                    {
                        _logger.LogInformation($"✓ Found HIERARCHICAL parent '{groupName}' with {analysis.SimpleFieldCount} fields and {analysis.RepeatingGroupCount} repeating child(s)");
                        return new GroupSearchResult
                        {
                            Found = true,
                            GroupName = groupName,
                            Items = itemList
                        };
                    }
                    // CASE 2: SIMPLE DATA RECORD
                    // Has only simple fields and no repeating children
                    // Example: Single object with basic properties
                    if (analysis.HasSimpleFields && !analysis.HasRepeatingGroups)
                    {
                        _logger.LogInformation($"✓ Found simple data record '{groupName}' with {analysis.SimpleFieldCount} fields");
                        return new GroupSearchResult
                        {
                            Found = true,
                            GroupName = groupName,
                            Items = itemList
                        };
                    }

                    // CASE 3: CONTAINER WRAPPER
                    // No direct fields, only contains nested lists
                    // Example: <Root><Patients>...</Patients></Root>
                    // → must recurse into children to find real group
                    if (!analysis.HasSimpleFields && analysis.HasNestedLists)
                    {
                        _logger.LogInformation($"'{groupName}' is a wrapper - checking children...");
                        // Recursively inspect each child list
                        foreach (var kvp in singleItem)
                        {
                            if (kvp.Value is List<ExpandoObject> childList)
                            {
                                var childResult = FindMainDataGroup(childList, kvp.Key);
                                // If a valid group is found in children, return immediately
                                if (childResult.Found)
                                    return childResult;
                            }
                        }
                    }
                }
            }

            return GroupSearchResult.NotFound;
        }

        /// <summary>
        /// Checks if group items contain nested repeating groups (done during analysis)
        /// This replaces the separate DetectNestedGroups call later
        /// </summary>
        private bool CheckForNestedGroups(List<ExpandoObject> groupItems)
        {
            if (groupItems == null || groupItems.Count == 0)
                return false;

            // Check the first item (structure should be consistent across all items)
            var firstItem = groupItems[0] as IDictionary<string, object>;
            if (firstItem == null)
                return false;

            foreach (var kvp in firstItem)
            {
                // Found a nested list with multiple items = nested repeating group
                if (kvp.Value is List<ExpandoObject> nestedList && nestedList.Count > 1)
                {
                    _logger.LogInformation($"  → Detected nested repeating group '{kvp.Key}' with {nestedList.Count} items");
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Analyzes item structure: fields, children, hierarchy
        /// </summary>
        private ItemAnalysis AnalyzeItemStructure(IDictionary<string, object> item)
        {
            var analysis = new ItemAnalysis();

            foreach (var kvp in item)
            {
                if (kvp.Value is string)
                {
                    analysis.SimpleFieldCount++;
                }
                else if (kvp.Value is List<ExpandoObject> childList)
                {
                    analysis.NestedListCount++;

                    // COUNT 2: Is this a repeating group (multiple items)?
                    if (childList.Count > 1)
                    {
                        analysis.RepeatingGroupCount++;
                    }
                }
            }
            // Decision 1: Does item have its own data fields?
            // Rule: If we found at least 1 simple field, it has data
            analysis.HasSimpleFields = analysis.SimpleFieldCount > 0;
            // Decision 2: Does item own repeating sub-groups?
            // Rule: If we found at least 1 repeating group, it has repeating children
            analysis.HasRepeatingGroups = analysis.RepeatingGroupCount > 0;
            // Decision 3: Does item have ANY nested structure?
            // Rule: If we found at least 1 nested list (repeating or single), it has nesting
            analysis.HasNestedLists = analysis.NestedListCount > 0;

            return analysis;
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
                            var dict = item as IDictionary<string, object>;
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
                // Determine if there are multiple records
                bool hasMultipleRecords = groupItems.Count > 1;
                // Create a MailMergeDataTable required for grouped mail merge execution
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
        /// Wrapper method to use your existing GetDataAsExpandoObject method
        /// </summary>
        private ExpandoObject ParseXmlToExpandoObject(XElement xElement)
        {
            // Convert XElement to XmlNode (required by your existing method)
            var xmlDoc = new XmlDocument();
            using (var xmlReader = xElement.CreateReader())
            {
                xmlDoc.Load(xmlReader);
            }

            var xmlNode = xmlDoc.DocumentElement;
            var expandoObject = new ExpandoObject();

            // Use your existing method
            GetDataAsExpandoObject(xmlNode, ref expandoObject);

            return expandoObject;
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

        /// <summary>
        /// Helper classes for clear results
        /// </summary>
        private class GroupSearchResult
        {
            public bool Found { get; set; }
            public string GroupName { get; set; }
            public List<ExpandoObject> Items { get; set; }

            public static GroupSearchResult NotFound => new GroupSearchResult { Found = false };
        }

        private class ItemAnalysis
        {
            public int SimpleFieldCount { get; set; }
            public int NestedListCount { get; set; }
            public int RepeatingGroupCount { get; set; }
            public bool HasSimpleFields { get; set; }
            public bool HasRepeatingGroups { get; set; }
            public bool HasNestedLists { get; set; }
        }
    }
}
