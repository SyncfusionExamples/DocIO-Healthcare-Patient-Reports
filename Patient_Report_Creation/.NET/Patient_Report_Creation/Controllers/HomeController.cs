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
                string defaultFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "Template-2.docx");
                return new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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

                    // For selective report, filter XML for the specific patient
                    if (reportType == "selective" && !string.IsNullOrWhiteSpace(patientId))
                    {
                        xmlStream.Position = 0;
                        XmlDocument masterXmlDoc = new XmlDocument();
                        masterXmlDoc.Load(xmlStream);

                        string primaryGroupName;
                        // Filter XML to contain only the specified patient's records
                        XmlDocument filteredXml = FilterXmlByPatientId(masterXmlDoc, patientId, out primaryGroupName);

                        if (filteredXml == null || filteredXml.DocumentElement.ChildNodes.Count == 0)
                        {
                            ViewBag.Message = $"No data found for Patient ID: {patientId}";
                            _logger.LogWarning($"No data found for Patient ID: {patientId}");
                            return View("Index");
                        }

                        // Create memory stream from filtered XML
                        using (MemoryStream filteredXmlStream = new MemoryStream())
                        {
                            filteredXml.Save(filteredXmlStream);
                            filteredXmlStream.Position = 0;

                            // Perform mail merge with filtered patient data

                            mergeSuccess = ExecuteMailMergeUsingKnownGroup(
                                document,
                                filteredXmlStream,
                                primaryGroupName);

                        }

                        _logger.LogInformation($"Generated selective report for Patient ID: {patientId}");
                    }
                    else
                    {
                        // Generate report for all records
                        mergeSuccess = ProcessXmlAndPerformMerge(document, xmlStream);
                        _logger.LogInformation("Generated report for all records");
                    }

                    if (!mergeSuccess)
                        return View("Index");
                }

                // Generate and return the output file based on requested format
                return GenerateOutputFile(document, outputFormat);
            }
        }
        /// <summary>
        /// Filters XML document to contain only records matching the specified PatientID.
        /// Preserves the original XML structure and root element.
        /// </summary>
        private XmlDocument FilterXmlByPatientId(XmlDocument xmlDoc, string patientId, out string primaryGroupName)
        {

            primaryGroupName = null;

            try
            {
                XmlDocument filteredDoc = new XmlDocument();
                XmlElement masterRoot = xmlDoc.DocumentElement;

                // Create new root element
                XmlElement newRoot = filteredDoc.CreateElement(masterRoot.LocalName);
                filteredDoc.AppendChild(newRoot);

                // ✅ IMPORTANT: detect PRIMARY GROUP ONCE (before filtering)
                var repeatingGroup = masterRoot.ChildNodes
                    .OfType<XmlNode>()
                    .Where(n => n.NodeType == XmlNodeType.Element)
                    .GroupBy(n => n.LocalName)
                    .FirstOrDefault(g => g.Count() > 1);

                if (repeatingGroup == null)
                    throw new Exception("No repeating group found in XML");

                primaryGroupName = repeatingGroup.Key; // ✅ STORE GROUP NAME

                // Copy root attributes
                foreach (XmlAttribute attr in masterRoot.Attributes)
                    newRoot.SetAttribute(attr.Name, attr.Value);

                // Find group nodes (Patients)
                XmlNodeList groupNodes = xmlDoc.GetElementsByTagName(primaryGroupName);
                int matchCount = 0;

                foreach (XmlNode groupNode in groupNodes)
                {
                    string nodePatientId = GetPatientIdFromNode(groupNode);

                    if (!string.IsNullOrEmpty(nodePatientId) &&
                        nodePatientId.Equals(patientId, StringComparison.OrdinalIgnoreCase))
                    {
                        XmlNode importedNode = filteredDoc.ImportNode(groupNode, true);
                        newRoot.AppendChild(importedNode);
                        matchCount++;
                    }
                }

                if (matchCount == 0)
                    return null;

                return filteredDoc;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error filtering XML");
                return null;
            }

        }

        /// <summary>
        /// Extracts PatientID from XML node by checking attributes or child elements.
        /// Supports multiple PatientID naming conventions (PatientID, MRN, etc.)
        /// </summary>
        private string GetPatientIdFromNode(XmlNode node)
        {
            // Check for PatientID attribute
            if (node.Attributes?["PatientID"] != null)
                return node.Attributes["PatientID"].Value;

            // Check for PatientID child element
            XmlNode patientIdNode = node.SelectSingleNode("PatientID");
            if (patientIdNode != null)
                return patientIdNode.InnerText;

            // Also check for common variations
            patientIdNode = node.SelectSingleNode("MRN");
            if (patientIdNode != null)
                return patientIdNode.InnerText;

            return null;
        }

        /// <summary>
        /// Creates a ZIP file from multiple PDF files
        /// </summary>
        private IActionResult CreateZipFromPdfs(List<byte[]> pdfFiles, List<string> fileNames)
        {
            var zipStream = new MemoryStream();
            using (var zip = new ZipArchive(zipStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 0; i < pdfFiles.Count; i++)
                {
                    var entry = zip.CreateEntry(fileNames[i], CompressionLevel.Fastest);
                    using (var entryStream = entry.Open())
                    {
                        entryStream.Write(pdfFiles[i], 0, pdfFiles[i].Length);
                    }
                }
            }

            zipStream.Position = 0;
            return File(zipStream.ToArray(), "application/zip", "patient_reports.zip");
        }
        /// <summary>
        /// Processes XML data, detects repeating groups, and performs mail merge on the document.
        /// Returns true if merge completed successfully, false otherwise.
        /// </summary>
        private bool ProcessXmlAndPerformMerge(WordDocument document, Stream xmlStream)
        {
            try
            {
                xmlStream.Position = 0;
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.Load(xmlStream);

                XmlNode documentElement = xmlDocument.DocumentElement;

                // Find repeating groups in the XML structure
                (XmlNode repeatingGroupNode, XmlNode repeatingGroupParent) = FindRepeatingGroupNode(documentElement);

                // Convert XML to dynamic ExpandoObject structure
                ExpandoObject parsedData = new ExpandoObject();

                GetDataAsExpandoObject(documentElement, ref parsedData);

                // Extract and execute appropriate mail merge
                if (repeatingGroupNode != null)
                {
                    return ExecuteGroupMailMerge(document, repeatingGroupNode, repeatingGroupParent,
                                                documentElement, parsedData);
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
        /// Finds repeating group node by recursively searching XML structure up to 5 levels deep.
        /// Returns tuple of (repeatingGroupNode, parentNode) or (null, null) if not found.
        /// </summary>

        private (XmlNode repeatingGroupNode, XmlNode repeatingGroupParent) FindRepeatingGroupNode(XmlNode documentElement)
        {
            // Will hold the first detected repeating node
            XmlNode repeatingGroupNode = null;
            // Will hold the parent of the repeating node
            XmlNode repeatingGroupParent = null;

            // Start recursive depth-first traversal from the document element
            FindRepeatingGroup(currentNode: documentElement, parentNode: null, ref repeatingGroupNode, ref repeatingGroupParent, maxDepth: 5);

            // Return the result as a tuple
            return (repeatingGroupNode, repeatingGroupParent);
        }
        /// <summary>
        /// Recursively traverses an XML node tree to locate the first element
        /// that appears more than once at the same hierarchy level.
        /// </summary>
        private bool FindRepeatingGroup(XmlNode currentNode, XmlNode parentNode, ref XmlNode repeatingGroupNode, ref XmlNode repeatingGroupParent, int maxDepth)
        {
            // Stop recursion if depth limit is reached
            if (maxDepth <= 0)
                return false;

            // Iterate through all child nodes of the current node
            foreach (XmlNode child in currentNode.ChildNodes)
            {
                // Process only element nodes (ignore text, comments, etc.)
                if (child.NodeType != XmlNodeType.Element)
                    continue;
                if (currentNode.SelectNodes(child.LocalName).Count > 1)
                {
                    // Store the repeating element
                    repeatingGroupNode = child;
                    // Store its parent
                    repeatingGroupParent = currentNode;
                    // Log where the repeating group was discovered
                    _logger.LogInformation(
                        $"Found repeating group '{child.LocalName}' under '{currentNode.LocalName}' " +
                        $"at depth {5 - maxDepth + 1}"
                    );
                    // Signal that the search is complete
                    return true;
                }

                //Recursively search deeper levels.
                //If found in a deeper call, immediately stop further traversal.
                if (FindRepeatingGroup(currentNode: child, parentNode: currentNode, ref repeatingGroupNode, ref repeatingGroupParent, maxDepth: maxDepth - 1))
                {
                    return true;
                }
            }
            // No repeating group found at this level
            return false;
        }

        private bool IsLogicalGroup(XmlNode node)
        {
            // Count only element children
            var elementChildren = node.ChildNodes
                .OfType<XmlNode>()
                .Where(n => n.NodeType == XmlNodeType.Element)
                .ToList();

            // A logical group:
            // 1. Has element children
            // 2. Is NOT a simple leaf
            return elementChildren.Count > 0 &&
                   elementChildren.Any(c => c.HasChildNodes);
        }


        /// <summary>
        /// Executes mail merge for documents with repeating groups (nested, grouped, or single record).
        /// </summary>
        private bool ExecuteGroupMailMerge(WordDocument document, XmlNode repeatingGroupNode,
                                           XmlNode repeatingGroupParent, XmlNode documentElement, ExpandoObject parsedData)
        {
            try
            {
                string groupName = repeatingGroupNode.LocalName;
                string parentNodeName = repeatingGroupParent.LocalName;
                string documentRootName = documentElement.LocalName;

                var rootDict = parsedData as IDictionary<string, object>;
                List<ExpandoObject> groupItems = NavigateToGroup(rootDict, documentElement,
                                                                 repeatingGroupParent, groupName, documentRootName);

                if (groupItems == null || groupItems.Count == 0)
                {
                    _logger.LogWarning($"No items found in repeating group '{groupName}'");
                    return false;
                }

                _logger.LogInformation($"Successfully found {groupItems.Count} items in repeating group '{groupName}'");

                // Determine merge strategy based on data structure
                bool hasNestedGroups = DetectNestedGroups(groupItems);
                bool hasMultipleRecords = groupItems.Count > 1;

                MailMergeDataTable dataTable = new MailMergeDataTable(groupName, groupItems);

                // Execute appropriate mail merge
                if (hasNestedGroups)
                {
                    document.MailMerge.StartAtNewPage = true;
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                }
                else if (hasMultipleRecords)
                {
                    document.MailMerge.StartAtNewPage = true;
                    document.MailMerge.ExecuteGroup(dataTable);
                }
                else
                {
                    var firstItem = groupItems[0] as IDictionary<string, object>;
                    string[] fieldNames = firstItem.Keys.ToArray();
                    string[] fieldValues = firstItem.Values
                        .Select(v => v?.ToString() ?? string.Empty)
                        .ToArray();
                    document.MailMerge.Execute(fieldNames, fieldValues);
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
        /// Navigates through nested ExpandoObject structure to locate repeating group items.
        /// </summary>
        private List<ExpandoObject> NavigateToGroup(IDictionary<string, object> rootDict, XmlNode documentElement,
                                                    XmlNode repeatingGroupParent, string groupName, string documentRootName)
        {
            try
            {
                // Build complete navigation path from root to repeating group
                List<string> navigationPath = BuildNavigationPath(documentElement, repeatingGroupParent, groupName, documentRootName);
                _logger.LogInformation($"Navigation path: {string.Join(" → ", navigationPath)}");

                // Navigate through the path to find the group items
                IDictionary<string, object> current = rootDict;
                foreach (var nodeName in navigationPath)
                {
                    if (!current.ContainsKey(nodeName))
                    {
                        _logger.LogWarning($"Navigation failed: '{nodeName}' not found in current path.");
                        return null;
                    }

                    var value = current[nodeName];

                    if (value is List<ExpandoObject> list)
                    {
                        if (list.Count == 0)
                        {
                            _logger.LogWarning($"Empty list found at '{nodeName}'.");
                            return null;
                        }

                        // If this is the last item in path, we've reached the repeating group
                        if (nodeName == navigationPath.Last())
                            return list;

                        // Otherwise navigate into the first item and continue
                        current = list[0] as IDictionary<string, object>;
                    }
                    else
                    {
                        _logger.LogWarning($"Expected list at '{nodeName}', found: {value?.GetType().Name}");
                        return null;
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to navigate to group");
                return null;
            }
        }

        /// <summary>
        /// Builds the navigation path from document root to repeating group.
        /// </summary>
        private List<string> BuildNavigationPath(XmlNode documentElement, XmlNode repeatingGroupParent,
                                                 string groupName, string documentRootName)
        {
            List<string> navigationPath = new List<string>();
            Stack<string> pathStack = new Stack<string>();

            // Trace back from repeating group parent to document root
            XmlNode currentXmlNode = repeatingGroupParent;
            while (currentXmlNode != null && currentXmlNode != documentElement)
            {
                pathStack.Push(currentXmlNode.LocalName);
                currentXmlNode = currentXmlNode.ParentNode;
            }

            // Start path from document root
            navigationPath.Add(documentRootName);

            // Add intermediate levels in correct order
            while (pathStack.Count > 0)
                navigationPath.Add(pathStack.Pop());

            // Add the repeating group name
            navigationPath.Add(groupName);

            return navigationPath;
        }
        /// <summary>
        /// Detects whether group items contain nested repeating groups.
        /// </summary>
        private bool DetectNestedGroups(List<ExpandoObject> groupItems)
        {
            foreach (var item in groupItems)
            {
                var itemDict = item as IDictionary<string, object>;

                foreach (var value in itemDict.Values)
                {
                    if (value is List<ExpandoObject> childList)
                    {
                        if (childList.Count > 1)
                            return true;

                        if (childList.Count == 1)
                        {
                            var nestedItem = childList[0] as IDictionary<string, object>;
                            if (nestedItem?.Values.OfType<List<ExpandoObject>>().Any() == true)
                                return true;
                        }
                    }
                }
            }
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
                var rootDict = parsedData as IDictionary<string, object>;
                string documentRootName = documentElement.LocalName;

                if (!rootDict.ContainsKey(documentRootName))
                {
                    ViewBag.Message = "XML root element not found.";
                    return false;
                }

                var documentRootList = rootDict[documentRootName] as List<ExpandoObject>;
                if (documentRootList == null || documentRootList.Count == 0)
                {
                    ViewBag.Message = "Invalid XML structure.";
                    return false;
                }

                var rootItem = documentRootList[0] as IDictionary<string, object>;
                var (fields, skippedLists) = FlattenRootFields(rootItem);

                if (fields.Count == 0)
                {
                    ViewBag.Message = "No merge fields found in XML.";
                    return false;
                }

                string[] fieldNames = fields.Keys.ToArray();
                string[] fieldValues = fields.Values.ToArray();
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
                string defaultXmlPath = Path.Combine(_hostingEnvironment.WebRootPath, "Data", "PatientDetails-2.xml");
                return new FileStream(defaultXmlPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
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

        private bool ExecuteMailMergeUsingKnownGroup(
    WordDocument document,
    Stream xmlStream,
    string primaryGroupName)
        {
            try
            {
                xmlStream.Position = 0;

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlStream);

                XmlNode rootNode = xmlDoc.DocumentElement;

                // Check nested repeating (Vitals, Medications, etc.)
                bool hasNestedRepeating = ContainsNestedRepeating(rootNode);

                // Convert XML → ExpandoObject
                ExpandoObject parsedData = new ExpandoObject();
                GetDataAsExpandoObject(rootNode, ref parsedData);
                var rootDict = parsedData as IDictionary<string, object>;

                // ✅ IMPORTANT: extract GROUP ITEMS (Patients), not root
                List<ExpandoObject> groupItems =
                    NavigateToGroup(
                        rootDict,
                        rootNode,
                        rootNode,
                        primaryGroupName,
                        rootNode.LocalName);

                if (groupItems == null || groupItems.Count == 0)
                {
                    _logger.LogWarning($"No items found for group '{primaryGroupName}'");
                    return false;
                }

                _logger.LogInformation(
                    $"Found {groupItems.Count} items for group '{primaryGroupName}'");

                MailMergeDataTable table =
                    new MailMergeDataTable(primaryGroupName, groupItems);

                document.MailMerge.StartAtNewPage = true;

                if (hasNestedRepeating)
                    document.MailMerge.ExecuteNestedGroup(table);
                else
                    document.MailMerge.ExecuteGroup(table);

                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to execute mail merge");
                return false;
            }
        }

        private bool ContainsNestedRepeating(XmlNode node)
        {
            var elementChildren = node.ChildNodes
                .OfType<XmlNode>()
                .Where(n => n.NodeType == XmlNodeType.Element)
                .ToList();

            var repeating = elementChildren
                .GroupBy(n => n.LocalName)
                .Any(g => g.Count() > 1);

            if (repeating)
                return true;

            foreach (var child in elementChildren)
                if (ContainsNestedRepeating(child))
                    return true;

            return false;
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
        private class XmlMergeStructure
        {
            public bool HasRepeating { get; set; }
            public bool HasNestedRepeating { get; set; }
            public int RepeatingLevelCount { get; set; }
            public string PrimaryGroupName { get; set; }
        }
    }
}
