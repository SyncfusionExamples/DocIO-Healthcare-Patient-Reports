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
        public IActionResult GenerateReport(IFormFile file, IFormFile xmlFile, string OutputType)
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

                return CreatePDF(wordStream, xmlStream, OutputType);
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
                return new FileStream(defaultFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            }
        }

        /// <summary>
        /// Creates a PDF document by performing a mail merge on a Word template using XML data.
        /// </summary>

        private IActionResult CreatePDF(Stream stream, Stream xmlStream, string type)
        {
            // Validate the output type parameter
            if (string.IsNullOrWhiteSpace(type))
            {
                ViewBag.Message = "Document type parameter is required.";
                return View("Index");
            }
            // Load the Word document from the provided stream
            using (WordDocument document = new WordDocument(stream, FormatType.Automatic))
            {
                // If XML data is provided, perform mail merge
                if (xmlStream != null && xmlStream.Length > 0)
                {
                    xmlStream.Position = 0;
                    // Load XML document from stream
                    XmlDocument xmlDocument = new XmlDocument();
                    xmlDocument.Load(xmlStream);

                    XmlNode documentElement = xmlDocument.DocumentElement;

                    // STEP 1: Find the repeating group node
                    XmlNode repeatingGroupNode = null;
                    XmlNode repeatingGroupParent = null;
                    // First attempt: repeating elements directly under root
                    foreach (XmlNode child in documentElement.ChildNodes)
                    {
                        if (child.NodeType == XmlNodeType.Element &&
                            documentElement.SelectNodes(child.LocalName).Count > 1)
                        {
                            repeatingGroupNode = child;
                            repeatingGroupParent = documentElement;
                            break;
                        }
                    }
                    // Second attempt: repeating elements under wrapper elements
                    if (repeatingGroupNode == null)
                    {
                        foreach (XmlNode wrapper in documentElement.ChildNodes)
                        {
                            if (wrapper.NodeType != XmlNodeType.Element)
                                continue;

                            foreach (XmlNode child in wrapper.ChildNodes)
                            {
                                if (child.NodeType == XmlNodeType.Element &&
                                    wrapper.SelectNodes(child.LocalName).Count > 1)
                                {
                                    repeatingGroupNode = child;
                                    repeatingGroupParent = wrapper;
                                    break;
                                }
                            }

                            if (repeatingGroupNode != null)
                                break;
                        }
                    }

                    // STEP 2: Convert XML into dynamic ExpandoObject structure
                    ExpandoObject parsedData = new ExpandoObject();
                    GetDataAsExpandoObject(documentElement, ref parsedData);

                    // STEP 3: Extract repeating group records for mail merge
                    List<ExpandoObject> groupItems = null;
                    string groupName = null;

                    if (repeatingGroupNode != null)
                    {
                        // Capture node names used for navigation
                        groupName = repeatingGroupNode.LocalName;  
                        string parentNodeName = repeatingGroupParent.LocalName; 
                        string documentRootName = documentElement.LocalName; 

                        try
                        {
                            var rootDict = parsedData as IDictionary<string, object>;

                            // Navigate from document root
                            if (rootDict.ContainsKey(documentRootName))
                            {
                                var documentRootList = rootDict[documentRootName] as List<ExpandoObject>;
                                if (documentRootList != null && documentRootList.Count > 0)
                                {
                                    var documentRootItem = documentRootList[0] as IDictionary<string, object>;

                                    // Check if parent is same as document root or nested
                                    if (documentRootName == parentNodeName)
                                    {
                                        // Structure 1: Components directly under Root
                                        if (documentRootItem.ContainsKey(groupName) && documentRootItem[groupName] is List<ExpandoObject>)
                                        {
                                            groupItems = documentRootItem[groupName] as List<ExpandoObject>;
                                        }
                                    }
                                    else
                                    {
                                        // Structure 2: Components under Validator under Root
                                        if (documentRootItem.ContainsKey(parentNodeName))
                                        {
                                            var parentList = documentRootItem[parentNodeName] as List<ExpandoObject>;
                                            if (parentList != null && parentList.Count > 0)
                                            {
                                                var parentItem = parentList[0] as IDictionary<string, object>;
                                                if (parentItem.ContainsKey(groupName) && parentItem[groupName] is List<ExpandoObject>)
                                                {
                                                    groupItems = parentItem[groupName] as List<ExpandoObject>;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Failed to extract group items");
                        }
                    }

                    // STEP 4: Determine merge strategy (nested, grouped, or simple)
                    if (groupItems != null && groupItems.Count > 0)
                    {
                        bool hasNestedGroups = false;
                        // Inspect whether child collections themselves contain lists
                        foreach (var item in groupItems)
                        {
                            var itemDict = item as IDictionary<string, object>;

                            foreach (var value in itemDict.Values)
                            {
                                if (value is List<ExpandoObject> childList)
                                {
                                    if (childList.Count > 1)
                                    {
                                        hasNestedGroups = true;
                                        break;
                                    }

                                    if (childList.Count == 1)
                                    {
                                        var nestedItem = childList[0] as IDictionary<string, object>;
                                        if (nestedItem?.Values.OfType<List<ExpandoObject>>().Any() == true)
                                        {
                                            hasNestedGroups = true;
                                            break;
                                        }
                                    }
                                }
                            }

                            if (hasNestedGroups) break;
                        }

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
                    }
                    else
                    {
                        // CASE 4: No groups found - simple merge with root fields
                        try
                        {
                            var rootDict = parsedData as IDictionary<string, object>;
                            string documentRootName = documentElement.LocalName;

                            if (rootDict.ContainsKey(documentRootName))
                            {
                                var documentRootList = rootDict[documentRootName] as List<ExpandoObject>;
                                if (documentRootList != null && documentRootList.Count > 0)
                                {
                                    var rootItem = documentRootList[0] as IDictionary<string, object>;

                                    // Flatten all fields from root element
                                    var fields = new Dictionary<string, string>();
                                    foreach (var kvp in rootItem)
                                    {
                                        // Handle simple string values
                                        if (kvp.Value is string stringValue)
                                        {
                                            fields[kvp.Key] = stringValue;
                                        }
                                        // Handle single-item lists (flatten them)
                                        else if (kvp.Value is List<ExpandoObject> list && list.Count == 1)
                                        {
                                            var nestedItem = list[0] as IDictionary<string, object>;
                                            if (nestedItem != null)
                                            {
                                                foreach (var nestedKvp in nestedItem)
                                                {
                                                    fields[nestedKvp.Key] = nestedKvp.Value?.ToString() ?? string.Empty;
                                                }
                                            }
                                        }
                                    }

                                    if (fields.Count > 0)
                                    {
                                        string[] fieldNames = fields.Keys.ToArray();
                                        string[] fieldValues = fields.Values.ToArray();
                                        document.MailMerge.Execute(fieldNames, fieldValues);
                                    }
                                    else
                                    {
                                        ViewBag.Message = "No merge fields found in XML.";
                                        return View("Index");
                                    }
                                }
                                else
                                {
                                    ViewBag.Message = "Invalid XML structure.";
                                    return View("Index");
                                }
                            }
                            else
                            {
                                ViewBag.Message = "XML root element not found.";
                                return View("Index");
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex, "Failed to perform simple mail merge");
                            ViewBag.Message = "Error processing simple merge fields.";
                            return View("Index");
                        }
                    }
                }
                // STEP 5: Generate output based on requested type
                if (string.Equals(type, "single", StringComparison.OrdinalIgnoreCase))
                {
                    MemoryStream pdfStream = SaveAsPDF(document);
                    return File(pdfStream, "application/pdf", "GeneratedDocument.pdf");
                }
                else if (string.Equals(type, "multiple", StringComparison.OrdinalIgnoreCase))
                {
                    // Return ZIP if split succeeded, otherwise fallback to single PDF
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
