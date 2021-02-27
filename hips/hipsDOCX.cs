using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using wp = DocumentFormat.OpenXml.Wordprocessing;

namespace hips
{
    /// <summary>
    /// Class to process DOCX documents to insert or retrieve covert data
    /// </summary>
    public static class hipsDOCX
    {
        const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mcNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        static string hipsNamespace = "http://schemas.microsoft.com/office/rfc/3514/true";

        /// <summary>
        /// Insert text into an existing document
        /// </summary>
        /// <param name="srcDocPath">Path of document to modify</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <param name="hipsNamespace">Optional parameter to specify namespace for covert text</param>
        public static void insertText(string srcDocPath, string covertText, string hipsNamespace = "")
        {
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(srcDocPath, true))
            {
                wdDoc.MainDocumentPart.Document = generateDocBodyXML(hipsNamespace, new StringReader(wdDoc.MainDocumentPart.Document.OuterXml), covertText);
            }
        }

        /// <summary>
        /// Insert text into a new document using existing document as template
        /// </summary>
        /// <param name="srcDocPath">Document to use as template</param>
        /// <param name="dstDocPath">Path of document to create</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <param name="hipsNamespace">Optional parameter to specify namespace for covert text</param>
        public static void insertText(string srcDocPath, string dstDocPath, string covertText, string hipsNamespace = "")
        {
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;
            StringReader docPartReader;
            using (WordprocessingDocument wdDocSrc = WordprocessingDocument.Open(srcDocPath, false))
            {
                docPartReader = new StringReader(wdDocSrc.MainDocumentPart.Document.OuterXml);
            }

            using (WordprocessingDocument wdDocDest = WordprocessingDocument.Open(dstDocPath, true))
            {
                wdDocDest.MainDocumentPart.Document = generateDocBodyXML(hipsNamespace, docPartReader, covertText);
            }
        }

        /// <summary>
        /// Given an existing 
        /// </summary>
        /// <param name="hipsNamespace">Namespace where covert text will be stored</param>
        /// <param name="docPartText">A reader that will provide the XML of the Document part</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <returns>Returns a DocumentFormat.OpenXml.Wordprocessing.Document object that has covert string inserted</returns>
        private static wp.Document generateDocBodyXML(string hipsNamespace, TextReader docPartText, string covertText)
        {
            NameTable nt = new NameTable(); // Manage namespaces to perform XPath queries.
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", wordmlNamespace);
            nsManager.AddNamespace("hips", hipsNamespace);
            nsManager.AddNamespace("mc", mcNamespace);

            XmlDocument xdoc = new XmlDocument(nt);
            xdoc.Load(docPartText);
            XmlNode mcIgnorable = xdoc.DocumentElement.Attributes.GetNamedItem("Ignorable", mcNamespace) ?? xdoc.CreateAttribute("mc", "Ignorable", mcNamespace);
            mcIgnorable.Value = (mcIgnorable.Value + " " + nsManager.LookupPrefix(hipsNamespace)).TrimStart();
            xdoc.DocumentElement.Attributes.SetNamedItem(mcIgnorable);

            XmlNode hiContent = xdoc.CreateNode(XmlNodeType.Element, "hips", "t", hipsNamespace);
            hiContent.InnerText = covertText;
            XmlNode firstParagraph = xdoc.SelectSingleNode("/w:document[1]/w:body[1]/w:p[1]", nsManager); //get first paragraph in document element
            firstParagraph.AppendChild(hiContent);

            return new wp.Document(xdoc.InnerXml);
        }

        /// <summary>
        /// Create a blank document and insert text into it
        /// </summary>
        /// <param name="filePath">Path and filename to create</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <param name="hipsNamespace">Optional parameter to specify namespace for covert text</param>
        public static void createFileInsertText(string filePath, string covertText, string hipsNamespace = "")
        {
            var blankDoc = new wp.Document(
                new wp.Body(
                    new wp.Paragraph(
                        new wp.Run())));
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;
            var xdoc = generateDocBodyXML(hipsNamespace, new StringReader(blankDoc.OuterXml), covertText);

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new wp.Document(xdoc.OuterXml);
            }
        }

        /// <summary>
        /// Extract covert data from overt file
        /// </summary>
        /// <param name="docxPath"></param>
        /// <param name="hipsNamespace">Optional parameter to specify namespace for covert text</param>
        /// <returns>Any text found in "t" element in the hipsNamespace</returns>
        public static string getText(string docxPath, string hipsNamespace = "")
        {
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docxPath, true))
            {
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);
                nsManager.AddNamespace("hips", hipsNamespace);
                nsManager.AddNamespace("mc", mcNamespace);

                XmlDocument xdoc = new XmlDocument(nt);
                xdoc.Load(new StringReader(wdDoc.MainDocumentPart.Document.OuterXml));
                XmlNodeList hipsNodes = xdoc.SelectNodes("//hips:t", nsManager);
                StringWriter sr = new StringWriter();

                foreach (XmlNode hipsText in hipsNodes)
                {
                    sr.WriteLine(hipsText.InnerText);
                }

                return sr.GetStringBuilder().ToString();
            }
        }
    }
}