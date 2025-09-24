using System;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

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
        public static void InsertText(string srcDocPath, string covertText, string hipsNamespace = "")
        {
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;

            using WordprocessingDocument wdDoc = WordprocessingDocument.Open(srcDocPath, true);
            var mainPart = wdDoc.MainDocumentPart ?? wdDoc.AddMainDocumentPart();
            mainPart.Document = GenerateDocBodyXML(hipsNamespace, new StringReader(mainPart.Document.OuterXml), covertText);
        }

        /// <summary>
        /// Insert text into a new document using existing document as template
        /// </summary>
        /// <param name="srcDocPath">Document to use as template</param>
        /// <param name="dstDocPath">Path of document to create</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <param name="hipsNamespace">Optional parameter to specify namespace for covert text</param>
        public static void InsertText(string srcDocPath, string dstDocPath, string covertText, string hipsNamespace = "")
        {
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;
            StringReader docPartReader;
            using (WordprocessingDocument wdDocSrc = WordprocessingDocument.Open(srcDocPath, false))
            {
                var mainPart = wdDocSrc.MainDocumentPart ?? wdDocSrc.AddMainDocumentPart();
                docPartReader = new StringReader(mainPart.Document.OuterXml);
            }

            using WordprocessingDocument wdDocDest = WordprocessingDocument.Create(dstDocPath, WordprocessingDocumentType.Document);
            var docmain = wdDocDest.AddMainDocumentPart();
            docmain.Document = GenerateDocBodyXML(hipsNamespace, docPartReader, covertText);
        }

        /// <summary>
        /// Given an existing 
        /// </summary>
        /// <param name="hipsNamespace">Namespace where covert text will be stored</param>
        /// <param name="docPartText">A reader that will provide the XML of the Document part</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <returns>Returns a DocumentFormat.OpenXml.Wordprocessing.Document object that has covert string inserted</returns>
        private static Wp.Document GenerateDocBodyXML(string hipsNamespace, TextReader docPartText, string covertText)
        {
            var nt = new NameTable(); // Manage namespaces to perform XPath queries.
            var nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", wordmlNamespace);
            nsManager.AddNamespace("hips", hipsNamespace);
            nsManager.AddNamespace("mc", mcNamespace);

            var xdoc = new XmlDocument(nt);
            xdoc.Load(docPartText);
            if (xdoc.DocumentElement == null)
            {
                throw new InvalidOperationException("DocumentElement was null");
            }
            XmlNode mcIgnorable = xdoc.DocumentElement.Attributes.GetNamedItem("Ignorable", mcNamespace) ?? xdoc.CreateAttribute("mc", "Ignorable", mcNamespace);
            mcIgnorable.Value = (mcIgnorable.Value + " " + nsManager.LookupPrefix(hipsNamespace)).TrimStart();
            xdoc.DocumentElement.Attributes.SetNamedItem(mcIgnorable);

            XmlNode hiContent = xdoc.CreateNode(XmlNodeType.Element, "hips", "t", hipsNamespace);
            hiContent.InnerText = covertText;
            var firstParagraph = xdoc.SelectSingleNode("/w:document[1]/w:body[1]/w:p[1]", nsManager) ?? throw new InvalidOperationException("Unable to find first paragraph"); //get first paragraph in document element
            firstParagraph.AppendChild(hiContent);

            return new Wp.Document(xdoc.InnerXml);
        }

        /// <summary>
        /// Create a blank document and insert text into it
        /// </summary>
        /// <param name="filePath">Path and filename to create</param>
        /// <param name="covertText">Text to be inserted</param>
        /// <param name="hipsNamespace">Optional parameter to specify namespace for covert text</param>
        public static void createFileInsertText(string filePath, string covertText, string hipsNamespace = "")
        {
            var blankDoc = new Wp.Document(
                new Wp.Body(
                    new Wp.Paragraph(
                        new Wp.Run())));
            hipsNamespace = string.IsNullOrEmpty(hipsNamespace) ? hipsDOCX.hipsNamespace : hipsNamespace;
            var xdoc = GenerateDocBodyXML(hipsNamespace, new StringReader(blankDoc.OuterXml), covertText);

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Wp.Document(xdoc.OuterXml);
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

            using WordprocessingDocument wdDoc = WordprocessingDocument.Open(docxPath, false);
            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", wordmlNamespace);
            nsManager.AddNamespace("hips", hipsNamespace);
            nsManager.AddNamespace("mc", mcNamespace);

            XmlDocument xdoc = new XmlDocument(nt);

            if (wdDoc.MainDocumentPart == null)
            {
                throw new InvalidOperationException("MainDocumentPart was null");
            }

            xdoc.Load(new StringReader(wdDoc.MainDocumentPart.Document.OuterXml));

            var hipsNodes = xdoc.SelectNodes("//hips:t", nsManager) ?? throw new InvalidOperationException("SelectNodes resulted in null XmlNodeList");
            var sr = new StringWriter();

            foreach (XmlNode hipsText in hipsNodes)
            {
                sr.Write(hipsText.InnerText);
            }

            return sr.GetStringBuilder().ToString();
        }
    }
}