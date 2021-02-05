using DocumentFormat.OpenXml.Packaging;
using System.IO;
using System.Xml;

namespace hips
{
    public static class hipsDOCX
    {
        // Given a document name, delete all the hidden text.
        const string wordmlNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        const string mcNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        static string hipsNamespace = "http://schemas.microsoft.com/office/rfc/3514/true";

        public static void insertText(string srcDocPath, string covertText, string hipsNS = "")
        {

            hipsNS = string.IsNullOrEmpty(hipsNS) ? hipsDOCX.hipsNamespace : hipsNS;
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(srcDocPath, true))
            {
                XmlDocument xdoc = genDocBodyXML(hipsNS, wdDoc, covertText);

                // Save the document XML back to its document part.
                xdoc.Save(wdDoc.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write));
            }
        }

        public static void insertText(string srcDocPath, string dstDocPath, string covertText, string hipsNS = "")
        {
            hipsNS = string.IsNullOrEmpty(hipsNS) ? hipsDOCX.hipsNamespace : hipsNS;
            XmlDocument xdoc;
            using (WordprocessingDocument wdDocSrc = WordprocessingDocument.Open(srcDocPath, false))
            {
                xdoc = genDocBodyXML(hipsNS, wdDocSrc, covertText);
            }

            // Save the document XML back to its document part.
            using (WordprocessingDocument wdDocDest = WordprocessingDocument.Open(dstDocPath, true))
            {
                xdoc.Save(wdDocDest.MainDocumentPart.GetStream(FileMode.Create, FileAccess.Write));
            }
        }

        private static XmlDocument genDocBodyXML(string hipsNS, WordprocessingDocument wdDoc, string covertText)
        {
            // Manage namespaces to perform XPath queries.
            NameTable nt = new NameTable();
            XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
            nsManager.AddNamespace("w", wordmlNamespace);
            nsManager.AddNamespace("hips", hipsNS);
            nsManager.AddNamespace("mc", mcNamespace);

            // Get the document part from the package.
            // Load the XML in the document part into an XmlDocument instance.
            XmlDocument xdoc = new XmlDocument(nt);
            var docstream = wdDoc.MainDocumentPart.GetStream(FileMode.Open, FileAccess.Read);
            xdoc.Load(docstream);
            docstream.Close();
            XmlNode firstParagraph = xdoc.SelectSingleNode("/w:document[1]/w:body[1]/w:p[1]", nsManager); //get first paragraph in document element
            XmlNode hiContent = xdoc.CreateNode(XmlNodeType.Element, "hips", "t", hipsNamespace);
            XmlNode mcIgnorable = xdoc.DocumentElement.Attributes.GetNamedItem("Ignorable", mcNamespace);
            mcIgnorable.Value += " " + nsManager.LookupPrefix(hipsNamespace);
            xdoc.DocumentElement.Attributes.SetNamedItem(mcIgnorable);
            hiContent.InnerText = covertText;
            firstParagraph.AppendChild(hiContent);

            return xdoc;
        }

        public static string getText(string docName, string hipsNS = "")
        {
            hipsNS = string.IsNullOrEmpty(hipsNS) ? hipsDOCX.hipsNamespace : hipsNS;

            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(docName, true))
            {
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("w", wordmlNamespace);
                nsManager.AddNamespace("hips", hipsNS);
                nsManager.AddNamespace("mc", mcNamespace);

                StringWriter sr = new StringWriter();

                // Get the document part from the package.
                // Load the XML in the document part into an XmlDocument instance.
                XmlDocument xdoc = new XmlDocument(nt);
                var docstream = wdDoc.MainDocumentPart.GetStream(FileMode.Open, FileAccess.Read);
                xdoc.Load(docstream);
                docstream.Close();
                XmlNodeList hipsNodes = xdoc.SelectNodes("//hips:t", nsManager); //get first paragraph in document element
                foreach (XmlNode hipsText in hipsNodes)
                {
                    sr.WriteLine(hipsText.InnerText);
                }

                return sr.GetStringBuilder().ToString();
            }
        }
    }
}