using System;
using System.CommandLine;
using inv = System.CommandLine.Invocation;

namespace hips
{
    class Program
    {
        static int Main(string[] args)
        {
            RootCommand cmd = new RootCommand
            {
                new Command("insert", "Insert covert data into overt file")
                {
                    new Command("docx", "Create DOCX file with covert data inserted")
                    {
                        new Argument<string>("documentPath", "Path to source document."),
                        new Argument<string>("covertText", "Covert text to insert.")
                    }.WithHandler(new Action<string,string>((documentPath, covertText) => {
                        hipsDOCX.insertText(documentPath, covertText);
                      }))
                },
                new Command("generate", "Create overt file with covert data inserted")
                {
                    new Command("docx", "Create DOCX file with covert data inserted")
                    {
                        new Argument<string>("sourceDocumentPath", "Path to source document."),
                        new Argument<string>("destinationDocumentPath", "Path of DOCX to generate."),
                        new Argument<string>("covertText", "Covert text to insert.")
                    }.WithHandler(new Action<string,string,string>((documentPath, destinationDocumentPath, covertText) => {
                        hipsDOCX.insertText(documentPath,destinationDocumentPath, covertText);
                      })),
                    new Command("html", "Create HTML file with covert data inserted")
                    {
                        new Argument<string>("sourceHTML", "Path to source HTML."),
                        new Argument<string>("destinationHTML", "Path of HTML to generate."),
                        new Argument<string>("covertPath", "Path to covert file to insert.")
                    }.WithHandler(new Action<string,string,string>((srcHTML, outHTML, srcBin) => {
                        hipsHTML.hideInHTML(srcHTML,outHTML,srcBin);
                     }))
                },
                new Command("extract", "extract covert data from overt file")
                {
                    new Command("docx", "Extract data from DOCX file")
                    {
                        new Argument<string>("overtDOCX", "Path to source document.")
                    }.WithHandler(new Action<string,IConsole>((overtDOCX, console)=>{
                        string val = hipsDOCX.getText(overtDOCX);
                        console.Out.Write(val);
                    })),
                    new Command("html", "Extract data from HTML file")
                    {
                        new Argument<string>("overtHTML", "Path to overt HTML."),
                        new Argument<string>("covertPath", "Path to extract covert file.")
                    }.WithHandler(new Action<string,string>((overtHTML,covertPath) => {
                            hipsHTML.getFromHTML(overtHTML,covertPath);
                        }))
                }
            };

            return cmd.Invoke(args);
        }
    }

    static class extensions
    {
        public static Command WithHandler(this Command command, Delegate handlerFunc)
        {
            command.Handler = inv.CommandHandler.Create(handlerFunc!);
            return command;
        }
    }
}