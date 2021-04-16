using System;
using System.CommandLine;
using inv = System.CommandLine.Invocation;

namespace hips
{
#nullable enable
    class Program
    {
        static int Main(string[] args)
        {
            RootCommand cmd = new RootCommand
            {
                new Command("insert", "Insert covert data into overt file")
                {
                    new Command("docx", "Modify DOCX file with covert data inserted")
                    {
                        new Argument<string>("documentPath", "Path to source document."),
                        new Argument<string>("covertText", "Covert text to insert."),
                        new Option<string?>(new[] { "--hipsNamespace", "-ns" }, "The namespace for covert text in OpenXML file.")
                    }.WithHandler(new Action<string,string>((documentPath, covertText) => {
                        hipsDOCX.insertText(documentPath, covertText);
                      }))
                },
                new Command("generate", "Create overt file with covert data inserted")
                {
                    new Command("from-existing", "Create overt file with covert data inserted using existing file")
                    {
                        new Command("docx", "Create DOCX file with covert data inserted")
                        {
                            new Argument<string>("sourceDocumentPath", "Path to source document."),
                            new Argument<string>("destinationDocumentPath", "Path of DOCX to generate."),
                            new Argument<string>("covertText", "Covert text to insert."),
                            new Option<string?>(new[] { "--hipsNamespace", "-ns" }, "The namespace for covert text in OpenXML file.")
                        }.WithHandler(new Action<string,string,string,string?>((sourceDocumentPath, destinationDocumentPath, covertText,hipsNamespace) =>
                            {
                                hipsDOCX.insertText(sourceDocumentPath,destinationDocumentPath, covertText,hipsNamespace);
                            })),
                        new Command("html", "Create HTML file with covert data inserted")
                        {
                            new Argument<string>("sourceHTML", "Path to source HTML."),
                            new Argument<string>("destinationHTML", "Path of HTML to generate."),
                            new Argument<string>("covertPath", "Path to covert file to insert.")
                        }.WithHandler(new Action<string,string,string>((sourceHTML, destinationHTML, covertPath) => {
                            hipsHTML.hideInHTML(sourceHTML,destinationHTML,covertPath);
                        }))
                    },
                    new Command("new", "Create overt file with covert data inserted")
                    {
                        new Command("docx", "Create DOCX file with covert data inserted")
                        {
                            new Argument<string>("destinationDocumentPath", "Path of DOCX to generate."),
                            new Argument<string>("covertText", "Covert text to insert."),
                            new Option<string?>(new[] { "--hipsNamespace", "-ns" }, "The namespace for covert text in OpenXML file.")
                        }.WithHandler(new Action<string,string,string?>((destinationDocumentPath, covertText,hipsNamespace) =>
                            {
                                hipsDOCX.createFileInsertText(destinationDocumentPath, covertText,hipsNamespace);
                            }))
                    }
                },
                new Command("extract", "Extract covert data from overt file")
                {
                    new Command("docx", "Extract data from DOCX file")
                    {
                        new Argument<string>("overtDOCX", "Path to source document."),
                        new Option<string?>(new[] { "--hipsNamespace", "-ns" }, "The namespace for covert text in OpenXML file.")
                    }.WithHandler(new Action<string,string,IConsole>((overtDOCX,hipsNamespace, console)=>
                        {
                            try
                            {
                                console.Out.Write(hipsDOCX.getText(overtDOCX,hipsNamespace));
                            }
                            catch (Exception ex)
                            {
                                console.Error.Write(ex.Message);
                            }
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
#nullable disable

    static class extensions
    {
        public static Command WithHandler(this Command command, Delegate handlerFunc)
        {
            command.Handler = inv.CommandHandler.Create(handlerFunc!);
            return command;
        }
    }
}