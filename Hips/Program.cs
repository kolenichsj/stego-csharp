using System;
using System.CommandLine;
using System.IO;

namespace Hips
{
    class Program
    {
        static int Main(string[] args)
        {
            var cmd = new RootCommand
            {
               new Command("insert", "Insert covert data into overt file")
               {
                   new Command("docx", "Modify DOCX file with covert data inserted")
                   {
                       new Argument<FileInfo>("documentPath"){ Description = "Path to source document."},
                       new Argument<string>("covertText") { Description = "Covert text to insert."},
                       new Option<string?>("--hipsNamespace", ["-ns"]){ Description = "The namespace for covert text in OpenXML file." }
                   }
                   .WithAction(parseResult =>
                   {
                       var documentPath = parseResult.GetValue<FileInfo>("documentPath");
                       var covertText = parseResult.GetValue<string>("covertText");
                       var hipsNamespace = parseResult.GetValue<string>("--hipsNamespace");
                       if (documentPath==null)
                       {
                           Console.Error.WriteLine("Value for argument documentPath is required");
                           return 1;
                       }

                       if (string.IsNullOrEmpty(covertText))
                       {
                           Console.Error.WriteLine("Value for argument covertText is required");
                           return 1;
                       }

                       try
                       {
                           HipsDOCX.InsertText(documentPath.FullName, covertText, hipsNamespace?? string.Empty);
                           return 0;
                       }
                       catch (Exception ex)
                       {
                           Console.Error.Write(ex.Message);
                           return 1;
                       }
                   })
               },
               new Command("generate", "Create overt file with covert data inserted")
               {
                   new Command("from-existing", "Create overt file with covert data inserted using existing file")
                   {
                       new Command("docx", "Create DOCX file with covert data inserted")
                       {
                           new Argument<FileInfo>("sourceDocumentPath") {Description = "Path to source document." },
                           new Argument<FileInfo>("destinationDocumentPath"){Description = "Path of DOCX to generate." },
                           new Argument<string>("covertText"){Description =  "Covert text to insert." },
                           new Option<string?>("--hipsNamespace", ["-ns"]){Description ="The namespace for covert text in OpenXML file." }
                       }
                       .WithAction(parseResult => {
                           var sourceDocumentPath = parseResult.GetValue<FileInfo>("sourceDocumentPath");
                           var destinationDocumentPath = parseResult.GetValue<FileInfo>("destinationDocumentPath");
                           var covertText = parseResult.GetValue<string>("covertText");
                           var hipsNamespace = parseResult.GetValue<string>("--hipsNamespace");

                           if (sourceDocumentPath==null ||destinationDocumentPath==null || string.IsNullOrEmpty(covertText))
                           {
                               Console.Error.WriteLine("Missing required value");
                               return 1;
                           }

                           try
                           {
                               HipsDOCX.InsertText(sourceDocumentPath.FullName,destinationDocumentPath.FullName, covertText,hipsNamespace??string.Empty);
                               return 0;
                           }
                           catch (Exception ex)
                           {
                               Console.Error.Write(ex.Message);
                               return 1;
                           }
                       }),
                       new Command("html", "Create HTML file with covert data inserted")
                       {
                           new Argument<FileInfo>("sourceHTML"){Description = "Path to source HTML." },
                           new Argument<FileInfo>("destinationHTML"){Description = "Path of HTML to generate." },
                           new Argument<FileInfo>("covertPath"){Description = "Path to covert file to insert." }
                       }
                       .WithAction(parseResult => {
                           var sourceHTML = parseResult.GetValue<FileInfo>("sourceHTML");
                           var destinationHTML = parseResult.GetValue<FileInfo>("destinationHTML");
                           var covertPath = parseResult.GetValue<FileInfo>("covertPath");
                           if (sourceHTML==null ||destinationHTML==null || covertPath==null)
                           {
                               Console.Error.WriteLine("Missing required value");
                               return 1;
                           }

                           try
                           {
                               HipsHTML.hideInHTML(sourceHTML.FullName,destinationHTML.FullName,covertPath.FullName);
                               return 0;
                           }
                           catch (Exception ex)
                           {
                               Console.Error.Write(ex.Message);
                               return 1;
                           }

                       })
                       },
                   new Command("new", "Create overt file with covert data inserted")
                   {
                       new Command("docx", "Create DOCX file with covert data inserted")
                       {
                           new Argument<FileInfo>("destinationDocumentPath"){Description = "Path of DOCX to generate." },
                           new Argument<string>("covertText"){Description = "Covert text to insert." },
                           new Option<string?>("--hipsNamespace", ["-ns"] ){Description = "The namespace for covert text in OpenXML file."}
                       }
                       .WithAction(parseResult => {
                           var destinationDocumentPath = parseResult.GetValue<FileInfo>("destinationDocumentPath");
                           var covertText = parseResult.GetValue<string>("covertText");
                           var hipsNamespace = parseResult.GetValue<string>("--hipsNamespace");
                           if (destinationDocumentPath==null ||string.IsNullOrEmpty(covertText))
                           {
                               Console.Error.WriteLine("Missing required value");
                               return 1;
                           }

                           try
                           {
                               HipsDOCX.CreateFileInsertText(destinationDocumentPath.FullName, covertText,hipsNamespace?? string.Empty);
                               return 0;
                           }
                           catch (Exception ex)
                           {
                               Console.Error.Write(ex.Message);
                               return 1;
                           }
                       })
                   }
               },
               new Command("extract", "Extract covert data from overt file")
               {
                   new Command("docx", "Extract data from DOCX file")
                   {
                       new Argument<FileInfo>("overtDOCX"){Description = "Path to source document." },
                       new Option<string?>("--hipsNamespace", ["-ns"]){Description = "The namespace for covert text in OpenXML file." }
                   }
                   .WithAction(parseResult => {
                       var overtDOCX = parseResult.GetValue<FileInfo>("overtDOCX");
                       var hipsNamespace = parseResult.GetValue<string>("--hipsNamespace");

                       if (overtDOCX==null)
                       {
                           Console.Error.WriteLine("Value for argument overtDOCX is required");
                           return 1;
                       }

                       try
                       {
                           Console.Out.Write(HipsDOCX.GetText(overtDOCX.FullName,hipsNamespace??string.Empty));
                           return 0;
                       }
                       catch (Exception ex)
                       {
                           Console.Error.Write(ex.Message);
                           return 1;
                       }
                   }),
                   new Command("html", "Extract data from HTML file")
                   {
                       new Argument<FileInfo>("overtHTML"){Description = "Path to overt HTML." },
                       new Argument<FileInfo>("covertPath"){Description = "Path to extract covert file." }
                   }
                  .WithAction(parseResult => {
                      var overtHTML = parseResult.GetValue<FileInfo>("overtHTML");
                      var covertPath = parseResult.GetValue<FileInfo>("covertPath");
                      if (overtHTML==null || covertPath == null)
                      {
                          Console.Error.WriteLine("Value for argument overtDOCX is required");
                          return 1;
                      }

                      try
                      {
                          HipsHTML.getFromHTML(overtHTML.FullName,covertPath.FullName);
                          return 0;
                      }
                      catch (Exception ex)
                      {
                          Console.Error.Write(ex.Message);
                          return 1;
                      }
                   })
               }
            };

            ParseResult parseResult = cmd.Parse(args);
            return parseResult.Invoke();
        }
    }


    static class Extensions
    {
        public static Command WithAction(this Command command, Action<ParseResult> action)
        {
            command.SetAction(action);
            return command;
        }

        public static Command WithAction(this Command command, Func<ParseResult, int> action)
        {
            command.SetAction(action);
            return command;
        }
    }
}