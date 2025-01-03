using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using System.Xml.Linq;
using CommandLine;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace DocxMerge
{
    class Program
    {
        static void Main(string[] args)
        {
            IEnumerable<string> inputFiles = null;
            string output = null;
            bool verbose = false;
            bool force = false;
            bool repairSpacing = false;

            var results = CommandLine.Parser.Default.ParseArguments<Options>(args);
            results.MapResult(options =>
            {
                inputFiles = options.InputFiles;
                output = options.Output ?? "output.docx";
                verbose = options.Verbose;
                force = options.Force;
                repairSpacing = options.RepairSpacing;
                return 0;
            }, errors =>
            {
                Environment.Exit(1);
                return 1; // ._.
            });


            if (inputFiles.Count() < 2) ExitWithError("There must be at least two input files");
            foreach (var file in inputFiles)
            {
                if (!File.Exists(file))
                    ExitWithError("Unable to find {0}", file);
            }

            try {
                if (verbose) Console.WriteLine("Creating initial document");
                File.Copy(inputFiles.First(), output, force);

                if (verbose) Console.WriteLine("Opening {0} for writing", output);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(output, true))
                {
                    foreach (var filepath in inputFiles.Skip(1))
                    {
                        if (verbose) Console.WriteLine("Adding {0} to {1}", filepath, output);

                        string altChuckId = "id" + Guid.NewGuid().ToString();
                        var mainPart = doc.MainDocumentPart;
                        var chunk = mainPart.AddAlternativeFormatImportPart(
                            AlternativeFormatImportPartType.WordprocessingML,
                            altChuckId);

                        if (repairSpacing)
                            RepairSentenceSpacing(filepath, verbose);

                        using (FileStream fileStream = File.Open(filepath, FileMode.Open))
                            chunk.FeedData(fileStream);                        

                        OpenXmlCompositeElement target = null;
                        
                        try 
                        {
                            target = mainPart.Document.Body.Elements<AltChunk>().Last();
                        }
                        catch
                        {
                            target = mainPart.Document.Body.Elements<Paragraph>().Last();
                        }

                        AltChunk altChunk = new AltChunk();
                        altChunk.Id = altChuckId;

                        mainPart.Document
                            .Body
                            .InsertAfter(altChunk, target);  
                        mainPart.Document.Save();        
                    }
                }
                
                if (verbose) Console.WriteLine("Successfully merged all documents");
            }
            catch (Exception ex)
            {
                if (verbose) ExitWithError(ex.ToString());
                else ExitWithError("DocxMerge failed to process the files: {0}", ex.Message);
            }
        }

        private static void RepairSentenceSpacing(string filepath, bool verbose)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filepath, true))
            {
                string docText = null;
                using (StreamReader sr = new StreamReader(wordDoc.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex(@"\. (?<a>[A-Z])");
                var refIndex = docText.IndexOf("References");
                if (refIndex > 0)
                {
                    var bodyText = docText.Substring(0, refIndex);
                    var references = docText.Substring(bodyText.Length);
                    docText = regexText.Replace(bodyText, ".  ${a}") + references;
                }
                else
                {
                    docText = regexText.Replace(docText, ".  ${a}");
                }               

                using (StreamWriter sw = new StreamWriter(wordDoc.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
            }
        }

        private static XDocument GetXDocument(WordprocessingDocument myDoc)
        {
            // Load the main document part into an XDocument
            XDocument mainDocumentXDoc;
            using (Stream str = myDoc.MainDocumentPart.GetStream())
            using (XmlReader xr = XmlReader.Create(str))
                mainDocumentXDoc = XDocument.Load(xr);
            return mainDocumentXDoc;
        }

        private static void SaveXDocument(WordprocessingDocument myDoc, XDocument mainDocumentXDoc)
        {
            // Serialize the XDocument back into the part
            using (Stream str = myDoc.MainDocumentPart.GetStream(
                FileMode.Create, FileAccess.Write))
            using (XmlWriter xw = XmlWriter.Create(str))
                mainDocumentXDoc.Save(xw);
        }

        private static void ExitWithError(string message, params object[] args)
        {
            var defaultColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            if (args.Any())
                Console.WriteLine(message, args);
            else
                Console.WriteLine(message);
            Console.ForegroundColor = defaultColor;
            Environment.Exit(1);
        }

    }
}
