using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using FormsDemo.ErrorHandlers;
using FormsDemo.Helpers;
using FormsDemo.Interfaces;
using FormsDemo.Models;
using Microsoft.Office.Interop.Word;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Document = Microsoft.Office.Interop.Word.Document;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Task = System.Threading.Tasks.Task;

namespace FormsDemo.Services
{
    public class WordToPdfService : IWordToPdfService
    {
        public async Task<byte[]> GenerateReport(ReportRequest param)
        {
            byte[] dataBytes;
            var templateFullPath = Path.Combine(Directory.GetCurrentDirectory(), @"DocumentTemplates", param.TemplateName);
            if (IsPdf(templateFullPath))
            {
                using (var stream = new FileStream(templateFullPath, FileMode.Open))
                {
                    dataBytes = ReadFully(stream);
                }
                return dataBytes;
            }

            var fileServerPhysical = "\\..\\{ProjectName}_cdn";
            var currentDirectory = new DirectoryInfo(Environment.CurrentDirectory);
            var cdnDirectory = new DirectoryInfo(
                currentDirectory.FullName +
                fileServerPhysical.Replace("{ProjectName}", currentDirectory.Name));

            // create if it doesnt exist
            Directory.CreateDirectory(cdnDirectory.FullName);

            fileServerPhysical = cdnDirectory.FullName;

            var outputFilePath = $"{fileServerPhysical}\\{Guid.NewGuid()}.docx";

            File.Copy(templateFullPath, outputFilePath, true);

            using (var doc = WordprocessingDocument.Open(outputFilePath, true))
            {
                try
                {
                    var headers = doc.MainDocumentPart.HeaderParts;
                    var footers = doc.MainDocumentPart.FooterParts;

                    var bookmarkList = new List<BookmarkStart>();

                    foreach (var eachFooter in footers)
                        bookmarkList.AddRange(eachFooter.RootElement
                            .Descendants<BookmarkStart>());

                    foreach (var eachHeader in headers)
                        bookmarkList.AddRange(eachHeader.RootElement
                            .Descendants<BookmarkStart>());

                    bookmarkList.AddRange(doc.MainDocumentPart.Document
                        .Descendants<BookmarkStart>());

                    foreach (var eachBookmarkStart in bookmarkList)
                        if (!eachBookmarkStart.Name.Value.StartsWith("_"))
                        {
                            var paragraph = eachBookmarkStart.Parent;
                            if (param.Values?.ContainsKey(RemoveNumbers(eachBookmarkStart.Name.Value)) == true)
                            {
                                var value = param.Values[RemoveNumbers(eachBookmarkStart.Name.Value)];

                                var runElement = new Run();
                                if (!string.IsNullOrEmpty(value) && value.Contains("\n"))
                                {
                                    var eachString = value.Split('\n');
                                    foreach (var eachLine in eachString)
                                        if (string.IsNullOrEmpty(eachLine))
                                        {
                                            runElement.AppendChild(new Break());
                                        }
                                        else
                                        {
                                            runElement.AppendChild(new Text(eachLine));
                                            runElement.AppendChild(new Break());
                                        }
                                }
                                else
                                {
                                    runElement.AppendChild(new Text(value));
                                }

                                if (paragraph
                                    .Descendants<ParagraphMarkRunProperties>()
                                    .Any())
                                {
                                    var paragraphMarkRunProperties =
                                        (ParagraphMarkRunProperties)paragraph
                                            .Descendants<ParagraphMarkRunProperties>()
                                            .ElementAt(0).Clone();
                                    runElement
                                        .PrependChild(
                                            paragraphMarkRunProperties);
                                }

                                paragraph.InsertAfter(runElement, eachBookmarkStart);
                            }
                        }

                    if (param.TableValues != null)
                        foreach (var eachTableValues in param.TableValues)
                        {
                            var bookMarkPrefixName = "Table_" + eachTableValues.Key;
                            var bookMark =
                                bookmarkList.FirstOrDefault(b => b.Name.Value.StartsWith(bookMarkPrefixName));
                            if (bookMark != null)
                            {
                                var tableElement = bookMark.Parent;
                                while (!(tableElement is Table)) tableElement = tableElement.Parent;

                                var listValues = eachTableValues.Value;
                                if (listValues.Count == 0)
                                {
                                    var previousElement = tableElement.PreviousSibling();
                                    previousElement?.Remove();
                                    tableElement.Remove();
                                }

                                if (listValues.Count > 0)
                                {
                                    var lastRow = tableElement.Elements<TableRow>().Last();
                                    foreach (var eachListValue in listValues)
                                    {
                                        var rowCopy = (TableRow)lastRow.CloneNode(true);
                                        foreach (var eachBookmarkStart in rowCopy
                                            .Descendants<BookmarkStart>())
                                            if (!eachBookmarkStart.Name.Value.StartsWith("_"))
                                            {
                                                var paragraph = eachBookmarkStart.Parent;
                                                if (eachBookmarkStart.Name.Value.StartsWith("Table_"))
                                                {
                                                    var keyName =
                                                        eachBookmarkStart.Name.Value.Replace(
                                                            "Table_" + eachTableValues.Key + "_", "");
                                                    var value = eachListValue[keyName];
                                                    var runElement = new Run(new Text(value));
                                                    if (paragraph
                                                        .Descendants<ParagraphMarkRunProperties>().Any())
                                                    {
                                                        var paragraphMarkRunProperties =
                                                            (ParagraphMarkRunProperties)paragraph
                                                                .Descendants<ParagraphMarkRunProperties>().ElementAt(0)
                                                                .Clone();
                                                        runElement
                                                            .PrependChild(paragraphMarkRunProperties);
                                                    }

                                                    paragraph.InsertAfter(runElement, eachBookmarkStart);
                                                }
                                            }

                                        tableElement.AppendChild(rowCopy);
                                    }

                                    tableElement.RemoveChild(lastRow);
                                }
                            }
                        }

                    doc.MainDocumentPart.Document.Save();

                }
                finally
                {
                    doc.Save();
                    doc.Close();
                }
            }
            string test = outputFilePath.Replace("docx", "pdf");
            var pdfPath = await ConvertWordToPdf(outputFilePath, outputFilePath.Replace("docx", "pdf"));

            if (!param.AuthorName.IsEmpty() || !param.Keywords.IsEmpty())
            {
                PdfDocument document = PdfReader.Open(pdfPath);
                document.Info.Author = param.AuthorName;
                document.Info.Keywords = param.Keywords;
                document.Save(pdfPath);
            }

            using (var stream = new FileStream(pdfPath, FileMode.Open))
            {
                dataBytes = ReadFully(stream);
            }

            File.Delete(pdfPath);

            if (dataBytes == null) throw new Exception("Error occurred upon generation of PDF report.");
            return dataBytes;
        }

        private bool IsPdf(string path)
        {
            var extension = Path.GetExtension(path);
            return extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        }

        private string RemoveNumbers(string input)
        {
            const string pattern = @"\d+$";
            var rgx = new Regex(pattern);
            return rgx.Replace(input, "");
        }

        private static byte[] ReadFully(Stream input)
        {
            var buffer = new byte[16 * 1024];
            using (var ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0) ms.Write(buffer, 0, read);
                return ms.ToArray();
            }
        }

        public async Task<string> ConvertWordToPdf(string source, string destination)
        {
            return await Task.Run(() =>
            {
                Document wordDocument = null;
                var appWord = new Application();

                try
                {
                    wordDocument = appWord.Documents.Open(source);
                    var directory = new FileInfo(destination).DirectoryName;
                    if (!Directory.Exists(directory))
                        new DirectoryInfo(directory).Create();
                    wordDocument.ExportAsFixedFormat(destination, WdExportFormat.wdExportFormatPDF);
                }
                finally
                {
                    FinalizeApplication(wordDocument, source, appWord);
                }

                return destination;
            });
        }

        private void FinalizeApplication(Document wordDocument, string fileCopy, Application appWord)
        {
            wordDocument?.Close(false);
            appWord?.Quit(false);
            File.Delete(fileCopy);
        }
    }
}
