using FormsDemo.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;
using FormsDemo.Models;
using FormsDemo.ErrorHandlers;

namespace FormsDemo.Services
{
    public class PacketBuilderService : IPacketBuilderService
    {
        private readonly IPacketReportService _packetReportService;
        public PacketBuilderService(IPacketReportService packetReportService)
        {
            _packetReportService = packetReportService;
        }
        public async Task<List<byte[]>> BuildPolicyPacket()
        {
            try
            {
                List<Form> forms = new List<Form>()
            {
               new Form()
               {
                  FormName = "MD Amendatory",
                  FormOrder = 1,
                  TemplateName = "MD Amendatory.docx",
                  IsSelected = true,
                  IsMandatory = false,
               },
               new Form()
               {
                  FormName = "ME Amendatory",
                  FormOrder = 2,
                  TemplateName = "ME Amendatory.docx",
                  IsSelected = true,
                  IsMandatory = false,
               },
                new Form()
               {
                  FormName = "NJ Amendatory",
                  FormOrder = 3,
                  TemplateName = "NJ Amendatory.docx",
                  IsSelected = true,
                  IsMandatory = false,
               }
            };

                var formPacketBytes = new List<byte[]>();
                var reportPages = new List<byte[]>();

                foreach (var form in forms)
                {
                    var dynamicPacketItemFormValue = _packetReportService.GetFormValues(form.FormName);
                    var dynamicPacketTableValue = _packetReportService.GetTableValues(form.FormName);
                    var dynamicPacketReportByte = await _packetReportService.GeneratePdfFileAsync(dynamicPacketTableValue, dynamicPacketItemFormValue, form.TemplateName);
                    formPacketBytes.Add(dynamicPacketReportByte);
                }

                if (formPacketBytes.Any(x => x.Length > 0))
                {
                    var documents = GetConsolidatedPagesBytes(formPacketBytes);
                    if (documents.Count() == 0)
                    {
                        throw new CustomErrorException("Page bytes is empty.");
                    }
                    reportPages.Add(documents);
                }

                return reportPages;
            }
            catch(Exception ex)
            {
                throw new CustomErrorException($"{ex.Message}");
            }
        }


        public byte[] GetConsolidatedPagesBytes(List<byte[]> packetLists)
        {
            PdfDocument outPdf = new PdfDocument();
            packetLists.ForEach(file =>
            {
                {
                    // Convert byte to stream
                    MemoryStream stream = new MemoryStream(file);

                    try
                    {
                        using (PdfDocument doc = PdfReader.Open(stream, PdfDocumentOpenMode.Import))
                        {
                            CopyPages(doc, outPdf);
                        }

                    }
                    catch (FileNotFoundException e)
                    {
                        throw e;
                    }
                    outPdf.Close();
                }
            });

            MemoryStream streams = new MemoryStream();
            outPdf.Save(streams, true);

            return streams.ToArray();
        }

        private void CopyPages(PdfDocument from, PdfDocument to)
        {
            for (int i = 0; i < from.PageCount; i++)
            {
                to.AddPage(from.Pages[i]);
            }
        }
    }
}
