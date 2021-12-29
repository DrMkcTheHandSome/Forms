using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FormsDemo.Helpers
{
	public class FileResultFromStream : ActionResult
	{
		public FileResultFromStream(string fileDownloadName, Stream fileStream, string contentType)
		{
			FileDownloadName = fileDownloadName;
			FileStream = fileStream;
			ContentType = contentType;
		}

		public string ContentType { get; }
		public string FileDownloadName { get; }
		public Stream FileStream { get; }

		public override async Task ExecuteResultAsync(ActionContext context)
		{
			var response = context.HttpContext.Response;
			response.ContentType = ContentType;
			context.HttpContext.Response.Headers.Add("Content-Disposition", new[] { "attachment; filename=" + FileDownloadName });

			await FileStream.CopyToAsync(context.HttpContext.Response.Body);
		}
	}
}
