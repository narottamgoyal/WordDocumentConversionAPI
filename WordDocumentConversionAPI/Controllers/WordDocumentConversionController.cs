using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OpenXmlPowerTools;

namespace WordDocumentConversionAPI.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WordDocumentConversionController : ControllerBase
    {
        private readonly ILogger<WordDocumentConversionController> _logger;
        private readonly string _tempPath;
        public WordDocumentConversionController(ILogger<WordDocumentConversionController> logger)
        {
            _logger = logger;
            _tempPath = Path.Combine(Path.GetTempPath(), "WordDocumentConversion");
            if (!Directory.Exists(_tempPath)) Directory.CreateDirectory(_tempPath);
        }

        /// <summary>
        /// Convert docx to html without image
        /// </summary>
        /// <returns></returns>
        [HttpPost("DocxToHtmlWithoutImage")]
        [Consumes("multipart/form-data")]
        [ProducesResponseType(typeof(byte[]), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(BadRequestObjectResult), 400)]
        public async Task<IActionResult> DocxToHtmlWithoutImage(IFormFile file)
        {
            if (file != null && file.Length > 0 && file.FileName.ToLower().EndsWith(".docx"))
            {
                using (var ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    var fileBytes = ms.ToArray();
                    //string base64String = Convert.ToBase64String(fileBytes);
                    return await ConvertDocxToHtmlWithoutImageAsync(file.FileName, fileBytes);
                }
            }
            return StatusCode(StatusCodes.Status400BadRequest, "Please upload docx file");
        }

        /// <summary>
        /// Convert docx to html with image
        /// </summary>
        /// <returns></returns>
        [HttpPost("DocxToHtmlWithImage")]
        [Consumes("multipart/form-data")]
        [ProducesResponseType(typeof(byte[]), StatusCodes.Status200OK)]
        [ProducesResponseType(typeof(BadRequestObjectResult), 400)]
        public async Task<IActionResult> DocxToHtmlWithImage(IFormFile file)
        {
            if (file != null && file.Length > 0 && file.FileName.ToLower().EndsWith(".docx"))
            {
                using (var ms = new MemoryStream())
                {
                    file.CopyTo(ms);
                    var fileBytes = ms.ToArray();
                    return await ConvertDocxToHtmlWithImageAsync(file.FileName, fileBytes);
                }
            }
            return StatusCode(StatusCodes.Status400BadRequest, "Please upload docx file");
        }

        private async Task<IActionResult> ConvertDocxToHtmlWithImageAsync(string fileName, byte[] fileBytes)
        {
            try
            {
                var docFilePath = Path.Combine(_tempPath, fileName);
                System.IO.File.WriteAllBytes(docFilePath, fileBytes);
                string htmlText = string.Empty;
                try
                {
                    htmlText = ParseDOCX(docFilePath);
                }
                catch (OpenXmlPackageException e)
                {
                    if (e.ToString().Contains("Invalid Hyperlink"))
                    {
                        using (FileStream fs = new FileStream(docFilePath, FileMode.Open, FileAccess.Read))
                        {
                            UriFixer.FixInvalidUri(fs, brokenUri => FixUri(brokenUri));
                        }
                        htmlText = ParseDOCX(docFilePath);
                    }
                }

                var outfile = Path.Combine(_tempPath, ("htmldata" + ".html"));
                var writer = System.IO.File.CreateText(outfile);
                writer.WriteLine(htmlText.ToString());
                writer.Dispose();
                System.IO.File.Delete(docFilePath);
                return File(await System.IO.File.ReadAllBytesAsync(outfile), "application/octet-stream", (Path.GetFileNameWithoutExtension(fileName) + ".html"));
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError);
            }
        }

        private async Task<IActionResult> ConvertDocxToHtmlWithoutImageAsync(string fileName, byte[] fileBytes)
        {
            try
            {

                var docFilePath = Path.Combine(_tempPath, fileName);
                System.IO.File.WriteAllBytes(docFilePath, fileBytes);
                var source = Package.Open(docFilePath);
                var document = WordprocessingDocument.Open(source);
                HtmlConverterSettings settings = new HtmlConverterSettings();
                XElement html = HtmlConverter.ConvertToHtml(document, settings);
                var outfile = Path.Combine(_tempPath, ("htmldata" + ".html"));
                var writer = System.IO.File.CreateText(outfile);
                writer.WriteLine(html.ToString());
                writer.Dispose();
                source.Close();
                System.IO.File.Delete(docFilePath);
                return File(await System.IO.File.ReadAllBytesAsync(outfile), "application/octet-stream", (Path.GetFileNameWithoutExtension(fileName) + ".html"));
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError);
            }
        }

        private static Uri FixUri(string brokenUri)
        {
            string newURI = string.Empty;
            if (brokenUri.Contains("mailto:"))
            {
                int mailToCount = "mailto:".Length;
                brokenUri = brokenUri.Remove(0, mailToCount);
                newURI = brokenUri;
            }
            else
            {
                newURI = " ";
            }
            return new Uri(newURI);
        }

        private static string ParseDOCX(string fileName)
        {
            try
            {
                byte[] byteArray = System.IO.File.ReadAllBytes(fileName);
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    memoryStream.Write(byteArray, 0, byteArray.Length);
                    using (WordprocessingDocument wDoc =
                                                WordprocessingDocument.Open(memoryStream, true))
                    {
                        int imageCounter = 0;
                        var pageTitle = fileName;
                        var part = wDoc.CoreFilePropertiesPart;
                        if (part != null)
                            pageTitle = (string)part.GetXDocument()
                                                    .Descendants(DC.title)
                                                    .FirstOrDefault() ?? fileName;

                        WmlToHtmlConverterSettings settings = new WmlToHtmlConverterSettings()
                        {
                            AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                            PageTitle = pageTitle,
                            FabricateCssClasses = true,
                            CssClassPrefix = "pt-",
                            RestrictToSupportedLanguages = false,
                            RestrictToSupportedNumberingFormats = false,
                            ImageHandler = imageInfo =>
                            {
                                ++imageCounter;
                                string extension = imageInfo.ContentType.Split('/')[1].ToLower();
                                ImageFormat imageFormat = null;
                                if (extension == "png") imageFormat = ImageFormat.Png;
                                else if (extension == "gif") imageFormat = ImageFormat.Gif;
                                else if (extension == "bmp") imageFormat = ImageFormat.Bmp;
                                else if (extension == "jpeg") imageFormat = ImageFormat.Jpeg;
                                else if (extension == "tiff")
                                {
                                    extension = "gif";
                                    imageFormat = ImageFormat.Gif;
                                }
                                else if (extension == "x-wmf")
                                {
                                    extension = "wmf";
                                    imageFormat = ImageFormat.Wmf;
                                }

                                if (imageFormat == null) return null;

                                string base64 = null;
                                try
                                {
                                    using (MemoryStream ms = new MemoryStream())
                                    {
                                        imageInfo.Bitmap.Save(ms, imageFormat);
                                        var ba = ms.ToArray();
                                        base64 = System.Convert.ToBase64String(ba);
                                    }
                                }
                                catch (System.Runtime.InteropServices.ExternalException)
                                { return null; }

                                ImageFormat format = imageInfo.Bitmap.RawFormat;
                                ImageCodecInfo codec = ImageCodecInfo.GetImageDecoders()
                                                            .First(c => c.FormatID == format.Guid);
                                string mimeType = codec.MimeType;

                                string imageSource =
                                        string.Format("data:{0};base64,{1}", mimeType, base64);

                                XElement img = new XElement(Xhtml.img,
                                        new XAttribute(NoNamespace.src, imageSource),
                                        imageInfo.ImgStyleAttribute,
                                        imageInfo.AltText != null ?
                                            new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                                return img;
                            }
                        };

                        XElement htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
                        var html = new XDocument(new XDocumentType("html", null, null, null),
                                                                                    htmlElement);
                        var htmlString = html.ToString(SaveOptions.DisableFormatting);
                        return htmlString;
                    }
                }
            }
            catch
            {
                return "The file is either open, please close it or contains corrupt data";
            }
        }
    }
}
