using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Drawing2D;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace MuPDFLib
{
    public static class MuPdfConverter
    {
        public static IDictionary<int, byte[]> ConvertPdfToPng(byte[] pdfBytes, RenderType type, bool antiAlias = false, float dpi = 150, Size size = new Size(), string password = "")
        {
            if (pdfBytes == null || pdfBytes.Length.Equals(0))
                throw new ArgumentNullException(nameof(pdfBytes));

            var output = new ConcurrentDictionary<int, byte[]>();
            var pageCount = 0;

            using (MuPDF pdfDoc = new MuPDF(pdfBytes, password))
            {
                pageCount = pdfDoc.PageCount+1;
            }

            if (pageCount > 0)
            {
                Parallel.For(1, pageCount, index =>
                {
                    using (MuPDF pdfDoc = new MuPDF(pdfBytes, password))
                    {
                        pdfDoc.Page = index;
                        pdfDoc.AntiAlias = antiAlias && !type.Equals(RenderType.Monochrome); // no point in anti-alias-ing with Monochrome

                        using (MemoryStream outputStream = new MemoryStream())
                        {
                            var width = 0;
                            var height = 0;
                            var maxSize = 1000;

                            using (var bitmap = ResizeImage(size, pdfDoc.GetBitmap(width, height, dpi, dpi, 0, type, false, false, maxSize)))
                            {
                                bitmap.Save(outputStream, ImageFormat.Png);
                                output.TryAdd(index, outputStream.ToArray());
                            }
                        }
                    }
                });
            }

            return output.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }
        
        public static byte[] ConvertPdfToTiff(byte[] image, float dpi, RenderType type, bool rotateLandscapePages, bool shrinkToLetter, int maxSizeInPdfPixels, string pdfPassword)
        {
            byte[] output = null;

            if (image == null)
                throw new ArgumentNullException("image");

            using (MuPDF pdfDoc = new MuPDF(image, pdfPassword))
            {
                using (MemoryStream outputStream = new MemoryStream())
                {
                    ImageCodecInfo info = null;
                    foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                        if (ice.MimeType == "image/tiff")
                            info = ice;

                    Bitmap saveTif = null;
                    pdfDoc.AntiAlias = false;
                    for (int i = 1; i <= pdfDoc.PageCount; i++)
                    {
                        int Width = 0;//Zero for no resize.
                        int Height = 0;//Zero for autofit height to width.

                        pdfDoc.Page = i;

                        Bitmap FirstImage = pdfDoc.GetBitmap(Width, Height, dpi, dpi, 0, type, rotateLandscapePages, shrinkToLetter, maxSizeInPdfPixels);
                        if (FirstImage == null)
                            throw new Exception("Unable to convert pdf to tiff!");
                        using (EncoderParameters ep = new EncoderParameters(2))
                        {
                            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                            ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionLZW);
                            if (type == RenderType.Monochrome)
                            {
                                ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                            }

                            if (i == 1)
                            {
                                saveTif = FirstImage;
                                saveTif.Save(outputStream, info, ep);
                            }
                            else
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                                saveTif.SaveAdd(FirstImage, ep);
                                FirstImage.Dispose();
                            }
                            if (i == pdfDoc.PageCount)
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                                saveTif.SaveAdd(ep);
                                saveTif.Dispose();
                            }
                        }
                    }
                    output = outputStream.ToArray();
                }
            }
            return output;
        }

        public static bool ConvertPdfToTiff(string sourceFile, string outputFile, float dpi, RenderType type, bool rotateLandscapePages, bool shrinkToLetter, int maxSizeInPdfPixels, string pdfPassword)
        {
            bool output = false;

            if (string.IsNullOrEmpty(sourceFile) || string.IsNullOrEmpty(outputFile))
                throw new ArgumentNullException();

            using (MuPDF pdfDoc = new MuPDF(sourceFile, pdfPassword))
            {
                using (FileStream outputStream = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    ImageCodecInfo info = null;
                    foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                        if (ice.MimeType == "image/tiff")
                            info = ice;

                    Bitmap saveTif = null;
                    pdfDoc.AntiAlias = false;
                    for (int i = 1; i <= pdfDoc.PageCount; i++)
                    {
                        pdfDoc.Page = i;

                        Bitmap FirstImage = pdfDoc.GetBitmap(0, 0, dpi, dpi, 0, type, rotateLandscapePages, shrinkToLetter, maxSizeInPdfPixels);
                        if (FirstImage == null)
                            throw new Exception("Unable to convert pdf to tiff!");

                        using (EncoderParameters ep = new EncoderParameters(2))
                        {
                            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                            ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionLZW);
                            if (type == RenderType.Monochrome)
                            {
                                ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                            }

                            if (i == 1)
                            {
                                saveTif = FirstImage;
                                saveTif.Save(outputStream, info, ep);
                            }
                            else
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                                saveTif.SaveAdd(FirstImage, ep);
                                FirstImage.Dispose();
                            }
                            if (i == pdfDoc.PageCount)
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                                saveTif.SaveAdd(ep);
                                saveTif.Dispose();
                            }
                        }
                    }
                }
                if (File.Exists(outputFile))
                    output = true;
            }
            return output;
        }

        public static byte[] ConvertPdfToFaxTiff(byte[] image, float dpi, bool shrinkToLetter, string pdfPassword)
        {
            byte[] output = null;
            const long Compression = (long)EncoderValue.CompressionCCITT4;

            using (MuPDF pdfDoc = new MuPDF(image, pdfPassword))
            {
                using (MemoryStream outputStream = new MemoryStream())
                {
                    ImageCodecInfo info = null;
                    foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                        if (ice.MimeType == "image/tiff")
                            info = ice;

                    Bitmap saveTif = null;
                    pdfDoc.AntiAlias = false;
                    for (int i = 1; i <= pdfDoc.PageCount; i++)
                    {
                        int Width = 0;//Zero for no resize.
                        //int Height = 0;//Zero for autofit height to width.
                        float DpiX = dpi;
                        float DpiY = dpi;

                        pdfDoc.Page = i;

                        if (dpi == 200)
                        {
                            Width = 1728;
                            DpiX = 204;
                            DpiY = 196;
                        }
                        else if (dpi == 300)
                            Width = 2592;
                        else if (dpi == 400)
                            Width = 3456;

                        Bitmap FirstImage = pdfDoc.GetBitmap(Width, 0, DpiX, DpiY, 0, RenderType.Monochrome, true, shrinkToLetter, 0);
                        if (FirstImage == null)
                            throw new Exception("Unable to convert pdf to tiff!");

                        using (EncoderParameters ep = new EncoderParameters(2))
                        {
                            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                            ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, Compression);

                            if (i == 1)
                            {
                                saveTif = FirstImage;
                                saveTif.Save(outputStream, info, ep);
                            }
                            else
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                                saveTif.SaveAdd(FirstImage, ep);
                                FirstImage.Dispose();
                            }
                            if (i == pdfDoc.PageCount)
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                                saveTif.SaveAdd(ep);
                                saveTif.Dispose();
                            }
                        }
                    }
                    output = outputStream.ToArray();
                }
            }
            return output;
        }

        public static bool ConvertPdfToFaxTiff(string sourceFile, string outputFile, float dpi, bool shrinkToLetter, string pdfPassword)
        {
            bool output = false;
            const long Compression = (long)EncoderValue.CompressionCCITT4;

            using (MuPDF pdfDoc = new MuPDF(sourceFile, pdfPassword))
            {
                using (FileStream outputStream = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    ImageCodecInfo info = null;
                    foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                        if (ice.MimeType == "image/tiff")
                            info = ice;

                    Bitmap saveTif = null;
                    pdfDoc.AntiAlias = false;
                    for (int i = 1; i <= pdfDoc.PageCount; i++)
                    {
                        int Width = 0;//Zero for no resize.
                        //int Height = 0;//Zero for autofit height to width.
                        float DpiX = dpi;
                        float DpiY = dpi;

                        pdfDoc.Page = i;

                        if (dpi == 200)
                        {
                            Width = 1728;
                            DpiX = 204;
                            DpiY = 196;
                        }
                        else if (dpi == 300)
                            Width = 2592;
                        else if (dpi == 400)
                            Width = 3456;

                        Bitmap FirstImage = pdfDoc.GetBitmap(Width, 0, DpiX, DpiY, 0, RenderType.Monochrome, true, shrinkToLetter, 0);
                        if (FirstImage == null)
                            throw new Exception("Unable to convert pdf to tiff!");
                        using (EncoderParameters ep = new EncoderParameters(2))
                        {
                            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
                            ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, Compression);

                            if (i == 1)
                            {
                                saveTif = FirstImage;
                                saveTif.Save(outputStream, info, ep);
                            }
                            else
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.FrameDimensionPage);
                                saveTif.SaveAdd(FirstImage, ep);
                                FirstImage.Dispose();
                            }
                            if (i == pdfDoc.PageCount)
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.Flush);
                                saveTif.SaveAdd(ep);
                                saveTif.Dispose();
                            }
                        }
                    }
                }
                if (File.Exists(outputFile))
                    output = true;
            }
            return output;
        }
        
        private static Bitmap ResizeImage(Size newSize, Bitmap img)
        {
            if ((newSize.Height + newSize.Width).Equals(0))
            {
                return img;
            }

            // Figure out the ratio
            double ratioX = (double)newSize.Width / (double)img.Size.Width;
            double ratioY = (double)newSize.Height / (double)img.Size.Height;

            // use whichever ratio is more than zero, less than one and the lower of the two
            ratioX = ratioX > 0 ? ratioX : 1;
            ratioY = ratioY > 0 ? ratioY : 1;

            double ratio = ratioX < ratioY ? ratioX : ratioY;

            // now we can get the new height and width
            int newHeight = Convert.ToInt32(img.Size.Height * ratio);
            int newWidth = Convert.ToInt32(img.Size.Width * ratio);

            using (Image thumbnail = new Bitmap(newWidth, newHeight))
            using (Graphics graphic = Graphics.FromImage(thumbnail))
            {
                graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphic.SmoothingMode = SmoothingMode.HighQuality;
                graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphic.CompositingQuality = CompositingQuality.HighQuality;

                int posX = 0;
                int posY = Convert.ToInt32(newSize.Height / 2);

                graphic.Clear(Color.White);
                graphic.DrawImage(img, posX, posY, newWidth, newHeight);

                return new Bitmap(thumbnail);
            }
        }
    }
}
