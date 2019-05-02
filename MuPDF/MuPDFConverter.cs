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
        public static Dictionary<int, byte[]> FastConvert(byte[] image, RenderType type)
        {
            if (image == null)
                throw new ArgumentNullException("image");

            var output = new ConcurrentDictionary<int, byte[]>();

            ImageCodecInfo info = null;
            foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                if (ice.MimeType == "image/tiff")
                    info = ice;

            var docList = new Dictionary<int, RenderType>();
            var pageCount = 0;

            using (MuPDF pdfDoc = new MuPDF(image, string.Empty))
            {
                pageCount = pdfDoc.PageCount;
            }

            if (pageCount > 0)
            {
                Parallel.For(1, pageCount+1, index =>
                {
                    using (MuPDF pdfDoc = new MuPDF(image, string.Empty))
                    {
                        pdfDoc.Page = index;
                        
                        using (MemoryStream outputStream = new MemoryStream())
                        {
                            var width = 0;
                            var height = 0;
                            var dpi = 150;
                            var maxSize = 1000;
                            
                            using (var bitmap = pdfDoc.GetBitmap(width, height, dpi, dpi, 0, type, false, false, maxSize))
                            {
                                using (EncoderParameters ep = new EncoderParameters(1))
                                {
                                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, type.Equals(RenderType.Monochrome) ? (long)EncoderValue.CompressionCCITT4 : (long)EncoderValue.CompressionLZW);
                                    bitmap.Save(outputStream, info, ep);
                                }
                            }

                            output.TryAdd(index, outputStream.ToArray());
                        }
                    }
                });
            }

            return output.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }


        public static Dictionary<int,byte[]> FastConvert(byte[] image)
        {
            if (image == null)
                throw new ArgumentNullException("image");

            var output = new ConcurrentDictionary<int, byte[]>();

            ImageCodecInfo info = null;
            foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                if (ice.MimeType == "image/tiff")
                    info = ice;

            var docList = new Dictionary<int, RenderType>();

            using (MuPDF pdfDoc = new MuPDF(image, string.Empty))
            {
                pdfDoc.AntiAlias = false;

                for (int i = 1; i <= pdfDoc.PageCount; i++)
                {
                    pdfDoc.Page = i;
                    using (var bitmap = pdfDoc.GetBitmap(100, 0, 50, 50, 0, RenderType.RGB, false, false, 0))
                    {
                        docList.Add(i, pdfDoc.Variance > 4 ? RenderType.RGB : RenderType.Monochrome);
                    }
                }
            }

            Parallel.ForEach(docList, page =>
            {
                using (MuPDF pdfDoc = new MuPDF(image, string.Empty))
                {
                    pdfDoc.Page = page.Key;
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        var width = 0;
                        var height = 0;
                        var dpi = 180;

                        if (page.Value.Equals(RenderType.RGB))
                        {
                            width = 1000;
                            dpi = 100;
                        }

                        using (var bitmap = pdfDoc.GetBitmap(width, height, dpi, dpi, 0, page.Value, false, false, 0))
                        {
                            using (EncoderParameters ep = new EncoderParameters(1))
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, page.Value.Equals(RenderType.Monochrome) ? (long)EncoderValue.CompressionCCITT4 : (long)EncoderValue.CompressionLZW);
                                bitmap.Save(outputStream, info, ep);
                            }
                        }

                        output.TryAdd(page.Key, outputStream.ToArray());
                    }
                }
            });

            return output.ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }


        public static List<byte[]> ConvertPdfToPng(byte[] image, int dpi, Size size, RenderType type, bool antiAlias, string pdfPassword)
        {
            if (image == null)
                throw new ArgumentNullException("image");

            var output = new List<byte[]>();

            using (MuPDF pdfDoc = new MuPDF(image, pdfPassword))
            {
                ImageCodecInfo info = null;
                foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                    if (ice.MimeType == "image/tiff")
                        info = ice;

                pdfDoc.AntiAlias = antiAlias;

                //Parallel.For(1, pdfDoc.PageCount,
                //i =>
                for (int i = 1; i <= pdfDoc.PageCount; i++)
                {
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        int Width = size.Width;//Zero for no resize.
                        int Height = size.Height;//Zero for autofit height to width.

                        pdfDoc.Page = i;
                    
                        using (var bitmap = pdfDoc.GetBitmap(Width, Height, dpi, dpi, 0, type, false, false, 0))
                        {
                            if (bitmap == null)
                                throw new Exception("Unable to convert pdf to png!");

                            if (type.Equals(RenderType.RGB) && pdfDoc.Variance < 4)
                            {
                                using (var mBitmap = pdfDoc.GetBitmap(Width, Height, dpi, dpi, 0, RenderType.Monochrome, false, false, 0))
                                {
                                    using (EncoderParameters ep = new EncoderParameters(1))
                                    {
                                        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

                                        mBitmap.Save(outputStream, info, ep);
                                    }
                                }
                            }
                            else
                            {
                                using (EncoderParameters ep = new EncoderParameters(1))
                                {
                                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionLZW);
                                    if (type == RenderType.Monochrome)
                                    {
                                        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                                    }

                                    bitmap.Save(outputStream, info, ep);
                                }
                            }
                        }

                        output.Add(outputStream.ToArray());
                    }
                }
                //});
            }
            return output;
        }


        public static List<byte[]> ConvertPdfToPng(byte[] image, int dpi, RenderType type, bool antiAlias, string pdfPassword)
        {
            if (image == null)
                throw new ArgumentNullException("image");

            var output = new List<byte[]>();

            using (MuPDF pdfDoc = new MuPDF(image, pdfPassword))
            {
                ImageCodecInfo info = null;
                foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
                    if (ice.MimeType == "image/tiff")
                        info = ice;
                    
                pdfDoc.AntiAlias = antiAlias;

                //Parallel.For(1, pdfDoc.PageCount,
                //i =>
                for(int i = 1; i <= pdfDoc.PageCount; i++)
                {
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        int Width = 0;//Zero for no resize.
                        int Height = 0;//Zero for autofit height to width.

                        pdfDoc.Page = i;

                        using (var bitmap = pdfDoc.GetBitmap(Width, Height, dpi, dpi, 0, type, false, false, 0))
                        {
                            if (bitmap == null)
                                throw new Exception("Unable to convert pdf to png!");

                            if (type.Equals(RenderType.RGB) && pdfDoc.Variance < 4)
                            {
                                using (var mBitmap = pdfDoc.GetBitmap(Width, Height, dpi, dpi, 0, RenderType.Monochrome, false, false, 0))
                                {
                                    using (EncoderParameters ep = new EncoderParameters(1))
                                    {
                                        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);

                                        mBitmap.Save(outputStream, info, ep);
                                    }
                                }
                            }
                            else
                            {
                                using (EncoderParameters ep = new EncoderParameters(1))
                                {
                                    ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionLZW);
                                    if (type == RenderType.Monochrome)
                                    {
                                        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)EncoderValue.CompressionCCITT4);
                                    }

                                    bitmap.Save(outputStream, info, ep);
                                }
                            }
                        }

                        output.Add(outputStream.ToArray());
                    }
                }
                    //});
            }
            return output;
        }

        private static Image ResizeImage(Image img, Size newSize)
        {
            Image thumbnail = new Bitmap(newSize.Width, newSize.Height);

            Graphics graphic = Graphics.FromImage(thumbnail);

            graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphic.SmoothingMode = SmoothingMode.HighQuality;
            graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
            graphic.CompositingQuality = CompositingQuality.HighQuality;

            // Figure out the ratio
            double ratioX = (double)newSize.Width / (double)img.Size.Width;
            double ratioY = (double)newSize.Height / (double)img.Size.Height;
            // use whichever multiplier is smaller
            double ratio = ratioX < ratioY ? ratioX : ratioY;

            // now we can get the new height and width
            int newHeight = Convert.ToInt32(img.Size.Height * ratio);
            int newWidth = Convert.ToInt32(img.Size.Width * ratio);

            // Now calculate the X,Y position of the upper-left corner 
            // (one of these will always be zero)
            int posX = Convert.ToInt32((newSize.Width - (img.Size.Width * ratio)) / 2);
            int posY = Convert.ToInt32((newSize.Height - (img.Size.Height * ratio)) / 2);

            graphic.Clear(Color.White); // white padding
            graphic.DrawImage(img, posX, posY, newWidth, newHeight);

            return thumbnail;
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
    }
}
