﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using LibOpenXml.Common;
using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace LibOpenXml.Word {

    // https://msdn.microsoft.com/EN-US/library/office/bb497430.aspx
    internal class ImageCreator {

        private WordprocessingDocument _package;

        public ImageCreator(WordprocessingDocument package) {

            this._package = package;
        }

        internal void InsertPicture(string picturePath, decimal resizablePercent, HorizontalAlignmentType alignment) {

            if (resizablePercent < 0.0M)
                throw new ArgumentOutOfRangeException("resizablePercent", "The resizable percent is lower than 0.0M. Define a value between 0.0M and 1.0M.");
            if (resizablePercent > 1.0M)
                throw new ArgumentOutOfRangeException("resizablePercent", "The resizable percent is grower than 1.0M. Define a value between 0.0M and 1.0M.");

            MainDocumentPart mainPart = this._package.MainDocumentPart;

            var imagePartId = CreateImagePart(mainPart, picturePath);

            ImageSize imageSize = GetOriginalSize(picturePath, resizablePercent);

            AddImageToBody(imagePartId, Path.GetFileName(picturePath), imageSize, alignment);
        }

        internal void InsertPicture(Bitmap pictureBmp, decimal resizablePercent, HorizontalAlignmentType alignment)
        {

            if (resizablePercent < 0.0M)
                throw new ArgumentOutOfRangeException("resizablePercent", "The resizable percent is lower than 0.0M. Define a value between 0.0M and 1.0M.");
            if (resizablePercent > 1.0M)
                throw new ArgumentOutOfRangeException("resizablePercent", "The resizable percent is grower than 1.0M. Define a value between 0.0M and 1.0M.");

            MainDocumentPart mainPart = this._package.MainDocumentPart;

            var imagePartId = CreateImagePart(mainPart, pictureBmp);

            ImageSize imageSize = GetOriginalSize(pictureBmp, resizablePercent);

            AddImageToBody(imagePartId, pictureBmp, imageSize, alignment);
        }

        private string CreateImagePart(MainDocumentPart mainPart, string picturePath) {

            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(picturePath, FileMode.Open)) {

                imagePart.FeedData(stream);
            }
            
            return mainPart.GetIdOfPart(imagePart);
        }

        
        /////////////////////////////////////////////////////////////////
        

        //[DllImport("gdi32.dll", EntryPoint = "DeleteObject")]
        //[return: MarshalAs(UnmanagedType.Bool)]
        //public static extern bool DeleteObject([In] IntPtr hObject);

        //public System.Windows.Media.ImageSource ImageSourceFromBitmap(Bitmap bmp)
        //{
        //    var handle = bmp.GetHbitmap();
            
        //    try
        //    {
        //        return CreateBitmapSourceFromHBitmap(handle, IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());
        //    }
        //    finally { DeleteObject(handle); }
        //}


        // https://stackoverflow.com/questions/26260654/wpf-converting-bitmap-to-imagesource

        public static System.Windows.Media.Imaging.BitmapSource CreateBitmapSourceFromGdiBitmap(Bitmap bitmap)
        {
            if (bitmap == null)
                throw new ArgumentNullException("bitmap");

            var rect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);

            var bitmapData = bitmap.LockBits(
                rect,
                System.Drawing.Imaging.ImageLockMode.ReadWrite,
                System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            try
            {
                var size = (rect.Width * rect.Height) * 4;

                return System.Windows.Media.Imaging.BitmapSource.Create(
                    bitmap.Width,
                    bitmap.Height,
                    bitmap.HorizontalResolution,
                    bitmap.VerticalResolution,
                    System.Windows.Media.PixelFormats.Bgra32,
                    null,
                    bitmapData.Scan0,
                    size,
                    bitmapData.Stride);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
        }


        ///////////////////////////////////////////////
        // http://www.nullskull.com/faq/1500/how-to-convert-bitmap-to-byte-array.aspx
        private string CreateImagePart(MainDocumentPart mainPart, Bitmap picturePath)
        {
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);
            byte[] data = default(byte[]);
            using (System.IO.MemoryStream sampleStream = new System.IO.MemoryStream())
            {
                picturePath.Save(sampleStream, System.Drawing.Imaging.ImageFormat.Bmp);
                data = sampleStream.ToArray();
                imagePart.FeedData(sampleStream);
            }
            return mainPart.GetIdOfPart(imagePart);
        }

        private void AddImageToBody(string relationshipId, string name, ImageSize imageSize, HorizontalAlignmentType alignment) {

            Body body = this._package.MainDocumentPart.Document.Body;

            var element = CreateDrawingElement(relationshipId, name, imageSize);

            var pictureParagraph = new Paragraph(new Run(element));

            // If the paragraph has no ParagraphProperties object, create one.
            WordHelper.CreateParagraphPropertiesIfNonExists(pictureParagraph);

            // Set paragraph alignment
            WordHelper.SetParagraphAlignment(pictureParagraph, alignment);

            // Append the reference to body, the element should be in a Run.
            body.AppendChild(pictureParagraph);
        }

        private void AddImageToBody(string relationshipId, Bitmap bmp, ImageSize imageSize, HorizontalAlignmentType alignment)
        {

            Body body = this._package.MainDocumentPart.Document.Body;

            var element = CreateDrawingElement(relationshipId, bmp, imageSize);

            var pictureParagraph = new Paragraph(new Run(element));

            // If the paragraph has no ParagraphProperties object, create one.
            WordHelper.CreateParagraphPropertiesIfNonExists(pictureParagraph);

            // Set paragraph alignment
            WordHelper.SetParagraphAlignment(pictureParagraph, alignment);

            // Append the reference to body, the element should be in a Run.
            body.AppendChild(pictureParagraph);
        }

        private ImageSize GetOriginalSize(string fileName, decimal percent) {

            ImageSize returnValue = new ImageSize();

            using (var img = new Bitmap(fileName)) {
                
                returnValue.Width = (long)((img.Width / img.HorizontalResolution) * 914400L);
                returnValue.Height = (long)((img.Height / img.VerticalResolution) * 914400L);

                returnValue.Width = (long)(returnValue.Width * percent);
                returnValue.Height = (long)(returnValue.Height * percent);
            }

            return returnValue;
        }

        private ImageSize GetOriginalSize(Bitmap img, decimal percent)
        {

            ImageSize returnValue = new ImageSize();

            //using (var img = new Bitmap(fileName))
            {

                returnValue.Width = (long)((img.Width / img.HorizontalResolution) * 914400L);
                returnValue.Height = (long)((img.Height / img.VerticalResolution) * 914400L);

                returnValue.Width = (long)(returnValue.Width * percent);
                returnValue.Height = (long)(returnValue.Height * percent);
            }

            return returnValue;
        }

        private Drawing CreateDrawingElement(string relationshipId, string name, ImageSize imageSize) {

            double imageWidthInInches = imageSize.Width / 914400.0;
            double imageHeightInInches = imageSize.Height / 914400.0;
            
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = imageSize.Width, Cy = imageSize.Height },
                         new DW.EffectExtent() {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties() {
                             Id = (UInt32Value)1U,
                             Name = name
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties() {
                                             Id = (UInt32Value)0U,
                                             Name = name
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension() {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         ) {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L }
                                             , new A.Extents() { 
                                                    Cx = imageSize.Width, //990000L, 
                                                    Cy = imageSize.Height, //792000L 
                                             }
                                         ),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         ) { Preset = A.ShapeTypeValues.Rectangle }))
                             ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     ) {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return element;
        }

        private Drawing CreateDrawingElement(string relationshipId, Bitmap name, ImageSize imageSize)
        {
            
            string nametemp = "temp__.bmp";
            if (File.Exists(nametemp)) File.Delete(nametemp);
            name.Save(nametemp);

            double imageWidthInInches = imageSize.Width / 914400.0;
            double imageHeightInInches = imageSize.Height / 914400.0;

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = imageSize.Width, Cy = imageSize.Height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = nametemp
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = nametemp
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L }
                                             , new A.Extents()
                                             {
                                                 Cx = imageSize.Width, //990000L, 
                                                 Cy = imageSize.Height, //792000L 
                                             }
                                         ),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return element;
        }
    }
}
