using OXmlCharts = DocumentFormat.OpenXml.Drawing.Charts;
using OXmlPackaging = DocumentFormat.OpenXml.Packaging;
using OXmlDrawing = DocumentFormat.OpenXml.Drawing;
using OXmlMath = DocumentFormat.OpenXml.Math;
using OXmlDrawing_W = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OXmlDrawing_PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using OXmlO_Drawing = DocumentFormat.OpenXml.Office.Drawing;
using OXmlWordprocessing = DocumentFormat.OpenXml.Wordprocessing;
using OXml = DocumentFormat.OpenXml;
using OXmlSpreadsheet = DocumentFormat.OpenXml.Spreadsheet;

using System;
using System.Data;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Media;
using System.Xml;
using System.Xml.Xsl;
using System.Text;
using System.Xml.Linq;

namespace LibOpenXml.Word
{

    public class Microsoft_Office_Interop_Word_Range
    {
        public string Text = "";
    }

    public class CounterFormula
    {
        public static int Count = 0;
        public static string fn = "eqerror.txt";
    }



    public class WordWriter2 : IDisposable
    {

        public enum WdStyle
        {
            wdNormal = 0,
            wdHeader_1 = 1,
            wdHeader_2 = 2,
            wdDebug = 4,
            wdDebugMono = 8
        }

        private OXmlPackaging.WordprocessingDocument _package;

        public void CreateDocument(string filePath)
        {

            var creator = new DocumentCreator();
            this._package = creator.CreatePackage(filePath);

            CreateCommonStyles();
        }

        public void OpenDocument(string filePath)
        {

            var creator = new DocumentOpener();
            this._package = creator.OpenPackage(filePath);
        }

        public void CreateCommonStyles()
        {

            StyleCreator styleCreator = new StyleCreator(this._package);
            styleCreator.CreateBasicStyles();
        }

        public void AddCustomStyle(FormatStyle customStyle)
        {

            if (customStyle == null)
                throw new ArgumentNullException("customStyle");
            if (string.IsNullOrEmpty(customStyle.StyleId))
                throw new ArgumentNullException("StyleId");
            if (string.IsNullOrEmpty(customStyle.Name))
                throw new ArgumentNullException("Name");

            StyleCreator styleCreator = new StyleCreator(this._package);
            styleCreator.AddCustomStyle(customStyle);
        }

        public void AddImage(string picturePath)
        {

            var imageCreator = new ImageCreator(this._package);
            imageCreator.InsertPicture(picturePath, 1.0M, HorizontalAlignmentType.Left);
        }

        public void AddImage(string picturePath, HorizontalAlignmentType alignment)
        {

            var imageCreator = new ImageCreator(this._package);
            imageCreator.InsertPicture(picturePath, 1.0M, alignment);
        }

        // ///////////////////////////////////////

        public void AddImage(System.Drawing.Bitmap bmp, HorizontalAlignmentType alignment)
        {

            var imageCreator = new ImageCreator(this._package);
            imageCreator.InsertPicture(bmp, 1.0M, alignment);
        }

        // ///////////////////////////////////////

        public void AddImage(string picturePath, decimal resizablePercent)
        {

            var imageCreator = new ImageCreator(this._package);
            imageCreator.InsertPicture(picturePath, resizablePercent, HorizontalAlignmentType.Left);
        }

        public void AddImage(string picturePath, decimal resizablePercent, HorizontalAlignmentType alignment)
        {

            var imageCreator = new ImageCreator(this._package);
            imageCreator.InsertPicture(picturePath, resizablePercent, alignment);
        }

        public void AddImage(System.Drawing.Bitmap pictureBmp, decimal resizablePercent, HorizontalAlignmentType alignment)
        {

            var imageCreator = new ImageCreator(this._package);
            imageCreator.InsertPicture(pictureBmp, resizablePercent, alignment);
        }

        public void CreateTextParagraph(TextParagraphType paragraphType, string text)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, paragraphType, HorizontalAlignmentType.Left, text);
        }

        public void CreateTextParagraph(TextParagraphType paragraphType, HorizontalAlignmentType alignment, string text)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, paragraphType, alignment, text);
        }

        public void CreateTextParagraph(TextParagraphType paragraphType, HorizontalAlignmentType alignment, string text, FormatStyle formatStyle)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, paragraphType, alignment, text, formatStyle);
        }

        public void CreateTextParagraph(string styleName, string text)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, HorizontalAlignmentType.Left, text, null);
        }

        public void CreateTextParagraph(string styleName, HorizontalAlignmentType alignment, string text)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, alignment, text, null);
        }

        public void CreateTextParagraph(string styleName, string text, FormatStyle formatStyle)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, HorizontalAlignmentType.Left, text, formatStyle);
        }

        public void CreateTextParagraph(string styleName, HorizontalAlignmentType alignment, string text, FormatStyle formatStyle)
        {

            var titleCreator = new ParagraphCreator();
            titleCreator.CreateTextParagraph(this._package, styleName, alignment, text, formatStyle);
        }

        public void CreateTable(DataTable table)
        {

            var tStyle = new TableStyle();

            CreateTable(table, tStyle);
        }

        //public void CreateTable(DocumentFormat.OpenXml.Drawing.Table table)
        //{

        //    var tStyle = new TableStyle();

        //    CreateTable(table, tStyle);
        //}

        public void CreateTable(DataTable table, TableStyle style)
        {

            TableCreator tableCreator = new TableCreator();
            tableCreator.AddTable(this._package, table, style);
        }

        //public void CreateTable(Table table, TableStyle style)
        //{

        //    TableCreator tableCreator = new TableCreator();
        //    tableCreator.AddTable(this._package, table, style);
        //}

        public void MergeDocument(string documentPath)
        {

            DocumentMerger merger = new DocumentMerger();
            merger.MergeDocument(this._package, documentPath);
        }

        public void AppendBreakPage()
        {

            PageBreaker pageBreaker = new PageBreaker();
            pageBreaker.BreakPage(this._package);
        }

        public void ReplaceBookmark(string bookmark, string content)
        {

            BookmarkReplacer replacer = new BookmarkReplacer();
            replacer.ReplaceBookmark(this._package, bookmark, content);
        }


        ///////////////////////////////////////////////// Word method ////////////////////////////////////////////

        public static System.Collections.Generic.Dictionary<string, string> replacingFields = new System.Collections.Generic.Dictionary<string, string>();

        public void zamena()
        {
            replacingFields["[client_bookmark]"] = "Zanxi";
            replacingFields["[contract_name_bookmark]"] = "Zinjao";
            replacingFields["[complect_name_bookmark]"] = "BeiJing";
            replacingFields["[object_shifr_bookmark]"] = "Asia";
            replacingFields["[DIRZAMPOSITION]"] = "Oil";

            //using (System.Security.Principal.WindowsIdentity.GetCurrent().Impersonate())
            //{
            //    File.Copy(templatePath, filePath, true);
            //}

            // Open the new Package
            //System.IO.Packaging.Package pkg = ; //System.IO.Packaging.Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);
            foreach (System.IO.Packaging.PackagePart part in this._package.Package.GetParts())
            {
                if (part.ContentType.ToLowerInvariant().EndsWith("xml"))
                {
                    XmlDocument xmlMainXMLDoc = new XmlDocument();
                    xmlMainXMLDoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));

                    foreach (var pair in replacingFields)
                    {
                        xmlMainXMLDoc.InnerXml = xmlMainXMLDoc.InnerXml.Replace(pair.Key, pair.Value);
                    }
                    // Open the stream to write document
                    StreamWriter partWrt = new StreamWriter(part.GetStream(FileMode.Open, FileAccess.Write));
                    xmlMainXMLDoc.Save(partWrt);

                    partWrt.Flush();
                    partWrt.Close();
                }
            }
            this._package.Package.Close();

        }


        //public Selection Selection
        //{ get { return _document.Application.Selection; } }

        public void FormulaLn(string formula, WdStyle style)
        {
            //Selection.Text = formula;
            //SetSeletionStyle(style);
            //var mathRange = Selection.Range.OMaths.Add(Selection.Range);
            //Selection.Range.OMaths.BuildUp();
            //Selection.EndKey();
            //Selection.MoveRight(WdUnits.wdCharacter);            
            MathML2Word(@"\begin{document}$" + formula + @"$\end{ document}");

            WriteLn();
        }

        public void FormulaLn(string formula, bool bold, bool italic, int fontSize, string FontName)
        {
            Formula(formula, bold, italic, fontSize, FontName);
            WriteLn();
        }

        public void Formula(string formula, bool bold, bool italic, int fontSize, string FontName)
        {
            MathML2Word(@"\begin{document}$" + formula + @"$ \end{ document}");
        }



        public void Write2(string text, string textReplace, bool bold, bool italic, int fontSize, string FontName, bool allCaps = false)
        {
            string docText = null;
            using (StreamReader sr = new StreamReader(this._package.MainDocumentPart.GetStream()))
            {
                docText = sr.ReadToEnd();
            }

            System.Text.RegularExpressions.Regex regexText = new System.Text.RegularExpressions.Regex(text);
            docText = regexText.Replace(docText, textReplace);

            using (StreamWriter sw = new StreamWriter(this._package.MainDocumentPart.GetStream(FileMode.Create)))
            {
                sw.Write(docText);
            }

            return;
            //replacingFields["[client_bookmark]"] = "Zanxi";
            //replacingFields["[contract_name_bookmark]"] = "Zinjao";
            //replacingFields["[complect_name_bookmark]"] = "BeiJing";
            //replacingFields["[object_shifr_bookmark]"] = "Asia";

            //replacingFields["[DIRZAMPOSITION]"] = "Oil";

            //using (System.Security.Principal.WindowsIdentity.GetCurrent().Impersonate())
            //{
            //    File.Copy(templatePath, filePath, true);
            //}

            // Open the new Package
            //System.IO.Packaging.Package pkg = ; //System.IO.Packaging.Package.Open(filePath, FileMode.Open, FileAccess.ReadWrite);

            foreach (System.IO.Packaging.PackagePart part in this._package.Package.GetParts())
            {
                if (part.ContentType.ToLowerInvariant().EndsWith("xml"))
                {
                    XmlDocument xmlMainXMLDoc = new XmlDocument();
                    //xmlMainXMLDoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));
                    xmlMainXMLDoc.Load(part.GetStream());

                    xmlMainXMLDoc.InnerXml = xmlMainXMLDoc.InnerXml.Replace(text, textReplace);

                    // Open the stream to write document
                    StreamWriter partWrt = new StreamWriter(part.GetStream(FileMode.Open, FileAccess.Write));
                    xmlMainXMLDoc.Save(partWrt);

                    partWrt.Flush();
                    partWrt.Close();
                }
            }
            //this._package.Package.Close();
        }

        public void Write(string text, bool bold, bool italic, int fontSize, string FontName,
        //Action<Selection, string> MethodTransformSelection,
        //Action<Selection, string> MethodReverseTransformSelection,
        params string[] tags)
        {
            if (string.IsNullOrEmpty(text))
                return;

            /// ///// tva - 173 output text
            /// 
            //CreateTextParagraph("Warning", HorizontalAlignmentType.Left, text);
            //CreateTextParagraph("Warning", HorizontalAlignmentType.Left, "");
            CreateTextParagraph(TextParagraphType.None, text);

            //LibASPIRCore.Model.Interconnection.TagSplitter ts = new LibASPIRCore.Model.Interconnection.TagSplitter(tags);
            //List<(string subString, string Tag)> textBlocks = ts.Parse(text);

            //for (int i = 0; i < textBlocks.Count; i++)
            //{
            //    Selection.Font.Bold = bold ? 1 : 0;
            //    Selection.Font.Italic = italic ? 1 : 0;
            //    Selection.Font.Name = FontName;
            //    Selection.Font.Size = fontSize;

            //    MethodTransformSelection(Selection, textBlocks[i].Tag);

            //    Selection.TypeText(textBlocks[i].subString);

            //    MethodReverseTransformSelection(Selection, textBlocks[i].Tag);
            //}
        }

        public void WriteLn(string text, bool bold, bool italic, int fontSize, string FontName)
        {
            Write(text, bold, italic, fontSize, FontName);
            WriteLn();
        }

        public void WriteLn(string text, WdStyle style)
        {
            Write(text, style);
            WriteLn();
        }

        public void WriteLn()
        {
            //Selection.TypeParagraph();
        }

        public void Write(string text, WdStyle style)
        {
            SetSeletionStyle(style);

            switch (style)
            {
                case WdStyle.wdHeader_1:
                    CreateTextParagraph(TextParagraphType.Heading1, text);
                    break;
                case WdStyle.wdHeader_2:
                    CreateTextParagraph(TextParagraphType.Heading2, text);
                    break;
                case WdStyle.wdNormal:
                    CreateTextParagraph(TextParagraphType.Normal, text);
                    break;
            }

            //CreateTextParagraph(TextParagraphType.Heading1, text);

            //Selection.TypeText(text);
            //CreateTextParagraph("Warning", HorizontalAlignmentType.Left, text);


            //Selection.TypeText(text);
            //CreateTextParagraph("Warning", HorizontalAlignmentType.Left, text);
        }

        public void Write(string text, bool bold, bool italic, int fontSize, string FontName, bool allCaps)
        {
            //Selection.Font.AllCaps = allCaps ? 1 : 0;
            Write(text, bold, italic, fontSize, FontName);
        }

        public void SetSelectionToBookmark(string text)
        {

        }

        public Microsoft_Office_Interop_Word_Range findRangeByString(string text)
        {
            return new Microsoft_Office_Interop_Word_Range();
        }

        public void AddContents()
        {
            //if (Closed) { throw new Exception("Ошибка при обращении к документу Word. Документ уже закрыт."); }

            //ActiveDocument.TablesOfContents.Add(Selection.Range, UseHeadingStyles: true, RightAlignPageNumbers: true,
            //    UpperHeadingLevel: 1, LowerHeadingLevel: 3,
            //    IncludePageNumbers: true, AddedStyles: "",
            //    UseHyperlinks: true, HidePageNumbersInWeb: true, UseOutlineLevels: true);

            //ActiveDocument.TablesOfContents[1].TabLeader = WdTabLeader.wdTabLeaderDots;
        }



        public void SetSeletionStyle(WdStyle style)
        {
            switch (style)
            {
                case WdStyle.wdNormal:
                    //Selection.set_Style("Обычный");
                    break;
                case WdStyle.wdHeader_1:
                    //Selection.set_Style("Заголовок 1");
                    break;
                case WdStyle.wdHeader_2:
                    //Selection.set_Style("Заголовок 2");
                    break;
                case WdStyle.wdDebug:
                    //Selection.set_Style("Debug");
                    break;
                case WdStyle.wdDebugMono:
                    //Selection.set_Style("DebugMono");
                    break;
            }
        }

        private void AddTable(string[,] source, int NumRows, int NumColumns)
        {

            using (var table = new DataTable())
            {

                //table.Columns.Add("Item", typeof(string));
                //table.Columns.Add("Min", typeof(string));
                //table.Columns.Add("Avg", typeof(string));
                //table.Columns.Add("Max", typeof(string));
                //table.Columns.Add("Zin", typeof(string));

                for (int i = 1; i <= NumColumns; i++)
                {
                    //tr.Append(new TableCell(new Paragraph(new Run(new Text(i.ToString())))));
                    table.Columns.Add(i.ToString(), typeof(string));
                }

                for (int i = 0; i < NumRows; i++)
                {

                    var r = table.NewRow();
                    //r[0] = i.ToString(); r[1] = 10M.ToString(); r[2] = 11M.ToString(); r[3] = 12M.ToString(); ; r[4] = 12M.ToString();
                    for (int j = 0; j < NumColumns; j++)
                    {
                        //tr.Append(new TableCell(new Paragraph(new Run(new Text((i * j).ToString())))));
                        r[j] = source[i, j].ToString();
                    }

                    table.Rows.Add(r);
                }

                CreateTable(table);

                //CreateTable(table, new TableStyle()
                //{
                //    Alignment = HorizontalAlignmentType.Center,
                //    ShowTitle = true,
                //    Title = "My Table"
                //});
            }
        }

        public void TableTagged(string[,] source,
            bool bold, bool italic, int fontSize, string FontName,
            //Action<Selection, string> MethodTransformSelection,
            //Action<Selection, string> MethodReverseTransformSelection,
            string[] tags,

            int captionRows
            //List<CellRangeRC> spans,
            //WdParagraphAlignment cellsHorizontalAlignment = WdParagraphAlignment.wdAlignParagraphCenter,
            //WdCellVerticalAlignment cellsVerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter,
            //WdLineSpacing lineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            )
        {
            //if (source == null)
            //    return;

            //if (source.GetLength(0) == 0 || source.GetLength(1) == 0)
            //    return;

            int NumRows = source.GetLength(0);
            int NumColumns = source.GetLength(1);

            //_table = ActiveDocument.Tables.Add(Selection.Range, NumRows, NumColumns);

            //_table.AllowAutoFit = true;
            //_table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed);
            ////_table.set_Style("Сетка таблицы");
            //_table.set_Style("Сетка таблицы компакт");

            //for (int i = 0; i < NumRows; i++)
            //{
            //    for (int j = 0; j < NumColumns; j++)
            //    {
            //        //_table.Cell(i + 1, j + 1).Range.InsertAfter(source[i, j]);

            //        _table.Cell(i + 1, j + 1).Range.Select();

            //        Write(source[i, j], bold, italic, fontSize, FontName, MethodTransformSelection, MethodReverseTransformSelection, tags);

            //        _table.Cell(i + 1, j + 1).Range.ParagraphFormat.Alignment = cellsHorizontalAlignment;
            //        _table.Cell(i + 1, j + 1).VerticalAlignment = cellsVerticalAlignment;
            //        _table.Cell(i + 1, j + 1).Range.ParagraphFormat.LineSpacingRule = lineSpacingRule;
            //    }
            //}

            //////MSS: фиксированные строк заголовка
            ////if (captionRows>0 && captionRows<= source.GetLength(0))
            ////{
            ////    for (int r=1;r<=captionRows;r++)
            ////        _table.Rows[r].HeadingFormat = 1;
            ////}

            ////объединение ячеек
            //if (spans != null)
            //{
            //    foreach (var c in spans)
            //    {
            //        object begCell = _table.Cell(c.R1 + 1, c.C1 + 1).Range.Start;
            //        object endCell = _table.Cell(c.R2 + 1, c.C2 + 1).Range.End;
            //        Range wordcellrange = _document.Range(ref begCell, ref endCell);
            //        wordcellrange.Select();
            //        _application.Selection.Cells.Merge();
            //    }
            //}

            ////автоподбор по содержимому
            //_table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
            //_table.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent;
            //_table.PreferredWidth = 100.0f;

            //Selection.EndKey(WdUnits.wdStory);

            //AddTable2(source, NumRows, NumColumns);
            AddTable(source, NumRows, NumColumns);
            WriteLn();
        }

        // https://stackoverflow.com/questions/1201518/convert-system-windows-media-imagesource-to-system-drawing-bitmap

        private System.Drawing.Image ImageWpfToGDI(System.Windows.Media.ImageSource image)
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            var encoder = new System.Windows.Media.Imaging.BmpBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(image as System.Windows.Media.Imaging.BitmapSource));
            encoder.Save(ms);
            ms.Flush();
            return System.Drawing.Image.FromStream(ms);
        }

        // https://stackoverflow.com/questions/8860725/cannot-implicitly-convert-type-system-drawing-image-to-system-drawing-bitmap

        public void ImageFromBitmap(ImageSource bitmap, double scaleX, double scaleY)
        {

            //bitmap.CreateBitmapSourceFromHBitmap(handle, IntPtr.Zero, System.Windows.Int32Rect.Empty, System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());

            //System.Drawing.Image image = ImageWpfToGDI(bitmap);

            //var imageBitmap = new System.Drawing.Bitmap(image.Width, image.Height);
            //using (var graphics = System.Drawing.Graphics.FromImage(imageBitmap))
            //    graphics.DrawImage(image, 0, 0, image.Width, image.Height);


            //System.Drawing.Bitmap bmp;
            //image.Save(bmp);
            //System.Drawing.Bitmap bmp = new;
            ///////////// tva image
            //AddImage("Super-IT.png", HorizontalAlignmentType.Center);

            System.Drawing.Image image = ImageWpfToGDI(bitmap);

            // 2022.03.31 tva
            //AddImage((System.Drawing.Bitmap)image, 0.4M, HorizontalAlignmentType.Center);
            AddImage(@"d:\_2022___\Src_3__\1.png", 0.4M, HorizontalAlignmentType.Center);



            ////Clipboard.Clear();

            //!!!DEBUG: временно закрывал - падает в Task использование Clipboard
            //Clipboard.SetImage((BitmapSource)bitmap);

            ////Clipboard.GetImage();
            ////Application.CutCopyMode = False
            ////if (!Clipboard.ContainsImage())
            ////    throw new Exception("Буфер не содержит изображения");

            //!!!DEBUG: временно блокировал - падает в Task
            //Selection.Paste();

            //Selection.MoveLeft(WdUnits.wdCharacter, 1, WdMovementType.wdExtend);
            //Selection.InlineShapes[1].ScaleHeight = (float)(scaleX * 100.0); //Word в процентах (%)
            //Selection.InlineShapes[1].ScaleWidth = (float)(scaleY * 100.0);
            //Selection.MoveRight(WdUnits.wdCharacter, 1);

            //MSS: попытка сделать через пересохранение в файл...
            ////временный файл
            //Random rnd = new Random();
            //string filePath = Ap.CFG.fTempDirectory + (rnd.Next(100000) + 1).ToString() + ".png";
            //using (var fileStream = new FileStream(filePath, FileMode.Create))
            //{
            //    BitmapEncoder encoder = new PngBitmapEncoder();
            //    encoder.Frames.Add(BitmapFrame.Create((BitmapSource)bitmap));
            //    encoder.Save(fileStream);
            //}
        }

        public void ImageFromBitmapCLn(ImageSource bitmap, double scaleX, double scaleY) //, bool newLine = true
        {
            //WdParagraphAlignment oldParagraphAlignment = Selection.ParagraphFormat.Alignment;
            //float OldFirstLineIndent = Selection.ParagraphFormat.FirstLineIndent;

            ////Selection.TypeParagraph();
            //Selection.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //Selection.ParagraphFormat.FirstLineIndent = 0;

            ImageFromBitmap(bitmap, scaleX, scaleY);

            //Selection.TypeParagraph();
            //Selection.ParagraphFormat.Alignment = oldParagraphAlignment;
            //Selection.ParagraphFormat.FirstLineIndent = OldFirstLineIndent;
        }

        //public void ImageToCellRangeFromPath(Range range, string path, bool isLeftCentered = true)
        public void ImageToCellRangeFromPath(object range, string path, bool isLeftCentered = true)
        {
            if (string.IsNullOrEmpty(path))
                return;

            if (range is null)
                return;

            //if (Closed) { throw new Exception("Ошибка при обращении к документу Word. Документ уже закрыт."); }

            //if (!File.Exists(path))
            //    return;

            //if (range.Cells is null)
            //    return;

            //InlineShape inlineShape = range.InlineShapes.AddPicture(path);
            //Shape shape = inlineShape.ConvertToShape();
            //shape.IncrementTop(Math.Abs(range.Cells.Height - shape.Height) / 2);//14.15f) / 2);
            //if (isLeftCentered)
            //{
            //    if (shape.Width > range.Cells.Width)
            //        shape.IncrementLeft(-(shape.Width - range.Cells.Width) / 2);
            //    //shape.IncrementLeft((range.Cells.Width - shape.Width) / 2);
            //}
        }

        public void AddTable2(string[,] source, int NumRows, int NumColumns)
        {
            // Create a Wordprocessing document.             
            {
                // Add a new main document part. 

                //MainDocumentPart mainPart = this._package.AddMainDocumentPart();

                //Create DOM tree for simple document. 

                //mainPart.Document = new DocumentFormat.OpenXml.Wordprocessing.Document();

                OXmlWordprocessing.Body body = new OXmlWordprocessing.Body();
                OXmlWordprocessing.Table table = new OXmlWordprocessing.Table();
                OXmlWordprocessing.TableProperties tblPr = new OXmlWordprocessing.TableProperties();
                OXmlWordprocessing.TableBorders tblBorders = new OXmlWordprocessing.TableBorders();
                tblBorders.TopBorder = new OXmlWordprocessing.TopBorder();
                tblBorders.TopBorder.Val = new OXml.EnumValue<OXmlWordprocessing.BorderValues>(OXmlWordprocessing.BorderValues.Single);
                tblBorders.BottomBorder = new OXmlWordprocessing.BottomBorder();
                tblBorders.BottomBorder.Val = new DocumentFormat.OpenXml.EnumValue<OXmlWordprocessing.BorderValues>(OXmlWordprocessing.BorderValues.Single);
                tblBorders.LeftBorder = new OXmlWordprocessing.LeftBorder();
                tblBorders.LeftBorder.Val = new DocumentFormat.OpenXml.EnumValue<OXmlWordprocessing.BorderValues>(OXmlWordprocessing.BorderValues.Single);
                tblBorders.RightBorder = new OXmlWordprocessing.RightBorder();
                tblBorders.RightBorder.Val = new DocumentFormat.OpenXml.EnumValue<OXmlWordprocessing.BorderValues>(OXmlWordprocessing.BorderValues.Single);
                tblBorders.InsideHorizontalBorder = new OXmlWordprocessing.InsideHorizontalBorder();
                tblBorders.InsideHorizontalBorder.Val = OXmlWordprocessing.BorderValues.Single;
                tblBorders.InsideVerticalBorder = new OXmlWordprocessing.InsideVerticalBorder();
                tblBorders.InsideVerticalBorder.Val = OXmlWordprocessing.BorderValues.Single;
                tblPr.Append(tblBorders);
                table.Append(tblPr);
                OXmlWordprocessing.TableRow tr;
                OXmlWordprocessing.TableCell tc;
                //first row - title
                tr = new OXmlWordprocessing.TableRow();
                tc = new OXmlWordprocessing.TableCell(
                    new OXmlWordprocessing.Paragraph(
                        new OXmlWordprocessing.Run(
                            new OXmlWordprocessing.Text("Multiplication table"))));
                OXmlWordprocessing.TableCellProperties tcp = new OXmlWordprocessing.TableCellProperties();
                OXmlWordprocessing.GridSpan gridSpan = new OXmlWordprocessing.GridSpan();
                gridSpan.Val = 11;
                tcp.Append(gridSpan);
                tc.Append(tcp);
                tr.Append(tc);
                table.Append(tr);
                //second row 
                tr = new OXmlWordprocessing.TableRow();
                tc = new OXmlWordprocessing.TableCell();
                tc.Append(
                    new OXmlWordprocessing.Paragraph(
                        new OXmlWordprocessing.Run(
                            new OXmlWordprocessing.Text("*"))));
                tr.Append(tc);
                for (int i = 1; i <= NumRows; i++)
                {
                    tr.Append(
                        new OXmlWordprocessing.TableCell(
                            new OXmlWordprocessing.Paragraph(
                                new OXmlWordprocessing.Run(new OXmlWordprocessing.Text(i.ToString())))));
                }
                table.Append(tr);
                for (int i = 1; i <= NumRows; i++)
                {
                    tr = new OXmlWordprocessing.TableRow();
                    tr.Append(new OXmlWordprocessing.TableCell(new OXmlWordprocessing.Paragraph(new OXmlWordprocessing.Run(new OXmlWordprocessing.Text(i.ToString())))));
                    for (int j = 1; j <= NumColumns; j++)
                    {
                        //tr.Append(new TableCell(new Paragraph(new Run(new Text((i * j).ToString())))));
                        tr.Append(new OXmlWordprocessing.TableCell(new OXmlWordprocessing.Paragraph(new OXmlWordprocessing.Run(new OXmlWordprocessing.Text(source[i - 1, j - 1])))));
                    }
                    table.Append(tr);
                }

                CreateTable(table);

                //CreateTable(table, new TableStyle()
                //{
                //    Alignment = HorizontalAlignmentType.Center,
                //    ShowTitle = true,
                //    Title = "My Table"
                //});

                //appending table to body
                //body.Append(table);
                // and body to the document
                //mainPart.Document.Append(body);
                // Save changes to the main document part. 
                //mainPart.Document.Save();
            }
        }

        public void CreateTable(OXmlWordprocessing.Table table)
        {

            var tStyle = new TableStyle();

            CreateTable(table, tStyle);
        }

        public void CreateTable(OXmlWordprocessing.Table table, TableStyle style)
        {

            TableCreator tableCreator = new TableCreator();
            tableCreator.AddTable(this._package, table, style);
        }


        ///////////////////////////////////////////////// Работа с колонтитулами //////////////////////////////////////////
        /////https://docs.microsoft.com/ru-ru/office/open-xml/how-to-remove-the-headers-and-footers-from-a-word-processing-document

        public void RemoveHeadersAndFooters(string filename)
        {
            var docPart = this._package.MainDocumentPart;

            // Count the header and footer parts and continue if there 
            // are any.
            if (((System.Collections.Generic.List<OXmlPackaging.HeaderPart>)docPart.HeaderParts).Count > 0 ||
                ((System.Collections.Generic.List<OXmlPackaging.FooterPart>)docPart.FooterParts).Count > 0)
            {
                // Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts);
                docPart.DeleteParts(docPart.FooterParts);

                // Get a reference to the root element of the main
                // document part.
                OXmlWordprocessing.Document document = docPart.Document;

                // Remove all references to the headers and footers.

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var headers =
                  (System.Collections.Generic.List<OXmlWordprocessing.HeaderReference>)document.Descendants<OXmlWordprocessing.HeaderReference>();
                foreach (var header in headers)
                {
                    header.Remove();
                }

                // First, create a list of all descendants of type
                // FooterReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var footers =
                  (System.Collections.Generic.List<OXmlWordprocessing.FooterReference>)document.Descendants<OXmlWordprocessing.FooterReference>();
                foreach (var footer in footers)
                {
                    footer.Remove();
                }

                // Save the changes.
                document.Save();
            }
        }

        public void ChangeTextInCell(string filepath, string txt)
        {
            // Use the file name and path passed in as an argument to 
            // open an existing document.                        
            // Find the first table in the document.
            DocumentFormat.OpenXml.Wordprocessing.Table table =
                    ((System.Collections.Generic.List<DocumentFormat.OpenXml.Wordprocessing.Table>)this._package.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>())[1];

            // Find the second row in the table.
            DocumentFormat.OpenXml.Wordprocessing.TableRow row = ((System.Collections.Generic.List<DocumentFormat.OpenXml.Wordprocessing.TableRow>)table.Elements<DocumentFormat.OpenXml.Wordprocessing.TableRow>()).ElementAt(1);

            // Find the third cell in the row.
            DocumentFormat.OpenXml.Wordprocessing.TableCell cell = row.Elements<DocumentFormat.OpenXml.Wordprocessing.TableCell>().ElementAt(2);

            // Find the first paragraph in the table cell.
            DocumentFormat.OpenXml.Wordprocessing.Paragraph p = cell.Elements<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().First();

            // Find the first run in the paragraph.
            DocumentFormat.OpenXml.Wordprocessing.Run r = p.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().First();

            // Set the text for the run.
            DocumentFormat.OpenXml.Wordprocessing.Text t = r.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>().First();
            t.Text = txt;

        }

        /////////////////// https://www.cyberforum.ru/windows-forms/thread1501514.html


        public void Colontituls(string tFileName)
        {
            string[] colontituls = new string[9] {"НазваниеПроекта", "ШифрПроекта", "РП", "Разраб",
                                                        "Компания", "Директор", "ГИП", "Н.Контр", "Date"};
            var zip = ZipFile.Open(tFileName, ZipArchiveMode.Read); // Открываем архив на чтение
            string path = @"C:\\Folder";
            zip.ExtractToDirectory(path); // Извлекаем все файлы в директорию
            string footer = ""; // Строка, куда будем считывать файл измененный колонтитул

            footer = FooterRead(path + "\\word\\footer1.xml"); // Вызываем функцию, которая вернет строку с измененным колонтитулом
            FooterReplace(footer, path + "\\word\\footer1.xml"); // Перезаписываем файл с колонтитулом

            ZipFile.CreateFromDirectory(path, @"C:\\1.docx"); // Упаковываем всю папку снова в архив
            Directory.Delete(path, true); // Удаляем папку

        }

        // Метод чтения из колонтиутла
        public string FooterRead(string footerPath)
        {
            string footer = ""; // Строка, в которой будет измененный текст
            using (StreamReader sr = new StreamReader(footerPath)) // Считываем в поток весь файл с колонтитулом
            {
                footer = sr.ReadToEnd(); // Записываем в строку этот поток
                sr.Close(); // Закрываем поток
            }
            /*
            ...
            Делаем замену нужного текста в строке
            ...
            */

            return footer; // Возвращаем строку с измененным колонтитулом
        }

        // Метод перезаписи файлов с колонтитулами
        public void FooterReplace(string footer, string footerPath)
        {
            using (StreamWriter sw = new StreamWriter(footerPath)) // Создаем поток для записи
            {
                sw.Write(footer); // Перезаписываем файл (обновляем колонтитул)
                sw.Close(); // Закрываем поток
            }
        }


        /////////////////////////// https://docs.microsoft.com/ru-ru/office/open-xml/how-to-replace-the-header-in-a-word-processing-document?redirectedfrom=MSDN

        public void AddHeaderFromTo(string filepathFrom, string filepathTo)
        {
            // Replace header in target document with header of source document.            
            {
                OXmlPackaging.MainDocumentPart mainPart = _package.MainDocumentPart;

                // Delete the existing header part.
                mainPart.DeleteParts(mainPart.HeaderParts);

                // Create a new header part.
                DocumentFormat.OpenXml.Packaging.HeaderPart headerPart =
            mainPart.AddNewPart<OXmlPackaging.HeaderPart>();

                // Get Id of the headerPart.
                string rId = mainPart.GetIdOfPart(headerPart);

                // Feed target headerPart with source headerPart.
                using (OXmlPackaging.WordprocessingDocument wdDocSource =
                    OXmlPackaging.WordprocessingDocument.Open(filepathFrom, true))
                {
                    DocumentFormat.OpenXml.Packaging.HeaderPart firstHeader =
            wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

                    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

                    if (firstHeader != null)
                    {
                        headerPart.FeedData(firstHeader.GetStream());
                    }
                }

                // Get SectionProperties and Replace HeaderReference with new Id.
                System.Collections.Generic.IEnumerable<DocumentFormat.OpenXml.Wordprocessing.SectionProperties> sectPrs =
            mainPart.Document.Body.Elements<OXmlWordprocessing.SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to headers.
                    sectPr.RemoveAllChildren<OXmlWordprocessing.HeaderReference>();

                    // Create the new header reference node.
                    sectPr.PrependChild<OXmlWordprocessing.HeaderReference>(new OXmlWordprocessing.HeaderReference() { Id = rId });
                }
            }
        }

        // // https://answacode.com/questions/17726196/kak-vstavit-izobrazhenie-v-zagolovok-dokumenta-openxml-word

        public void AddImageToHeader(string imagePath, string tagName)
        {
            System.Drawing.Bitmap image = new System.Drawing.Bitmap(imagePath);
            OXmlWordprocessing.SdtElement controlBlock = _package.MainDocumentPart.HeaderParts.First().Header.Descendants<OXmlWordprocessing.SdtElement>().Where
                                            (r => r.SdtProperties.GetFirstChild<OXmlWordprocessing.Tag>().Val == tagName).SingleOrDefault();
            //find the Blip element of the content control
            OXmlDrawing.Blip blip = controlBlock.Descendants<OXmlDrawing.Blip>().FirstOrDefault();

            //add image and change embeded id
            OXmlPackaging.ImagePart imagePart = _package.MainDocumentPart.AddImagePart(OXmlPackaging.ImagePartType.Jpeg);
            using (MemoryStream stream = new MemoryStream())
            {
                image.Save(stream, ImageFormat.Jpeg);
                stream.Position = 0;
                imagePart.FeedData(stream);
            }
            blip.Embed = _package.MainDocumentPart.GetIdOfPart(imagePart);
        }

        //

        private void btnMergeWordDocs_Click()
        {
            string sourceFolder = @"C:\Test\MergeDocs\";
            string targetFolder = @"C:\Test\";

            string altChunkIdBase = "acID";
            int altChunkCounter = 1;
            string altChunkId = altChunkIdBase + altChunkCounter.ToString();

            OXmlPackaging.MainDocumentPart wdDocTargetMainPart = null;
            OXmlWordprocessing.Document docTarget = null;
            OXmlPackaging.AlternativeFormatImportPartType afType;
            OXmlPackaging.AlternativeFormatImportPart chunk = null;
            OXmlWordprocessing.AltChunk ac = null;
            //using (WordprocessingDocument wdPkgTarget = WordprocessingDocument.Create(targetFolder + "mergedDoc.docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                //Will create document in 2007 Compatibility Mode.
                //In order to make it 2010 a Settings part must be created and a CompatMode element for the Office version set.
                wdDocTargetMainPart = _package.MainDocumentPart;
                if (wdDocTargetMainPart == null)
                {
                    wdDocTargetMainPart = _package.AddMainDocumentPart();
                    OXmlWordprocessing.Document wdDoc = new OXmlWordprocessing.Document(
                        new OXmlWordprocessing.Body(
                            new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                                new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text() { Text = "First Para" })),
                                new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text() { Text = "Second para" })),
                                new OXmlWordprocessing.SectionProperties(
                                    new OXmlWordprocessing.SectionType() { Val = OXmlWordprocessing.SectionMarkValues.NextPage },
                                    new OXmlWordprocessing.PageSize() { Code = 9 },
                                    new OXmlWordprocessing.PageMargin() { Gutter = 0, Bottom = 1134, Top = 1134, Left = 1318, Right = 1318, Footer = 709, Header = 709 },
                                    new OXmlWordprocessing.Columns() { Space = "708" },
                                    new OXmlWordprocessing.TitlePage())));
                    wdDocTargetMainPart.Document = wdDoc;
                }
                docTarget = wdDocTargetMainPart.Document;
                OXmlWordprocessing.SectionProperties secPropLast = docTarget.Body.Descendants<OXmlWordprocessing.SectionProperties>().Last();
                OXmlWordprocessing.SectionProperties secPropNew = (OXmlWordprocessing.SectionProperties)secPropLast.CloneNode(true);
                //A section break must be in a ParagraphProperty
                DocumentFormat.OpenXml.Wordprocessing.Paragraph lastParaTarget = (DocumentFormat.OpenXml.Wordprocessing.Paragraph)docTarget.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last();
                DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties paraPropTarget = lastParaTarget.ParagraphProperties;
                if (paraPropTarget == null)
                {
                    paraPropTarget = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                }
                paraPropTarget.Append(secPropNew);
                DocumentFormat.OpenXml.Wordprocessing.Run paraRun = lastParaTarget.Descendants<DocumentFormat.OpenXml.Wordprocessing.Run>().FirstOrDefault();
                //lastParaTarget.InsertBefore(paraPropTarget, paraRun);
                lastParaTarget.InsertAt(paraPropTarget, 0);

                //Process the individual files in the source folder.
                //Note that this process will permanently change the files by adding a section break.
                System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(sourceFolder);
                System.Collections.Generic.IEnumerable<System.IO.FileInfo> docFiles = di.EnumerateFiles();
                foreach (System.IO.FileInfo fi in docFiles)
                {
                    using (OXmlPackaging.WordprocessingDocument pkgSourceDoc = OXmlPackaging.WordprocessingDocument.Open(fi.FullName, true))
                    {
                        System.Collections.Generic.IEnumerable<OXmlPackaging.HeaderPart> partsHeader = pkgSourceDoc.MainDocumentPart.GetPartsOfType<OXmlPackaging.HeaderPart>();
                        System.Collections.Generic.IEnumerable<OXmlPackaging.FooterPart> partsFooter = pkgSourceDoc.MainDocumentPart.GetPartsOfType<OXmlPackaging.FooterPart>();
                        //If the source document has headers or footers we want to retain them.
                        //This requires inserting a section break at the end of the document.
                        if (partsHeader.Count() > 0 || partsFooter.Count() > 0)
                        {
                            OXmlWordprocessing.Body sourceBody = pkgSourceDoc.MainDocumentPart.Document.Body;
                            OXmlWordprocessing.SectionProperties docSectionBreak = sourceBody.Descendants<OXmlWordprocessing.SectionProperties>().Last();
                            //Make a copy of the document section break as this won't be imported into the target document.
                            //It needs to be appended to the last paragraph of the document
                            OXmlWordprocessing.SectionProperties copySectionBreak = (OXmlWordprocessing.SectionProperties)docSectionBreak.CloneNode(true);
                            OXmlWordprocessing.Paragraph lastpara = sourceBody.Descendants<OXmlWordprocessing.Paragraph>().Last();
                            OXmlWordprocessing.ParagraphProperties paraProps = lastpara.ParagraphProperties;
                            if (paraProps == null)
                            {
                                paraProps = new DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties();
                                lastpara.Append(paraProps);
                            }
                            paraProps.Append(copySectionBreak);
                        }
                        pkgSourceDoc.MainDocumentPart.Document.Save();
                    }
                    //Insert the source file into the target file using AltChunk
                    afType = OXmlPackaging.AlternativeFormatImportPartType.WordprocessingML;
                    chunk = wdDocTargetMainPart.AddAlternativeFormatImportPart(afType, altChunkId);
                    System.IO.FileStream fsSourceDocument = new System.IO.FileStream(fi.FullName, System.IO.FileMode.Open);
                    chunk.FeedData(fsSourceDocument);
                    //Create the chunk
                    ac = new OXmlWordprocessing.AltChunk();
                    //Link it to the part
                    ac.Id = altChunkId;
                    docTarget.Body.InsertAfter(ac, docTarget.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().Last());
                    docTarget.Save();
                    altChunkCounter += 1;
                    altChunkId = altChunkIdBase + altChunkCounter.ToString();
                    chunk = null;
                    ac = null;
                }
            }
        }

        // /////////////////////// TABLE OF CONTENTS (TOC - оглавление содержание)

        //public OXmlWordprocessing.Paragraph SetHeading1(this OXmlWordprocessing.Paragraph p)
        //{
        //    var pPr = p.Descendants<OXmlWordprocessing.ParagraphProperties>().First();
        //    pPr.ParagraphStyleId = new ParagraphStyleId() { Val = "Heading1" };
        //    return p;
        //}

        public void AddTocToPage2()
        {
            //DocumentFormat.OpenXml.Packaging.par headerPart =
            //mainPart.AddNewPart<HeaderPart>();
            var docPart = this._package.MainDocumentPart;
            OXmlWordprocessing.Document document = docPart.Document;
            var paragraphToc =
                  (System.Collections.Generic.List<OXmlWordprocessing.Paragraph>)document.Descendants<OXmlWordprocessing.Paragraph>();
            //paragraphToc.add

        }

        public void AddToc()
        {
            var sdtBlock = new OXmlWordprocessing.SdtBlock();
            //sdtBlock.InnerXml = GetTOC(Translations.ResultsBooksTableOfContentsTitle, 16);
            sdtBlock.InnerXml = GetTOC("Содержание", 16);
            _package.MainDocumentPart.Document.Body.AppendChild(sdtBlock);

            DocumentFormat.OpenXml.Wordprocessing.SimpleField f;
            f = new OXmlWordprocessing.SimpleField();
            f.Instruction = "sdtContent";
            f.Dirty = true;
            _package.MainDocumentPart.Document.Body.Append(f);


            var setting = _package.MainDocumentPart.DocumentSettingsPart;
            if (setting != null)
            {
                _package.MainDocumentPart.DocumentSettingsPart.Settings.Append(new OXmlWordprocessing.UpdateFieldsOnOpen() { Val = new OXml.OnOffValue(true) });
                _package.MainDocumentPart.DocumentSettingsPart.Settings.Save();
            }

            //var settingsPart = _package.MainDocumentPart.AddNewPart<ContentPart>().First();

            //var settingsPart = _package.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
            //settingsPart.Settings = new Settings { BordersDoNotSurroundFooter = new BordersDoNotSurroundFooter() { Val = true } };
            //settingsPart.Settings.Append(new UpdateFieldsOnOpen() { Val = true });

        }

        private string GetTOC(string title, int titleFontSize)
        {
            return $@"<w:sdt>
     <w:sdtPr>
        <w:id w:val=""-493258456"" />
        <w:docPartObj>
           <w:docPartGallery w:val=""Table of Contents"" />
           <w:docPartUnique />
        </w:docPartObj>
     </w:sdtPr>
     <w:sdtEndPr>
        <w:rPr>
           <w:rFonts w:asciiTheme=""minorHAnsi"" w:eastAsiaTheme=""minorHAnsi"" w:hAnsiTheme=""minorHAnsi"" w:cstheme=""minorBidi"" />
           <w:b />
           <w:bCs />
           <w:noProof />
           <w:color w:val=""auto"" />
           <w:sz w:val=""22"" />
           <w:szCs w:val=""22"" />
        </w:rPr>
     </w:sdtEndPr>
     <w:sdtContent>
        <w:p w:rsidR=""00095C65"" w:rsidRDefault=""00095C65"">
           <w:pPr>
              <w:pStyle w:val=""TOCHeading"" />
              <w:jc w:val=""center"" /> 
           </w:pPr>
           <w:r>
                <w:rPr>
                  <w:b /> 
                  <w:color w:val=""2E74B5"" w:themeColor=""accent1"" w:themeShade=""BF"" /> 
                  <w:sz w:val=""{titleFontSize * 2}"" /> 
                  <w:szCs w:val=""{titleFontSize * 2}"" /> 
              </w:rPr>
              <w:t>{title}</w:t>
           </w:r>
        </w:p>
        <w:p w:rsidR=""00095C65"" w:rsidRDefault=""00095C65"">
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType=""begin"" />
           </w:r>
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:instrText xml:space=""preserve""> TOC \o ""1-3"" \h \z \u </w:instrText>
           </w:r>
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType=""separate"" />
           </w:r>
           <w:r>
              <w:rPr>
                 <w:noProof />
              </w:rPr>
              <w:t>No table of contents entries found.</w:t>
           </w:r>
           <w:r>
              <w:rPr>
                 <w:b />
                 <w:bCs />
                 <w:noProof />
              </w:rPr>
              <w:fldChar w:fldCharType=""end"" />
           </w:r>
        </w:p>
     </w:sdtContent>
  </w:sdt>";
        }


        // ////////////////////////////////////////////////////////////


        /////////////////////////////////////////////
        /// <summary>
        /// Запись формул в документ Word
        /// </summary>
        /// <param name="documentParent"></param>
        /// <param name="documentChild"></param>
        /// 

        public void MathML2Word(string EqMathML, string docx)
        {
            XslCompiledTransform xslTransform = new XslCompiledTransform();
            //xslTransform.Load(@"C:\Program Files (x86)\Microsoft Office\Office14\MML2OMML.xsl");
            xslTransform.Load("MML2OMML.XSL");

            // Load the file containing your MathML presentation markup.
            //using (XmlReader reader = XmlReader.Create(File.Open("../../../test1.xml", FileMode.Open)))
            //using (XmlReader reader = XmlReader.Create(File.Open(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\test2.xml", FileMode.Open)))
            using (XmlReader reader = XmlReader.Create(File.Open(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\OpenXML-SDK\mathMM.xml", FileMode.Open)))
            //using (XmlReader reader = XmlReader.Create(File.Open(@"d:\_2022___\Src_3__\mathml007.htm", FileMode.Open)))

            //using (XmlReader reader = XmlReader.Create(File.Open(EqMathML, FileMode.Open)))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlWriterSettings settings = xslTransform.OutputSettings.Clone();

                    // Configure xml writer to omit xml declaration.
                    settings.ConformanceLevel = ConformanceLevel.Fragment;
                    settings.OmitXmlDeclaration = true;
                    XmlWriter xw = XmlWriter.Create(ms, settings);
                    // Transform our MathML to OfficeMathML
                    xslTransform.Transform(reader, xw);
                    ms.Seek(0, SeekOrigin.Begin);
                    StreamReader sr = new StreamReader(ms, Encoding.UTF8);

                    string officeML = sr.ReadToEnd();
                    Console.Out.WriteLine(officeML);

                    // Create a OfficeMath instance from the OfficeMathML xml.
                    DocumentFormat.OpenXml.Math.OfficeMath om = new DocumentFormat.OpenXml.Math.OfficeMath(officeML);

                    //创建Word文档(Microsoft.Office.Interop.Word)  
                    //Microsoft.Office.Interop.Word._Application WordApp = new Application();
                    //WordApp.Visible = true;
                    using (OXmlPackaging.WordprocessingDocument _package = OXmlPackaging.WordprocessingDocument.Create(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\template.docx", OXml.WordprocessingDocumentType.Document))
                    {
                        // Add a new main document part. 
                        _package.AddMainDocumentPart();

                        // Create the Document DOM. 
                        _package.MainDocumentPart.Document =
                          new OXmlWordprocessing.Document(
                            new OXmlWordprocessing.Body(
                              new OXmlWordprocessing.Paragraph(
                                new OXmlWordprocessing.Run(
                                  new OXmlWordprocessing.Text("  ")))));

                        // Save changes to the main document part. 
                        _package.MainDocumentPart.Document.Save();
                    }

                    //using (OXmlPackaging.WordprocessingDocument wordDoc = OXmlPackaging.WordprocessingDocument.Open(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\template.docx", true))
                    {
                        DocumentFormat.OpenXml.Wordprocessing.Paragraph par =
                          _package.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().FirstOrDefault();

                        foreach (var currentRun in om.Descendants<DocumentFormat.OpenXml.Math.Run>())
                        {
                            // Add font information to every run.
                            DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties2 =
                              new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                            currentRun.InsertAt(runProperties2, 0);
                        }
                        par.Append(om);
                    }
                }
            }
            System.Diagnostics.Process.Start(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\template.docx");
        }

        //public void MathML2Word__(string EqMathML)
        public void MathML2Word(string latexExpression)
        {
            Latex2MathML.LatexToMathMLConverter lmm = new Latex2MathML.LatexToMathMLConverter();
            lmm.ValidateResult = true;
            //lmm.BeforeXmlFormat += MyEventListener;
            //lmm.ExceptionEvent += ExceptionListener;
            lmm.Convert(latexExpression);

            string EqMathML = lmm.Output;

            XslCompiledTransform xslTransform = new XslCompiledTransform();
            //xslTransform.Load(@"C:\Program Files (x86)\Microsoft Office\Office14\MML2OMML.xsl");
            xslTransform.Load(@"d:\_2022___\Src_3__\libopenxml\LibOpenXML\MML2OMML.XSL");

            // Load the file containing your MathML presentation markup.
            //using (XmlReader reader = XmlReader.Create(File.Open("../../../test1.xml", FileMode.Open)))
            //using (XmlReader reader = XmlReader.Create(File.Open(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\test2.xml", FileMode.Open)))
            //using (XmlReader reader = XmlReader.Create(File.Open(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\OpenXML-SDK\mathMM.xml", FileMode.Open)))
            //using (XmlReader reader = XmlReader.Create(File.Open(@"d:\_2022___\Src_3__\mathml007.htm", FileMode.Open)))

            XDocument docxml = null;
            try
            {
                docxml = XDocument.Parse(EqMathML);
            }
            catch (Exception ex)
            {
                CounterFormula.Count++;
                File.AppendAllText(CounterFormula.fn, "[" + CounterFormula.Count + "]  " + latexExpression + "   \n");
                return;

            }
            //XmlReader reader = docxml.CreateReader();
            using (XmlReader reader = docxml.CreateReader())
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlWriterSettings settings = xslTransform.OutputSettings.Clone();

                    // Configure xml writer to omit xml declaration.
                    settings.ConformanceLevel = ConformanceLevel.Fragment;
                    settings.OmitXmlDeclaration = true;
                    XmlWriter xw = XmlWriter.Create(ms, settings);
                    // Transform our MathML to OfficeMathML
                    xslTransform.Transform(reader, xw);
                    ms.Seek(0, SeekOrigin.Begin);
                    StreamReader sr = new StreamReader(ms, Encoding.UTF8);

                    string officeML = sr.ReadToEnd();
                    Console.Out.WriteLine(officeML);

                    // Create a OfficeMath instance from the OfficeMathML xml.
                    DocumentFormat.OpenXml.Math.OfficeMath om = new DocumentFormat.OpenXml.Math.OfficeMath(officeML);

                    DocumentFormat.OpenXml.Wordprocessing.Paragraph par =
                      _package.MainDocumentPart.Document.Body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Paragraph>().LastOrDefault(); //FirstOrDefault();

                    foreach (var currentRun in om.Descendants<DocumentFormat.OpenXml.Math.Run>())
                    {
                        // Add font information to every run.
                        DocumentFormat.OpenXml.Wordprocessing.RunProperties runProperties2 =
                          new DocumentFormat.OpenXml.Wordprocessing.RunProperties();
                        currentRun.InsertAt(runProperties2, 0);
                    }
                    par.Append(om);

                }
            }
            //System.Diagnostics.Process.Start(@"d:\_2022___\_Src_inet__\_OpenXml___\MathML2Word-master\template.docx");

        }


        ////////////////////////////////////////////////// index substring
        
        public string fontName = "Microsoft Yahei";
        public string fontSize = "21";
        public string fontSizeComplexScript = "24";

        public void AddMath(string UpStr = "X", string DownStr = "i")
        {
            AddMath(GetSubScript(UpStr,DownStr));
        }
        public void AddMath(OXmlMath.OfficeMath oMath)
        {
            OXmlWordprocessing.Paragraph paragraph = new OXmlWordprocessing.Paragraph();
            OXmlWordprocessing.ParagraphProperties paragraphProperties1 = new OXmlWordprocessing.ParagraphProperties();
            OXmlWordprocessing.ParagraphMarkRunProperties paragraphMarkRunProperties = new OXmlWordprocessing.ParagraphMarkRunProperties();
            OXmlWordprocessing.RunFonts runFonts1 = new OXmlWordprocessing.RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName };
            OXmlWordprocessing.FontSize fontSize1 = new OXmlWordprocessing.FontSize() { Val = fontSize };
            paragraphMarkRunProperties.Append(runFonts1);
            paragraphMarkRunProperties.Append(fontSize1);
            OXmlWordprocessing.Justification justification = new OXmlWordprocessing.Justification() { Val = OXmlWordprocessing.JustificationValues.Center };
            paragraphProperties1.Append(paragraphMarkRunProperties);
            paragraphProperties1.Append(justification);
            paragraph.Append(paragraphProperties1);
            paragraph.Append(oMath);

            OXmlWordprocessing.Paragraph par =
                      _package.MainDocumentPart.Document.Body.Descendants< OXmlWordprocessing.Paragraph >().LastOrDefault(); //FirstOrDefault();
            par.Append(paragraph);
        }

        public OXmlMath.OfficeMath GetSubScript(string baseString, string downString)
        {
            OXmlMath.OfficeMath officeMath = new OXmlMath.OfficeMath();
            OXmlMath.Subscript subScript = GenerateSubscript(baseString, downString);
            officeMath.Append(subScript);
            return officeMath;
        }

        public OXmlMath.Subscript GenerateSubscript(string baseString, string downString)
        {
            OXmlMath.Subscript subscript1 = new OXmlMath.Subscript();
            OXmlMath.Base base1 = new OXmlMath.Base();
            OXmlMath.Run run1 = GenerateMathRun(baseString);
            base1.Append(run1);
            OXmlMath.SubArgument subArgument = new OXmlMath.SubArgument();
            OXmlMath.Run run = GenerateMathRun(downString);
            subArgument.Append(run);
            subscript1.Append(base1);
            subscript1.Append(subArgument);
            return subscript1;
        }

        private OXmlMath.Run GenerateMathRun(string baseString)
        {
            OXmlMath.Run run1 = new OXmlMath.Run();
            OXmlWordprocessing.RunProperties runProperties = new OXmlWordprocessing.RunProperties();
            OXmlWordprocessing.RunFonts runFonts1 = new OXmlWordprocessing.RunFonts() { Ascii = fontName, HighAnsi = fontName, EastAsia = fontName };
            OXmlWordprocessing.FontSize fontSize1 = new OXmlWordprocessing.FontSize() { Val = fontSize };
            OXmlWordprocessing.FontSizeComplexScript fsc = new OXmlWordprocessing.FontSizeComplexScript { Val = fontSizeComplexScript };
            run1.Append(new OXmlWordprocessing.RunProperties(runFonts1, fontSize1, fsc));

            run1.Append(new OXmlWordprocessing.Text() { Text = baseString });
            return run1;
        }
        //////////////////////////////////////////////////        


        public void MergeDoc(string documentParent, string documentChild)
        {
            using (WordWriter2 builder = new WordWriter2())
            {
                OpenDocument(documentParent);
                MergeDocument(documentChild);                                               
            }
        }

        public void CloseDontSave()
        {
            Dispose();
        }


        public void Dispose()
        {

            if (this._package != null)
            {

                this._package.Close();
                this._package.Dispose();
            }
        }
    }
}
