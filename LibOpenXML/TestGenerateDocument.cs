using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LibOpenXml.Charts;
using LibOpenXml.Common;
using LibOpenXml.Excel;
using LibOpenXml.Word;

namespace LibOpenXml
{
    public class TestGenerateDocument
    {
        public static string figure = @"d:\_2022___\gazonokosilshik.jpg";

        // ---------------------------------------------------- Generate WORD Documents ------------------------------------------

        public void GenerateWordDocMerge()
        {
            string document = "test.docx";
            string document2 = "SvayaSAPR.dotx";

            if (File.Exists(document))
                File.Delete(document);

            using (WordWriter2 builder = new WordWriter2())
            {

                builder.CreateDocument(document);

                builder.AddCustomStyle(new FormatStyle()
                {
                    Color = System.Drawing.Color.Maroon,
                    FontName = "Courier New",
                    FontSize = 14,
                    IsBold = true,
                    IsItalic = true,
                    Name = "Warning",
                    StyleId = "Warning",
                    HighlightColor = FormatStyle.HighlightColors.Black
                });
            }

            using (WordWriter2 builder = new WordWriter2())
            {
                builder.MergeDoc(document, document2);
                
            }


            using (WordWriter2 builder = new WordWriter2())
            {
                builder.OpenDocument(document);

                builder.AddCustomStyle(new FormatStyle()
                {
                    Color = System.Drawing.Color.Maroon,
                    FontName = "Courier New",
                    FontSize = 14,
                    IsBold = true,
                    IsItalic = true,
                    Name = "Warning",
                    StyleId = "Warning",
                    HighlightColor = FormatStyle.HighlightColors.Black
                });

                builder.AddToc();
                builder.AppendBreakPage();
                builder.AppendBreakPage();

                builder.CreateTextParagraph(TextParagraphType.Title, "Owl");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Owl is a custom framework to make easier the document's development using OpenXML. Enjoy it!");
                builder.CreateTextParagraph(TextParagraphType.Normal, string.Empty);

                builder.CreateTextParagraph(TextParagraphType.Title, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, "Heading 1");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading2, "Heading 2");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading3, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(document);
            Process.Start(startInfo);
                        
        }


        public void GenerateWordDoc()
        {
            string document = "test.docx";

            if (File.Exists(document))
                File.Delete(document);

            using (WordWriter2 builder = new WordWriter2())
            {

                builder.CreateDocument(document);

                builder.AddCustomStyle(new FormatStyle()
                {
                    Color = System.Drawing.Color.Maroon,
                    FontName = "Courier New",
                    FontSize = 14,
                    IsBold = true,
                    IsItalic = true,
                    Name = "Warning",
                    StyleId = "Warning",
                    HighlightColor = FormatStyle.HighlightColors.Black
                });

                builder.AddToc();

                builder.AppendBreakPage();

                builder.CreateTextParagraph(TextParagraphType.Title, "Owl");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Owl is a custom framework to make easier the document's development using OpenXML. Enjoy it!");
                builder.CreateTextParagraph(TextParagraphType.Normal, string.Empty);

                builder.CreateTextParagraph(TextParagraphType.Title, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, "Heading 1");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading2, "Heading 2");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading3, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, "None style text.");
                builder.CreateTextParagraph("Warning", "Warning to this point. Pay attention!");
                

                builder.CreateTextParagraph(TextParagraphType.Title, HorizontalAlignmentType.Center, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, HorizontalAlignmentType.Center, "Heading 1");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading2, HorizontalAlignmentType.Center, "Heading 2");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading3, HorizontalAlignmentType.Center, "Heading 3");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Normal, HorizontalAlignmentType.Center, "Normal text.");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.None, HorizontalAlignmentType.Center, "None style text.");
                builder.CreateTextParagraph("Warning", HorizontalAlignmentType.Center, "Warning to this point. Pay attention!");

                builder.CreateTextParagraph(TextParagraphType.Title, HorizontalAlignmentType.Right, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, HorizontalAlignmentType.Right, "Heading 1");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading2, HorizontalAlignmentType.Right, "Heading 2");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading3, HorizontalAlignmentType.Right, "Heading 3");
                builder.AppendBreakPage();
                builder.CreateTextParagraph(TextParagraphType.Heading3, HorizontalAlignmentType.Right, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, HorizontalAlignmentType.Right, "None style text.");
                builder.CreateTextParagraph("Warning", HorizontalAlignmentType.Right, "Warning to this point. Pay attention!");

                //builder.AddImage(figure);
                //builder.AddImage(figure, HorizontalAlignmentType.Center);
                //builder.AddImage(figure, HorizontalAlignmentType.Right);
                builder.AddImage(figure, 0.75M);
                builder.AddImage(figure, 0.50M);
                builder.AddImage(figure, 0.4M, HorizontalAlignmentType.Center);
                builder.AddImage(figure, 0.25M, HorizontalAlignmentType.Right);
                                
            }

            using (WordWriter2 builder = new WordWriter2())
            {

                builder.OpenDocument(document);

                builder.CreateTextParagraph(TextParagraphType.Title, "Title");
                builder.CreateTextParagraph(TextParagraphType.Heading1, "Heading 1");
                builder.CreateTextParagraph(TextParagraphType.Heading2, "Heading 2");
                builder.CreateTextParagraph(TextParagraphType.Heading3, "Heading 3");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.Normal, "Normal text.");
                builder.CreateTextParagraph(TextParagraphType.None, "None style text.");
                builder.CreateTextParagraph("Warning", "Warning to this point. Pay attention!");

                //builder.AddImage(figure);
                //builder.AddImage(figure, HorizontalAlignmentType.Center);
                //builder.AddImage(figure, HorizontalAlignmentType.Right);

                //AddTable(builder);

                builder.CreateTextParagraph(TextParagraphType.Heading3, @"\begin{document} $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$ \end{document}   -----     ");
                builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                builder.MathML2Word(@"\begin{document} $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$ \end{document}");

                builder.CreateTextParagraph(TextParagraphType.Heading3, @"\begin{document} $\sqrt{x^2+1}$ \end{document}  -----    ");
                builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                builder.MathML2Word(@"\begin{document} $\sqrt{x^2+1}$ \end{document}");

                builder.CreateTextParagraph(TextParagraphType.Heading3, @"\begin{document} $\[ \sum_{i=1}^{\infty} \frac{1}{n^s} = \prod_p \frac{ 1} { 1 - p ^{ -s} } \]$ \end{document}  -----    ");
                builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                builder.MathML2Word(@"\begin{document} $\[ \sum_{i=1}^{\infty} \frac{1}{n^s} = \prod_p \frac{ 1} { 1 - p ^{ -s} } \]$ \end{document}");

                //"F_u = γ_t γ_c \sum_{ }^{ } R_{ (af, i)} \dot A_{ (af, i)}"
                builder.CreateTextParagraph(TextParagraphType.Normal, @"\begin{document} $F_u = γ_t γ_c \sum_{}^{}R_{af, i} A_{af, i}$ \end{document}  -----    ");
                builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                builder.MathML2Word(@"\begin{document} $F_u = γ_t γ_c \sum_{}^{}R_{af, i} A_{af, i}$ \end{document}");

                //builder.CreateTextParagraph(TextParagraphType.Normal, @"\begin{document} $$ \end{document}  -----    ");
                //builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                //builder.MathML2Word__(@"\begin{document} $$ \end{document}");

                //builder.CreateTextParagraph(TextParagraphType.Heading3, @"\begin{document} $x = a_0 + \cfrac{ 1}  {  a_1 + \cfrac{ 1}   {  a_2  + \cfrac{ 1} { a_3 + \cfrac{ 1} { a_4} } } }$ \end{document}  -----    ");
                //builder.MathML2Word__(@"\begin{document} \begin{equation}$x = a_0 + \cfrac{ 1}  {  a_1 + \cfrac{ 1}   {  a_2  + \cfrac{ 1} { a_3 + \cfrac{ 1} { a_4} } } }$\end{equation} \end{document}");

                //builder.CreateTextParagraph(TextParagraphType.Normal, @"\begin{document} $\frac{ \begin{ array}[b]{r} \left(x_1 x_2 \right)\\\times \left(x'_1 x'_2 \right) \end{array}}{\left(y_1y_2y_3y_4 \right)}}$ \end{document}  -----    ");
                //builder.MathML2Word__(@"\begin{document} $\frac{ \begin{ array}[b]{r} \left(x_1 x_2 \right)\\\times \left(x'_1 x'_2 \right) \end{array}}{\left(y_1y_2y_3y_4 \right)}$ \end{document}");

                // \begin{equation} x = a_0 + \cfrac{ 1}  {  a_1 + \cfrac{ 1}   {  a_2  + \cfrac{ 1} { a_3 + \cfrac{ 1} { a_4} } } }  \end{ equation}
                // \frac{ \begin{ array}[b]{r} \left(x_1 x_2 \right)\\\times \left(x'_1 x'_2 \right) \end{array}}{\left(y_1y_2y_3y_4 \right)}

                //builder.CreateTextParagraph(TextParagraphType.Heading3, @"\begin{document} $ν=(0,45∙[(T_bf-T_0^' )/A]^(1/3)∙σ∙D_e)/(T_bf-T_e-C∙√(T_bf-T_e )  )=(0,45∙[(-0,72-(-2))/22]^(1/3)∙1,33∙0,537)/(-0,72-(-1,681)-0,24∙√((-0,72-(-1,681)))  )=0,171  $ \end{document}  -----    ");
                //builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                //builder.MathML2Word__(@"\begin{document} $ν=\frac{(0,45∙\[(T_bf-T_0^' )/A\]^(1/3)∙σ∙D_e)}{(T_bf-T_e-C∙\sqrt(T_bf-T_e )}  )=$ \end{document}");
                //builder.MathML2Word__(@"\begin{document} $ν=(0,45∙[(T_bf-T_0^' )/A]^(1/3)∙σ∙D_e)/(T_bf-T_e-C∙\sqrt(T_bf-T_e )  )=(0,45∙[(-0,72-(-2))/22]^(1/3)∙1,33∙0,537)/(-0,72-(-1,681)-0,24∙\sqrt((-0,72-(-1,681)))  )=0,171  $ \end{document}");


                //builder.CreateTextParagraph(TextParagraphType.Normal, @"\begin{document} $$ \end{document}  -----    ");
                //builder.CreateTextParagraph(TextParagraphType.None, @"   ");
                //builder.MathML2Word__(@"\begin{document} $$ \end{document}");


                // ν=(0,45∙[(T_bf-T_0^' )/A]^(1/3)∙σ∙D_e)/(T_bf-T_e-C∙√(T_bf-T_e )  )=(0,45∙[(-0,72-(-2))/22]^(1/3)∙1,33∙0,537)/(-0,72-(-1,681)-0,24∙√((-0,72-(-1,681)))  )=0,171  
                // \sum_{\substack{  0 < i < m \\ 0 < j < n }} P(i, j)
                // P\left(A=2\middle|\frac{A^2}{B}>4\right)

                builder.CreateTextParagraph(TextParagraphType.Heading3, HorizontalAlignmentType.Right, "R(a,b)");
                builder.CreateTextParagraph(TextParagraphType.None, @"  **************************  ");
                builder.AddMath("R","a,b"); builder.AddMath("S", "t"); builder.AddMath("Q", "t"); builder.AddMath("D", "z"); builder.CreateTextParagraph(TextParagraphType.None, "  **************************  ");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(document);
            Process.Start(startInfo);
        }


        private static void AddTable(WordWriter2 builder)
        {

            using (var table = new DataTable())
            {

                table.Columns.Add("Item", typeof(string));
                table.Columns.Add("Min", typeof(string));
                table.Columns.Add("Avg", typeof(string));
                table.Columns.Add("Max", typeof(string));

                for (int i = 0; i < 10; i++)
                {

                    var r = table.NewRow();
                    r[0] = i.ToString(); r[1] = 10M.ToString(); r[2] = 11M.ToString(); r[3] = 12M.ToString();

                    table.Rows.Add(r);
                }

                builder.CreateTable(table, new TableStyle()
                {
                    Alignment = HorizontalAlignmentType.Center,
                    ShowTitle = true,
                    Title = "My Table"
                });
            }
        }

        // ---------------------------------------------------- Generate EXCEL Documents ------------------------------------------

        public void GenerateExcelDoc()
        {
            string document = "test.xlsx";

            if (File.Exists(document))
                File.Delete(document);

            var data = GetData();

            using (ExcelBuilder builder = new ExcelBuilder())
            {

                builder.CreateDocument(document);

                builder.ImportData(data);

                var table1 = GetTable("Simple import 1");
                builder.ImportData(table1);

                var table2 = GetTable("Simple import 2");
                builder.ImportData(table2, "Table Alias");
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(document);
            Process.Start(startInfo);
        }

        static DataSet GetData()
        {

            DataSet returnValue = new DataSet();

            for (int i = 0; i < 100; i++)
            {

                var table = GetTable("My Sample Chart " + i);
                returnValue.Tables.Add(table);
            }

            return returnValue;
        }

        static DataTable GetTable(string name)
        {

            DataTable data = new DataTable(name);

            data.Columns.AddRange(new DataColumn[]{
                    new DataColumn("Time", typeof(DateTime)),
                    new DataColumn("Group 1", typeof(int)),
                    new DataColumn("Group 2", typeof(int)),
                    new DataColumn("Group 3", typeof(int)),
                    new DataColumn("Group 4", typeof(int)),
                });

            DateTime baseDate = new DateTime(2015, 01, 01);
            Random rnd = new Random();
            for (int j = 0; j < 1000; j++)
            {

                int limiar = j * 10;

                data.Rows.Add(baseDate.AddDays(j),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30));
            }

            return data;
        }

        // ---------------------------------------------------- Generate Chart Documents ------------------------------------------

        public void GenerateCharts()
        {

            string image = "pieChart.png";

            if (File.Exists(image))
                File.Delete(image);

            var data = new PieChartItem[] {
              new PieChartItem() { Label= "Group 1", Value = 100 },
              new PieChartItem() { Label= "Group 2", Value = 102 },
              new PieChartItem() { Label= "Group 3", Value = 103.565M },
              new PieChartItem() { Label= "Group 4", Value = 203.47M },
              new PieChartItem() { Label= "Group 5", Value = 100 },
              new PieChartItem() { Label= "Group 6", Value = 102 },
              new PieChartItem() { Label= "Group 7", Value = 103.565M },
              new PieChartItem() { Label= "Group 8", Value = 203.47M },
              new PieChartItem() { Label= "Group 9", Value = 100 },
              new PieChartItem() { Label= "Group 10", Value = 120 },
              new PieChartItem() { Label= "Group 11", Value = 130.565M },
              new PieChartItem() { Label= "Group 12", Value = 2300.47M }
            };

            ChartBuilder chartBuilder = new ChartBuilder();
            chartBuilder.BuildPieChart(image,
                                       data,
                                       new PieChartContext()
                                       {
                                           Title = "My sample chart!",
                                           IsLabelOutside = true,
                                           LabelStyle = PieChartLabelStyle.LabelPercent
                                       });

            ProcessStartInfo startInfo = new ProcessStartInfo(image);
            Process.Start(startInfo);
        }

        private static void CreateLineChart()
        {

            string image = "lineChart.png";

            if (File.Exists(image))
                File.Delete(image);

            DataTable data = new DataTable("My Sample Chart");
            data.Columns.AddRange(new DataColumn[]{
                new DataColumn("Time", typeof(DateTime)),
                new DataColumn("Group 1", typeof(int)),
                new DataColumn("Group 2", typeof(int)),
                new DataColumn("Group 3", typeof(int)),
                new DataColumn("Group 4", typeof(int)),
            });

            DateTime baseDate = new DateTime(2015, 01, 01);
            Random rnd = new Random();
            for (int i = 0; i < 10; i++)
            {

                int limiar = i * 10;

                data.Rows.Add(baseDate.AddDays(i),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30),
                              rnd.Next(limiar, limiar + 30));
            }

            ChartBuilder chartBuilder = new ChartBuilder();
            chartBuilder.BuildLineChart(image,
                                        data,
                                        new LineChartContext()
                                        {
                                            Title = "My sample chart!",
                                            ImageSize = new ImageSize() { Width = 1390, Height = 464 }
                                        });

            ProcessStartInfo startInfo = new ProcessStartInfo(image);
            Process.Start(startInfo);
        }
    }
}
