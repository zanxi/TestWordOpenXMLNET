﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Data;

namespace LibOpenXml.Word {

    internal class TableCreator {

        private enum RowIdentification {

            Title, 
            Header, 
            Row
        }

        public void AddTable(WordprocessingDocument package, DataTable table, TableStyle tableStyle) {

            Body body = package.MainDocumentPart.Document.Body;

            // Create an empty table.
            Table tbl = new Table();

            // Set table width
            SetTableWidth(tableStyle, tbl);

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = CreateTableProperties();

            // Set table alignment
            SetTableAlignment(tableStyle.Alignment, tblProp);

            // Append the TableProperties object to the empty table.
            tbl.AppendChild<TableProperties>(tblProp);

            //if (tableStyle.ShowTitle) {
            //    AddTitleRow(table, tableStyle, tbl);
            //}

            //if (tableStyle.ShowHeader) {
            //    AddHeaderRow(table, tbl, tableStyle);
            //}

            AddRows(table, tableStyle, tbl, tblProp);

            // Append the final table to the document body
            body.Append(tbl);
        }

        public void AddTable(WordprocessingDocument package, Table tbl, TableStyle tableStyle)
        {

            Body body = package.MainDocumentPart.Document.Body;

            // Create an empty table.
            //Table tbl = new Table();

            // Set table width
            SetTableWidth(tableStyle, tbl);

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = CreateTableProperties();

            // Set table alignment
            SetTableAlignment(tableStyle.Alignment, tblProp);

            // Append the TableProperties object to the empty table.
            tbl.AppendChild<TableProperties>(tblProp);

            //if (tableStyle.ShowTitle)
            //{
            //    AddTitleRow(table, tableStyle, tbl);
            //}

            //if (tableStyle.ShowHeader)
            //{
            //    AddHeaderRow(table, tbl, tableStyle);
            //}

            //AddRows(table, tableStyle, tbl, tblProp);

            // Append the final table to the document body
            body.Append(tbl);
        }

        private static TableProperties CreateTableProperties() {
            return new TableProperties(
                            new TableBorders(
                                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 }
                            )
                        );
        }

        private static void SetTableWidth(TableStyle tableStyle, Table tbl) {

            if (tableStyle.WidthUnit == TableStyle.TableWidthUnit.Percent) {

                double percent = tableStyle.Width / 100D;
                int wordPercent = (int)(5000 * (percent));

                TableWidth width = new TableWidth() {
                    Type = TableWidthUnitValues.Pct,
                    Width = wordPercent.ToString()
                };
                tbl.AppendChild<TableWidth>(width);
            }
        }

        private static void AddRows(DataTable table, TableStyle tableStyle, Table wordTable, TableProperties tableProperties) {

            for (int r = 0; r < table.Rows.Count; r++) {

                // Create a row.
                TableRow tRow = new TableRow();

                for (int c = 0; c < table.Columns.Count; c++) {

                    // Create a cell.
                    TableCell tCell = new TableCell();

                    if ((r % 2 == 0) && (tableStyle.EnableAlternativeBackgroundColor)) {

                        var tCellProperties = new TableCellProperties();

                        // Set Cell Color
                        Shading tCellColor = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = System.Drawing.ColorTranslator.ToHtml(tableStyle.AlternativeBackgroundColor) };
                        tCellProperties.Append(tCellColor);

                        tCell.Append(tCellProperties);
                    }

                    string rowContent = table.Rows[r][c] == null ? string.Empty : table.Rows[r][c].ToString();

                    ParagraphProperties paragProperties = new ParagraphProperties();
                    SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
                    paragProperties.Append(spacingBetweenLines1);

                    Run run = new Run();
                    if (rowContent.Contains(Environment.NewLine)) {

                        string[] lines = rowContent.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                        for (int i = 0; i < lines.Length; i++) {

                            var line = lines[i];
                            if (i > 0) {
                                run.AppendChild(new Break());
                            }
                            Text newText = new Text(line);
                            run.AppendChild<Text>(newText);
                        }
                    }
                    else {
                        Text newText = new Text(rowContent);
                        run.AppendChild<Text>(newText);
                    }

                    ApplyFontProperties(tableStyle, RowIdentification.Row, run);

                    var parag = new Paragraph();
                    parag.Append(paragProperties);
                    parag.Append(run);

                    // Specify the table cell content.
                    tCell.Append(parag);

                    // Append the table cell to the table row.
                    tRow.Append(tCell);
                }

                // Append the table row to the table.
                wordTable.Append(tRow);
            }
        }

        private static void ApplyFontProperties(TableStyle tableStyle, RowIdentification rowIdentification, Run run) {

            int fontSize = 0;

            switch (rowIdentification) {
                case RowIdentification.Title:
                    fontSize = tableStyle.TitleFontSize;
                    break;
                case RowIdentification.Header:
                    fontSize = tableStyle.HeaderFontSize;
                    break;
                case RowIdentification.Row:                    
                default:
                    fontSize = tableStyle.RowFontSize;
                    break;
            }

            RunProperties runProp = new RunProperties();
            FontSize size = new FontSize();
            size.Val = new StringValue((fontSize * 2).ToString());
            runProp.Append(size);

            run.PrependChild<RunProperties>(runProp);
        }

        private static void AddHeaderRow(DataTable table, Table wordTable, TableStyle tableStyle) {

            // Create a row.
            TableRow tRow = new TableRow();

            foreach (DataColumn iColumn in table.Columns) {

                // Create a cell.
                TableCell tCell = new TableCell();

                TableCellProperties tCellProperties = new TableCellProperties();

                // Set Cell Color
                Shading tCellColor = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = System.Drawing.ColorTranslator.ToHtml(tableStyle.HeaderBackgroundColor) };
                tCellProperties.Append(tCellColor);

                // Append properties to the cell
                tCell.Append(tCellProperties);

                ParagraphProperties paragProperties = new ParagraphProperties();
                Justification justification1 = new Justification() { Val = JustificationValues.Center };
                SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
                paragProperties.Append(spacingBetweenLines1);
                paragProperties.Append(justification1);

                var parag = new Paragraph();
                parag.Append(paragProperties);

                var run = new Run(new Text(iColumn.ColumnName));
                ApplyFontProperties(tableStyle, RowIdentification.Header, run);
                parag.Append(run);

                // Specify the table cell content.
                tCell.Append(parag);

                // Append the table cell to the table row.
                tRow.Append(tCell);
            }

            // Append the table row to the table.
            wordTable.Append(tRow);
        }

        private static void AddTitleRow(DataTable table, TableStyle tableStyle, Table wordTable) {

            // Create a row.
            TableRow tRow = new TableRow();

            // Create a cell.
            TableCell tCell = new TableCell();

            TableCellProperties tCellProperties = new TableCellProperties();

            // Set Cell Color
            Shading tCellColor = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = System.Drawing.ColorTranslator.ToHtml(tableStyle.TitleBackgroundColor) };
            tCellProperties.Append(tCellColor);

            // Set Cell Span
            GridSpan tCellSpan = new GridSpan() { Val = table.Columns.Count };
            tCellProperties.Append(tCellSpan);

            // Append properties to the cell
            tCell.Append(tCellProperties);

            ParagraphProperties paragProperties = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "0" };
            paragProperties.Append(justification1);
            paragProperties.Append(spacingBetweenLines1);

            var titleParagraph = new Paragraph();
            titleParagraph.Append(paragProperties);

            var run = new Run(new Text(tableStyle.Title));
            ApplyFontProperties(tableStyle, RowIdentification.Title, run);

            titleParagraph.Append(run);

            // Specify the table cell content.
            tCell.Append(titleParagraph);

            // Append the table cell to the table row.
            tRow.Append(tCell);

            // Append the table row to the table.
            wordTable.Append(tRow);
        }

        private static void SetTableAlignment(HorizontalAlignmentType alignment, TableProperties tblProp) {

            TableJustification tblJustification = null;

            if (alignment == HorizontalAlignmentType.Center) {

                tblJustification = new TableJustification() { Val = TableRowAlignmentValues.Center };
            }
            else if (alignment == HorizontalAlignmentType.Right) {

                tblJustification = new TableJustification() { Val = TableRowAlignmentValues.Right };
            }
            else {

                tblJustification = new TableJustification() { Val = TableRowAlignmentValues.Left };
            }

            tblProp.Append(tblJustification);
        }
    }
}
