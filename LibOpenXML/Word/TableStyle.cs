
using System.Drawing;

namespace LibOpenXml.Word {

    public class TableStyle {

        public enum TableWidthUnit {
            Auto,
            Percent,
        }

        public TableWidthUnit WidthUnit { get; set; }

        public int Width { get; set; }

        public HorizontalAlignmentType Alignment { get; set; }

        public int RowFontSize { get; set; }

        public int TitleFontSize { get; set; }

        public int HeaderFontSize { get; set; }

        public string Title { get; set; }

        public bool ShowTitle { get; set; }

        public bool ShowHeader { get; set; }

        public bool EnableAlternativeBackgroundColor { set; get; }

        public Color TitleBackgroundColor { get; set; }

        public Color HeaderBackgroundColor { get; set; }

        public Color AlternativeBackgroundColor { get; set; }

        public TableStyle() {

            ShowHeader = false;
            ShowTitle = false;
            EnableAlternativeBackgroundColor = false;

            TitleFontSize = HeaderFontSize = RowFontSize = 9;

            Alignment = HorizontalAlignmentType.Left;

            WidthUnit = TableWidthUnit.Auto;
            Width = 100;

            
            //TitleBackgroundColor = Color.FromArgb(128, 128, 128);
            TitleBackgroundColor = Color.White;
            //HeaderBackgroundColor = Color.FromArgb(192, 192, 192);
            HeaderBackgroundColor = Color.White;
            //AlternativeBackgroundColor = Color.FromArgb(225, 230, 235);
            AlternativeBackgroundColor = Color.White;
        }
    }
}
