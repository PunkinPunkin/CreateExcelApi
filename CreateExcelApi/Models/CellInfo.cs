using CreateExcelApi.Models.Enums;
using NPOI.SS.UserModel;

namespace CreateExcelApi.Models
{
    public class CellInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public short FontSize { get; set; } = 12;
        public CellFontColor FontColor { get; set; } = CellFontColor.Black;
        public HorizontalAlignment HorizontalAlign { get; set; } = HorizontalAlignment.General;
        public VerticalAlignment VerticalAlign { get; set; } = VerticalAlignment.Center;
    }
}
