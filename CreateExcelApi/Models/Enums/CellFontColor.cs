using NPOI.HSSF.Util;

namespace CreateExcelApi.Models.Enums
{
    public enum CellFontColor
    {
        Black = HSSFColor.Black.Index,
        White = HSSFColor.White.Index,
        Red = HSSFColor.Red.Index,
        Blue = HSSFColor.Blue.Index,
        Green = HSSFColor.Green.Index
    }
}
