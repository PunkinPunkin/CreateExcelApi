using System.Text.Json.Serialization;
using CreateExcelApi.Models.Enums;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace CreateExcelApi.Models
{
    public class ExcelInfo : IDisposable
    {
        private readonly ILogger<ExcelInfo> _logger;

        private short _fontSize;
        private short _titleFontSize;
        private short _subTitleFontSize;
        private string _fontName;
        private Dictionary<string, ICellStyle> _cellStyles;

        private string STYLE_DEFAULT;
        private string STYLE_DEFAULT_BOLD;
        private string STYLE_TITLE;
        private string STYLE_SUB_TITLE;
        private string STYLE_PRINT_INFO;

        private IWorkbook _workbook;
        private MemoryStream _excelStream;

        [JsonConstructor]
        public ExcelInfo()
        {
            Initial();
        }
        
        public ExcelInfo(ILogger<ExcelInfo> logger)
        {
            _logger = logger;
            Initial();
        }

        public DateTime PrintTime { get; set; } = DateTime.Now;
        public string PrintEmployee { get; set; } = "user";
        public ICollection<SheetInfo> Sheets { get; set; } = Array.Empty<SheetInfo>();

        public string PrintInfo => $"{PrintTime:yyyy-MM-dd HH:mm:ss}    {PrintEmployee}";

        private void Initial()
        {
            _fontSize = 12;
            _titleFontSize = 18;
            _subTitleFontSize = 16;
            _fontName = "微軟正黑體";
            _cellStyles = new Dictionary<string, ICellStyle>();

            STYLE_DEFAULT = $"{(int)HorizontalAlignment.Left}:{(int)VerticalAlignment.Center}:{(int)CellFontColor.Black}:{_fontSize}";
            STYLE_DEFAULT_BOLD = $"{(int)HorizontalAlignment.Left}:{(int)VerticalAlignment.Center}:{(int)CellFontColor.Black}:{_fontSize}:B";
            STYLE_TITLE = $"{(int)HorizontalAlignment.Center}:{(int)VerticalAlignment.Center}:{(int)CellFontColor.Black}:{_titleFontSize}:B";
            STYLE_SUB_TITLE = $"{(int)HorizontalAlignment.Center}:{(int)VerticalAlignment.Center}:{(int)CellFontColor.Black}:{_subTitleFontSize}";
            STYLE_PRINT_INFO = $"{(int)HorizontalAlignment.Right}:{(int)VerticalAlignment.Center}:{(int)CellFontColor.Black}:{_fontSize}";
        }

        private void InitialCellStyle()
        {
            _cellStyles.Clear();

            var style = _workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Left;
            style.VerticalAlignment = VerticalAlignment.Center;
            var font = _workbook.CreateFont();
            font.FontName = _fontName;
            font.FontHeightInPoints = _fontSize;
            font.IsBold = false;
            style.SetFont(font);
            _cellStyles.Add(STYLE_DEFAULT, style); //STYLE_DEFAULT
            GetCellStyles(_workbook, HorizontalAlignment.Left, VerticalAlignment.Center, CellFontColor.Black, _fontSize, true); //STYLE_DEFAULT_BOLD
            GetCellStyles(_workbook, HorizontalAlignment.Center, VerticalAlignment.Center, CellFontColor.Black, _titleFontSize, true); //STYLE_TITLE
            GetCellStyles(_workbook, HorizontalAlignment.Center, VerticalAlignment.Center, CellFontColor.Black, _subTitleFontSize); //STYLE_SUB_TITLE
            GetCellStyles(_workbook, HorizontalAlignment.Right, VerticalAlignment.Center, CellFontColor.Black, _fontSize); //STYLE_PRINT_INFO
        }

        private ICellStyle GetCellStyles(IWorkbook workbook, HorizontalAlignment horizontalAlign, VerticalAlignment verticalAlign, CellFontColor cellFontColor, short fontSize, bool isBold = false)
        {
            string key = $"{(int)horizontalAlign}:{(int)verticalAlign}:{(int)cellFontColor}:{fontSize}" + (isBold ? ":B" : string.Empty);
            if (_cellStyles.ContainsKey(key))
            {
                return _cellStyles[key];
            }
            try
            {
                ICellStyle style = workbook.CreateCellStyle();
                style.Alignment = horizontalAlign;
                style.VerticalAlignment = verticalAlign;
                IFont font = workbook.CreateFont();
                font.FontName = _fontName;
                font.IsBold = isBold;
                font.FontHeightInPoints = fontSize;
                font.Color = (short)cellFontColor;
                style.SetFont(font);
                _cellStyles.Add(key, style);
                return style;
            }
            catch (Exception e)
            {
                _logger.LogError("GetCellStyles, key: " + key, e);
                return _cellStyles[STYLE_DEFAULT];
            }
        }

        protected int CreateSheetTitle(ISheet sheet, SheetInfo sheetInfo)
        {
            if (sheet == null || sheetInfo == null)
            {
                return 0;
            }
            int columnCount = sheetInfo.TableHeaders == null ? 1 : sheetInfo.TableHeaders.Count;
            int rowNum = 0;
            int position = 0;
            bool loop = true;
            while (loop)
            {
                IRow row = sheet.GetRow(rowNum);
                if (row == null)
                {
                    row = sheet.CreateRow(rowNum);
                    sheet.AddMergedRegion(new CellRangeAddress(rowNum, rowNum, 0, columnCount - 1));
                    row.CreateCell(0);
                }
                ICell cell = row.GetCell(0);
                switch (position)
                {
                    case 0:
                        cell.CellStyle = _cellStyles[STYLE_TITLE];
                        cell.SetCellValue(sheetInfo.Title);
                        rowNum++;
                        break;
                    case 1:
                        if (string.IsNullOrWhiteSpace(sheetInfo.SubTitle))
                        {
                            break;
                        }
                        cell.CellStyle = _cellStyles[STYLE_SUB_TITLE];
                        cell.SetCellValue(sheetInfo.SubTitle);
                        rowNum++;
                        break;
                    case 2:
                        if (string.IsNullOrWhiteSpace(sheetInfo.SearchCondition))
                        {
                            break;
                        }
                        cell.CellStyle = _cellStyles[STYLE_PRINT_INFO];
                        cell.SetCellValue(sheetInfo.SearchCondition);
                        rowNum++;
                        break;
                    default:
                        cell.CellStyle = _cellStyles[STYLE_PRINT_INFO];
                        cell.SetCellValue(PrintInfo);
                        loop = false;
                        break;
                }
                position++;
            }
            return rowNum;
        }

        public MemoryStream GetStream()
        {
            _workbook = new XSSFWorkbook();
            if (Sheets == null || !Sheets.Any())
            {
                ISheet sheet = _workbook.CreateSheet("Sheet1");
                sheet.CreateRow(1);
                sheet.GetRow(1).CreateCell(0).SetCellValue(PrintInfo);

                using (_excelStream = new())
                {
                    _workbook.Write(_excelStream);
                }
                return _excelStream;
            }

            ICreationHelper factory = _workbook.GetCreationHelper();
            IClientAnchor anchor = factory.CreateClientAnchor();
            InitialCellStyle();
            int sheetNumber = 1;
            foreach (var sheetInfo in Sheets)
            {
                if (sheetInfo == null)
                {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(sheetInfo.Name))
                {
                    sheetInfo.Name = $"Sheet{sheetNumber++}";
                }
                ISheet sheet = _workbook.CreateSheet(sheetInfo.Name);
                IDrawing drawing = sheet.CreateDrawingPatriarch();
                int rowNum = CreateSheetTitle(sheet, sheetInfo) + 2;
                if (sheetInfo.TableHeaders != null)
                {
                    int columnCount = sheetInfo.TableHeaders.Count;
                    IRow header = sheet.CreateRow(rowNum);
                    for (int c = 0; c < columnCount; c++)
                    {
                        ICell cell = header.CreateCell(c);
                        var cellInfo = sheetInfo.TableHeaders.ElementAt(c);
                        cell.SetCellValue(cellInfo.Name);
                        if (!string.IsNullOrWhiteSpace(cellInfo.Comment))
                        {
                            IComment comment = drawing.CreateCellComment(anchor);
                            comment.Author = PrintEmployee;
                            comment.String = factory.CreateRichTextString(cellInfo.Comment);
                            comment.Visible = false;
                            cell.CellComment = comment;
                        }
                        cell.CellStyle = _cellStyles[STYLE_DEFAULT_BOLD];
                        if (sheetInfo.Data == null)
                        {
                            break;
                        }
                        for (int d = 0; d < sheetInfo.Data.Count; d++)
                        {
                            var values = sheetInfo.Data.ElementAt(d);
                            IRow dataRow = sheet.GetRow(rowNum + d + 1);
                            if (dataRow == null)
                            {
                                dataRow = sheet.CreateRow(rowNum + d + 1);
                            }
                            ICell dataCell = dataRow.CreateCell(c);
                            dataCell.SetCellValue(values.ElementAt(c));
                            dataCell.CellStyle = GetCellStyles(_workbook, cellInfo.HorizontalAlign, cellInfo.VerticalAlign, cellInfo.FontColor, cellInfo.FontSize);
                        }
                    }

                    for (int i = 0; i < columnCount; i++)
                    {
                        sheet.AutoSizeColumn(i);
                        GC.Collect();
                    }
                }
            }

            using (_excelStream = new())
            {
                _workbook.Write(_excelStream);
            }
            return _excelStream;
        }

        public void Dispose()
        {
            if (_excelStream != null)
                _excelStream.Dispose();
        }
    }
}
