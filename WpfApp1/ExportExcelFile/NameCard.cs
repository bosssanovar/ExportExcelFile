using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExportExcelFile
{
    public class NameCard(NameCardParams parameters, string path)
    {
        private const string TemplateFileName = @"NameCardTemplate.xlsx";
        private readonly string[] NameCells = ["B2", "F2", "J2", "N2", "R2"];
        private readonly string[] ButtonLabelLeftCells = ["B9", "F9", "J9", "N9", "R9"];
        private readonly string[] ButtonLabelCenterCells = ["C9", "G9", "K9", "O9", "S9"];
        private readonly string[] ButtonLabelRightCells = ["D9", "H9", "L9", "P9", "T9"];

        private readonly NameCardParams _params = parameters;
        private readonly string path = path;

        public void ExportXlsx()
        {
            var workbook = new XLWorkbook(TemplateFileName);
            var worksheet = workbook.Worksheets.First();

            SetName(worksheet, 0, _params.Name);
            SetNameColor(worksheet, 0, _params.Color);
            SetButtonLabel(worksheet, 0, _params.NameCardType);

            workbook.SaveAs(path);
        }

        private void SetName(IXLWorksheet worksheet, int index, string name)
        {
            worksheet.Cell(NameCells[index]).Value = name;
        }

        private void SetNameColor(IXLWorksheet worksheet, int index, System.Drawing.Color color)
        {
            worksheet.Cell(NameCells[index]).Style.Fill.SetBackgroundColor(XLColor.FromColor(color));
        }

        private void SetButtonLabel(IXLWorksheet worksheet, int index, NameCardType type)
        {
            worksheet.Cell(ButtonLabelLeftCells[index]).Value = type.GetButtonLabelLeft();
            worksheet.Cell(ButtonLabelCenterCells[index]).Value = type.GetButtonLabelCenter();
            worksheet.Cell(ButtonLabelRightCells[index]).Value = type.GetButtonLabelRight();
        }
    }
}
