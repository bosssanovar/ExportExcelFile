using ClosedXML.Excel;

namespace ExportExcelFile
{
    public class NameCard(NameCardParams parameters, string path)
    {
        private const string TemplateFileName = @"NameCardTemplate.xlsx";

        private readonly NameCardParams _params = parameters;
        private readonly string path = path;

        public ExportErrorType ExportXlsx()
        {
            var workbook = new XLWorkbook(TemplateFileName);

            var worksheet = workbook.Worksheets.First();

            worksheet.Cell("B2").Value = "Hello world!";

            workbook.SaveAs(path);

            return ExportErrorType.None;
        }
    }
}
