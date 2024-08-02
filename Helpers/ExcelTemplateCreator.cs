using OfficeOpenXml;

namespace HastaKayitProjesi.Helpers
{
    public class ExcelTemplateCreator : IExcelTemplateCreator
    {
        public void CreateExcelTemplate(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Hasta Kayıtları");
                CreateExcelTemplate(worksheet);
                FileInfo fi = new FileInfo(filePath);
                package.SaveAs(fi);
            }
        }

        public void CreateExcelTemplate(ExcelWorksheet worksheet)
        {
            worksheet.Cells[1, 1].Value = "Hasta Id";
            worksheet.Cells[1, 2].Value = "Kimlik No";
            worksheet.Cells[1, 3].Value = "Adı";
            worksheet.Cells[1, 4].Value = "Soyadı";
            worksheet.Cells[1, 5].Value = "Doğum Tarihi";
            worksheet.Cells[1, 6].Value = "Telefon Numarası";
            worksheet.Cells[1, 7].Value = "Cinsiyet";
            worksheet.Cells[1, 8].Value = "Adres";
            worksheet.Cells[1, 9].Value = "İlçe";
            worksheet.Cells[1, 10].Value = "İl";
            worksheet.Cells[1, 11].Value = "Ülke";
            worksheet.Cells[1, 12].Value = "Anne Adı";
            worksheet.Cells[1, 13].Value = "Baba Adı";
            worksheet.Cells[1, 14].Value = "E-Posta";
            worksheet.Cells[1, 15].Value = "Kan Grubu";
            worksheet.Cells[1, 16].Value = "Meslek";
            worksheet.Cells[1, 17].Value = "Pasaport Numarası";
        }
    }
}
