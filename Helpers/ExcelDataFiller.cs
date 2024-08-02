using OfficeOpenXml;

namespace HastaKayitProjesi.Helpers
{
    public class ExcelDataFiller : IExcelDataFiller
    {
        public void FillExcelWithRandomData(string filePath, int dataCount, int existingRecordCount)
        {
            Random random = new Random();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                for (int i = 2; i <= dataCount + 1; i++)
                {
                    worksheet.Cells[i, 1].Value = existingRecordCount + i - 1;
                    worksheet.Cells[i, 2].Value = random.Next(100000000, 999999999).ToString();
                    worksheet.Cells[i, 3].Value = $"Ad{i - 1}";
                    worksheet.Cells[i, 4].Value = $"Soyad{i - 1}";
                    worksheet.Cells[i, 5].Value = DateTime.Now.AddYears(-random.Next(18, 90)).ToString("yyyy-MM-dd");
                    worksheet.Cells[i, 6].Value = $"90{random.Next(500000000, 599999999)}";
                    worksheet.Cells[i, 7].Value = random.Next(0, 2) == 0 ? "Erkek" : "Kadın";
                    worksheet.Cells[i, 8].Value = $"Adres{i - 1}";
                    worksheet.Cells[i, 9].Value = $"İlçe{i - 1}";
                    worksheet.Cells[i, 10].Value = $"İl{i - 1}";
                    worksheet.Cells[i, 11].Value = "Türkiye";
                    worksheet.Cells[i, 12].Value = $"Anne{i - 1}";
                    worksheet.Cells[i, 13].Value = $"Baba{i - 1}";
                    worksheet.Cells[i, 14].Value = $"email{i - 1}@gmail.com";
                    worksheet.Cells[i, 15].Value = "A+";
                    worksheet.Cells[i, 16].Value = $"Meslek{i - 1}";
                    worksheet.Cells[i, 17].Value = random.Next(100000000, 999999999).ToString();
                }

                package.Save();
            }
        }
    }
}
