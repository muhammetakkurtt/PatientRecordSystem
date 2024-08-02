using HastaKayıtProjesi.Data;
using HastaKayitProjesi.Helpers;
using HastaKayitProjesi.Models;
using OfficeOpenXml;

namespace HastaKayitProjesi.Data
{
    public class DataExporter : IDataExporter
    {
        private readonly IExcelTemplateCreator _excelTemplateCreator;

        public DataExporter(IExcelTemplateCreator excelTemplateCreator)
        {
            _excelTemplateCreator = excelTemplateCreator;
        }

        public void ExportDataToExcel(string filePath, List<HastaKayit> data)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("Hasta Kayıtları");

            _excelTemplateCreator.CreateExcelTemplate(worksheet);

            for (var i = 0; i < data.Count; i++)
            {
                var kayit = data[i];
                worksheet.Cells[i + 2, 1].Value = kayit.HastaId;
                worksheet.Cells[i + 2, 2].Value = kayit.KimlikNo;
                worksheet.Cells[i + 2, 3].Value = kayit.Adi;
                worksheet.Cells[i + 2, 4].Value = kayit.Soyadi;
                worksheet.Cells[i + 2, 5].Value = kayit.DogumTarihi.ToString("yyyy-MM-dd");
                worksheet.Cells[i + 2, 6].Value = kayit.TelefonNumarasi;
                worksheet.Cells[i + 2, 7].Value = kayit.Cinsiyet;
                worksheet.Cells[i + 2, 8].Value = kayit.Adres;
                worksheet.Cells[i + 2, 9].Value = kayit.Ilce;
                worksheet.Cells[i + 2, 10].Value = kayit.Il;
                worksheet.Cells[i + 2, 11].Value = kayit.Ulke;
                worksheet.Cells[i + 2, 12].Value = kayit.AnneAdi;
                worksheet.Cells[i + 2, 13].Value = kayit.BabaAdi;
                worksheet.Cells[i + 2, 14].Value = kayit.EPosta;
                worksheet.Cells[i + 2, 15].Value = kayit.KanGrubu;
                worksheet.Cells[i + 2, 16].Value = kayit.Meslek;
                worksheet.Cells[i + 2, 17].Value = kayit.PasaportNumarasi;
            }

            package.SaveAs(new FileInfo(filePath));
        }
    }
}
