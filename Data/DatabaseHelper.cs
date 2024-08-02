using HastaKayıtProjesi.Configuration;
using HastaKayıtProjesi.Data;
using HastaKayitProjesi.Helpers;
using HastaKayıtProjesi.Helpers;
using Npgsql;
using OfficeOpenXml;

namespace HastaKayitProjesi.Data
{
    public class DatabaseHelper : IDatabaseHelper
    {
        private readonly IConfigurationService _configurationService;
        private readonly IExcelTemplateCreator _excelTemplateCreator;
        private readonly IExcelDataFiller _excelDataFiller;
        private readonly ICommandBuilder _commandBuilder;
        private readonly IDatabaseReader _databaseReader;
        private readonly IDataExporter _dataExporter;
        private readonly string _connString;

        public DatabaseHelper(IConfigurationService configurationService, IExcelTemplateCreator excelTemplateCreator, IExcelDataFiller excelDataFiller, ICommandBuilder commandBuilder, IDatabaseReader databaseReader, IDataExporter dataExporter)
        {
            _configurationService = configurationService;
            _excelTemplateCreator = excelTemplateCreator;
            _excelDataFiller = excelDataFiller;
            _commandBuilder = commandBuilder;
            _databaseReader = databaseReader;
            _dataExporter = dataExporter;
            _connString = _configurationService.GetConnectionString();
        }

        public int GetExistingRecordCount()
        {
            using var conn = new NpgsqlConnection(_connString);
            conn.Open();
            using var cmd = new NpgsqlCommand("SELECT COUNT(*) FROM hasta_kayit", conn);
            return Convert.ToInt32(cmd.ExecuteScalar());
        }

        public void UploadExcelToDatabase(string filePath)
        {
            var existingRecordCount = GetExistingRecordCount();
            using var conn = new NpgsqlConnection(_connString);
            conn.Open();
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];
            for (var i = 2; i <= worksheet.Dimension.End.Row; i++)
            {
                var hastaId = existingRecordCount + i - 1;
                var cmd = _commandBuilder.CreateInsertCommand(conn, worksheet, i, hastaId);
                cmd.ExecuteNonQuery();
            }
        }

        public void ExportHastaKayitlariToExcel(string filePath)
        {
            var hastaKayitlari = _databaseReader.GetHastaKayitlari();
            _dataExporter.ExportDataToExcel(filePath, hastaKayitlari);
        }

        public void ClearHastaKayitlari()
        {
            using var conn = new NpgsqlConnection(_connString);
            conn.Open();
            using var cmd = new NpgsqlCommand("DELETE FROM hasta_kayit", conn);
            cmd.ExecuteNonQuery();
        }
    }
}
