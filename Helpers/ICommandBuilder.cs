using Npgsql;

namespace HastaKayıtProjesi.Helpers
{
    public interface ICommandBuilder
    {
        NpgsqlCommand CreateInsertCommand(NpgsqlConnection conn, OfficeOpenXml.ExcelWorksheet worksheet, int rowIndex, int hastaId);
    }
}
