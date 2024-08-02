using HastaKayıtProjesi.Helpers;
using Npgsql;
using OfficeOpenXml;

public class CommandBuilder : ICommandBuilder
{
    public NpgsqlCommand CreateInsertCommand(NpgsqlConnection conn, ExcelWorksheet worksheet, int rowIndex, int hastaId)
    {
        var cmd = new NpgsqlCommand
        {
            Connection = conn,
            CommandText = @"INSERT INTO hasta_kayit 
            (hasta_id, kimlik_no, adi, soyadi, dogum_tarihi, telefon_numarasi, cinsiyet, adres, ilce, il, ulke, anne_adi, baba_adi, eposta, kan_grubu, meslek, pasaport_numarasi) 
            VALUES (@hasta_id, @kimlik_no, @adi, @soyadi, @dogum_tarihi, @telefon_numarasi, @cinsiyet, @adres, @ilce, @il, @ulke, @anne_adi, @baba_adi, @eposta, @kan_grubu, @meslek, @pasaport_numarasi)"
        };

        cmd.Parameters.AddWithValue("hasta_id", hastaId);
        cmd.Parameters.AddWithValue("kimlik_no", worksheet.Cells[rowIndex, 2].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("adi", worksheet.Cells[rowIndex, 3].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("soyadi", worksheet.Cells[rowIndex, 4].Value?.ToString() ?? (object)DBNull.Value);

        DateTime dogumTarihi;
        if (worksheet.Cells[rowIndex, 5].Value != null && DateTime.TryParse(worksheet.Cells[rowIndex, 5].Value.ToString(), out dogumTarihi))
        {
            cmd.Parameters.AddWithValue("dogum_tarihi", dogumTarihi);
        }
        else
        {
            cmd.Parameters.AddWithValue("dogum_tarihi", DBNull.Value);
        }

        cmd.Parameters.AddWithValue("telefon_numarasi", worksheet.Cells[rowIndex, 6].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("cinsiyet", worksheet.Cells[rowIndex, 7].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("adres", worksheet.Cells[rowIndex, 8].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("ilce", worksheet.Cells[rowIndex, 9].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("il", worksheet.Cells[rowIndex, 10].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("ulke", worksheet.Cells[rowIndex, 11].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("anne_adi", worksheet.Cells[rowIndex, 12].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("baba_adi", worksheet.Cells[rowIndex, 13].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("eposta", worksheet.Cells[rowIndex, 14].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("kan_grubu", worksheet.Cells[rowIndex, 15].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("meslek", worksheet.Cells[rowIndex, 16].Value?.ToString() ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("pasaport_numarasi", worksheet.Cells[rowIndex, 17].Value?.ToString() ?? (object)DBNull.Value);

        return cmd;
    }
}
