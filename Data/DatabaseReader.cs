using HastaKayitProjesi.Models;
using Npgsql;

public class DatabaseReader : IDatabaseReader
{
    private readonly string _connString;

    public DatabaseReader(string connString)
    {
        _connString = connString;
    }

    public List<HastaKayit> GetHastaKayitlari()
    {
        var hastaKayitlari = new List<HastaKayit>();

        using var conn = new NpgsqlConnection(_connString);
        conn.Open();
        using var cmd = new NpgsqlCommand("SELECT * FROM hasta_kayit", conn);
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var kayit = new HastaKayit
            {
                HastaId = reader.GetInt32(0),
                KimlikNo = reader.GetString(1),
                Adi = reader.GetString(2),
                Soyadi = reader.GetString(3),
                DogumTarihi = reader.GetDateTime(4),
                TelefonNumarasi = reader.GetString(5),
                Cinsiyet = reader.GetString(6),
                Adres = reader.GetString(7),
                Ilce = reader.GetString(8),
                Il = reader.GetString(9),
                Ulke = reader.GetString(10),
                AnneAdi = reader.GetString(11),
                BabaAdi = reader.GetString(12),
                EPosta = reader.GetString(13),
                KanGrubu = reader.GetString(14),
                Meslek = reader.GetString(15),
                PasaportNumarasi = reader.IsDBNull(16) ? null : reader.GetString(16)
            };
            hastaKayitlari.Add(kayit);
        }

        return hastaKayitlari;
    }
}
