using HastaKayitProjesi.Models;

namespace HastaKayıtProjesi.Data
{
    public interface IDataExporter
    {
        void ExportDataToExcel(string filePath, List<HastaKayit> data);
    }
}
