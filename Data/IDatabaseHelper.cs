namespace HastaKayıtProjesi.Data
{
    public interface IDatabaseHelper
    {
        int GetExistingRecordCount();
        void UploadExcelToDatabase(string filePath);
        void ExportHastaKayitlariToExcel(string filePath);
        void ClearHastaKayitlari();
    }
}
