namespace HastaKayitProjesi.Helpers
{
    public interface IExcelDataFiller
    {
        void FillExcelWithRandomData(string filePath, int dataCount, int existingRecordCount);
    }
}
