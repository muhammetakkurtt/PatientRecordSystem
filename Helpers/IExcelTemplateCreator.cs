namespace HastaKayitProjesi.Helpers
{
    public interface IExcelTemplateCreator
    {
        void CreateExcelTemplate(string filePath);
        void CreateExcelTemplate(OfficeOpenXml.ExcelWorksheet worksheet);
    }
}
