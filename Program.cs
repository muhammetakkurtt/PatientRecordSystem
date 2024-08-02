using HastaKayitProjesi.Configuration;
using HastaKayitProjesi.Data;
using HastaKayitProjesi.Forms;
using HastaKayitProjesi.Helpers;

namespace HastaKayitProjesi
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var configurationService = new ConfigurationService();
            var excelTemplateCreator = new ExcelTemplateCreator();
            var excelDataFiller = new ExcelDataFiller();
            var commandBuilder = new CommandBuilder();
            var databaseReader = new DatabaseReader(configurationService.GetConnectionString());
            var dataExporter = new DataExporter(excelTemplateCreator);
            var databaseHelper = new DatabaseHelper(configurationService, excelTemplateCreator, excelDataFiller, commandBuilder, databaseReader, dataExporter);
            var mainForm = new MainForm(excelTemplateCreator, excelDataFiller, databaseHelper);

            Application.Run(mainForm);
        }
    }
}
