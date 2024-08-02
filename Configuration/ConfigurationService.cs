using HastaKayıtProjesi.Configuration;
using Microsoft.Extensions.Configuration;

namespace HastaKayitProjesi.Configuration
{
    public class ConfigurationService : IConfigurationService
    {
        private readonly IConfiguration _configuration;

        public ConfigurationService()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            _configuration = builder.Build();
        }

        public string GetConnectionString()
        {
            var connString = _configuration.GetConnectionString("DefaultConnection");
            if (string.IsNullOrEmpty(connString))
            {
                throw new ArgumentException("Veritabanı bağlantı bilgileri eksik. Lütfen appsettings.json dosyasını kontrol edin.");
            }
            return connString;
        }
    }
}