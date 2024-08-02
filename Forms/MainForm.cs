using HastaKayıtProjesi.Data;
using HastaKayitProjesi.Helpers;

namespace HastaKayitProjesi.Forms
{
    public partial class MainForm : Form
    {
        private readonly IExcelTemplateCreator _excelTemplateCreator;
        private readonly IExcelDataFiller _excelDataFiller;
        private readonly IDatabaseHelper _databaseHelper;
        private readonly string _directoryPath;
        private TextBox? _dataCountTextBox;

        public MainForm(IExcelTemplateCreator excelTemplateCreator, IExcelDataFiller excelDataFiller, IDatabaseHelper databaseHelper)
        {
            InitializeComponent();
            InitializeCustomComponents();
            _excelTemplateCreator = excelTemplateCreator;
            _excelDataFiller = excelDataFiller;
            _databaseHelper = databaseHelper;
            _directoryPath = AppDomain.CurrentDomain.BaseDirectory;
        }

        private void InitializeCustomComponents()
        {
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10)
            };
            Controls.Add(mainPanel);

            var tableLayoutPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6,
                AutoSize = true,
                Padding = new Padding(10),
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30F));
            tableLayoutPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70F));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            tableLayoutPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainPanel.Controls.Add(tableLayoutPanel);

            var createTemplateButton = new Button
            {
                Text = "Boş Excel Şablonu Oluştur",
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                AutoSize = true
            };
            createTemplateButton.Click += CreateTemplateButton_Click;
            tableLayoutPanel.Controls.Add(createTemplateButton, 0, 0);
            tableLayoutPanel.SetColumnSpan(createTemplateButton, 2);

            var fillDataButton = new Button
            {
                Text = "Excel'i Rastgele Bilgilerle Doldur",
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                AutoSize = true
            };
            fillDataButton.Click += FillDataButton_Click;
            tableLayoutPanel.Controls.Add(fillDataButton, 0, 1);
            tableLayoutPanel.SetColumnSpan(fillDataButton, 2);

            var dataCountLabel = new Label
            {
                Text = "Hasta Sayısı:",
                Anchor = AnchorStyles.Left,
                AutoSize = true
            };
            tableLayoutPanel.Controls.Add(dataCountLabel, 0, 2);

            _dataCountTextBox = new TextBox
            {
                Anchor = AnchorStyles.Left,
                Width = 50
            };
            tableLayoutPanel.Controls.Add(_dataCountTextBox, 1, 2);

            var uploadButton = new Button
            {
                Text = "Veritabanına Yükle",
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                AutoSize = true
            };
            uploadButton.Click += UploadButton_Click;
            tableLayoutPanel.Controls.Add(uploadButton, 0, 3);
            tableLayoutPanel.SetColumnSpan(uploadButton, 2);

            var exportButton = new Button
            {
                Text = "Hasta Verilerini İndir",
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                AutoSize = true
            };
            exportButton.Click += ExportButton_Click;
            tableLayoutPanel.Controls.Add(exportButton, 0, 4);
            tableLayoutPanel.SetColumnSpan(exportButton, 2);

            var clearButton = new Button
            {
                Text = "Veritabanını Temizle",
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                AutoSize = true
            };
            clearButton.Click += ClearButton_Click;
            tableLayoutPanel.Controls.Add(clearButton, 0, 5);
            tableLayoutPanel.SetColumnSpan(clearButton, 2);
        }

        private void CreateTemplateButton_Click(object? sender, EventArgs e)
        {
            var filePath = Path.Combine(_directoryPath, "HastaKayitTemplate.xlsx");
            _excelTemplateCreator.CreateExcelTemplate(filePath);
            MessageBox.Show($"Boş Excel şablonu oluşturuldu: {filePath}");
        }

        private void FillDataButton_Click(object? sender, EventArgs e)
        {
            var filePath = Path.Combine(_directoryPath, "HastaKayitTemplate.xlsx");
            if (int.TryParse(_dataCountTextBox?.Text, out var dataCount) && dataCount > 0)
            {
                var existingRecordCount = _databaseHelper.GetExistingRecordCount();
                _excelDataFiller.FillExcelWithRandomData(filePath, dataCount, existingRecordCount);
                MessageBox.Show($"Excel {dataCount} rastgele veri ile dolduruldu: {filePath}");
            }
            else
            {
                MessageBox.Show("Lütfen geçerli bir veri sayısı girin.");
            }
        }

        private void UploadButton_Click(object? sender, EventArgs e)
        {
            var filePath = Path.Combine(_directoryPath, "HastaKayitTemplate.xlsx");
            _databaseHelper.UploadExcelToDatabase(filePath);
            MessageBox.Show("Veriler veritabanına yüklendi.");
        }

        private void ExportButton_Click(object? sender, EventArgs e)
        {
            var filePath = Path.Combine(_directoryPath, "ExportedHastaKayitlar.xlsx");
            _databaseHelper.ExportHastaKayitlariToExcel(filePath);
            MessageBox.Show($"Hasta verileri Excel dosyasına aktarıldı: {filePath}");
        }

        private void ClearButton_Click(object? sender, EventArgs e)
        {
            _databaseHelper.ClearHastaKayitlari();
            MessageBox.Show("Veritabanındaki tüm kayıtlar silindi.");
        }
    }
}
