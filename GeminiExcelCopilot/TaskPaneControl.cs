using GenerativeAI;
using System;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace GeminiExcelCopilot
{
    public partial class TaskPaneControl : UserControl
    {
        private GeminiService geminiService;

        public TaskPaneControl()
        {
            InitializeComponent();

            if (cmbActionType.Items.Count == 0)
            {
                cmbActionType.Items.AddRange(new object[] {
                    "Soru Sor (Genel)",
                    "Formül Üret",
                    "Makro (VBA) Üret",
                    "Seçili Alanı Analiz Et",
                    "Otomatik Grafik Oluştur",
                    "Halka Grafik Oluştur"
                });
            }
            if (cmbActionType.SelectedIndex == -1)
            {
                cmbActionType.SelectedIndex = 0;
            }

            // API Anahtarını Yükle
            txtApiKey.Text = Properties.Settings.Default.GeminiApiKey;

            InitializeGemini();
        }

        private void InitializeGemini()
        {
            try
            {
                geminiService = new GeminiService();
                txtResult.ForeColor = System.Drawing.Color.Green;
                txtResult.Text = "Veri Asistanı hazır. Lütfen bir eylem seçin.";
                button1.Enabled = true;
                btnApply.Enabled = true;
                cmbActionType.Enabled = true;
                textBox1.Enabled = true;
            }
            catch (Exception ex)
            {
                txtResult.ForeColor = System.Drawing.Color.Red;
                txtResult.Text = ex.Message;
                button1.Enabled = false;
                btnApply.Enabled = false;
                cmbActionType.Enabled = false;
                textBox1.Enabled = false;
            }
        }

        private void btnSaveApiKey_Click(object sender, EventArgs e)
        {
            string newKey = txtApiKey.Text.Trim();
            Properties.Settings.Default.GeminiApiKey = newKey;
            Properties.Settings.Default.Save();
            InitializeGemini();

            if (geminiService != null)
            {
                MessageBox.Show("API Anahtarı başarıyla kaydedildi!", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Anahtar kaydedildi ancak servis başlatılamadı.", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button1.Text = "İnceleniyor..."; // "Düşünüyor" yerine daha akademik bir ifade
            txtResult.Text = "";

            string userPrompt = textBox1.Text;
            string action = cmbActionType.SelectedItem.ToString();
            string language = Globals.ThisAddIn.ExcelLanguageName;
            string finalPrompt = "";
            string csvData = "";

            switch (action)
            {
                case "Formül Üret":
                    finalPrompt = $"Talimat: Kullanıcının '{language}' dilindeki şu isteğini anla: \"{userPrompt}\". " +
                                  $"Bu isteği yerine getiren **İngilizce (US English) / Invariant** Excel formülünü yaz. " +
                                  $"Sadece formülü döndür. Örn: =SUM(A1:A10).";
                    break;

                case "Makro (VBA) Üret":
                    finalPrompt = $"Talimat: Kullanıcının şu isteğini yerine getiren tam bir Excel VBA Sub...End Sub kodu yaz. Sadece kodu ver. Kullanıcı İsteği: \"{userPrompt}\"";
                    break;

                case "Seçili Alanı Analiz Et":
                    csvData = GetSelectedRangeAsCsv();
                    if (string.IsNullOrEmpty(csvData))
                    {
                        txtResult.Text = "Lütfen analiz edilecek veri alanını seçiniz.";
                        button1.Enabled = true;
                        button1.Text = "Gönder";
                        return;
                    }

                    // AKADEMİK VE TDK UYUMLU YENİ PROMPT
                    finalPrompt = $"Sen akademik düzeyde çalışan kıdemli bir veri bilimcisisin. Sana sunulan CSV veri setini ve soruyu analiz et.\n\n" +
                                  $"TALİMATLAR:\n" +
                                  $"1. DİL VE ÜSLUP: Yanıtını kusursuz bir Türkçe ile, Türk Dil Kurumu (TDK) yazım ve noktalama kurallarına tam uyarak yaz. Sokak ağzı, günlük konuşma dili veya 'ben' dili kullanma. Akademik, nesnel ve resmi bir üslup benimse (Edilgen çatı kullan: örn: 'baktım' yerine 'incelendiğinde', 'görüyoruz' yerine 'gözlemlenmektedir').\n" +
                                  $"2. FORMAT: Markdown formatı (kalın **, italik *, # başlık, kod bloğu ```) KESİNLİKLE KULLANMA. Sadece düz metin (plain text) kullan. Başlıkları belirtmek için BÜYÜK HARF kullan. Listeleme için sadece tire (-) işareti kullan.\n" +
                                  $"3. İÇERİK: Verileri yorumlarken neden-sonuç ilişkisi kur ve sayısal verilerle destekle. Yabancı terimlerden kaçın, Türkçe karşılıklarını kullan (Örn: 'Trend' yerine 'Eğilim', 'Data' yerine 'Veri').\n\n" +
                                  $"VERİ SETİ:\n---\n{csvData}\n---\n\nSORU/YÖNERGE: \"{userPrompt}\"";
                    break;

                case "Otomatik Grafik Oluştur":
                case "Halka Grafik Oluştur":
                    csvData = GetSelectedRangeAsCsv();
                    if (string.IsNullOrEmpty(csvData))
                    {
                        txtResult.Text = "Lütfen grafik için veri alanını seçiniz.";
                        button1.Enabled = true;
                        button1.Text = "Gönder";
                        return;
                    }
                    finalPrompt = $"Talimat: Veri seti: \n{csvData}\n Bu veri için en uygun Excel grafik türünü öner. Cevabın SADECE Excel'in C# Interop sabitinin adı olmalı (örn: xlColumnClustered, xlLine, xlPie, xlBarClustered, xlXYScatter, xlDoughnut).";
                    break;

                default:
                    finalPrompt = userPrompt;
                    break;
            }

            string response = "";
            bool isError = false;

            try
            {
                response = await geminiService.GenerateContentAsync(finalPrompt);
            }
            catch (Exception ex)
            {
                isError = true;
                response = $"Hata oluştu: {ex.Message}";
            }

            this.Invoke(new Action(() =>
            {
                if (isError)
                {
                    txtResult.ForeColor = System.Drawing.Color.Red;
                    txtResult.Text = response;
                }
                else
                {
                    if (action == "Otomatik Grafik Oluştur" || action == "Halka Grafik Oluştur")
                    {
                        string chartTypeString = response.Trim();
                        bool success = CreateChartFromGemini(chartTypeString, userPrompt);
                        if (success)
                            txtResult.Text = $"'{userPrompt}' başlıklı grafik başarıyla oluşturulmuştur.";
                        else
                        {
                            txtResult.ForeColor = System.Drawing.Color.Red;
                            txtResult.Text = $"Grafik oluşturulamadı. Beklenmeyen yanıt: {chartTypeString}";
                        }
                    }
                    else
                    {
                        // AKADEMİK TEMİZLİK: Her ihtimale karşı Markdown karakterlerini kodla da temizle
                        string cleanResponse = response.Replace("**", "").Replace("##", "").Replace("```", "").Replace("`", "");

                        txtResult.ForeColor = System.Drawing.Color.Black;
                        txtResult.Text = cleanResponse;
                    }
                }

                button1.Enabled = true;
                button1.Text = "Gönder";
            }));
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            string formula = txtResult.Text;
            if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith("="))
            {
                MessageBox.Show("Uygulanacak geçerli bir formül bulunamamıştır.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            try
            {
                Excel.Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
                if (activeCell == null)
                {
                    MessageBox.Show("Lütfen önce bir hücre seçiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                activeCell.Formula = formula;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Formül uygulanırken hata oluştu: {ex.Message}", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetSelectedRangeAsCsv()
        {
            Excel.Range selection;
            try
            {
                selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (selection == null || selection.Cells.Count <= 1) return null;
            }
            catch (Exception) { return null; }

            StringBuilder csvBuilder = new StringBuilder();
            int rowCount = selection.Rows.Count;
            int colCount = selection.Columns.Count;
            if (rowCount > 1000) rowCount = 1000;
            if (colCount > 50) colCount = 50;

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    Excel.Range cell = selection.Cells[i, j] as Excel.Range;
                    string cellValue = cell.Value2?.ToString() ?? "";
                    cellValue = cellValue.Replace("\"", "").Replace(",", ";").Replace("\n", " ");
                    csvBuilder.Append(cellValue);
                    if (j < colCount) csvBuilder.Append(",");
                }
                csvBuilder.AppendLine();
            }
            return csvBuilder.ToString();
        }

        private bool CreateChartFromGemini(string chartTypeString, string title)
        {
            Excel.XlChartType chartType;
            switch (chartTypeString)
            {
                case "xlColumnClustered": chartType = Excel.XlChartType.xlColumnClustered; break;
                case "xlLine": chartType = Excel.XlChartType.xlLine; break;
                case "xlPie": chartType = Excel.XlChartType.xlPie; break;
                case "xlBarClustered": chartType = Excel.XlChartType.xlBarClustered; break;
                case "xlArea": chartType = Excel.XlChartType.xlArea; break;
                case "xlXYScatter": chartType = Excel.XlChartType.xlXYScatter; break;
                case "xlDoughnut": chartType = Excel.XlChartType.xlDoughnut; break;
                default: return false;
            }

            try
            {
                Excel.Range selection = Globals.ThisAddIn.Application.Selection as Excel.Range;
                if (selection == null) return false;
                Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;

                Excel.ChartObject chartObj = sheet.ChartObjects().Add(
                    Left: selection.Left + selection.Width + 20,
                    Top: selection.Top,
                    Width: 450,
                    Height: 300);

                chartObj.Chart.SetSourceData(selection);
                chartObj.Chart.ChartType = chartType;
                chartObj.Chart.HasTitle = true;
                chartObj.Chart.ChartTitle.Text = title;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Grafik oluşturma hatası: {ex.Message}");
                return false;
            }
        }
    }
}