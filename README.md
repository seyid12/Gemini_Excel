# Gemini_Excel
# Veri Asistanı (Gemini Powered Excel Copilot)

**Veri Asistanı**, Microsoft Excel içerisine Google Gemini yapay zeka modelini entegre eden, C# ve VSTO (Visual Studio Tools for Office) ile geliştirilmiş gelişmiş bir eklentidir.

Bu proje, Excel kullanıcılarının doğal dil kullanarak karmaşık formüller oluşturmasını, veri analizi yapmasını, otomatik grafikler çizmesini ve VBA makroları yazmasını sağlar. Özellikle **Türkçe dil desteği**, **akademik analiz dili** ve **TDK uyumluluğu** ile öne çıkar.

![Proje Durumu](https://img.shields.io/badge/Durum-Aktif-success)
![Lisans](https://img.shields.io/badge/Lisans-MIT-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Excel-lightgrey)

## 🚀 Özellikler

### 1. Akademik Düzeyde Veri Analizi
* Seçili veri setini analiz eder ve sonuçları **TDK kurallarına uygun, akademik ve resmi bir dille** raporlar.
* Markdown formatı yerine temiz, düz metin (plain text) çıktısı verir.

### 2. Akıllı Formül Üretimi ve Uygulama
* Kullanıcının doğal dildeki isteğini (Örn: *"A sütununu B ile topla"*) anlar.
* Gemini'den **İngilizce (Invariant)** formül alır ve bunu Excel'in kurulu olduğu dile (Örn: Türkçe `=TOPLA`) otomatik olarak çevirerek hücreye uygular.
* `@` işareti hatasını ve `#AD?` hatalarını engelleyen güvenli uygulama yöntemi kullanır.

### 3. Otomatik Grafik Oluşturma
* Veri setinin içeriğine göre en uygun grafik türünü (Sütun, Çizgi, Pasta, Dağılım, Halka vb.) önerir.
* Excel sayfasında grafiği otomatik olarak çizer ve başlığını ayarlar.

### 4. VBA Makro Desteği
* Tekrarlayan işler için doğal dil komutlarını çalıştırılabilir VBA kodlarına dönüştürür.

### 5. Güvenli API Anahtarı Yönetimi
* API anahtarı kod içinde saklanmaz.
* Kullanıcı dostu arayüz üzerinden girilen anahtar, kullanıcının yerel ayarlarında (`User Settings`) şifreli olmasa da güvenli bir şekilde saklanır.

---

## 🛠️ Kurulum ve Geliştirme

Bu projeyi kendi bilgisayarınızda çalıştırmak veya geliştirmek için aşağıdaki adımları izleyin.

### Gereksinimler
* **İşletim Sistemi:** Windows 10 veya 11
* **Yazılım:** Microsoft Excel (2016, 2019 veya Office 365)
* **IDE:** Visual Studio 2022 (Community, Professional veya Enterprise)
* **Workload:** Visual Studio Installer'da *"Office/SharePoint development"* seçili olmalıdır.
* **API:** Google AI Studio'dan alınmış bir [Gemini API Anahtarı](https://aistudio.google.com/).

### Adım Adım Kurulum

1.  **Repoyu Klonlayın:**
    ```bash
    git clone [https://github.com/KULLANICI_ADINIZ/Veri-Asistani.git](https://github.com/KULLANICI_ADINIZ/Veri-Asistani.git)
    ```

2.  **Projeyi Açın:**
    `GeminiExcelCopilot.sln` dosyasını Visual Studio ile açın.

3.  **Paketleri Yükleyin:**
    Solution Explorer'da projeye sağ tıklayın ve **"Manage NuGet Packages"** seçeneğine gidin. Şu paketin yüklü olduğundan emin olun (yüklü değilse "Restore" yapın):
    * `Google.Ai.GenerativeLanguage`

4.  **Derleyin ve Çalıştırın:**
    `F5` tuşuna basarak projeyi başlatın. Excel otomatik olarak açılacak ve sağ tarafta **"Veri Asistanı"** bölmesi görünecektir.

---

## 📖 Kullanım Kılavuzu

### 1. Başlangıç
Eklenti ilk açıldığında API anahtarı soracaktır.
* Google AI Studio'dan aldığınız anahtarı `API Anahtarı` kutusuna yapıştırın.
* **"Kaydet"** butonuna basın. Bağlantı başarılıysa arayüz aktif olacaktır.

### 2. Formül Üretme
* İşlem menüsünden **"Formül Üret"**i seçin.
* Kutuya isteğinizi yazın: *"C2 ile C10 arasındaki en büyük değeri bul."*
* **"Gönder"**e basın. Sonuç kutusunda formül görünecektir.
* Excel'de bir hücre seçip **"Hücreye Uygula"** butonuna basarak formülü aktarın.

### 3. Veri Analizi
* Excel'de analiz etmek istediğiniz tabloyu seçin.
* Menüden **"Seçili Alanı Analiz Et"**i seçin.
* Sorunuzu sorun: *"Bu satış verilerindeki genel eğilim nedir?"*
* Asistan, akademik bir dille veriyi yorumlayacaktır.

### 4. Grafik Çizme
* Veri tablosunu seçin.
* Menüden **"Otomatik Grafik Oluştur"**u seçin.
* Kutuya grafik başlığını yazın ve gönderin.

---

## 🏗️ Teknoloji Yığını

* **Dil:** C# (.NET Framework 4.8)
* **Platform:** VSTO (Visual Studio Tools for Office) Excel Add-in
* **Yapay Zeka:** Google Gemini 2.5 Flash (`Google.Ai.GenerativeLanguage`)
* **Arayüz:** Windows Forms (WinForms)

---

## 🤝 Katkıda Bulunma

Katkılarınızı bekliyoruz! Lütfen önce bir "Issue" açarak yapmak istediğiniz değişikliği tartışın.

1.  Bu repoyu Fork'layın.
2.  Kendi branch'inizi oluşturun (`git checkout -b feature/YeniOzellik`).
3.  Değişikliklerinizi commit yapın (`git commit -m 'Yeni özellik eklendi'`).
4.  Branch'inizi Push yapın (`git push origin feature/YeniOzellik`).
5.  Bir Pull Request oluşturun.

## 📄 Lisans

Bu proje [MIT Lisansı](LICENSE) altında lisanslanmıştır.

---
**Geliştirici Notu:** Bu proje, Excel'in yerel dil ayarlarını (Localization) otomatik algılayarak formülleri dönüştüren özel bir yapıya sahiptir.
