# Spreadox Report Studio (MVC Sample)

[![Live Demo](https://img.shields.io/badge/Live%20Demo-Vercel-black?logo=vercel)](https://spreadox-report-mvc.vercel.app/)

> 🌐 **Canlı Demo (Live Web App):** [https://spreadox-report-mvc.vercel.app/](https://spreadox-report-mvc.vercel.app/)

![Spreadox Report Studio](https://raw.githubusercontent.com/FatihBaycu/Conq.Noble.Spreadox/refs/heads/main/docs/images/spreadox-report-studio.png)

Bu proje, **Spreadox** kütüphanesini kullanarak dinamik veri üretimi, interaktif grafik tipi yapılandırması, Excel ve PDF rapor oluşturma, üstbilgi/altbilgi (Header/Footer) özelleştirme, çoklu dil desteği (English / Türkçe) ve özel font desteği sunan bir ASP.NET Core MVC örnek uygulamasıdır.

## Özellikler

1. **Çoklu Dil Desteği (English & Turkish Localization)**
   - **Default Language:** English (`en-US`).
   - **Alternatif Dil:** Türkçe (`tr-TR`).
   - Navbar'daki dil seçici (🇬🇧 English / 🇹🇷 Türkçe) ile tek tıkla anında geçiş.
   - Sayfa başlıkları, butonlar, açıklamalar, uyarı mesajları ve JavaScript dinamik etiketleri seçilen dile göre otomatik uyarlanır.

2. **Dinamik Veri Üretimi (Random Data Generator)**
   - İstenilen sayıda veri başlığı üretme (örn: 10 başlık, 5 başlık, 15 başlık vb.).
   - Her başlık için otomatik kategoriler (Jan..Dec) ve gerçekçi değerler.
   - Üretilen verileri doğrudan arayüzden düzenleyebilme (Başlık adı, kategoriler, değerler, renk).

3. **Gelişmiş Grafik Seçenekleri (Chart Types)**
   - Her veri başlığı için bağımsız grafik tipi seçimi:
     - `ColumnClustered`, `ColumnStacked`
     - `BarClustered`, `BarStacked`
     - `Line`, `LineMarker`, `LineStacked`
     - `Area`, `AreaStacked`
     - `Pie`, `Doughnut`
     - `Scatter`, `Radar`, `Combo`
   - Her grafik serisi için özel HEX / ColorPicker renk seçimi.

4. **Sayfa Düzeni (Layout Options)**
   - **Ayrı Sayfalarda (Separate Pages / Sheets):** Her grafik kendi PDF sayfasında veya ayrı Excel çalışma sayfasında yer alır, otomatik özet istatistikler (Min, Max, Mean, Sum) eklenir.
   - **Aynı Sayfada (Single Page / Sheet):** Tüm grafikler tek bir sayfada sıralı olarak konumlandırılır.

5. **Header (Üstbilgi) Düzenleme Desteği**
   - Header Başlığı (Title)
   - Alt Açıklama Metni (Subtitle / Text)
   - Header Logosu / Resim Ekleme (Aspect-fit PNG, JPEG)
   - Yazı Tipi (Font Family) seçimi
   - Yazı Boyutu (Font Size) ve Kalın (Bold) seçeneği
   - Başlık Rengi seçimi (Color Picker)
   - İyileştirilmiş yükseklik ve ferah alt çizgi boşluğu.

6. **Footer (Altbilgi) Düzenleme Desteği**
   - İstenilen Özel Metin (Footer Text)
   - Sayfa Numarası Formatı (`{0}` geçerli sayfa, `{1}` toplam sayfa - örn: `Page {0} of {1}`)
   - Footer Logosu / Küçük Resim (PNG, JPEG)
   - Yazı Tipi (Font Family) seçimi

7. **Özel Font Yükleme Desteği (Custom TTF Font Registration)**
   - Kütüphaneye eklenen `PdfFontRegistry` ve `PdfFontResolver` sayesinde sisteme çalışma zamanında `.ttf` TrueType font dosyası yüklenebilir.
   - Yüklenen fontlar anında Header ve Footer yazı tipi listelerine eklenir ve PDF üretiminde kullanılır.

8. **Çift Çıktı Desteği (Excel & PDF)**
   - **PDF:** Spreadox'un yüksek performanslı vektör tabanlı A4 Landscape PDF motoru ile vektörel grafikler, kapak sayfası, header/footer ve özet istatistikler.
   - **Excel:** Yerel OOXML DrawingML grafikleri, stillendirilmiş tablolar, dondurulmuş bölmeler (freeze panes) ile tam uyumlu `.xlsx` dosyası oluşturulur.

## Çalıştırma

Projeyi çalıştırmak için:

```bash
cd samples/ReportBuilderMvc
dotnet run
```

Tarayıcınızda açın:
```
http://localhost:5000 veya konsolda belirtilen port
```
Sayfanın sağ üst köşesinden dilinizi (🇬🇧 English / 🇹🇷 Türkçe) dilediğiniz zaman değiştirebilirsiniz.
