# PDFconverter
The PDF converter is a fully client-side web application that processes files directly in the user's browser.

<b>ENG:</b>
<h2> Technical Overview of the PDF Converter </h2><br>
The PDF converter is a fully client-side web application that processes files directly in the user's browser. Here are the main components and technical details:
Core Technologies

HTML5: Page structure and user interface
CSS3: Styling, animations, and responsive design
JavaScript: All conversion logic and user interactions

<h3>Libraries Used</h3><br>

<b>PDF.js (v3.4.120):</b> Mozilla's library for parsing and extracting content from PDF files using JavaScript<br>
<b>JSZip (v3.10.1):</b> Used to create ZIP-formatted files necessary for generating Word (.docx) documents<br>
<b>SheetJS (XLSX, v0.18.5):</b> Library for creating Excel spreadsheets<br>
<b>FileSaver.js (v2.0.5):</b> Helper library for downloading the generated files to the user's computer<br>

<h3>Technical Workflow</h3><br>

<b>File Input:</b><br>

PDF file is uploaded via drag-and-drop or file selection
The file is read into an ArrayBuffer using the FileReader API


<b>PDF Processing:</b>

PDF.js converts the ArrayBuffer into JavaScript objects representing the PDF content
Text content and position information (x, y coordinates) are extracted from each page
These position details are used to maintain the original layout of text on the page


<b>Word Conversion:</b>

Extracted text data is transformed into XML structures compatible with Office Open XML (OOXML) format
JSZip is used to create the file structure required for a DOCX file (various .xml files and folders)
This structure is packaged as a ZIP file conforming to the OOXML format


<b>Excel Conversion:</b>

Text and positioning from the PDF are analyzed to reconstruct table structures
Texts with the same y-coordinate form rows, while x-coordinates determine columns
This data is converted to XLSX format using the SheetJS library


<h3>File Download:</h3><br>

The generated DOCX or XLSX file is downloaded to the user's computer using FileSaver.js



<h3>Security Features</h3>

All processing occurs in the browser; no files are sent to any server
PDF content is temporarily processed in browser memory and automatically cleaned up when finished
Only uses libraries loaded from CDNs that serve over HTTPS

<h3>Performance Optimizations</h3>

Progress bar and incremental processing for handling large files
Asynchronous operations (Promises and async/await) to prevent browser freezing
Preview of the first page to provide quick feedback to the user

This application offers a cost-effective and easily deployable solution as it runs entirely in the browser without requiring server infrastructure. It provides a seamless document conversion experience while maintaining user privacy and data security.<br>

<hr>
<br>
<b>TR:</b><br>
Teknik olarak, PDF dönüştürücü sayfası tamamen istemci tarafında (client-side) çalışan bir web uygulamasıdır. İşte ana bileşenleri ve teknik detayları:
Temel Teknolojiler

HTML5: Sayfa yapısı ve kullanıcı arayüzü
CSS3: Stillemeler, animasyonlar ve responsive tasarım
JavaScript: Tüm dönüştürme mantığı ve kullanıcı etkileşimi

<h3>Kullanılan Kütüphaneler</h3>

<b>PDF.js (v3.4.120):</b> Mozilla tarafından geliştirilen, PDF dosyalarını JavaScript ile işleyip içeriğini çıkarmak için kullanılan kütüphane<br>
<b>JSZip (v3.10.1):</b> Word (.docx) oluşturmak için gereken ZIP formatında dosya oluşturma kütüphanesi<br>
<b>SheetJS (XLSX, v0.18.5):</b> Excel dosyaları oluşturmak için kullanılan kütüphane<br>
<b>FileSaver.js (v2.0.5):</b> Oluşturulan dosyaları kullanıcının bilgisayarına indirmek için kullanılan yardımcı kütüphane<br>

<h3>Teknik İş Akışı</h3>

<b>Dosya Alımı:</b>

Sürükle-bırak veya dosya seçme yoluyla PDF dosyası yüklenir
Dosya, FileReader API kullanılarak bir ArrayBuffer'a okunur


<b>PDF İşleme:</b>

PDF.js, ArrayBuffer'ı alıp PDF içeriğini JavaScript nesnelerine çevirir
Her sayfadan metin içeriği ve pozisyon bilgileri (x, y koordinatları) çıkarılır
Bu pozisyon bilgileri, metinlerin sayfa üzerindeki düzenini korumak için kullanılır


<b>Word Dönüştürme:</b>

Çıkarılan metin verisi, Office Open XML (OOXML) formatına uygun XML yapılarına dönüştürülür
JSZip kullanılarak bir DOCX dosyasının gerektirdiği dosya yapısı (.xml dosyaları ve klasörler) oluşturulur
Bu yapı, OOXML formatına uygun bir ZIP dosyası olarak paketlenir


<b>Excel Dönüştürme:</b>

PDF'den çıkarılan metin ve konumlar, tablo yapısını yeniden oluşturmak için analiz edilir
Aynı y-koordinatına sahip metinler satırları, x-koordinatları ise sütunları belirler
Bu veriler SheetJS kütüphanesi kullanılarak XLSX formatına dönüştürülür


<h3>Dosya İndirme:</h3>

Oluşturulan DOCX veya XLSX dosyası, FileSaver.js kullanılarak kullanıcının bilgisayarına indirilir



<h3>Güvenlik Özellikleri</h3>

Tüm işlem tarayıcıda gerçekleşir, dosyalar hiçbir sunucuya gönderilmez
PDF içeriği geçici olarak tarayıcı belleğinde işlenir ve işlem bitince otomatik olarak temizlenir
Sadece HTTPS üzerinden hizmet veren CDN'lerden yüklenen kütüphaneler kullanılır

<h3>Performans Optimizasyonları</h3>

Büyük dosyaları işlerken, ilerleme çubuğu ve aşamalı işleme kullanılır
Asenkron işlemler (Promise ve async/await) ile tarayıcının donması engellenir
Tek sayfayı önizleme olarak göstererek kullanıcıya hızlı geri bildirim sağlanır

Bu uygulama, sunucu altyapısına ihtiyaç duymadan tamamen tarayıcıda çalışabilmesi sayesinde düşük maliyetli ve kolay dağıtılabilir bir çözüm sunmaktadır.
