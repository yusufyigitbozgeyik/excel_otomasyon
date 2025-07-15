# Stok Yönetim Paneli (Python + Excel)

Bu proje, küçük ve orta ölçekli işletmeler için profesyonel bir stok yönetim uygulamasıdır. Python ile geliştirilmiş, verileri Excel dosyasında saklar ve modern bir arayüz sunar.

## Özellikler
- **Ürün Ekle / Güncelle / Sil**
- **Stok Giriş/Çıkış İşlemleri** (miktar artır/azalt)
- **Gelişmiş Tablo Görünümü** (arama, kategoriye göre filtreleme)
- **Kritik Stok Seviyesi Uyarısı**
- **Raporlama** (toplam ürün, toplam stok, kritik ürün sayısı)
- **Excel Dosyasını Yedekleme**
- **Modern ve kullanıcı dostu arayüz**
- **Tüm işlemler Excel dosyasına otomatik kaydedilir**

## Ürün Bilgileri (Excel Sütunları)
- Stok Kodu (benzersiz)
- Ürün Adı
- Kategori
- Miktar
- Birim (adet, kg, litre vs.)
- Kritik Seviye (miktar azaldığında uyarı için)
- Açıklama
- Son Güncelleme Tarihi

## Kurulum
1. Proje klasörüne gelin:
   ```
   cd excel_otomasyon
   ```
2. Gerekli kütüphaneleri yükleyin:
   ```
   pip install -r requirements.txt
   ```

## Kullanım
1. Uygulamayı başlatın:
   ```
   python app.py
   ```
2. Açılan arayüzde sağ üstten yeni bir Excel dosyası oluşturun veya mevcut bir dosyayı seçin.
3. Ürün ekleyin, güncelleyin, silin veya stok giriş/çıkış işlemleri yapın.
4. Arama ve kategori filtrelerini kullanarak ürünleri kolayca bulun.
5. Rapor ve yedekleme butonlarını kullanarak stok durumunu analiz edin ve dosyanızı yedekleyin.

## Ekran Görüntüsü
> Arayüzün ekran görüntüsünü buraya ekleyebilirsiniz.

## Katkı ve Lisans
Bu proje örnek ve eğitim amaçlıdır. Dilediğiniz gibi geliştirebilir ve paylaşabilirsiniz.

---

**Hazırlayan:** github.com/kullanici_adi 