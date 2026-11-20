# ⚔️ RapidConverter - Egzotik Lila PDF Dönüştürücü

Python ve Streamlit ile geliştirilmiş, **egzotik lila** temalı, arkasında **ufak kılıçlarla savaşan sevimli çöp adam animasyonu** barındıran modern ve esprili bir belge dönüştürme uygulaması.

---

## 🌟 Öne Çıkan Özellikler

1. **📄 PDF'ten Word'e (.docx) Dönüştürme:**
   - `pdf2docx` altyapısıyla PDF dosyalarını düzenlenebilir Word formatına dönüştürür.
2. **📚 Birden Fazla PDF'i Birleştirme (Merge):**
   - `pypdf` kütüphanesiyle birden çok PDF'i tek bir dosyada birleştirir.
3. **🖼️ Farklı Formatları PDF'e Dönüştürme:**
   - **Görseller (JPG, PNG, WEBP):** `Pillow` ile tek tıkla çok sayfalı PDF'e çevrilir.
   - **Word (.docx):** `python-docx` ve `reportlab` ile Render uyumlu, saf Python altyapısıyla yüksek çözünürlüklü PDF'e aktarılır.
4. **🎨 Egzotik Lila Tasarım & Dinamik Modlar:**
   - Sağ üst köşede akıllı tema değiştirici:
     - Gündüz Modunda: **☀️ Güneş simgesi**
     - Gece Modunda: **🇹🇷 Türk bayrağı simgesi**
5. **⚔️ Arka Planda Kılıçlı Çöp Adam Arenası:**
   - Özel HTML5 Canvas motoruyla çizilen, kılıç darbeleriyle kıvılcımlar saçan, taklalar atan ve akrobatik dövüş hareketleri yapan sevimli çöp adam animasyonu!
6. **😄 Esprili ve Samimi Türkçe Hitap Tarzı:**
   - *"Üstad PDF hazır!"*, *"Hepsini tek potada erittik kral!"*, *"Bizim çöp adamlar kılıç biledi..."*

---

## 🚀 Yerel Olarak Çalıştırma

1. Gerekli kütüphaneleri yükleyin:
   ```bash
   pip install -r requirements.txt
   ```

2. Uygulamayı başlatın:
   ```bash
   streamlit run app.py
   ```

3. Tarayıcınızda açılan `http://localhost:8501` adresine gidin.

---

## 🌐 Render Üzerinde Canlıya Alma (Deployment)

1. Projeyi bir GitHub veya GitLab deposuna (repository) yükleyin.
2. [Render.com](https://render.com) hesabınıza giriş yapın.
3. **New +** -> **Web Service** seçeneğine tıklayın.
4. GitHub deponuzu bağlayın.
5. Aşağıdaki ayarları uygulayın:
   - **Environment:** `Python`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
6. **Deploy Web Service** butonuna basın. Birkaç dakika içinde uygulamanız canlıda!

