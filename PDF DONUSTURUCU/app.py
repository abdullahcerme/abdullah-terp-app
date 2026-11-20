import os
import io
import time
import tempfile
from xml.sax.saxutils import escape
import streamlit as st
from PIL import Image

# PDF işleme kütüphaneleri
from pdf2docx import Converter
try:
    from pypdf import PdfWriter as PDF_COMBINER
except ImportError:
    try:
        from pypdf import PdfMerger as PDF_COMBINER
    except ImportError:
        from PyPDF2 import PdfMerger as PDF_COMBINER

# DOCX'ten PDF'e dönüştürme kütüphaneleri (Platform bağımsız Render uyumlu)
from docx import Document as DocxDocument
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# Sayfa Yapılandırması
# ---------------------------------------------------------
st.set_page_config(
    page_title="RapidConverter | Egzotik PDF Dönüştürücü",
    page_icon="⚔️",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ---------------------------------------------------------
# Session State (Koyu / Aydınlık Mod ve Durumlar)
# ---------------------------------------------------------
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = True  # Varsayılan egzotik gece modu

def toggle_theme():
    st.session_state.dark_mode = not st.session_state.dark_mode

# ---------------------------------------------------------
# Egzotik Lila Tema & Dinamik CSS Stilleri
# ---------------------------------------------------------
if st.session_state.dark_mode:
    # KOYU EGZOTİK LİLA MOD
    bg_gradient = "radial-gradient(circle at 20% 20%, #1c0b2b 0%, #0d0417 60%, #07020d 100%)"
    card_bg = "rgba(28, 14, 46, 0.78)"
    card_border = "1px solid rgba(192, 132, 252, 0.3)"
    text_primary = "#f3e8ff"
    text_secondary = "#d8b4fe"
    accent_color = "#c084fc"
    accent_hover = "#d8b4fe"
    button_bg = "linear-gradient(135deg, #9333ea 0%, #7e22ce 100%)"
    button_shadow = "0 0 20px rgba(168, 85, 247, 0.55)"
    glass_nav = "rgba(23, 10, 38, 0.85)"
    canvas_opacity = "0.7"
    theme_btn_label = "🇹🇷 Koyu Mod (Aydınlığa Geç)"
    theme_btn_help = "Tıkla ve gündüz moduna ışınlan üstad!"
else:
    # AYDINLIK EGZOTİK LİLA MOD
    bg_gradient = "radial-gradient(circle at 20% 20%, #faf5ff 0%, #f3e8ff 50%, #e9d5ff 100%)"
    card_bg = "rgba(255, 255, 255, 0.82)"
    card_border = "1px solid rgba(168, 85, 247, 0.35)"
    text_primary = "#3b0764"
    text_secondary = "#6b21a8"
    accent_color = "#9333ea"
    accent_hover = "#7e22ce"
    button_bg = "linear-gradient(135deg, #a855f7 0%, #9333ea 100%)"
    button_shadow = "0 8px 25px rgba(147, 51, 234, 0.3)"
    glass_nav = "rgba(255, 255, 255, 0.9)"
    canvas_opacity = "0.55"
    theme_btn_label = "☀️ Aydınlık Mod (Karanlığa Geç)"
    theme_btn_help = "Tıkla ve gece egzotik moduna geç kral!"

st.markdown(f"""
<style>
    /* Ana Arka Plan ve Yazı Tipleri */
    .stApp {{
        background: {bg_gradient} !important;
        background-attachment: fixed !important;
        color: {text_primary} !important;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    }}

    /* Üst Boşluk Düzeltme */
    .block-container {{
        padding-top: 1.8rem;
        padding-bottom: 3rem;
        max-width: 1100px;
        position: relative;
        z-index: 2;
    }}

    /* Cam Efektli (Glassmorphism) Kartlar */
    .rapid-card {{
        background: {card_bg};
        backdrop-filter: blur(14px);
        -webkit-backdrop-filter: blur(14px);
        border: {card_border};
        border-radius: 20px;
        padding: 26px;
        margin-bottom: 24px;
        box-shadow: 0 10px 35px rgba(0, 0, 0, 0.15);
        transition: transform 0.25s ease, box-shadow 0.25s ease;
    }}
    .rapid-card:hover {{
        transform: translateY(-2px);
        box-shadow: 0 16px 40px rgba(168, 85, 247, 0.2);
    }}

    /* Başlık ve Rozetler */
    .rapid-title {{
        font-size: 2.6rem;
        font-weight: 800;
        background: linear-gradient(135deg, #d8b4fe 0%, #c084fc 50%, #f472b6 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: -0.5px;
        margin-bottom: 4px;
        display: inline-block;
    }}
    .rapid-subtitle {{
        color: {text_secondary};
        font-size: 1.05rem;
        font-weight: 500;
        margin-bottom: 12px;
    }}
    .rapid-badge {{
        display: inline-flex;
        align-items: center;
        gap: 6px;
        background: rgba(168, 85, 247, 0.18);
        border: 1px solid rgba(192, 132, 252, 0.4);
        color: {accent_color};
        padding: 4px 14px;
        border-radius: 50px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-bottom: 10px;
    }}

    /* Buton Tasarımları */
    .stButton > button {{
        background: {button_bg} !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.65rem 1.6rem !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        box-shadow: {button_shadow} !important;
        transition: all 0.25s cubic-bezier(0.4, 0, 0.2, 1) !important;
        cursor: pointer !important;
    }}
    .stButton > button:hover {{
        transform: translateY(-2px) scale(1.02) !important;
        box-shadow: 0 0 25px rgba(192, 132, 252, 0.75) !important;
    }}
    .stButton > button:active {{
        transform: translateY(0) scale(0.98) !important;
    }}

    /* İndirme Butonu */
    .stDownloadButton > button {{
        background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.4) !important;
        animation: pulseDownload 2s infinite;
    }}
    @keyframes pulseDownload {{
        0% {{ box-shadow: 0 0 0 0 rgba(16, 185, 129, 0.6); }}
        70% {{ box-shadow: 0 0 0 12px rgba(16, 185, 129, 0); }}
        100% {{ box-shadow: 0 0 0 0 rgba(16, 185, 129, 0); }}
    }}

    /* Sekme Çubuğu (Tabs) */
    .stTabs [data-baseweb="tab-list"] {{
        background: {card_bg};
        padding: 6px;
        border-radius: 14px;
        gap: 8px;
        border: {card_border};
    }}
    .stTabs [data-baseweb="tab"] {{
        border-radius: 10px;
        color: {text_secondary};
        font-weight: 600;
        padding: 10px 20px;
        transition: all 0.2s ease;
    }}
    .stTabs [aria-selected="true"] {{
        background: rgba(168, 85, 247, 0.25) !important;
        color: {accent_color} !important;
        border-bottom: 2px solid {accent_color} !important;
    }}

    /* Dosya Yükleme Kutuları */
    [data-testid="stFileUploader"] {{
        background: {card_bg};
        border-radius: 16px;
        border: 2px dashed rgba(168, 85, 247, 0.4);
        padding: 16px;
        transition: border 0.3s ease;
    }}
    [data-testid="stFileUploader"]:hover {{
        border-color: {accent_color};
    }}

    /* Bilgi ve Uyarı Kutuları */
    .stAlert {{
        border-radius: 14px !important;
        backdrop-filter: blur(8px) !important;
    }}

    /* Arka Plan Çöp Adam Arena Canvas Konteyneri */
    #stickman-canvas-host {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        pointer-events: none;
        z-index: 1;
        opacity: {canvas_opacity};
        overflow: hidden;
    }}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# Havalı Kılıçlı Çöp Adam Arka Plan Animasyonu (HTML5 Canvas)
# ---------------------------------------------------------
st.components.v1.html(f"""
<!DOCTYPE html>
<html>
<head>
<style>
    body, html {{
        margin: 0;
        padding: 0;
        overflow: hidden;
        background: transparent;
        width: 100vw;
        height: 100vh;
    }}
    canvas {{
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;
        pointer-events: none;
    }}
</style>
</head>
<body>
<canvas id="stickmanCanvas"></canvas>
<script>
    const canvas = document.getElementById('stickmanCanvas');
    const ctx = canvas.getContext('2d');

    function resize() {{
        canvas.width = window.innerWidth;
        canvas.height = window.innerHeight;
    }}
    window.addEventListener('resize', resize);
    resize();

    // Zaman Ölçeği (Hızlı Dövüş & Sinematik Ağır Çekim)
    let timeScale = 1.0;
    let slowMoTimer = 0;

    function triggerSlowMo(duration, scale) {{
        slowMoTimer = duration || 30;
        timeScale = scale || 0.24;
    }}

    // Kıvılcım parçacıkları ve Şok Dalgaları
    const sparks = [];
    const shockwaves = [];

    function addSparks(x, y, color, count, speedMult) {{
        const cnt = count || 16;
        const spdMult = speedMult || 1.0;
        for(let i=0; i<cnt; i++) {{
            const angle = Math.random() * Math.PI * 2;
            const speed = (Math.random() * 5.5 + 2) * spdMult;
            sparks.push({{
                x: x,
                y: y,
                vx: Math.cos(angle) * speed,
                vy: Math.sin(angle) * speed - 1.5,
                alpha: 1,
                decay: Math.random() * 0.025 + 0.015,
                color: color || '#facc15',
                radius: Math.random() * 2.5 + 1.2
            }});
        }}
    }}

    function addShockwave(x, y, color) {{
        shockwaves.push({{
            x: x,
            y: y,
            radius: 6,
            maxRadius: 90,
            alpha: 0.85,
            color: color || '#c084fc'
        }});
    }}

    // 10 Farklı Efsanevi Savaş Hareketi Listesi
    const COMBAT_MOVES = [
        'slash_combo',      // 1. Seri 3'lü çapraz kılıç darbeleri
        'heavy_strike',     // 2. Ağır efsanevi kılıç savurması (Ağır Çekim)
        'aerial_dive',      // 3. Havadan kılıçla dalış (Ağır Çekim)
        'flash_dash',       // 4. Işık hızında öne atılma ve kılıç saplama
        'tornado_spin',     // 5. Kendi etrafında 360 dönen kasırga savuruşu
        'backflip_slash',   // 6. Geriye akrobatik ters takla ile hilal kesişi
        'parry_counter',    // 7. Kılıç savuşturma ve ani karşı taarruz
        'thrust_charge',    // 8. Derin ve keskin düz mızrak hamlesi
        'somersault_strike',// 9. Rakibin üzerinden öne takla atarak havadan vuruş
        'blade_dance'       // 10. Efsanevi kılıç dansı ve parıltılı bitirici (Ağır Çekim)
    ];

    // Çöp Adam Savaşçı Sınıfı
    class StickFighter {{
        constructor(x, y, dir, swordColor, name) {{
            this.baseX = x;
            this.x = x;
            this.y = y;
            this.dir = dir; // 1 = sağa, -1 = sola
            this.swordColor = swordColor;
            this.name = name;
            this.state = 'walk';
            this.animTick = Math.random() * 100;
            this.swordAngle = 0;
            this.legAngle = 0;
            this.armAngle = 0;
            this.jumpY = 0;
            this.flipAngle = 0;
            this.opponent = null;
        }}

        pickNextMove() {{
            const idx = Math.floor(Math.random() * COMBAT_MOVES.length);
            this.state = COMBAT_MOVES[idx];
            this.animTick = 0;
        }}

        update(dt) {{
            this.animTick += dt;

            // Rakibe göre yön ayarla
            if (this.opponent) {{
                const dist = this.opponent.x - this.x;
                if (this.state === 'walk' || this.state === 'idle') {{
                    this.dir = dist >= 0 ? 1 : -1;
                }}

                // Yakınlaşınca 10 efsanevi hareketten birini seç
                if (Math.abs(dist) < 75 && this.state === 'walk') {{
                    this.pickNextMove();
                }}
            }}

            // 10 FARKLI SAVAŞ HAREKETİ MOTORU
            if (this.state === 'walk') {{
                this.x += this.dir * 1.8 * dt;
                this.legAngle = Math.sin(this.animTick * 0.18) * 0.45;
                this.swordAngle = -0.3 + Math.sin(this.animTick * 0.1) * 0.15;
                this.jumpY = Math.abs(Math.sin(this.animTick * 0.18)) * -6;
                this.flipAngle = 0;

                if (this.animTick > 50) {{
                    this.pickNextMove();
                }}
            }}
            // HAREKET 1: Seri Çapraz Kılıç Kombosu
            else if (this.state === 'slash_combo') {{
                this.swordAngle = Math.sin(this.animTick * 0.42) * 2.0;
                this.legAngle = 0.3;
                this.jumpY = -4;
                this.x += this.dir * 1.2 * dt;

                if (this.animTick >= 10 && this.animTick <= 12 && this.opponent) {{
                    addSparks((this.x + this.opponent.x)/2, this.y - 45, this.swordColor, 8);
                }}
                if (this.animTick > 25) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}
            // HAREKET 2: Ağır Efsanevi Vuruş (AĞIR ÇEKİM)
            else if (this.state === 'heavy_strike') {{
                const p = Math.min(1, this.animTick / 30);
                this.swordAngle = -1.5 + p * 3.2;
                this.jumpY = -Math.sin(p * Math.PI) * 12;
                this.x += this.dir * 1.4 * dt;

                if (this.animTick >= 14 && this.animTick <= 16 && this.opponent) {{
                    triggerSlowMo(36, 0.22); // Sinematik Ağır Çekim
                    const clashX = (this.x + this.opponent.x) / 2;
                    addSparks(clashX, this.y - 45, '#ffffff', 20, 1.4);
                    addSparks(clashX, this.y - 45, this.swordColor, 16, 1.2);
                    addShockwave(clashX, this.y - 45, this.swordColor);
                }}
                if (this.animTick > 30) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}
            // HAREKET 3: Havadan Kılıçla Dalış (AĞIR ÇEKİM)
            else if (this.state === 'aerial_dive') {{
                const p = Math.min(1, this.animTick / 36);
                this.jumpY = -Math.sin(p * Math.PI) * 62;
                this.swordAngle = 1.5;
                this.legAngle = -0.4;
                this.x += this.dir * 2.2 * dt;

                if (this.animTick >= 16 && this.animTick <= 18) {{
                    triggerSlowMo(32, 0.25); // Havada asılı kalma ağır çekimi
                }}
                if (this.animTick >= 33 && this.animTick <= 35) {{
                    addSparks(this.x + this.dir * 25, this.y - 15, this.swordColor, 18, 1.2);
                    addShockwave(this.x + this.dir * 25, this.y - 15, this.swordColor);
                }}
                if (this.animTick > 36) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}
            // HAREKET 4: Işık Hızında Öne Atılma (Flash Dash)
            else if (this.state === 'flash_dash') {{
                this.x += this.dir * 5.8 * dt;
                this.jumpY = -2;
                this.swordAngle = 0.85;
                this.legAngle = 0.5;

                if (this.animTick >= 12 && this.opponent) {{
                    addSparks((this.x + this.opponent.x)/2, this.y - 40, this.swordColor, 12, 1.3);
                }}
                if (this.animTick > 20) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}
            // HAREKET 5: Kendi Etrafında 360 Kasırga Savuruşu
            else if (this.state === 'tornado_spin') {{
                const p = Math.min(1, this.animTick / 28);
                this.flipAngle = this.dir * (p * Math.PI * 4); // 720 derece çift tur
                this.swordAngle = 1.3;
                this.jumpY = -Math.sin(p * Math.PI) * 16;
                this.x += this.dir * 2.2 * dt;

                if (this.animTick >= 12 && this.animTick <= 14) {{
                    addSparks(this.x, this.y - 40, this.swordColor, 12, 1.1);
                }}
                if (this.animTick > 28) {{
                    this.state = 'walk';
                    this.animTick = 0;
                    this.flipAngle = 0;
                }}
            }}
            // HAREKET 6: Geriye Akrobatik Ters Takla & Hilal Savuruşu
            else if (this.state === 'backflip_slash') {{
                const p = Math.min(1, this.animTick / 34);
                this.jumpY = -Math.sin(p * Math.PI) * 48;
                this.flipAngle = -this.dir * (p * Math.PI * 2);
                this.swordAngle = -0.5 + p * 2.4;
                this.x -= this.dir * 2.8 * dt;

                if (this.animTick > 34) {{
                    this.state = 'walk';
                    this.animTick = 0;
                    this.flipAngle = 0;
                    this.jumpY = 0;
                }}
            }}
            // HAREKET 7: Kılıç Savuşturma (Parry) & Karşı Darbe
            else if (this.state === 'parry_counter') {{
                if (this.animTick < 14) {{
                    // Savunma ve savuşturma
                    this.swordAngle = 0.2;
                    this.jumpY = -2;
                    if (this.animTick >= 6 && this.animTick <= 8) {{
                        addSparks(this.x + this.dir * 20, this.y - 45, '#fbbf24', 10);
                    }}
                }} else {{
                    // Şimşek karşı saldırı
                    this.swordAngle = -1.3;
                    this.x += this.dir * 2.6 * dt;
                    if (this.animTick >= 16 && this.animTick <= 18 && this.opponent) {{
                        addSparks((this.x + this.opponent.x)/2, this.y - 45, this.swordColor, 14);
                    }}
                }}
                if (this.animTick > 28) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}
            // HAREKET 8: Derin Düz Mızrak Hamlesi (Stinger Thrust)
            else if (this.state === 'thrust_charge') {{
                this.swordAngle = 0.05;
                this.legAngle = -0.3;
                this.jumpY = -1;
                this.x += this.dir * 3.6 * dt;

                if (this.animTick >= 10 && this.animTick <= 12 && this.opponent) {{
                    addSparks((this.x + this.opponent.x)/2, this.y - 38, this.swordColor, 12, 1.2);
                    addShockwave((this.x + this.opponent.x)/2, this.y - 38, this.swordColor);
                }}
                if (this.animTick > 24) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}
            // HAREKET 9: Rakibin Üzerinden Takla Atarak Havadan Vuruş
            else if (this.state === 'somersault_strike') {{
                const p = Math.min(1, this.animTick / 36);
                this.jumpY = -Math.sin(p * Math.PI) * 54;
                this.flipAngle = this.dir * (p * Math.PI * 2);
                this.swordAngle = -0.6 + p * 2.5;
                this.x += this.dir * 3.4 * dt;

                if (this.animTick >= 22 && this.animTick <= 24 && this.opponent) {{
                    addSparks((this.x + this.opponent.x)/2, this.y - 42, this.swordColor, 15);
                }}
                if (this.animTick > 36) {{
                    this.state = 'walk';
                    this.animTick = 0;
                    this.flipAngle = 0;
                }}
            }}
            // HAREKET 10: Efsanevi Kılıç Dansı ve Bitirici (AĞIR ÇEKİM)
            else if (this.state === 'blade_dance') {{
                const p = Math.min(1, this.animTick / 40);
                this.swordAngle = Math.sin(p * Math.PI * 4) * 1.6;
                this.jumpY = -Math.abs(Math.sin(p * Math.PI * 2)) * 14;
                this.x += this.dir * 1.5 * dt;

                if (this.animTick >= 20 && this.animTick <= 22 && this.opponent) {{
                    triggerSlowMo(40, 0.20); // Efsanevi Ağır Çekim
                    const clashX = (this.x + this.opponent.x) / 2;
                    addSparks(clashX, this.y - 45, '#ffffff', 24, 1.5);
                    addSparks(clashX, this.y - 45, this.swordColor, 20, 1.3);
                    addShockwave(clashX, this.y - 45, this.swordColor);
                }}
                if (this.animTick > 40) {{
                    this.state = 'walk';
                    this.animTick = 0;
                }}
            }}

            // Ekrandan çıkmayı engelleyen yumuşak sınır kontrolü
            if (this.x < 40) {{
                this.x = 40;
                this.dir = 1;
            }}
            if (this.x > canvas.width - 40) {{
                this.x = canvas.width - 40;
                this.dir = -1;
            }}
        }}

        draw() {{
            ctx.save();
            ctx.translate(this.x, this.y + this.jumpY);
            if (this.flipAngle !== 0) {{
                ctx.rotate(this.flipAngle);
            }}

            const strokeStyle = '#c084fc';
            ctx.strokeStyle = strokeStyle;
            ctx.lineWidth = 3.2;
            ctx.lineCap = 'round';
            ctx.lineJoin = 'round';

            // Kafa (Tatlı küçük yuvarlak)
            ctx.beginPath();
            ctx.arc(0, -56, 12, 0, Math.PI * 2);
            ctx.stroke();

            // Havalı Ninja Saç Bandı / Fular ucu
            ctx.beginPath();
            ctx.strokeStyle = '#f472b6';
            ctx.lineWidth = 2.4;
            const ribbonWave = Math.sin(this.animTick * 0.2) * 5;
            ctx.moveTo(-this.dir * 8, -56);
            ctx.quadraticCurveTo(-this.dir * 22, -56 + ribbonWave, -this.dir * 32, -52 - ribbonWave);
            ctx.stroke();

            // Gövde
            ctx.beginPath();
            ctx.strokeStyle = strokeStyle;
            ctx.lineWidth = 3.2;
            ctx.moveTo(0, -44);
            ctx.lineTo(0, -18);
            ctx.stroke();

            // Bacaklar
            const l1 = Math.sin(this.legAngle) * 22;
            const l2 = -Math.sin(this.legAngle) * 22;
            // Sol bacak
            ctx.beginPath();
            ctx.moveTo(0, -18);
            ctx.lineTo(-l1 * 0.5, -9);
            ctx.lineTo(-l1, 0);
            ctx.stroke();
            // Sağ bacak
            ctx.beginPath();
            ctx.moveTo(0, -18);
            ctx.lineTo(l2 * 0.5, -9);
            ctx.lineTo(l2, 0);
            ctx.stroke();

            // Kollar ve Havalı Kılıç
            // Sol kol (Denge kolu)
            ctx.beginPath();
            ctx.moveTo(0, -38);
            ctx.lineTo(-this.dir * 14, -28);
            ctx.lineTo(-this.dir * 18, -16);
            ctx.stroke();

            // Sağ kol (Kılıcı tutan ana kol)
            const handX = this.dir * 16;
            const handY = -34;
            ctx.beginPath();
            ctx.moveTo(0, -38);
            ctx.lineTo(this.dir * 10, -36);
            ctx.lineTo(handX, handY);
            ctx.stroke();

            // Egzotik Işıldayan Kılıç (Katana)
            ctx.save();
            ctx.translate(handX, handY);
            ctx.rotate(this.dir * this.swordAngle);

            // Kılıç Kabzası
            ctx.strokeStyle = '#e2e8f0';
            ctx.lineWidth = 2.8;
            ctx.beginPath();
            ctx.moveTo(-this.dir * 4, 3);
            ctx.lineTo(this.dir * 6, -5);
            ctx.stroke();

            // Kılıç Koruması (Tsuba)
            ctx.strokeStyle = '#fbbf24';
            ctx.lineWidth = 3.5;
            ctx.beginPath();
            ctx.moveTo(this.dir * 4, -8);
            ctx.lineTo(this.dir * 8, -2);
            ctx.stroke();

            // Kılıç Namlusu (Neon Işıltılı Lila / Camgöbeği)
            ctx.shadowColor = this.swordColor;
            ctx.shadowBlur = 14;
            ctx.strokeStyle = this.swordColor;
            ctx.lineWidth = 3.5;
            ctx.beginPath();
            ctx.moveTo(this.dir * 6, -5);
            ctx.lineTo(this.dir * 38, -32);
            ctx.stroke();

            // Kılıç keskin uç parıltısı
            ctx.fillStyle = '#ffffff';
            ctx.beginPath();
            ctx.arc(this.dir * 38, -32, 2.2, 0, Math.PI * 2);
            ctx.fill();

            ctx.restore();
            ctx.restore();
        }}
    }}

    // Çöp Adam Arenası: 2 Çift Savaşçı (Üst ve Alt Katlar)
    const fighters = [];

    // Alt Zemin Düellocuları (Mor Ninja vs Camgöbeği Samuray)
    const y1 = Math.min(window.innerHeight - 80, 680);
    const f1 = new StickFighter(window.innerWidth * 0.22, y1, 1, '#d8b4fe', 'MorNinja');
    const f2 = new StickFighter(window.innerWidth * 0.42, y1, -1, '#38bdf8', 'GökSamuray');
    f1.opponent = f2;
    f2.opponent = f1;
    fighters.push(f1, f2);

    // Üst Havada Akrobat Düellocuları (Pembe Kılıç vs Altın Kılıç)
    const y2 = Math.min(window.innerHeight * 0.48, 380);
    const f3 = new StickFighter(window.innerWidth * 0.65, y2, 1, '#f472b6', 'AteşKılıç');
    const f4 = new StickFighter(window.innerWidth * 0.85, y2, -1, '#facc15', 'AltınKılıç');
    f3.opponent = f4;
    f4.opponent = f3;
    fighters.push(f3, f4);

    // Animasyon Döngüsü
    function animate() {{
        // Ağır Çekim Zamanlayıcı Kontrolü
        if (slowMoTimer > 0) {{
            slowMoTimer--;
            if (slowMoTimer <= 0) {{
                timeScale = 1.0;
            }}
        }}
        const dt = timeScale;

        ctx.clearRect(0, 0, canvas.width, canvas.height);

        // Şok Dalgalarını Güncelle ve Çiz
        for(let i = shockwaves.length - 1; i >= 0; i--) {{
            const sw = shockwaves[i];
            sw.radius += (sw.maxRadius - sw.radius) * 0.15 * dt + 1.2;
            sw.alpha -= 0.04 * dt;
            if (sw.alpha <= 0 || sw.radius >= sw.maxRadius) {{
                shockwaves.splice(i, 1);
                continue;
            }}
            ctx.save();
            ctx.globalAlpha = sw.alpha;
            ctx.strokeStyle = sw.color;
            ctx.shadowColor = sw.color;
            ctx.shadowBlur = 12;
            ctx.lineWidth = 2.8;
            ctx.beginPath();
            ctx.arc(sw.x, sw.y, sw.radius, 0, Math.PI * 2);
            ctx.stroke();
            ctx.restore();
        }}

        // Kıvılcımları güncelle ve çiz
        for(let i = sparks.length - 1; i >= 0; i--) {{
            const s = sparks[i];
            s.x += s.vx * dt;
            s.y += s.vy * dt;
            s.vy += 0.15 * dt; // Yerçekimi
            s.alpha -= s.decay * dt;

            if (s.alpha <= 0) {{
                sparks.splice(i, 1);
                continue;
            }}

            ctx.save();
            ctx.globalAlpha = s.alpha;
            ctx.fillStyle = s.color;
            ctx.shadowColor = s.color;
            ctx.shadowBlur = 8;
            ctx.beginPath();
            ctx.arc(s.x, s.y, s.radius, 0, Math.PI*2);
            ctx.fill();
            ctx.restore();
        }}

        // Çöp adamları güncelle ve çiz
        fighters.forEach(f => {{
            f.update(dt);
            f.draw();
        }});

        requestAnimationFrame(animate);
    }}

    animate();
</script>
</body>
</html>
""", height=220)

# ---------------------------------------------------------
# Üst Başlık & Sağ Üst Tema Değiştirme Butonu
# ---------------------------------------------------------
col_brand, col_theme = st.columns([0.76, 0.24])

with col_brand:
    st.markdown('<div class="rapid-badge">⚔️ Egzotik Lila Muhafızları Devrede &nbsp;|&nbsp; ⚡ Işık Hızında Dönüştürme</div>', unsafe_allow_html=True)
    st.markdown('<h1 class="rapid-title">RapidConverter</h1>', unsafe_allow_html=True)
    st.markdown(
        f'<p class="rapid-subtitle">Belgelerini zahmetsizce dönüştür üstad! Arka planda çöp adamlar kılıçlarıyla sayfaları dilimliyor, sen sadece keyfine bak.</p>',
        unsafe_allow_html=True
    )

with col_theme:
    st.write("") # Görsel dengeleme boşluğu
    st.button(
        theme_btn_label,
        key="theme_toggle_btn",
        on_click=toggle_theme,
        help=theme_btn_help,
        use_container_width=True
    )
    if st.session_state.dark_mode:
        st.caption("🌙 Gece bekçileri kılıç biledi, gözler dinleniyor.")
    else:
        st.caption("☀️ Güneş parıldıyor, çöp adamlar antrenmanda!")

st.markdown("<hr style='border: none; height: 1px; background: rgba(168, 85, 247, 0.25); margin: 8px 0 24px 0;'/>", unsafe_allow_html=True)

# ---------------------------------------------------------
# Yardımcı Dönüştürme Fonksiyonları
# ---------------------------------------------------------
def format_bytes(num_bytes):
    """Bayt miktarını okunabilir boyuta (KB/MB) dönüştürür."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    elif num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    else:
        return f"{num_bytes / (1024 * 1024):.2f} MB"

def convert_pdf_to_docx_file(pdf_bytes):
    """pdf2docx kullanarak PDF'i Word (.docx) dosyasına dönüştürür."""
    tmp_pdf_path = None
    tmp_docx_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(pdf_bytes)
            tmp_pdf_path = tmp_pdf.name
        tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")

        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None)
        cv.close()

        with open(tmp_docx_path, "rb") as f:
            docx_data = f.read()
        return docx_data, None
    except Exception as e:
        return None, str(e)
    finally:
        if tmp_pdf_path and os.path.exists(tmp_pdf_path):
            try:
                os.remove(tmp_pdf_path)
            except Exception:
                pass
        if tmp_docx_path and os.path.exists(tmp_docx_path):
            try:
                os.remove(tmp_docx_path)
            except Exception:
                pass

def merge_multiple_pdfs(uploaded_pdf_list):
    """pypdf kullanarak birden fazla PDF dosyasını tek bir PDF'te birleştirir."""
    try:
        merger = PDF_COMBINER()
        for pdf_file in uploaded_pdf_list:
            pdf_bytes_io = io.BytesIO(pdf_file.getvalue())
            merger.append(pdf_bytes_io)
        output_stream = io.BytesIO()
        merger.write(output_stream)
        merger.close()
        return output_stream.getvalue(), None
    except Exception as e:
        return None, str(e)

def convert_images_to_pdf_bytes(image_files):
    """Pillow kullanarak görselleri (JPG/PNG) tek bir PDF'e dönüştürür."""
    try:
        pil_images = []
        for img_file in image_files:
            img = Image.open(io.BytesIO(img_file.getvalue()))
            # RGB moduna dönüştür (PNG şeffaflığı ve CMYK uyumluluğu için)
            if img.mode in ("RGBA", "P", "LA"):
                bg = Image.new("RGB", img.size, (255, 255, 255))
                if img.mode == "P":
                    img = img.convert("RGBA")
                bg.paste(img, mask=img.split()[-1])
                img = bg
            elif img.mode != "RGB":
                img = img.convert("RGB")
            pil_images.append(img)

        if not pil_images:
            return None, "Hiç geçerli görsel bulunamadı üstadım!"

        out_buffer = io.BytesIO()
        first_img = pil_images[0]
        extra_imgs = pil_images[1:] if len(pil_images) > 1 else []
        first_img.save(out_buffer, format="PDF", save_all=True, append_images=extra_imgs)
        return out_buffer.getvalue(), None
    except Exception as e:
        return None, str(e)

def convert_docx_to_pdf_bytes(docx_bytes):
    """
    python-docx ve ReportLab kullanarak Word (.docx) dosyasını PDF'e dönüştürür.
    Render üzerinde harici LibreOffice veya MS Word gerektirmeden %100 saf Python ile çalışır.
    """
    try:
        doc = DocxDocument(io.BytesIO(docx_bytes))
        pdf_buffer = io.BytesIO()
        pdf = SimpleDocTemplate(
            pdf_buffer,
            pagesize=letter,
            rightMargin=40, leftMargin=40,
            topMargin=40, bottomMargin=40
        )
        styles = getSampleStyleSheet()

        # Egzotik Başlık ve Paragraf Stilleri
        h1_style = ParagraphStyle(
            'RapidH1',
            parent=styles['Heading1'],
            fontName='Helvetica-Bold',
            fontSize=18,
            leading=22,
            textColor=colors.HexColor('#7e22ce'),
            spaceAfter=12
        )
        h2_style = ParagraphStyle(
            'RapidH2',
            parent=styles['Heading2'],
            fontName='Helvetica-Bold',
            fontSize=14,
            leading=18,
            textColor=colors.HexColor('#9333ea'),
            spaceAfter=8
        )
        body_style = ParagraphStyle(
            'RapidBody',
            parent=styles['Normal'],
            fontName='Helvetica',
            fontSize=10.5,
            leading=15,
            textColor=colors.HexColor('#1f2937'),
            spaceAfter=6
        )

        story = []

        # Paragrafları dönüştür
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                story.append(Spacer(1, 6))
                continue

            safe_text = escape(text)
            style_name = p.style.name.lower() if p.style and p.style.name else ""

            if "heading 1" in style_name or "title" in style_name:
                story.append(Paragraph(safe_text, h1_style))
            elif "heading 2" in style_name:
                story.append(Paragraph(safe_text, h2_style))
            else:
                story.append(Paragraph(safe_text, body_style))

        # Tabloları dönüştür
        for table in doc.tables:
            table_data = []
            for row in table.rows:
                row_cells = [Paragraph(escape(cell.text.strip()), body_style) for cell in row.cells]
                table_data.append(row_cells)

            if table_data:
                t = Table(table_data)
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#f3e8ff')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor('#581c87')),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                    ('TOPPADDING', (0, 0), (-1, -1), 6),
                    ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#d8b4fe')),
                    ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#9333ea')),
                ]))
                story.append(Spacer(1, 8))
                story.append(t)
                story.append(Spacer(1, 10))

        if not story:
            story.append(Paragraph("Boş Belge", body_style))

        pdf.build(story)
        return pdf_buffer.getvalue(), None
    except Exception as e:
        return None, str(e)


# ---------------------------------------------------------
# Ana Sekmeler (Tabs)
# ---------------------------------------------------------
tab_pdf2word, tab_merge, tab_to_pdf = st.tabs([
    "📄 PDF'i Word'e (.docx) Çevir",
    "📚 Birden Fazla PDF'i Birleştir",
    "🖼️ Word / Görselleri PDF'e Çevir"
])

# =========================================================
# SEKME 1: PDF -> Word (.docx)
# =========================================================
with tab_pdf2word:
    st.markdown("""
    <div class="rapid-card">
        <h3>📄 PDF'ini Düzenlenebilir Word Belgesine Çevir</h3>
        <p>PDF dosyan kilitli sandık gibi mi duruyor üstad? Bize bırak, kılıç darbeleriyle metinleri ve tabloları Word formatına dönüştürelim!</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_pdf = st.file_uploader(
        "Dönüştürmek istediğin PDF dosyasını buraya sürükle veya seç kral:",
        type=["pdf"],
        key="pdf2docx_uploader",
        help="Yalnızca .pdf formatında dosyalar kabul edilir."
    )

    if uploaded_pdf is not None:
        file_size_str = format_bytes(len(uploaded_pdf.getvalue()))
        st.info(f"📁 **Seçilen Dosya:** `{uploaded_pdf.name}` &nbsp;|&nbsp; 📊 **Boyut:** `{file_size_str}` &nbsp;|&nbsp; *Maşallah dosyan canavar gibi!*")

        if st.button("🚀 Word'e (.docx) Dönüştür Üstad!", key="btn_convert_pdf2docx", use_container_width=True):
            with st.spinner("⚔️ Çöp adamlar kılıç sallayarak PDF sayfalarını Word'e çeviriyor... Çayından bir yudum al üstad..."):
                start_time = time.time()
                docx_bytes, err = convert_pdf_to_docx_file(uploaded_pdf.getvalue())
                elapsed = time.time() - start_time

            if err:
                st.error(f"😅 **Eyvah üstadım! Bir terslik oldu:** `{err}`\n\n*Bizim çöp adamların kılıcı dosyanın şifresine çarpmış olabilir. Dosyanın şifreli veya bozuk olmadığından emin ol kral!*")
            else:
                st.balloons()
                out_name = os.path.splitext(uploaded_pdf.name)[0] + "_rapid.docx"
                st.success(f"🎉 **Üstad PDF hazır!** Sıcak fırından yeni çıktı, sadece **{elapsed:.1f} saniyede** taptaze .docx belgen hazırlandı!")
                st.download_button(
                    label=f"📥 {out_name} Belgesini İndir",
                    data=docx_bytes,
                    file_name=out_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

# =========================================================
# SEKME 2: Birden Fazla PDF'i Birleştir (Merge)
# =========================================================
with tab_merge:
    st.markdown("""
    <div class="rapid-card">
        <h3>📚 Çoklu PDF Birleştirici (Potada Eritme)</h3>
        <p>Parça parça PDF'lerden bıktın mı üstad? İster 2 tane, ister 50 tane PDF yükle; hepsini tek bir güçlü ve nizami PDF olarak birleştirelim!</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded_pdfs = st.file_uploader(
        "Birleştirmek istediğin tüm PDF dosyalarını seç üstad:",
        type=["pdf"],
        accept_multiple_files=True,
        key="merge_uploader",
        help="Birden fazla dosya seçebilir veya sürükleyip bırakabilirsin."
    )

    if uploaded_pdfs:
        st.write(f"📋 **Kuyruktaki Belgeler ({len(uploaded_pdfs)} adet):**")
        total_size = sum(len(f.getvalue()) for f in uploaded_pdfs)
        cols_summary = st.columns(min(len(uploaded_pdfs), 4))
        for idx, f in enumerate(uploaded_pdfs):
            col_target = cols_summary[idx % len(cols_summary)]
            col_target.caption(f"🔹 {idx+1}. `{f.name}` ({format_bytes(len(f.getvalue()))})")

        st.caption(f"📦 **Toplam Boyut:** `{format_bytes(total_size)}`")

        col_merge_name, col_merge_action = st.columns([0.65, 0.35])
        with col_merge_name:
            custom_merge_name = st.text_input(
                "Oluşturulacak dosyanın adı ne olsun kral?",
                value="rapid_birlestirilmis_belge.pdf",
                help="Sonuna .pdf koymayı unutma üstad!"
            )
        with col_merge_action:
            st.write("")
            st.write("")
            merge_btn = st.button("🔥 Hepsini Birleştir Kral!", key="btn_merge_pdfs", use_container_width=True)

        if merge_btn:
            if not custom_merge_name.lower().endswith(".pdf"):
                custom_merge_name += ".pdf"

            with st.spinner("⚔️ Çöp adamlar sayfaları sıraya dizip çelik gibi birbirine perçinliyor..."):
                start_time = time.time()
                merged_bytes, err = merge_multiple_pdfs(uploaded_pdfs)
                elapsed = time.time() - start_time

            if err:
                st.error(f"😅 **Eyvah üstadım! Sayfaları birleştirirken kılıçlar karıştı:** `{err}`\n\n*Lütfen PDF'lerin hasarlı veya bozuk olmadığını kontrol et kral.*")
            else:
                st.balloons()
                st.success(f"🔥 **Hepsini tek potada erittik üstad!** Nur topu gibi birleşmiş **{custom_merge_name}** hazır! ({elapsed:.1f} sn sürdü)")
                st.download_button(
                    label=f"📥 {custom_merge_name} Dosyasını İndir",
                    data=merged_bytes,
                    file_name=custom_merge_name,
                    mime="application/pdf",
                    use_container_width=True
                )

# =========================================================
# SEKME 3: Farklı Formatları PDF'e Çevirme (Görsel & Word -> PDF)
# =========================================================
with tab_to_pdf:
    st.markdown("""
    <div class="rapid-card">
        <h3>🖼️ Farklı Formatları Tek Tıkla PDF Yap</h3>
        <p>İster fotoğrafları (JPG, PNG) sırayla albüm gibi PDF'e çevir, ister Word (.docx) belgelerini şık bir PDF'e dönüştür üstad!</p>
    </div>
    """, unsafe_allow_html=True)

    convert_choice = st.radio(
        "Hangi kaynaktan PDF üretmek istersin üstad?",
        ["📸 Görsellerden PDF'e (JPG / PNG / WEBP)", "📝 Word (.docx) Belgesinden PDF'e"],
        horizontal=True
    )

    if "Görsellerden" in convert_choice:
        uploaded_images = st.file_uploader(
            "PDF'e dönüştürmek istediğin fotoğrafları seç üstad:",
            type=["jpg", "jpeg", "png", "webp"],
            accept_multiple_files=True,
            key="images_uploader"
        )

        if uploaded_images:
            st.info(f"🖼️ Toplam **{len(uploaded_images)}** adet görsel seçildi. Sayfa sırası yüklediğin sıraya göre ayarlanacaktır.")
            img_out_name = st.text_input("Çıktı PDF dosyasının adı:", value="rapid_fotograflar.pdf")

            if st.button("✨ Görselleri PDF'e Dönüştür Üstad!", key="btn_convert_images", use_container_width=True):
                if not img_out_name.lower().endswith(".pdf"):
                    img_out_name += ".pdf"

                with st.spinner("⚔️ Çöp adamlar fotoğrafları sayfalara inci gibi yerleştiriyor..."):
                    img_pdf_bytes, err = convert_images_to_pdf_bytes(uploaded_images)

                if err:
                    st.error(f"😅 **Eyvah kral! Görselleri işlerken bir aksilik oldu:** `{err}`")
                else:
                    st.balloons()
                    st.success(f"🎉 **Görseller inci gibi dizildi üstad!** PDF dosyan fırından sıcak sıcak çıktı!")
                    st.download_button(
                        label=f"📥 {img_out_name} Belgesini İndir",
                        data=img_pdf_bytes,
                        file_name=img_out_name,
                        mime="application/pdf",
                        use_container_width=True
                    )

    else:
        # Word -> PDF
        uploaded_docx = st.file_uploader(
            "PDF'e dönüştürmek istediğin Word (.docx) belgesini seç kral:",
            type=["docx"],
            key="docx2pdf_uploader"
        )

        if uploaded_docx:
            st.info(f"📄 **Seçilen Belge:** `{uploaded_docx.name}` ({format_bytes(len(uploaded_docx.getvalue()))})")
            default_pdf_name = os.path.splitext(uploaded_docx.name)[0] + "_rapid.pdf"

            if st.button("🚀 Word'ü PDF'e Çevir Üstad!", key="btn_convert_docx", use_container_width=True):
                with st.spinner("⚔️ Çöp adamlar Word paragraflarını PDF mühürleriyle damgalıyor..."):
                    start_time = time.time()
                    pdf_data, err = convert_docx_to_pdf_bytes(uploaded_docx.getvalue())
                    elapsed = time.time() - start_time

                if err:
                    st.error(f"😅 **Eyvah üstadım! Word belgesi okunurken bir hata oluştu:** `{err}`")
                else:
                    st.balloons()
                    st.success(f"🎉 **Üstad PDF hazır!** Word belgen **{elapsed:.1f} saniyede** yüksek çözünürlüklü PDF formatına dönüştürüldü!")
                    st.download_button(
                        label=f"📥 {default_pdf_name} Belgesini İndir",
                        data=pdf_data,
                        file_name=default_pdf_name,
                        mime="application/pdf",
                        use_container_width=True
                    )

# ---------------------------------------------------------
# Alt Bilgi (Footer)
# ---------------------------------------------------------
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(f"""
<div style="text-align: center; color: {text_secondary}; font-size: 0.9rem; padding: 20px 0; border-top: 1px solid rgba(168, 85, 247, 0.2);">
    ⚔️ <b>RapidConverter</b> &bull; Egzotik Lila Temalı Güçlü PDF Asistanı &bull; Kılıç darbeleriyle ışık hızında dönüştürme.<br>
    <span style="font-size: 0.8rem; opacity: 0.8;">Made with ❤️ for Render & Web &bull; Çöp adamlar görev başında!</span>
</div>
""", unsafe_allow_html=True)
