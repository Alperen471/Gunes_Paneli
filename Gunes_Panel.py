import pandas as pd
import json
import time
import os
import sys
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- AYARLAR ---
KLASOR_ADI = "Growatt_Gunluk_Veriler"
ISTENMEYEN_SERI = "Photovoltaic Output" 

def growatt_v16_acik_kal():
    if not os.path.exists(KLASOR_ADI):
        os.makedirs(KLASOR_ADI)
    
    print(f" Tarayıcı başlatılıyor...")
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    
    options.add_experimental_option("detach", True)
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        driver.get("https://server.growatt.com/login")
        
        print("\n" + "="*70)
        print(" BEKLEME MODU")
        print("1. Giriş yapın.")
        print("2. İstediğiniz sayfaya gidin.")
        print("3. Tarihi seçin ve grafiğin yüklenmesini bekleyin.")
        print("👉 HER ŞEY HAZIRSA BURAYA GELİP 'ENTER' TUŞUNA BASIN.")
        print("="*70)
        
        input("Hazır mısınız? (ENTER): ")
        

        print(" Tarih okunuyor...")
        js_date_finder = r"""
        var inputs = document.getElementsByTagName('input');
        var foundDate = "";
        var dateRegex = /^\d{4}-\d{2}-\d{2}$/;
        for (var i = 0; i < inputs.length; i++) {
            var el = inputs[i];
            if (el.offsetParent !== null && el.value && el.value.match(dateRegex)) {
                foundDate = el.value;
                break;
            }
        }
        return foundDate;
        """
        algilanan_tarih = driver.execute_script(js_date_finder)
        
        if not algilanan_tarih:
            print(" Otomatik tarih okunamadı, manuel girin.")
            algilanan_tarih = input("Tarih (Yıl-Ay-Gün): ").strip()
        else:
            print(f" Tarih: {algilanan_tarih}")

        # --- 2. DOSYA KONTROLÜ ---
        dosya_adi = f"{algilanan_tarih}.xlsx"
        dosya_yolu = os.path.join(KLASOR_ADI, dosya_adi)

        if os.path.exists(dosya_yolu):
            print(f" '{dosya_adi}' zaten var.")
            secim = input("Üzerine yazılsın mı? (e/h): ").lower()
            if secim != 'e':
                print("İptal edildi.")
                return

        # --- 3. GRAFİK VERİSİNİ ÇEK ---
        print("Grafikler taranıyor...")
        js_code = r"""
        var results = [];
        try {
            // ECharts
            var allDivs = document.querySelectorAll('div[_echarts_instance_]');
            allDivs.forEach(function(div) {
                var instance = echarts.getInstanceByDom(div);
                if (instance) {
                    var options = instance.getOption();
                    if(options.series){
                        options.series.forEach(function(s){
                            if(s.data && s.data.length > 0){
                                var cleanData = s.data.map(d => {
                                    var t = Array.isArray(d) ? d[0] : '';
                                    var v = Array.isArray(d) ? d[1] : d;
                                    return {zaman: t, deger: v};
                                });
                                results.push({name: s.name || 'ECharts Data', data: cleanData});
                            }
                        });
                    }
                }
            });

            // Highcharts
            if (window.Highcharts && Highcharts.charts) {
                Highcharts.charts.forEach(function(chart) {
                    if(chart && chart.series){
                        var categories = chart.xAxis && chart.xAxis[0] ? chart.xAxis[0].categories : [];
                        chart.series.forEach(function(s){
                            if(s.visible && s.data && s.data.length > 0){
                                var cleanData = [];
                                s.data.forEach((point, index) => {
                                    var val = null;
                                    var timeLabel = "";
                                    if (typeof point === 'number') val = point;
                                    else if (Array.isArray(point)) val = point[1];
                                    else if (point && point.y !== undefined) val = point.y;
                                    
                                    if (categories && categories.length > index) timeLabel = categories[index];
                                    else if (point.name) timeLabel = point.name;
                                    else timeLabel = index;
                                    cleanData.push({zaman: timeLabel, deger: val});
                                });
                                results.push({name: s.name, data: cleanData});
                            }
                        });
                    }
                });
            }
        } catch(e) { return JSON.stringify([{error: e.message}]); }
        return JSON.stringify(results);
        """
        
        json_str = driver.execute_script(js_code)
        data_list = json.loads(json_str)
        
        if not data_list or (len(data_list) == 1 and 'error' in data_list[0]):
            print("Veri bulunamadı.")
            return

        # --- 4. VERİ İŞLEME VE TEMİZLİK ---
        all_rows = []
        for item in data_list:
            
            # İSTEMEDİĞİN SERİYİ ATLIYORUZ
            if item['name'] == ISTENMEYEN_SERI:
                continue 

            for v in item['data']:
                all_rows.append({
                    "Tarih": algilanan_tarih, 
                    "Saat": v['zaman'], 
                    "Değer": v['deger']
                })
        
        if all_rows:
            df = pd.DataFrame(all_rows)
            
            # Saat Filtresi (:00)
            try:
                df['Saat'] = df['Saat'].astype(str)
                if df['Saat'].str.contains(':').any():
                    df = df[df['Saat'].str.endswith(':00')] 
            except: pass 

            # Temizlik (Boşlar ve Sıfırlar)
            df = df.dropna(subset=['Değer']) 
            df['Değer'] = pd.to_numeric(df['Değer'], errors='coerce')
            df = df.dropna(subset=['Değer']) 
            df = df[df['Değer'] > 0] 

            if df.empty:
                print("Uyarı: Tüm veriler temizlendi (veri yok).")
                return

            try:
                df.to_excel(dosya_yolu, index=False)
                print("\n" + "*"*50)
                print(f"DOSYA HAZIR: {dosya_yolu}")
                print("*"*50)
                print(df.head())
            except PermissionError:
                print(f"Dosya açık! Kapatıp tekrar dene.")
        else:
            print("⚠️ Veri listesi boş.")

    except Exception as e:
        print(f"Hata: {e}")
    finally:

        print("İşlem tamamlandı.")
        # driver.quit() 

if __name__ == "__main__":
    growatt_v16_acik_kal()