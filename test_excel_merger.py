import os
import sys
import pandas as pd
from functions.excel_bridge import ExcelBridge

def create_test_excel_files():
    """
    Test için örnek Excel dosyaları oluştur
    """
    # Test klasörü oluştur
    test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_files')
    os.makedirs(test_dir, exist_ok=True)
    
    # Test dosyası 1
    df1 = pd.DataFrame({
        'Ad': ['Ali', 'Ayşe', 'Mehmet'],
        'Soyad': ['Yılmaz', 'Kaya', 'Demir'],
        'Yaş': [25, 30, 35]
    })
    file1_path = os.path.join(test_dir, 'test1.xlsx')
    df1.to_excel(file1_path, sheet_name='Sayfa1', index=False)
    
    # Test dosyası 2
    df2 = pd.DataFrame({
        'Şehir': ['İstanbul', 'Ankara', 'İzmir'],
        'Nüfus': [15000000, 5000000, 4000000],
        'Bölge': ['Marmara', 'İç Anadolu', 'Ege']
    })
    file2_path = os.path.join(test_dir, 'test2.xlsx')
    with pd.ExcelWriter(file2_path) as writer:
        df2.to_excel(writer, sheet_name='Şehirler', index=False)
        
        # İkinci bir sayfa ekle
        df2_2 = pd.DataFrame({
            'Ülke': ['Türkiye', 'Almanya', 'Fransa'],
            'Başkent': ['Ankara', 'Berlin', 'Paris']
        })
        df2_2.to_excel(writer, sheet_name='Ülkeler', index=False)
    
    return [file1_path, file2_path]

def test_excel_merger():
    """
    Excel birleştirme fonksiyonunu test et
    """
    print("Excel birleştirme testi başlatılıyor...")
    
    # Test dosyalarını oluştur
    input_files = create_test_excel_files()
    print(f"Test dosyaları oluşturuldu: {input_files}")
    
    # Çıktı dosyası yolu
    output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'test_files', 'birlestirilmis.xlsx')
    
    # Excel dosyalarını birleştir
    success, message = ExcelBridge.merge_excel_files(input_files, output_file)
    
    # Sonucu göster
    if success:
        print(f"Test başarılı: {message}")
        print(f"Birleştirilmiş dosya: {output_file}")
        
        # Birleştirilmiş dosyayı kontrol et
        dfs = pd.read_excel(output_file, sheet_name=None)
        print(f"Birleştirilmiş dosyadaki sayfalar: {list(dfs.keys())}")
        
        for sheet_name, df in dfs.items():
            print(f"\nSayfa: {sheet_name}")
            print(df)
    else:
        print(f"Test başarısız: {message}")

if __name__ == "__main__":
    test_excel_merger()