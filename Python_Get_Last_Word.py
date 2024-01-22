import openpyxl

def extract_and_write(file_path, sheet_name):
    # Türkiye illeri listesi
    cities = [
        "ADANA", "ADIYAMAN", "AFYONKARAHİSAR", "AĞRI", "AMASYA", "ANKARA", "ANTALYA", "ARTVİN", 
        "AYDIN", "BALIKESİR", "BİLECİKK", "BİNGÖL", "BİTLİS", "BOLU", "BURDUR", "BURSA", "ÇANAKKALE",
        "ÇANKIRI", "ÇORUM", "DENİZLİ", "DİYARBAKIR", "EDİRNE", "ELAZIĞ", "ERZİNCAN", "ERZURUM", 
        "ESKİŞEHİR", "GAZİANTEP", "GİRESUN", "GÜMÜŞHANE", "HAKKARİ", "HATAY", "ISPARTA", "MERSİN",
        "İSTANBUL", "İZMİR", "KARS", "KASTAMONU", "KAYSERİ", "KIRKLARELİ", "KIRŞEHİR", "KOCAELİ", 
        "KONYA", "KÜTAHYA", "MALATYA", "MANİSA", "KAHRAMANMARAŞ", "MARDİN", "MUĞLA", "MUŞ", 
        "NEVŞEHİR", "NİĞDE", "ORDU", "RİZE", "SAKARYA", "SAMSUN", "SİİRT", "SİNOP", "SİVAS", 
        "TEKİRDAĞ", "TOKAT", "TRABZON", "TUNCELİ", "ŞANLIURFA", "UŞAK", "VAN", "YOZGAT", "ZONGULDAK", 
        "AKSARAY", "BAYBURT", "KARAMAN", "KIRIKKALE", "BATMAN", "ŞIRNAK", "BARTIN", "ARDAHAN", 
        "IĞDIR", "YALOVA", "KARABÜK", "KİLİS", "OSMANİYE", "DÜZCE"
    ]

    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook(file_path)

    # Belirli bir sayfayı seç
    sheet = workbook[sheet_name]

    # 2. satırdan 1297. satıra kadar döngü
    for row in range(2, 1298):
        # E sütunundaki değeri al
        e_cell_value = sheet[f'E{row}'].value

        # Değer boş değilse işlem yap
        if e_cell_value:
            # Kelimelere ayır ve her birini kontrol et
            for word in e_cell_value.upper().split():
                clean_word = word.split('/')[-1].split('-')[-1]
                if clean_word in cities:
                    # Eşleşen kelimeyi D sütunundaki hücreye yaz
                    sheet[f'F{row}'] = clean_word
                    break  # Eşleşme bulunduğunda döngüden çık

    # Değişiklikleri kaydet
    workbook.save(file_path)

# Fonksiyonu çağır, dosya yolu ve sayfa adı ile birlikte
extract_and_write("C:\\Users\\furkan.cakir\\Desktop\\Kopya UKZ ŞUBELERİ - 60_sevk_adresleri (002).xlsx", "Sayfa2")
