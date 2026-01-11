import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
from docx import Document
from docx2pdf import convert
import win32com.client as win32
import os
import pythoncom

class MailGondericiUygulamasi:
    def __init__(self, root):
        self.root = root
        self.root.title("Toplu Mail ve Belge Gönderici (Full Versiyon)")
        self.root.geometry("650x750") # Pencere boyunu biraz uzattık

        # Değişkenler
        self.word_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.konu_basligi = tk.StringVar()
        self.dosya_eki = tk.StringVar() # Yeni eklediğimiz değişken
        
        self.arayuz_olustur()

    def arayuz_olustur(self):
        # 1. Word Dosyası Seçimi
        tk.Label(self.root, text="Word Şablonunu Seç:", font=("Arial", 10, "bold")).pack(pady=5)
        tk.Entry(self.root, textvariable=self.word_path, width=60).pack(pady=2)
        tk.Button(self.root, text="Gözat", command=self.word_sec).pack(pady=2)

        # 2. Excel Dosyası Seçimi
        tk.Label(self.root, text="Excel Listesini Seç:", font=("Arial", 10, "bold")).pack(pady=5)
        tk.Entry(self.root, textvariable=self.excel_path, width=60).pack(pady=2)
        tk.Button(self.root, text="Gözat", command=self.excel_sec).pack(pady=2)

        # 3. Mail Başlığı
        tk.Label(self.root, text="Mail Konu Başlığı:", font=("Arial", 10, "bold")).pack(pady=5)
        tk.Entry(self.root, textvariable=self.konu_basligi, width=60).pack(pady=2)

        # --- YENİ EKLENEN KISIM: DOSYA İSMİ EKİ ---
        tk.Label(self.root, text="PDF Dosya İsmi Eki (Örn: Mutabakat, Sertifika):", font=("Arial", 10, "bold")).pack(pady=5)
        tk.Entry(self.root, textvariable=self.dosya_eki, width=60).pack(pady=2)
        tk.Label(self.root, text="* Boş bırakırsanız sadece isim olur (Örn: Ahmet Yılmaz.pdf)", font=("Arial", 8, "italic"), fg="gray").pack()
        # ------------------------------------------
        
        # 4. Mail Gövdesi
        tk.Label(self.root, text="Mail İçeriği (Gövde):", font=("Arial", 10, "bold")).pack(pady=5)
        self.text_area_body = tk.Text(self.root, height=8, width=60)
        self.text_area_body.pack(pady=2)

        # 5. Başlat Butonu
        tk.Button(self.root, text="İŞLEMİ BAŞLAT", command=self.islemi_baslat, bg="green", fg="white", font=("Arial", 12, "bold")).pack(pady=20)

        # 6. Log Ekranı
        tk.Label(self.root, text="İşlem Durumu:", font=("Arial", 10)).pack(pady=2)
        self.log_area = scrolledtext.ScrolledText(self.root, height=10, width=75, state='disabled')
        self.log_area.pack(pady=5)

    def log_yaz(self, mesaj):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, mesaj + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')
        self.root.update()

    def word_sec(self):
        dosya = filedialog.askopenfilename(filetypes=[("Word Dosyaları", "*.docx")])
        if dosya: self.word_path.set(dosya)

    def excel_sec(self):
        dosya = filedialog.askopenfilename(filetypes=[("Excel Dosyaları", "*.xlsx")])
        if dosya: self.excel_path.set(dosya)

    def word_degistir(self, doc_obj, eski_metin, yeni_metin):
        for p in doc_obj.paragraphs:
            if eski_metin in p.text:
                p.text = p.text.replace(eski_metin, yeni_metin)
        for table in doc_obj.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if eski_metin in p.text:
                            p.text = p.text.replace(eski_metin, yeni_metin)

    def islemi_baslat(self):
        word_dosyasi = self.word_path.get()
        excel_dosyasi = self.excel_path.get()
        konu = self.konu_basligi.get()
        ek_isim = self.dosya_eki.get().strip() # Kullanıcının girdiği dosya ekini al
        govde_metni = self.text_area_body.get("1.0", tk.END)

        if not word_dosyasi or not excel_dosyasi:
            messagebox.showerror("Hata", "Lütfen Word ve Excel dosyalarını seçiniz.")
            return

        try:
            df = pd.read_excel(excel_dosyasi)
            df.columns = [c.strip().upper() for c in df.columns]

            if 'ISIM' not in df.columns or 'MAIL' not in df.columns:
                messagebox.showerror("Hata", "Excel dosyasında 'ISIM' ve 'MAIL' sütunları bulunamadı.")
                return
            
            pythoncom.CoInitialize() 
            outlook = win32.Dispatch('outlook.application')

            self.log_yaz("--- İşlem Başlıyor ---")

            for index, row in df.iterrows():
                kisi_ismi = str(row['ISIM'])
                kisi_mail = str(row['MAIL'])
                
                kisi_cc = ""
                if 'CC' in df.columns:
                    val = row['CC']
                    if not pd.isna(val):
                        kisi_cc = str(val)

                self.log_yaz(f"{kisi_ismi} için işlem yapılıyor...")

                # 1. Word Şablonunu Aç ve Düzenle
                doc = Document(word_dosyasi)
                self.word_degistir(doc, "{ISIM}", kisi_ismi)

                # --- DOSYA İSMİ BELİRLEME ---
                if ek_isim:
                    pdf_adi = f"{kisi_ismi} {ek_isim}.pdf"
                else:
                    pdf_adi = f"{kisi_ismi}.pdf"
                # ----------------------------

                temp_docx = os.path.abspath(f"temp_{index}.docx")
                temp_pdf = os.path.abspath(pdf_adi)

                doc.save(temp_docx)
                convert(temp_docx, temp_pdf)

                # 2. Mail Hazırla
                mail = outlook.CreateItem(0)
                mail.To = kisi_mail
                
                if kisi_cc:
                    mail.CC = kisi_cc
                
                guncel_konu = konu.replace("{ISIM}", kisi_ismi)
                mail.Subject = guncel_konu
                
                mail.Display() 
                imzali_govde = mail.HTMLBody

                guncel_govde = govde_metni.replace("{ISIM}", kisi_ismi)
                guncel_govde_html = guncel_govde.replace("\n", "<br>")

                yeni_html_icerik = f"<div style='font-family: Calibri, Arial; font-size: 11pt;'>{guncel_govde_html}</div><br>" + imzali_govde
                
                mail.HTMLBody = yeni_html_icerik
                
                # PDF Ekle
                mail.Attachments.Add(temp_pdf)

                mail.Send()
                
                log_mesaji = f"-> Mail gönderildi: {kisi_mail} (Dosya: {pdf_adi})"
                if kisi_cc:
                    log_mesaji += f" (CC: {kisi_cc})"
                self.log_yaz(log_mesaji)

                try:
                    os.remove(temp_docx)
                    os.remove(temp_pdf) 
                except:
                    pass

            self.log_yaz("--- Tümü Başarıyla Tamamlandı! ---")
            messagebox.showinfo("Başarılı", "Tüm mailler gönderildi.")

        except Exception as e:
            self.log_yaz(f"HATA OLUŞTU: {str(e)}")
            messagebox.showerror("Hata", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = MailGondericiUygulamasi(root)
    root.mainloop()