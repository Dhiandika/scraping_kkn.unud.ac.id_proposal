import os
import requests
from PyPDF2 import PdfMerger
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import csv

# Fungsi untuk membuat folder jika belum ada
def make_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

# Fungsi untuk mengunduh file PDF
def download_pdf(url, path):
    response = requests.get(url)
    with open(path, 'wb') as file:
        file.write(response.content)

# URL halaman yang ingin diambil datanya
url = "https://kkn.unud.ac.id/proposal"

# Batas jumlah halaman yang akan diambil file-nya
page_limit = 14

# Set up Selenium WebDriver untuk Edge
service = EdgeService('msedgedriver.exe')  # Ganti dengan path ke msedgedriver.exe
options = webdriver.EdgeOptions()
driver = webdriver.Edge(service=service, options=options)
driver.get(url)

# Tunggu hingga dropdown jumlah catatan per halaman muncul
wait = WebDriverWait(driver, 20)  # Tingkatkan waktu tunggu hingga 20 detik
try:
    # Tunggu hingga elemen dropdown terlihat
    records_dropdown = wait.until(EC.visibility_of_element_located((By.CLASS_NAME, 'chosen-container')))
    print("Dropdown ditemukan, mencoba mengklik elemen.")

    # Klik dropdown untuk menampilkan pilihan
    records_dropdown.click()
    
    # Tunggu hingga opsi dropdown terlihat dan pilih opsi '100'
    option = wait.until(EC.visibility_of_element_located((By.XPATH, "//li[@class='active-result' and text()='100']")))
    option.click()
    print("Nilai dropdown berhasil diatur ke 100.")
except Exception as e:
    print(f"Exception: {e}")
    driver.quit()
    exit()

# Tunggu hingga tabel dimuat ulang
wait.until(EC.presence_of_element_located((By.ID, 'tabel-laporan-proposal')))

# Ambil konten HTML setelah halaman dimuat sepenuhnya
html = driver.page_source
driver.quit()

# Parsing HTML dengan BeautifulSoup
soup = BeautifulSoup(html, 'html.parser')

# Temukan tabel yang berisi data
table = soup.find('table', id='tabel-laporan-proposal')

if table:
    print("Tabel ditemukan.")
else:
    print("Tabel tidak ditemukan.")
    exit()

# Ambil semua baris data dari tabel
rows = table.find('tbody').find_all('tr')
print(f"Jumlah baris yang ditemukan: {len(rows)}")
page_count = 0

# Buat file CSV dan tulis headernya
csv_file = 'download_log.csv'
csv_columns = ['No', 'Periode', 'Desa', 'File']
with open(csv_file, 'w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(csv_columns)

for tr in rows:
    if page_count >= page_limit:
        break

    cells = tr.find_all('td')
    if len(cells) < 4:
        print(f"Baris tidak memiliki cukup sel: {tr}")
        continue

    no = cells[0].text.strip()
    periode = cells[1].text.strip()
    desa = cells[2].text.strip()
    files = cells[3].find_all('a')

    print(f"Mengolah baris nomor: {no}, Periode: {periode}, Desa: {desa}")

    # Buat nama folder berdasarkan periode dan desa
    folder_name = f"{periode}_{desa.replace(' ', '_').replace('(', '').replace(')', '').replace('-', '').replace(',', '')}"
    make_dir(folder_name)

    # Path untuk file PDF gabungan
    combined_pdf_path = os.path.join(folder_name, f"{folder_name}_combined.pdf")

    # Gabungkan semua file PDF dalam satu baris
    merger = PdfMerger()

    file_list = []
    for file_link in files:
        file_url = file_link['href']
        file_name = file_url.split('/')[-1]
        pdf_path = os.path.join(folder_name, file_name)
        print(f"Mengunduh file: {file_name} dari {file_url}")
        download_pdf(file_url, pdf_path)
        merger.append(pdf_path)
        file_list.append(pdf_path)

    merger.write(combined_pdf_path)
    merger.close()

    # Hapus file PDF individu setelah digabungkan
    for pdf_path in file_list:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
            print(f"File {pdf_path} telah dihapus.")

    # Tulis informasi ke file CSV
    with open(csv_file, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for f in file_list:
            writer.writerow([no, periode, desa, os.path.basename(f)])

    page_count += 1

print(f"Unduhan selesai untuk {page_count} halaman pertama.")
