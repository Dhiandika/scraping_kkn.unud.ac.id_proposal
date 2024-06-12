import os
import requests
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup

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
page_limit = 3

# Set up Selenium WebDriver untuk Edge
service = EdgeService('msedgedriver.exe')  # Ganti dengan path ke msedgedriver.exe
options = webdriver.EdgeOptions()
driver = webdriver.Edge(service=service, options=options)
driver.get(url)

# Tunggu hingga tabel dimuat
wait = WebDriverWait(driver, 10)
table = wait.until(EC.presence_of_element_located((By.ID, 'tabel-laporan-proposal')))

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

    # Unduh semua file PDF dalam satu baris
    for file_link in files:
        file_url = file_link['href']
        file_name = file_url.split('/')[-1]
        print(f"Mengunduh file: {file_name} dari {file_url}")
        download_pdf(file_url, os.path.join(folder_name, file_name))
    
    page_count += 1

print(f"Unduhan selesai untuk {page_count} halaman pertama.")
