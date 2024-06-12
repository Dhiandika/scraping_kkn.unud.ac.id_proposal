import os
import requests
import threading
from PyPDF2 import PdfMerger, PdfReader
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementClickInterceptedException
from bs4 import BeautifulSoup
import csv
import time

# Fungsi untuk membuat folder jika belum ada
def make_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

# Fungsi untuk mengunduh file PDF dengan retry mechanism
def download_pdf_thread(url, path, retries=5):
    def download():
        attempt = 0
        while attempt < retries:
            try:
                response = requests.get(url, timeout=10)
                response.raise_for_status()  # Raise HTTPError for bad responses
                with open(path, 'wb') as file:
                    file.write(response.content)
                print(f"File {path} telah diunduh.")
                return
            except (requests.exceptions.RequestException, requests.exceptions.Timeout) as e:
                print(f"Error downloading {url}: {e}. Retrying ({attempt + 1}/{retries})...")
                attempt += 1
                time.sleep(2)  # Wait before retrying
        print(f"Failed to download {url} after {retries} retries.")
    
    thread = threading.Thread(target=download)
    thread.start()
    return thread

# Fungsi untuk memproses tabel di halaman saat ini
def process_table(soup, writer):
    table = soup.find('table', id='tabel-laporan-proposal')
    if not table:
        print("Tabel tidak ditemukan.")
        return 0

    # Ambil semua baris data dari tabel
    rows = table.find('tbody').find_all('tr')
    print(f"Jumlah baris yang ditemukan: {len(rows)}")

    for tr in rows:
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
        threads = []
        for file_link in files:
            file_url = file_link['href']
            file_name = file_url.split('/')[-1]
            pdf_path = os.path.join(folder_name, file_name)
            print(f"Mengunduh file: {file_name} dari {file_url}")
            thread = download_pdf_thread(file_url, pdf_path)
            threads.append(thread)
            file_list.append(pdf_path)

        # Tunggu semua thread selesai
        for thread in threads:
            thread.join()

        # Cek apakah file PDF valid dan tidak kosong sebelum mencoba menambahkannya
        for pdf_path in file_list:
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                try:
                    with open(pdf_path, 'rb') as f:
                        reader = PdfReader(f)
                        for page_num in range(len(reader.pages)):
                            page = reader.pages[page_num]
                            merger.add_page(page)
                except Exception as e:
                    print(f"Error processing PDF file {pdf_path}: {e}")
            else:
                print(f"Invalid or empty PDF file: {pdf_path}")

        with open(combined_pdf_path, 'wb') as f_out:
            merger.write(f_out)

        merger.close()

        # Tulis informasi ke file CSV setelah file PDF gabungan dibuat
        writer.writerow([no, periode, desa, combined_pdf_path])

        # Hapus file PDF individu setelah digabungkan
        for pdf_path in file_list:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
                print(f"File {pdf_path} telah dihapus.")

    return len(rows)

# Fungsi untuk menunggu spinner menghilang
def wait_for_spinner_to_disappear(wait):
    try:
        wait.until(EC.invisibility_of_element((By.ID, 'spinnerxxx')))
        print("Spinner hilang.")
    except TimeoutException:
        print("Spinner tidak menghilang dalam waktu yang ditentukan.")

# URL halaman yang ingin diambil datanya
url = "https://kkn.unud.ac.id/proposal"

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

# Buat file CSV dan tulis headernya
csv_file = 'download_log.csv'
csv_columns = ['No', 'Periode', 'Desa', 'File']
with open(csv_file, 'w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(csv_columns)

    # Ambil data dari semua halaman
    while True:
        # Tunggu hingga tabel dimuat
        wait.until(EC.presence_of_element_located((By.ID, 'tabel-laporan-proposal')))

        # Ambil konten HTML setelah halaman dimuat sepenuhnya
        html = driver.page_source

        # Parsing HTML dengan BeautifulSoup
        soup = BeautifulSoup(html, 'html.parser')

        # Proses tabel di halaman saat ini
        rows_processed = process_table(soup, writer)

        # Cek apakah ada halaman berikutnya
        next_button = soup.find('li', class_='next')
        if next_button and 'disabled' not in next_button['class']:
            print("Navigasi ke halaman berikutnya.")
            try:
                # Tunggu spinner menghilang
                wait_for_spinner_to_disappear(wait)

                # Scroll ke tampilan elemen
                next_link = driver.find_element(By.LINK_TEXT, '>')
                driver.execute_script("arguments[0].scrollIntoView();", next_link)

                # Klik tombol "Next" menggunakan JavaScript
                driver.execute_script("arguments[0].click();", next_link)
            except ElementClickInterceptedException as e:
                print(f"Error saat mengklik tombol 'Next': {e}")
                wait_for_spinner_to_disappear(wait)
                continue
        else:
            print("Tidak ada halaman berikutnya atau tombol 'Next' dinonaktifkan.")
            break

print("Unduhan selesai untuk semua data.")
driver.quit()
