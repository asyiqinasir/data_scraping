import os
import pdfplumber
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

COLUMNS = ['MP', 'Kod', 'Maksud Perbelanjaan', 'Jenis Perbelanjaan', 'Anggaran']

def extract_data_from_pdf(pdf_path, mp):
    data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    continue
                lines = text.split('\n')
                maksud_perbelanjaan = ''
                anggaran_section = False

                for line in lines:
                    line = line.strip()
                    words = line.split()
                    if not words:
                        continue

                    if "ANGGARAN PERBELANJAAN PEMBANGUNAN BAGI TAHUN" in line :
                        anggaran_section = True
                        continue
                    if "JUMLAH ANGGARAN PERBELANJAAN" in line:
                        anggaran_section = False
                        continue

                    if anggaran_section:
                        logging.debug(f"Processing line in anggaran section: {line}")
                        if len(words) >= 2 and words[0].isdigit() and len(words[0]) == 5:
                            kod = words[0]
                            jenis_perbelanjaan = " ".join(words[1:-2])
                            anggaran = words[-1]
                            data.append({
                                "MP": mp,
                                "Kod": kod,
                                "Maksud Perbelanjaan": maksud_perbelanjaan,
                                "Jenis Perbelanjaan": jenis_perbelanjaan,
                                "Anggaran": anggaran
                            })

                    if line.startswith("Maksud Bekalan/Pembangunan"):
                        maksud_perbelanjaan = line.split(" - ", 1)[1] if " - " in line else line

        logging.info(f"Extracted data from {pdf_path}")
    except Exception as e:
        logging.error(f"Failed to extract data from {pdf_path}: {e}")
        raise
    return data

def save_to_excel(data, output_path):
    try:
        if not data:  # ensure there's data to save
            logging.warning(f"No data to save for {output_path}")
            return

        df = pd.DataFrame(data, columns=COLUMNS)
        df.to_excel(output_path, index=False)
        logging.info(f"Saved data to {output_path}")
    except Exception as e:
        logging.error(f"Failed to save data to {output_path}: {e}")
        raise

def main():
    pdf_dir = 'pdf-to-ocr/2013/3-OCR-pdf'  # update with your directory
    output_dir = 'pdf-to-ocr/2013/output/DE-oa'  # your output directory

    os.makedirs(output_dir, exist_ok=True)

    try:
        for pdf_name in os.listdir(pdf_dir):
            pdf_path = os.path.join(pdf_dir, pdf_name)
            output_path = os.path.join(output_dir, f'{os.path.splitext(pdf_name)[0]}_EXTRACTED_DE_objek_am.xlsx')
            mp = os.path.splitext(pdf_name)[0].replace('.', '').upper()

            if os.path.exists(pdf_path):
                extracted_data = extract_data_from_pdf(pdf_path, mp)
                if extracted_data:  # save if there's data
                    save_to_excel(extracted_data, output_path)
            else:
                logging.error(f"File does not exist: {pdf_path}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
