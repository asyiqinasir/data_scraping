import os
import pdfplumber
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

COLUMNS = ['MP', 'Maksud Perbelanjaan', 'Butiran', 'Tajuk', 'Jumlah Anggaran Harga Projek', 'Cara Langsung', 'Pinjaman']

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
                pembangunan_section = False

                for line in lines:
                    line = line.strip()
                    words = line.split()
                    if not words:
                        continue

                    if "Maksud Pembangunan" in line:
                        pembangunan_section = True
                        maksud_perbelanjaan = line.split(" - ", 1)[1] if " - " in line else line
                        continue
                    if "JUMLAH PEMBANGUNAN" in line:
                        pembangunan_section = False
                        continue

                    if pembangunan_section:
                        logging.debug(f"Processing line in pembangunan section: {line}")
                        if len(words) >= 7 and words[0].isdigit() and len(words[0]) == 5:
                            butiran = words[0]
                            tajuk = " ".join(words[1:-5])
                            jumlah_anggaran_harga_projek = words[-6]
                            cara_langsung = words[-2]
                            pinjaman = words[-1]
                            data.append({
                                "MP": mp,
                                "Maksud Perbelanjaan": maksud_perbelanjaan,
                                "Butiran": butiran,
                                "Tajuk": tajuk,
                                "Jumlah Anggaran Harga Projek": jumlah_anggaran_harga_projek,
                                "Cara Langsung": cara_langsung,
                                "Pinjaman": pinjaman
                            })

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
    output_dir = 'pdf-to-ocr/2013/output/DE'  # your output directory

    os.makedirs(output_dir, exist_ok=True)

    try:
        for pdf_name in os.listdir(pdf_dir):
            pdf_path = os.path.join(pdf_dir, pdf_name)
            output_path = os.path.join(output_dir, f'{os.path.splitext(pdf_name)[0]}_EXTRACTED_DE.xlsx')
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
