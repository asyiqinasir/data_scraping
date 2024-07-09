import os
import pdfplumber
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

COLUMNS = [
    'MP', 'Maksud Perbelanjaan', 'Kod Program', 'Program',
    'Kod Aktiviti', 'Aktiviti', 'Objek Am', 'Anggaran 2014',
    'Anggaran 2015', 'Bil. Jawatan 2014', 'Bil. Jawatan 2015'
]

def extract_data_from_pdf(pdf_path, mp):
    data = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text is None:
                    continue
                lines = text.split('\n')
                current_program = ''
                current_kod_program = ''
                current_kod_aktiviti = ''
                maksud_perbelanjaan = ''
                current_aktiviti = ''
                skip_section = False

                for line in lines:
                    line = line.strip()
                    words = line.split()
                    if not words:
                        continue

                    if "RINGKASAN ANGGARAN PERBELANJAAN" in line:
                        skip_section = True
                    if "JUMLAH ANGGARAN PERBELANJAAN" in line:
                        skip_section = False
                        continue
                    if skip_section:
                        continue

                    if line.startswith("Maksud"):
                        maksud_perbelanjaan = line.split(" - ", 1)[1] if " - " in line else line

                    # Extract Program
                    if line.isupper() and len(words) >= 2 and words[0].isdigit() and len(words[0]) == 6:
                        current_kod_program = words[0]
                        program_part = " ".join(words[1:])
                        potential_program = program_part.split(".")[0].strip()
                        # Only take characters before the first digit
                        current_program = ''.join(c for c in potential_program if not c.isdigit()).strip()
                        current_program = current_program[:50]  # Ensure max 50 characters
                        current_kod_aktiviti = ''
                        current_aktiviti = ''

                    # Extract Activity
                    elif len(words) >= 2 and words[0].isdigit() and len(words[0]) == 6:
                        current_kod_aktiviti = words[0]
                        activity_part = " ".join(words[1:])
                        potential_aktiviti = activity_part.split(".")[0].strip()
                        if potential_aktiviti.isdigit():
                            continue
                        # Only take characters before the first digit
                        current_aktiviti = ''.join(c for c in potential_aktiviti if not c.isdigit()).strip()
                        current_aktiviti = current_aktiviti[:50]  # Ensure max 50 characters

                    elif len(words) > 4 and words[0] in ["10000", "20000", "30000", "40000", "50000"]:
                        objek_am = words[0]
                        anggaran_2014 = words[-4]
                        anggaran_2015 = words[-3]
                        bil_jawatan_2014 = words[-2]
                        bil_jawatan_2015 = words[-1]

                        data.append({
                            "MP": mp,
                            "Maksud Perbelanjaan": maksud_perbelanjaan,
                            "Kod Program": current_kod_program,
                            "Program": current_program,
                            "Kod Aktiviti": current_kod_aktiviti,
                            "Aktiviti": current_aktiviti,
                            "Objek Am": objek_am,
                            "Anggaran 2014": anggaran_2014,
                            "Anggaran 2015": anggaran_2015,
                            "Bil. Jawatan 2014": bil_jawatan_2014,
                            "Bil. Jawatan 2015": bil_jawatan_2015
                        })
                    elif len(words) > 2 and words[0] in ["10000", "20000", "30000", "40000", "50000"]:  # In case bil_jawatan is not present
                        objek_am = words[0]
                        anggaran_2014 = words[-2]
                        anggaran_2015 = words[-1]
                        bil_jawatan_2014 = bil_jawatan_2015 = "0"

                        data.append({
                            "MP": mp,
                            "Maksud Perbelanjaan": maksud_perbelanjaan,
                            "Kod Program": current_kod_program,
                            "Program": current_program,
                            "Kod Aktiviti": current_kod_aktiviti,
                            "Aktiviti": current_aktiviti,
                            "Objek Am": objek_am,
                            "Anggaran 2014": anggaran_2014,
                            "Anggaran 2015": anggaran_2015,
                            "Bil. Jawatan 2014": bil_jawatan_2014,
                            "Bil. Jawatan 2015": bil_jawatan_2015
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
<<<<<<< HEAD
    pdf_dir = 'pdf-to-ocr/2015/3-OCR-pdf' 
    output_dir = 'pdf-to-ocr/2015/output' # your output directory
=======
    pdf_dir = 'input' # your input directory
    output_dir = 'output' # your output directory
>>>>>>> 18936af74c53a09065d0170cde48b72f8039eb72

    os.makedirs(output_dir, exist_ok=True)

    try:
        for pdf_name in os.listdir(pdf_dir):
            pdf_path = os.path.join(pdf_dir, pdf_name)
            output_path = os.path.join(output_dir, f'{os.path.splitext(pdf_name)[0]}_EXTRACTED.xlsx')
            mp = os.path.splitext(pdf_name)[0].replace('.', '').upper()

            if os.path.exists(pdf_path):
                extracted_data = extract_data_from_pdf(pdf_path, mp)
                if extracted_data:  #save if there's data
                    save_to_excel(extracted_data, output_path)
            else:
                logging.error(f"File does not exist: {pdf_path}")

    except Exception as e:
        logging.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
