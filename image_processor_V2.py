import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import pikepdf
import io
from concurrent.futures import ProcessPoolExecutor, as_completed
import threading
import logging
import tempfile
from PyPDF2 import PdfMerger

CANVAS_SIZES = {
    'CSAT (272 x 394 mm)': (272, 394),
    'A4 (210 x 297 mm)': (210, 297)
}
SUPPORTED_IMAGE_EXTENSIONS = ('.tiff', '.tif', '.jpg', '.jpeg', '.png')

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

processing = False

def create_pdf(image_data_list, canvas_option, pdf_progress_var, status_label):
    try:
        total_images = len(image_data_list)
        pdf_paths = [None] * total_images

        with ProcessPoolExecutor() as executor:
            futures = [executor.submit(create_pdf_page, img_data, i, canvas_option) for i, img_data in enumerate(image_data_list)]
            for i, future in enumerate(as_completed(futures)):
                pdf_path = future.result()
                pdf_paths[i] = pdf_path
                progress = (i + 1) / total_images * 100
                pdf_progress_var.set(progress)
                status_label.config(text=f"PDF 생성 중: {i+1}/{total_images}")
                root.update_idletasks()

        output_pdf_path = os.path.join(tempfile.gettempdir(), "combined_output.pdf")
        combine_pdfs(pdf_paths, output_pdf_path)
        logging.info(f"PDF 생성 완료: {output_pdf_path}")

        for pdf_path in pdf_paths:
            os.remove(pdf_path)

        return output_pdf_path

    except Exception as e:
        logging.error(f"PDF 생성 중 오류 발생: {e}")
        raise

def process_images(input_folder, canvas_option, convert_progress_var, pdf_progress_var, compress_progress_var, status_label):
    image_files = [f for f in os.listdir(input_folder) if f.lower().endswith(SUPPORTED_IMAGE_EXTENSIONS)]
    total_images = len(image_files)
    
    if total_images == 0:
        messagebox.showinfo("정보", "선택한 폴더에 이미지가 없습니다.")
        return
    
    image_data_list = [None] * total_images
    
    for i, img in enumerate(image_files):
        try:
            with open(os.path.join(input_folder, img), 'rb') as f:
                image_data_list[i] = f.read()
            progress = (i + 1) / total_images * 100
            convert_progress_var.set(progress)
            status_label.config(text=f"이미지 불러오는 중: {i+1}/{total_images}")
            root.update_idletasks()
        except Exception as e:
            logging.error(f"이미지 불러오는 중 오류 발생: {e}")
    
    pdf_name = os.path.basename(input_folder) + ".pdf"
    pdf_path = os.path.join(os.path.dirname(input_folder), pdf_name)
    
    try:
        temp_pdf_path = create_pdf(image_data_list, canvas_option, pdf_progress_var, status_label)
    except Exception as e:
        messagebox.showerror("오류", f"PDF 생성 실패: {e}")
        return
    
    try:
        compress_pdf(temp_pdf_path, pdf_path, compress_progress_var, status_label)
    except Exception as e:
        messagebox.showerror("오류", f"PDF 압축 실패: {e}")
        return
    
    messagebox.showinfo(
        "성공",
        f"{total_images}개의 이미지가 변환되어 '{pdf_name}' 이름의 PDF로 결합 및 압축되었습니다."
    )

def create_pdf_page(img_data, page_index, canvas_option):
    try:
        page_width_mm, page_height_mm = CANVAS_SIZES[canvas_option]
        if canvas_option == 'A4 (210 x 297 mm)':
            a4_aspect_ratio = page_height_mm / page_width_mm

        temp_pdf_path = os.path.join(tempfile.gettempdir(), f"temp_page_{page_index}.pdf")
        c = canvas.Canvas(temp_pdf_path, pagesize=(page_width_mm * mm, page_height_mm * mm))

        img = Image.open(io.BytesIO(img_data))
        img_reader = ImageReader(io.BytesIO(img_data))
        width, height = img.size
        aspect_ratio = height / width

        if canvas_option == 'A4 (210 x 297 mm)':
            if aspect_ratio > a4_aspect_ratio:
                target_height_mm = page_height_mm
                target_width_mm = target_height_mm / aspect_ratio
            else:
                target_width_mm = page_width_mm
                target_height_mm = target_width_mm * aspect_ratio
        else:
            target_width_mm = 235
            target_height_mm = target_width_mm * aspect_ratio

        paste_x = (page_width_mm * mm - target_width_mm * mm) / 2
        paste_y = (page_height_mm * mm - target_height_mm * mm) / 2

        c.drawImage(img_reader, paste_x, paste_y, width=target_width_mm * mm, height=target_height_mm * mm)
        c.save()

        return temp_pdf_path
    except Exception as e:
        logging.error(f"PDF 페이지 {page_index} 생성 중 오류 발생: {e}")
        raise

def combine_pdfs(pdf_paths, output_path):
    try:
        merger = PdfMerger()
        for pdf_path in sorted(pdf_paths, key=lambda x: int(os.path.splitext(os.path.basename(x))[0].split('_')[-1])):
            merger.append(pdf_path)
        merger.write(output_path)
        merger.close()
        logging.info(f"결합된 PDF 저장: {output_path}")
    except Exception as e:
        logging.error(f"PDF 결합 중 오류 발생: {e}")
        raise

def compress_pdf(input_pdf, output_pdf, compress_progress_var, status_label):
    try:
        status_label.config(text="PDF 압축 중...")
        root.update_idletasks()

        if not os.path.exists(input_pdf):
            raise FileNotFoundError(f"{input_pdf} 파일이 존재하지 않습니다.")

        with pikepdf.Pdf.open(input_pdf) as pdf:
            pdf.save(output_pdf, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.generate)
        
        os.remove(input_pdf)
        compress_progress_var.set(100)
        logging.info(f"PDF 압축 완료 및 원본 삭제: {input_pdf}")
    except Exception as e:
        logging.error(f"PDF 압축 중 오류 발생: {e}")
        raise

def select_input_folder():
    folder = filedialog.askdirectory()
    if folder:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, folder)

def start_processing():
    global processing
    if processing:
        messagebox.showinfo("정보", "처리가 이미 진행 중입니다.")
        return

    input_folder = input_entry.get()
    canvas_option = canvas_size_option.get()
    if not input_folder:
        messagebox.showerror("오류", "입력 폴더를 선택하세요.")
        return
    
    processing = True
    process_button.config(state=tk.DISABLED)
    convert_progress_var.set(0)
    pdf_progress_var.set(0)
    compress_progress_var.set(0)
    status_label.config(text="시작 중...")
    
    def process_thread():
        global processing
        try:
            process_images(input_folder, canvas_option, convert_progress_var, pdf_progress_var, compress_progress_var, status_label)
        except Exception as e:
            messagebox.showerror("오류", f"오류 발생: {e}")
        finally:
            convert_progress_var.set(0)
            pdf_progress_var.set(0)
            compress_progress_var.set(0)
            status_label.config(text="")
            process_button.config(state=tk.NORMAL)
            processing = False

    threading.Thread(target=process_thread, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("이미지 PDF 변환기")

    input_label = tk.Label(root, text="입력 폴더:")
    input_label.grid(row=0, column=0, padx=5, pady=5)
    input_entry = tk.Entry(root, width=50)
    input_entry.grid(row=0, column=1, padx=5, pady=5)
    input_button = tk.Button(root, text="찾아보기", command=select_input_folder)
    input_button.grid(row=0, column=2, padx=5, pady=5)

    canvas_size_option = tk.StringVar(value='CSAT (272 x 394 mm)')
    canvas_size_label = tk.Label(root, text="캔버스 크기 옵션:")
    canvas_size_label.grid(row=1, column=0, padx=5, pady=5)
    canvas_size_combobox = ttk.Combobox(root, textvariable=canvas_size_option, values=['CSAT (272 x 394 mm)', 'A4 (210 x 297 mm)'], state="readonly")
    canvas_size_combobox.grid(row=1, column=1, padx=5, pady=5)
    canvas_size_combobox.current(0)

    process_button = tk.Button(root, text="이미지 처리", command=start_processing)
    process_button.grid(row=2, column=1, padx=5, pady=10)

    convert_progress_var = tk.DoubleVar()
    convert_progress_bar = ttk.Progressbar(root, variable=convert_progress_var, maximum=100)
    convert_progress_bar.grid(row=3, column=0, columnspan=3, padx=5, pady=5, sticky="we")

    pdf_progress_var = tk.DoubleVar()
    pdf_progress_bar = ttk.Progressbar(root, variable=pdf_progress_var, maximum=100)
    pdf_progress_bar.grid(row=4, column=0, columnspan=3, padx=5, pady=5, sticky="we")

    compress_progress_var = tk.DoubleVar()
    compress_progress_bar = ttk.Progressbar(root, variable=compress_progress_var, maximum=100)
    compress_progress_bar.grid(row=5, column=0, columnspan=3, padx=5, pady=5, sticky="we")

    status_label = tk.Label(root, text="")
    status_label.grid(row=6, column=0, columnspan=3, padx=5, pady=5)

    root.mainloop()
