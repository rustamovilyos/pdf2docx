import io
import sys
import threading
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import PyPDF2
from typing import Tuple
from tkinter import messagebox as mbox

from pdf2docx import parse
from tqdm import tqdm


class Interface(Frame):

    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.upload()

    def upload(self):
        btn_open = ttk.Button(text="Выбрать файл", command=self.open_file)
        btn_open.pack(expand=True, padx=15, pady=15)

        self.progress = ttk.Progressbar(self.master, orient="horizontal", length=200, mode="determinate")
        self.progress.pack(pady=10)

        self.text_output = Text(self.master, wrap=WORD, state=DISABLED)
        self.text_output.pack(expand=True, fill=BOTH, padx=10, pady=10)

    def get_num_pages(self, input_file: str) -> int:
        """Get number of pages in the PDF"""
        with open(input_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            return len(pdf_reader.pages)

    def convert_pdf2docx(self, input_file: str, output_file: str, pages: Tuple = None):
        """Converts pdf to docx"""
        try:
            num_pages = self.get_num_pages(input_file)
            self.progress["maximum"] = num_pages
            self.progress["value"] = 0
            # Redirecting print to a StringIO object
            output_buffer = io.StringIO()
            sys.stdout = output_buffer

            if pages:
                pages = [int(i) for i in list(pages) if i.isnumeric()]
            result = parse(pdf_file=input_file,
                           docx_with_path=output_file, pages=pages)
            try:
                summary = {
                    "Файл": input_file, "Страницы": num_pages, "Выходной файл": output_file
                }
            except FileNotFoundError:
                self.on_error_file_not_found_error()

            # Restoring sys.stdout
            try:
                sys.stdout = sys.__stdout__

                # Printing Summary
                output_text = output_buffer.getvalue()
                self.text_output.config(state=NORMAL)
                self.text_output.delete(1.0, END)
                self.text_output.insert(END, "\n".join(f"{i}:{j}" for i, j in summary.items()))
                self.text_output.insert(END, output_text)
                self.text_output.config(state=DISABLED)

                # Close the StringIO object
                output_buffer.close()

                # Perform the actual conversion
                with tqdm(total=num_pages, desc="Конвертация", unit="стр.", leave=False) as pbar:
                    for i in range(num_pages):
                        # Simulate some processing time for each page (you can replace this with actual conversion code)
                        import time
                        time.sleep(0.1)
                        pbar.update(1)  # Update the progress bar
                        self.progress["value"] = i + 1
                        self.progress.update()

                return self.on_info()
            except UnboundLocalError:
                self.on_error_unbound_local_error()

        except KeyboardInterrupt:
            return self.on_error()

    def on_error(self):
        mbox.showerror("Ошибка!!!", "Не могу открыть файл!")

    def on_error_file_not_found_error(self):
        mbox.showerror("Ошибка!!!", "Файл не выбран!")

    def on_error_unbound_local_error(self):
        mbox.showwarning("Внемание!", "Пожалуйста выбирайте .pdf файл!")

    def on_info(self):
        mbox.showinfo("Информация", "Конвертация завершено!")

    def open_file(self):
        """Открывает файл для редактирования"""
        filepath = askopenfilename(
            filetypes=[("Text Files", "*.pdf"), ("All files", "*")]
        )
        if filepath:
            t = threading.Thread(target=self.convert_pdf2docx, args=(filepath, f'{filepath}.docx'))
            t.start()


def main():
    window = Tk()
    user = Interface(master=window)
    window.title("Конвертер PDF на WORD")
    window.geometry('680x330+350+100')

    window.mainloop()


if __name__ == "__main__":
    main()
