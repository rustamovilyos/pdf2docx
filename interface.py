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


class Interface(Frame):

    def __init__(self):
        super().__init__()
        self.upload()

    def upload(self):
        btn_open = ttk.Button(text="Выбрать файл", command=self.open_file)

        btn_open.pack(expand=True, padx=15, pady=15)

    def show_conversion_animation(self):
        self.label_animation = Label(self, text="идет...")
        self.label_animation.pack(expand=True, padx=20, pady=5, fill=BOTH)

    def stop_conversion_animation(self):
        self.label_animation.pack_forget()

    def get_num_pages(self, input_file: str) -> int:
        """Get number of pages in the PDF"""
        with open(input_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            return len(pdf_reader.pages)

    def convert_pdf2docx(self, input_file: str, output_file: str, pages: Tuple = None):
        """Converts pdf to docx"""
        try:
            self.show_conversion_animation()
            # if self.show_conversion_animation():
            #     self.stop_conversion_animation()
            # else:
            #     self.show_conversion_animation()
            # Redirecting print to a StringIO object
            output_buffer = io.StringIO()
            sys.stdout = output_buffer

            if pages:
                pages = [int(i) for i in list(pages) if i.isnumeric()]
            result = parse(pdf_file=input_file,
                           docx_with_path=output_file, pages=pages)
            try:
                summary = {
                    "Файл": input_file, "Страницы": self.get_num_pages(input_file), "Выходной файл": output_file
                }
            except FileNotFoundError:
                self.onError_FileNotFoundError()

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

                return self.onInfo()
            except UnboundLocalError:
                self.onError_UnboundLocalError()

        except KeyboardInterrupt:
            return self.onError()

    def onError(self):
        mbox.showerror("Ошибка!!!", "Не могу открыть файл!")

    def onError_FileNotFoundError(self):
        mbox.showerror("Ошибка!!!", "Файл не выбран!")

    def onError_UnboundLocalError(self):
        mbox.showwarning("Внемание!", "Пожалуйста выбирайте .pdf файл!")

    def onInfo(self):
        mbox.showinfo("Информация", "Конвертация завершено!")

    def open_file(self):
        """Открывает файл для редактирования"""
        filepath = askopenfilename(
            filetypes=[("Text Files", "*.pdf"), ("All files", "*")]
        )
        t = threading.Thread(target=self.convert_pdf2docx, args=(filepath, f'{filepath}.docx'))
        t.start()


def main():
    window = Tk()
    user = Interface()
    window.title("Добро пожаловать в приложение")
    window.geometry('680x330+350+100')


    # Create a Text widget for text output
    user.text_output = Text(window, wrap=WORD)
    user.text_output.pack(expand=True, padx=20, pady=15, fill=BOTH)

    window.mainloop()


if __name__ == "__main__":
    main()
