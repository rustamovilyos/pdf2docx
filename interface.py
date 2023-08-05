import io
import sys
from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
import PyPDF2
from typing import Tuple
from tkinter import messagebox as mbox


class Interface(Frame):

    def __init__(self):
        super().__init__()
        self.upload()

    def open_file(self):
        """Открывает файл для редактирования"""
        filepath = askopenfilename(
            filetypes=[("Text Files", "*.pdf")]
        )
        print(filepath)
        return self.convert_pdf2docx(f'{filepath}', output_file=f'{filepath}.docx')

    def upload(self):
        btn_open = ttk.Button(text="Выбрать файл", command=self.open_file)

        btn_open.grid(row=0, column=0, sticky="ew", padx=100, pady=50)

    def get_num_pages(self, input_file: str) -> int:
        """Get number of pages in the PDF"""
        with open(input_file, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            return len(pdf_reader.pages)

    def convert_pdf2docx(self, input_file: str, output_file: str, pages: Tuple = None):
        """Converts pdf to docx"""
        # Redirecting print to a StringIO object
        output_buffer = io.StringIO()
        sys.stdout = output_buffer

        summary = {
            "Файл": input_file, "Страницы": self.get_num_pages(input_file), "Выходной файл": output_file
        }

        # Restoring sys.stdout
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

    def onError(self):
        mbox.showerror("Ошибка", "Не могу открыть файл")

    def onInfo(self):
        mbox.showinfo("Информация", "Конвертация завершено")


def main():
    window = Tk()
    user = Interface()
    window.title("Добро пожаловать в приложение")
    window.geometry('300x150+500+300')

    # Create a Toplevel window for text output
    text_window = Toplevel(window)
    text_window.title("Результаты конвертации")
    text_window.geometry("600x400+600+200")

    # Create a Text widget for text output
    user.text_output = Text(text_window, wrap=WORD, state=DISABLED)
    user.text_output.pack(expand=True, fill=BOTH)

    window.mainloop()


if __name__ == "__main__":
    main()
