from pdf2docx import parse
from typing import Tuple


def convert_pdf2docx(input_file: str, output_file: str, pages: Tuple = None):
    """Converts pdf to docx"""
    if pages:
        pages = [int(i) for i in list(pages) if i.isnumeric()]
    result = parse(pdf_file=input_file,
                   docx_with_path=output_file, pages=pages)
    summary = {
        "File": input_file, "Pages": str(pages), "Output File": output_file
    }
    # Printing Summary
    print("## Summary ########################################################")
    print("\n".join("{}:{}".format(i, j) for i, j in summary.items()))
    print("###################################################################")
    return result


if __name__ == "__main__":
    print(
        "Siz word faylda konvertatsiya qilmoqchi bo'lgan pdf faylingiz ushbu programma bilan "
        "bitta papkada turganligiga ishonch xosil qiling.\n"

        "===========================================================================================================\n"

        "Убедитесь, что файл PDF, который вы хотите преобразовать в файл Word, "
        "находится в одной папке с этой программой.\n"
        "==========================================================================================================="
    )
    convert_pdf2docx(input('Input filename with .pfd: '), input('\n Output filename with .docx: '))
