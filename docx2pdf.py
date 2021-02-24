import comtypes.client
import pathlib
import sys,glob
import PyPDF2
from multiprocessing import Pool


pp = pathlib.Path(sys.argv[1])
docx_path = pp.absolute()

def render(d, p):
    word = comtypes.client.CreateObject("Word.Application")
    word.Visible = 0
    ow = word.Documents.Open(d)
    ow.SaveAs(p, FileFormat=17)
    ow.close()
    pdf_reader=PyPDF2.PdfFileReader(p)
    pdf_writer=PyPDF2.PdfFileWriter()
    for pageNum in range(pdf_reader.numPages):
        pdf_writer.addPage(pdf_reader.getPage(pageNum))
    pdf_writer.encrypt(user_pwd="",owner_pwd="1")
    with open(p, 'wb') as outPDF:
        pdf_writer.write(outPDF)

if __name__ == "__main__":
    p = Pool(1)
    for f in glob.glob(str(docx_path / "*.docx")):
        if f.find('~') == -1:
            filename = f.split('.')[0]
            print(filename)
            docx = str(docx_path / f'{filename}.docx')
            pdf = str(docx_path / f'{filename}.pdf')
            p.apply_async(render, (docx, pdf))
    p.close()
    p.join()