from tkinter.filedialog import askdirectory, askopenfilename
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfReader,PdfWriter
import tkinter
from tkinter import *

# a = askopenfilename()
pdf_file = open(askopenfilename(title="Select PDF Files",filetypes=(("PDF Files","*.pdf"),)),'rb')
pdf_reader = PdfFileReader(pdf_file)
pdf_writer = PdfFileWriter()
len_pdf = pdf_reader.numPages
pages = []
def act():
    vr.destroy()
    for i,j in enumerate(pages):
        print(j.get())
        nm = j.get()
        if(nm==1):
            pdf_writer.addPage(pdf_reader.getPage(i))
        else:
            continue
    split_file = open(askdirectory()+"\\demo.pdf",'wb')
    pdf_writer.write(split_file)


vr = tkinter.Tk()
v1 = 0
v2 = 0
for i in range(len_pdf):
    pages.append(0)
    pages[i]=tkinter.IntVar()
    c = tkinter.Checkbutton(vr,text=f"{i+1}",variable=pages[i])
    c.grid(row=v1,column=v2)
    if(v2==10):
        v1=v1+1
        v2=0
    else:
        v2 = v2+1
        
b = tkinter.Button(vr,command=act,text= "Done !!")
b.grid(column=0,columnspan=2)
vr.mainloop()

