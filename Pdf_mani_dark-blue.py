import random
import string
import tkinter
from tkinter import *
from customtkinter import *
from tkinter.filedialog import askdirectory, askopenfilename, askopenfilenames

from docx2pdf import convert
from PyPDF2 import PdfFileReader, PdfFileWriter, PdfMerger
import glob # to select all specific files in folder



# https://github.com/TomSchimansky/CustomTkinter
set_appearance_mode("light")  # Modes: system (default), light, dark
set_default_color_theme("dark-blue")  # Themes: blue (default), dark-blue, green



app = CTk()  # create CTk window like you do with the Tk window
app.attributes('-fullscreen', False)
app.state("zoom")


Heading = "PDF Manipulator"
radio_var = tkinter.IntVar(value=1)
filepath_list = []
savepath = ""
pages_range = []
pages = []
check = False

def change_mode():
    if switch_theme.get() == 1:
        set_appearance_mode("dark")
    else:
        set_appearance_mode("light")



def change_screen():
    if full_screen.get() == 1:
        app.attributes('-fullscreen', True)
    else:
        app.attributes('-fullscreen', False)

def random_string():
    chr_set = string.ascii_letters
    return ''.join(random.choice(chr_set) for i in range(3))


def clear_textbox():
    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    global savepath
    savepath = ""
    filepath_list.clear()
    pages.clear()
    pages_range.clear()
    global check
    check = False

def remove_textbox():
    for widget in frame_center.winfo_children():
        widget.destroy()
    
# **************** choosing different files and path *****************

def docx_chooser():
    path = askopenfilenames(title="Select docx files ",filetypes=(("DOCx Files","*.docx"),)) 
    for i in list(path):
        box.insert(END,"File path : "+i+"\n")
        filepath_list.append(i)
    

def pdf_chooser():
    path = askopenfilenames(title="Select PDF Files",filetypes=(("PDF Files","*.pdf"),))
    for i in list(path):
        box.insert(END,"File path : "+i+"\n")
        filepath_list.append(i)

def path_chooser():
    pathname = askdirectory(title="Select path to save output")
    global savepath
    savepath = pathname
    box.insert(END,"Path To save pdf : "+savepath+"\n")

# ********************** Repetitive code ******************************
def clear_all():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()

# ************************* Conversion ********************************
def convert_chosen_docx():
    if(len(filepath_list)==0):
        box.insert(END,"Please choose File First !!\n")
        return
    if(len(savepath)==0):
        box.insert(END,"Please choose destination directory first !!\n")
        return
    for i in filepath_list:
        convert(i,savepath)
    box.insert(END,"Conversion successful !!")

def convert_mass_docx():
    if(len(savepath)==0):
        box.insert(END,"Please provide path first !!\n")
        return
    convert(savepath)
    box.insert(END,"Conversion successful !!")


def coversion_options():
    clear_all()

    convert_single_btn = CTkButton(master=frame_center, text="Chosen Conversion", command=convert_chosen)
    convert_single_btn.place(relx = 0.4, rely=0.2)

    convert_mass_btn = CTkButton(master=frame_center, text="Mass Conversion", command=convert_mass)
    convert_mass_btn.place(relx = 0.4, rely=0.4)

def convert_mass():
    clear_all()
    
    choose_folder = CTkButton(master=frame_center, text="Mass Conversion path", command=path_chooser)
    choose_folder.place(relx = 0.4, rely=0.2)

    convert_mass_btn = CTkButton(master=frame_center, text="Convert All", command=convert_mass_docx,fg_color="#4bad6b",hover_color="#9dccac")
    convert_mass_btn.place(relx = 0.4, rely=0.4)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92) 


def convert_chosen():
    clear_all()

    file_chooser_btn = CTkButton(master=frame_center, text="Choose DOCx", command=docx_chooser)
    file_chooser_btn.place(relx = 0.4, rely=0.2)

    choose_path_btn = CTkButton(master=frame_center, text="Path to save", command=path_chooser)
    choose_path_btn.place(relx = 0.4, rely=0.3)

    convert_btn = CTkButton(master=frame_center, text="convert", command=convert_chosen_docx,fg_color="#4bad6b",hover_color="#9dccac")
    convert_btn.place(relx = 0.4, rely=0.45)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92)    


# ***************** Merging Pdfs *************************
def merge_chosen_pdfs():
    if(len(filepath_list)==0):
        box.insert(END,"Please choose File First !!\n")
        return
    if(len(savepath)==0):
        box.insert(END,"Please choose destination directory first !!\n")
        return
    merger = PdfMerger()
    for i in filepath_list:
        merger.append(i)
    merger.write(savepath+"\\"+random_string()+"Merged.pdf")
    merger.close()
    box.insert(END,"Merging successful !!")
    pass

def merge_mass_pdfs():
    if(len(savepath)==0):
        box.insert(END,"Please provide path first !!\n")
        return
    merger = PdfMerger()
    files = glob.glob(savepath+"\\*.pdf")
    for i in files:
        merger.append(i)
    merger.write(savepath+"\\"+random_string()+"Merged.pdf")
    merger.close()
    box.insert(END,"Mass Merging successful !!")
    pass

def merge_options():
    clear_all()

    merge_chosen_btn = CTkButton(master=frame_center, text="Chosen Merging", command=merge_chosen)
    merge_chosen_btn.place(relx = 0.4, rely=0.2)

    merge_mass_btn = CTkButton(master=frame_center, text="Mass Merging", command=merge_mass)
    merge_mass_btn.place(relx = 0.4, rely=0.4)



def merge_mass():
    clear_all()
    
    choose_folder = CTkButton(master=frame_center, text="Mass Merger path", command=path_chooser)
    choose_folder.place(relx = 0.4, rely=0.2)

    convert_mass_btn = CTkButton(master=frame_center, text="Merge All", command=merge_mass_pdfs,fg_color="#4bad6b",hover_color="#9dccac")
    convert_mass_btn.place(relx = 0.4, rely=0.4)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92) 


def merge_chosen():
    clear_all()
    
    file_chooser_btn = CTkButton(master=frame_center, text="Choose Multiple Pdfs", command=pdf_chooser)
    file_chooser_btn.place(relx = 0.4, rely=0.2)

    choose_path_btn = CTkButton(master=frame_center, text="Path to save", command=path_chooser)
    choose_path_btn.place(relx = 0.4, rely=0.3)

    file_chooser_btn = CTkButton(master=frame_center, text="Merge All", command=merge_chosen_pdfs,fg_color="#4bad6b",hover_color="#9dccac")
    file_chooser_btn.place(relx = 0.4, rely=0.45)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92) 

# **************** PDF split *****************************

def split_chosen():
    os.system('python Split.py')


def split():
    if(len(filepath_list)==0):
        box.insert(END,"Provide file first...\n")
        return
    if(len(savepath)==0):
        box.insert(END,"Provid path first...\n")
        return
    pdf_writer = PdfFileWriter()
    for i in range(len_pdf):
        if(i in pages):
            pdf_writer.addPage(pdf_reader.getPage(i))
    split_file = open(savepath+"\\splitByRange"+random_string()+".pdf",'wb')
    pdf_writer.write(split_file)
    box.insert(END,"Pdf split sucessfully!!")
    



def page_selected():
    if(len(filepath_list)==0):
        box.insert(END,"Provide path first...\n")
        return

    pdf_file = filepath_list[0]
    global pdf_reader
    pdf_reader = PdfFileReader(pdf_file)
    global len_pdf
    len_pdf = pdf_reader.numPages
    print(len_pdf)
    if(int(Ent.get())>len_pdf or int(Ent.get())<=0):
        box.insert(END,"Wrong input, Try again..")
        return
    pages_range.append(int(Ent.get())-1)
    box.insert(END,Ent.get(),"  ")
    Ent.delete(0,END)
    if(len(pages_range)==2):
        st = pages_range[0]
        ep = pages_range[1]
        if(st>ep):
            box.insert(END,"\nWrong input pages, Try again..")
            pages_range.clear()
        else:
            box.insert(END,"\nPages : ")
            for i in range(st,ep+1):
                pages.append(i)
                box.insert(END,str(i+1)+", ")
            box.insert(END,"\n")

def page_selected_for_Selected():
    if(len(filepath_list)==0):
        box.insert(END,"Provide path first...\n")
        return

    pdf_file = filepath_list[0]
    global pdf_reader
    pdf_reader = PdfFileReader(pdf_file)
    global len_pdf
    len_pdf = pdf_reader.numPages
    global check 
    if(check==False):
        box.insert(END,"total No. of pages : "+str(len_pdf)+"\n")
        box.insert(END,"x : Entered page is out of range..\n")
        check = True
    if(int(Ent.get())>len_pdf or int(Ent.get())<=0):
        box.insert(END,"x ")
        Ent.delete(0,END)
    else:
        pages.append(int(Ent.get())-1)
        box.insert(END,str(Ent.get())+" ")
        Ent.delete(0,END)


def page_deselect():
    Ent.delete(0,END)
    pages_range.clear()
    box.insert(END,"\nEnter pages again....")

def split_options():
    clear_all()

    selected_pages_btn = CTkButton(master=frame_center, text="Specific pages split", command=split_chosen)
    selected_pages_btn.place(relx = 0.4, rely=0.2)

    range_split_btn = CTkButton(master=frame_center, text="Range split", command=split_range)
    range_split_btn.place(relx = 0.4, rely=0.4)

def split_range():
    clear_all()
    
    choose_folder = CTkButton(master=frame_center, text="Select pdf", command=pdf_chooser)
    choose_folder.place(relx = 0.4, rely=0.2)

    global Ent
    Ent = Entry(master=frame_center,width=5)
    Ent.place(relx=0.4,rely=0.3)

    btn = tkinter.Button(master=frame_center, text="OK", command=page_selected)
    btn.place(relx = 0.5, rely=0.3)

    btn = tkinter.Button(master=frame_center, text="Clear", command=page_deselect)
    btn.place(relx = 0.6, rely=0.3)

    path_btn = CTkButton(master=frame_center, text="Save to", command=path_chooser)
    path_btn.place(relx = 0.4, rely=0.4)

    split_range_btn = CTkButton(master=frame_center, text="Split by range", command=split,fg_color="#4bad6b",hover_color="#9dccac")
    split_range_btn.place(relx = 0.4, rely=0.5)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92) 


#****************** Deletion ********************************

def delete_chosen():
    choose_folder = CTkButton(master=frame_center, text="Select pdf", command=pdf_chooser)
    choose_folder.place(relx = 0.4, rely=0.2)

    global Ent
    Ent = Entry(master=frame_center,width=5)
    Ent.place(relx=0.4,rely=0.3)

    btn = tkinter.Button(master=frame_center, text="OK", command=page_selected_for_Selected)
    btn.place(relx = 0.5, rely=0.3)

    btn = tkinter.Button(master=frame_center, text="Clear", command=page_deselect)
    btn.place(relx = 0.6, rely=0.3)

    path_btn = CTkButton(master=frame_center, text="Save to", command=path_chooser)
    path_btn.place(relx = 0.4, rely=0.4)

    split_range_btn = CTkButton(master=frame_center, text="Delete", command=delete,fg_color="#4bad6b",hover_color="#9dccac")
    split_range_btn.place(relx = 0.4, rely=0.5)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92)


def delete():
    if(len(filepath_list)==0):
        box.insert(END,"Provide file first...\n")
        return
    if(len(savepath)==0):
        box.insert(END,"Provid path first...\n")
        return
    pdf_writer = PdfFileWriter()
    for i in range(len_pdf):
        if(i not in pages):
            pdf_writer.addPage(pdf_reader.getPage(i))
    split_file = open(savepath+"\\del_chosen_"+random_string()+".pdf",'wb')
    pdf_writer.write(split_file)
    box.insert(END,"Pages deleted sucessfully!!")
    

def page_selected():
    if(len(filepath_list)==0):
        box.insert(END,"Provide path first...\n")
        return

    pdf_file = filepath_list[0]
    global pdf_reader
    pdf_reader = PdfFileReader(pdf_file)
    global len_pdf
    len_pdf = pdf_reader.numPages
    print(len_pdf)
    if(int(Ent.get())>len_pdf or int(Ent.get())<=0):
        box.insert(END,"Wrong input, Try again..")
        return
    pages_range.append(int(Ent.get())-1)
    box.insert(END,Ent.get(),"  ")
    Ent.delete(0,END)
    if(len(pages_range)==2):
        st = pages_range[0]
        ep = pages_range[1]
        if(st>ep):
            box.insert(END,"\nWrong input pages, Try again..")
            pages_range.clear()
        else:
            box.insert(END,"\nPages : ")
            for i in range(st,ep+1):
                pages.append(i)
                box.insert(END,str(i+1)+", ")
            box.insert(END,"\n")

    
def page_deselect():
    Ent.delete(0,END)
    pages_range.clear()
    box.insert(END,"\nEnter pages again....")

def delete_options():
    clear_all()

    del_selected_pages_btn = CTkButton(master=frame_center, text="delete Specific pages ", command=delete_chosen)
    del_selected_pages_btn.place(relx = 0.4, rely=0.2)

    del_range_btn = CTkButton(master=frame_center, text="Delete in Range ", command=delete_range)
    del_range_btn.place(relx = 0.4, rely=0.4)

def delete_range():
    clear_all()
    
    choose_folder = CTkButton(master=frame_center, text="Select pdf", command=pdf_chooser)
    choose_folder.place(relx = 0.4, rely=0.2)

    global Ent
    Ent = Entry(master=frame_center,width=5)
    Ent.place(relx=0.4,rely=0.3)

    btn = tkinter.Button(master=frame_center, text="OK", command=page_selected)
    btn.place(relx = 0.5, rely=0.3)

    btn = tkinter.Button(master=frame_center, text="Clear", command=page_deselect)
    btn.place(relx = 0.6, rely=0.3)

    path_btn = CTkButton(master=frame_center, text="Save to", command=path_chooser)
    path_btn.place(relx = 0.4, rely=0.4)

    split_range_btn = CTkButton(master=frame_center, text="Delete by range", command=delete,fg_color="#4bad6b",hover_color="#9dccac")
    split_range_btn.place(relx = 0.4, rely=0.5)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92) 

# ***************** About Section ***************************

def About():
    remove_textbox()
    
    lb = CTkLabel(master=frame_center,text="ABOUT",text_font=("Times New Roman", 22,"bold"))
    lb.place(relx = 0.35, rely= 0.05)

    lb1 = CTkLabel(master=frame_center,text="Project : PDF Manipulator \
            \n\nObjective : PDF Operations \
            \n\nDeveloper : Rajshekhar, Benjamin\
            \n\nTechnology : Python-Tkinter\
            "
            ,text_font=("Times New Roman", 15,"italic"))
    lb1.place(relx = 0.25, rely= 0.17)

    lb2 = CTkLabel(master=frame_center,text="Liabraries",text_font=("Times New Roman", 17,"bold"))
    lb2.place(relx = 0.35, rely= 0.56)

    lb3 = CTkLabel(master=frame_center,text="        PyPDF2 : Split and Delete \n\n\
        docx2PDF : Conversion \n\n\
        Customtkinter : User Interface",text_font=("Times New Roman", 15,"italic"))
    lb3.place(relx = 0.17, rely= 0.65)


# **************** Changing Colors ***************************

def change_color():
    x = radio_var.get()
    if(x==0):
        app.destroy()
        os.system('python Pdf_mani_green.py')
        
    elif(x==2):
        app.destroy()
        os.system('python Pdf_mani_blue.py')

    else:
        pass

# ********** Left Frame ********************

frame_left = CTkFrame(master=app)
frame_left.pack(side=LEFT,fill="y")

switch_theme = CTkSwitch(master=frame_left,text="     Dark Mode    ",command=change_mode)
switch_theme.place(relx=0.5, rely= 0.1, anchor=tkinter.CENTER)

full_screen = CTkSwitch(master=frame_left,text="     Full Screen Mode    ",command=change_screen)
full_screen.place(relx=0.58, rely= 0.15, anchor=tkinter.CENTER)

button = CTkButton(master=frame_left, text="Convert To PDF", command=coversion_options)
button.place(relx=0.5, rely=0.25, anchor=tkinter.CENTER)

button1 = CTkButton(master=frame_left, text="Merge PDF", command=merge_options)
button1.place(relx=0.5, rely=0.35, anchor=tkinter.CENTER)

button2 = CTkButton(master=frame_left, text="Delete PDF pages", command=delete_options)
button2.place(relx=0.5, rely=0.45, anchor=tkinter.CENTER)

button3 = CTkButton(master=frame_left, text="Split PDF", command=split_options)
button3.place(relx=0.5, rely=0.55, anchor=tkinter.CENTER)

button4 = CTkButton(master=frame_left, text="About", command=About)
button4.place(relx=0.5, rely=0.65, anchor=tkinter.CENTER)


label_radio_group = CTkLabel(master=frame_left,text=" Colors :",text_font=("Times New Roman", 15))
label_radio_group.place(relx=0.4, rely= 0.75, anchor=tkinter.CENTER)


radio_button_1 = CTkRadioButton(master=frame_left,text="green",variable=radio_var,value=0,command=change_color)
radio_button_1.place(relx=0.45, rely= 0.8, anchor=tkinter.CENTER)

radio_button_2 =CTkRadioButton(master=frame_left,text="dark-blue",variable=radio_var,value=1,command=change_color)
radio_button_2.place(relx=0.5, rely= 0.85, anchor=tkinter.CENTER)

radio_button_3 = CTkRadioButton(master=frame_left,text="blue",variable=radio_var,value=2,command=change_color)
radio_button_3.place(relx=0.43, rely= 0.90, anchor=tkinter.CENTER)


# ************* Title Frame *********************

title_frame = CTkFrame(master=app)
title_frame.pack(side=TOP,fill="x")

lb = CTkLabel(title_frame,text=f"{Heading}",text_font=("Times New Roman", 22))
lb.pack(pady=20)

# ************* Center Frame *********************
frame_center = CTkFrame(master=app,width=500,height=500)
frame_center.place(relx = 0.4, rely=0.2)

lb = CTkLabel(master=frame_center,text="ABOUT",text_font=("Times New Roman", 22,"bold"))
lb.place(relx = 0.35, rely= 0.05)

lb1 = CTkLabel(master=frame_center,text="Project : PDF Manipulator \
            \n\nObjective : PDF Operations \
            \n\nDeveloper : Rajshekhar, Benjamin\
            \n\nTechnology : Python-Tkinter\
            "
            ,text_font=("Times New Roman", 15,"italic"))
lb1.place(relx = 0.25, rely= 0.17)

lb2 = CTkLabel(master=frame_center,text="Liabraries",text_font=("Times New Roman", 17,"bold"))
lb2.place(relx = 0.35, rely= 0.56)

lb3 = CTkLabel(master=frame_center,text="        PyPDF2 : Split and Delete \n\n\
        docx2PDF : Conversion \n\n\
        Customtkinter : User Interface",text_font=("Times New Roman", 15,"italic"))
lb3.place(relx = 0.17, rely= 0.65)

app.mainloop()
