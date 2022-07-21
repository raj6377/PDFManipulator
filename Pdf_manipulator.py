import random
import string
import tkinter
from tkinter import *
from customtkinter import *
from tkinter.filedialog import askdirectory, askopenfilenames

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
radio_var = tkinter.IntVar(value=0)
filepath_list = []
savepath = ""
box = CTkTextbox()

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


# ************************* Conversion ********************************
def convert_chosen_docx():
    for i in filepath_list:
        convert(i,savepath)
    box.insert(END,"Conversion successful !!")

def convert_mass_docx():
    convert(savepath)
    box.insert(END,"Conversion successful !!")


def coversion_options():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()

    convert_single_btn = CTkButton(master=frame_center, text="Chosen Conversion", command=convert_chosen)
    convert_single_btn.place(relx = 0.4, rely=0.2)

    convert_mass_btn = CTkButton(master=frame_center, text="Mass Conversion", command=convert_mass)
    convert_mass_btn.place(relx = 0.4, rely=0.4)

def convert_mass():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()
    
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
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()

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
    merger = PdfMerger()
    for i in filepath_list:
        merger.append(i)
    merger.write(savepath+"\\"+random_string()+"Merged.pdf")
    merger.close()
    box.insert(END,"Merging successful !!")
    pass

def merge_mass_pdfs():
    merger = PdfMerger()
    files = glob.glob(savepath+"\\*.pdf")
    for i in files:
        merger.append(i)
    merger.write(savepath+"\\"+random_string()+"Merged.pdf")
    merger.close()
    box.insert(END,"Mass Merging successful !!")
    pass

def merge_options():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()

    merge_chosen_btn = CTkButton(master=frame_center, text="Chosen Merging", command=merge_chosen)
    merge_chosen_btn.place(relx = 0.4, rely=0.2)

    merge_mass_btn = CTkButton(master=frame_center, text="Mass Merging", command=merge_mass)
    merge_mass_btn.place(relx = 0.4, rely=0.4)



def merge_mass():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()
    
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
    # deleting already existing widgets
    for widget in frame_center.winfo_children():
        widget.destroy()
    
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



def split_options():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()

    selected_pages_btn = CTkButton(master=frame_center, text="Specific pages split", command=split_chosen)
    selected_pages_btn.place(relx = 0.4, rely=0.2)

    range_split_btn = CTkButton(master=frame_center, text="Range split", command=split_range)
    range_split_btn.place(relx = 0.4, rely=0.4)

def split_range():
    # deleting already existing elements on center frame
    for widget in frame_center.winfo_children():
        widget.destroy()
    clear_textbox()
    
    choose_folder = CTkButton(master=frame_center, text="Mass Merger path", command=path_chooser)
    choose_folder.place(relx = 0.4, rely=0.2)

    convert_mass_btn = CTkButton(master=frame_center, text="Merge All", command=merge_mass_pdfs,fg_color="#4bad6b",hover_color="#9dccac")
    convert_mass_btn.place(relx = 0.4, rely=0.4)

    global box
    box= CTkTextbox(master=frame_center,height=90,width=400,relief=SUNKEN)
    box.place(relx = 0.5, rely=0.75,anchor=CENTER)

    clear_btn = CTkButton(master=frame_center, text="Clear", command=clear_textbox,fg_color="#f24e71",hover_color="#e88ea1")
    clear_btn.place(relx = 0.4, rely=0.92) 




# **************** Changing Colors ***************************

def change_color():
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

button2 = CTkButton(master=frame_left, text="Compress PDF", command=None)
button2.place(relx=0.5, rely=0.45, anchor=tkinter.CENTER)

button2 = CTkButton(master=frame_left, text="Split PDF", command=split_options)
button2.place(relx=0.5, rely=0.55, anchor=tkinter.CENTER)

button3 = CTkButton(master=frame_left, text="Delete PDF Pages", command=None)
button3.place(relx=0.5, rely=0.65, anchor=tkinter.CENTER)


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




app.mainloop()
