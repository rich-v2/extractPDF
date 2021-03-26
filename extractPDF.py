# This script extracts figures and tables from a PDF file
import fitz # PyMuPDF
import io
import PIL.Image, PIL.ImageTk
import os
import re
import camelot
import tabula
import pandas as pd
from pandastable import Table
import threading
import winsound

from tkinter import *
from tkinter import ttk

# Figures
def extract_figures(file, pages):
    # open the file
    pdf_file = fitz.open(file)
    file = file.replace(".pdf","").strip()
    os.makedirs(os.path.join(file,"Figures"), exist_ok=True)
    if pages == "all" or pages == "":
        for page_index in range(len(pdf_file)):
            # get the page itself
            page = pdf_file[page_index]
            image_list = page.getImageList()
            # printing number of images found in this page
            if image_list:
                print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
            else:
                print("[!] No images found on page", page_index)
            for image_index, img in enumerate(page.getImageList(), start=1):
                # get the XREF of the image
                xref = img[0]
                # extract the image bytes
                base_image = pdf_file.extractImage(xref)
                image_bytes = base_image["image"]
                # get the image extension
                image_ext = base_image["ext"]
                # load it to PIL
                image = PIL.Image.open(io.BytesIO(image_bytes))
                # save it to local disk
                image.save(open(os.path.join(file, "Figures",f"image{page_index+1}_{image_index}.{image_ext}"), "wb"))
                figure_list.insert(END,f"image{page_index+1}_{image_index}.{image_ext}")
    elif re.search(",",pages):
        pages = [int(x)-1 for x in pages.split(",")]

        for page_index in pages:
            # get the page itself
            page = pdf_file[page_index]
            image_list = page.getImageList()
            # printing number of images found in this page
            if image_list:
                print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
            else:
                print("[!] No images found on page", page_index)
            for image_index, img in enumerate(page.getImageList(), start=1):
                # get the XREF of the image
                xref = img[0]
                # extract the image bytes
                base_image = pdf_file.extractImage(xref)
                image_bytes = base_image["image"]
                # get the image extension
                image_ext = base_image["ext"]
                # load it to PIL
                image = PIL.Image.open(io.BytesIO(image_bytes))
                # save it to local disk
                image.save(open(os.path.join(file, "Figures",f"image{page_index+1}_{image_index}.{image_ext}"), "wb"))
                figure_list.insert(END,f"image{page_index+1}_{image_index}.{image_ext}")
    elif re.search("-",pages):
        pages = [int(x) for x in pages.split("-")]

        for page_index in range(pages[0]-1,pages[1]):
            # get the page itself
            page = pdf_file[page_index]
            image_list = page.getImageList()
            # printing number of images found in this page
            if image_list:
                print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
            else:
                print("[!] No images found on page", page_index)
            for image_index, img in enumerate(page.getImageList(), start=1):
                # get the XREF of the image
                xref = img[0]
                # extract the image bytes
                base_image = pdf_file.extractImage(xref)
                image_bytes = base_image["image"]
                # get the image extension
                image_ext = base_image["ext"]
                # load it to PIL
                image = PIL.Image.open(io.BytesIO(image_bytes))
                # save it to local disk
                image.save(open(os.path.join(file, "Figures",f"image{page_index+1}_{image_index}.{image_ext}"), "wb"))
                figure_list.insert(END,f"image{page_index+1}_{image_index}.{image_ext}")
    else:
        page_index = int(pages)

        # get the page itself
        page = pdf_file[page_index]
        image_list = page.getImageList()
        # printing number of images found in this page
        if image_list:
            print(f"[+] Found a total of {len(image_list)} images in page {page_index}")
        else:
            print("[!] No images found on page", page_index)
        for image_index, img in enumerate(page.getImageList(), start=1):
            # get the XREF of the image
            xref = img[0]
            # extract the image bytes
            base_image = pdf_file.extractImage(xref)
            image_bytes = base_image["image"]
            # get the image extension
            image_ext = base_image["ext"]
            # load it to PIL
            image = PIL.Image.open(io.BytesIO(image_bytes))
            # save it to local disk
            image.save(open(os.path.join(file, "Figures",f"image{page_index+1}_{image_index}.{image_ext}"), "wb"))
            figure_list.insert(END,f"image{page_index+1}_{image_index}.{image_ext}")
        

def thread_figures(file,pages):
    t = threading.Thread(target=extract_figures,args=[file,pages])
    t.start()

    
def plot_figure(file, fig):
    file = file.replace(".pdf","").strip()
    figpath = os.path.join(file,"Figures",fig)
    img1 = PIL.Image.open(figpath)
    img1 = img1.resize((int(WIDTH*0.45),int(HEIGHT*0.3)))
    logo = PIL.ImageTk.PhotoImage(img1)  # PIL solution
    label_fig.configure(image=logo)
    label_fig.image = logo

def del_figure(file,fig):
    file = file.replace(".pdf","")
    figpath = os.path.join(file,"Figures",fig)

    figure_list.delete(figure_list.curselection())

    os.remove(figpath)

# Tables
def extract_tables(file, pages, flavor, mod):
    if mod == "Camelot":
        table_info_text.insert(END,"Extracting tables...\n\n")
        flavor = flavor.lower()
        # extract all the tables in the PDF file
        if flavor == "stream":
            tables = camelot.read_pdf(file, pages=pages, flavor=flavor)
        else:
            tables = camelot.read_pdf(file, pages=pages, flavor=flavor)

        file = file.replace(".pdf","")
        os.makedirs(os.path.join(file,"Tables"), exist_ok=True)

        if tables:
            for ind, tab in enumerate(tables):
                for key,value in tab.parsing_report.items():
                    table_info_text.insert(END,key + ": " + str(value) + "\n")
                #table_list.insert(END,"Table " + str(ind) + ": " + str(tab.parsing_report['accuracy'])+"% Accuracy")
                if tab.parsing_report['accuracy'] > 80 and tab.shape[1] > 1:
                    tab.to_csv(os.path.join(file,"Tables","table_" + str(pages) + "_" + str(ind) + ".csv"))
                    table_list.insert(END,"table_" + str(pages) + "_" + str(ind) + ".csv")
                    table_info_text.insert(END,"Table added." + "\n\n")
            #tables = camelot.TableList([x for x in tables if x.parsing_report['accuracy'] > 90 and x.parsing_report['whitespace'] < 25])
            #tables.export(os.path.join(file,"Tables","tables.xlsx"), f = "excel") 
        table_info_text.insert(END,"Table extraction complete.\n\n")
        winsound.Beep(500,1000)
    elif mod == "Tabula":
        table_info_text.insert(END,"Extracting tables...\n\n")
        flavor = flavor.lower()
        
        # extract all the tables in the PDF file
        if flavor == "stream":
            tables = tabula.read_pdf(file, pages=pages, stream=True)
        elif flavor == "lattice":
            tables = tabula.read_pdf(file, pages=pages, lattice=True)
        
        file = file.replace(".pdf","")
        os.makedirs(os.path.join(file,"Tables"), exist_ok=True)

        if tables:
            for ind, tab in enumerate(tables):
                tab.to_csv(os.path.join(file,"Tables","table_" + str(pages) + "_" + str(ind) + ".csv"))
                table_list.insert(END,"table_" + str(pages) + "_" + str(ind) + ".csv")
                table_info_text.insert(END,"Table added." + "\n\n")
        table_info_text.insert(END,"Table extraction complete.\n\n")
        winsound.Beep(500,1000)


def thread_tables(file,pages, flavor, mod):
    t = threading.Thread(target=extract_tables,args=[file,pages,flavor,mod])
    t.start()


def show_table(file,tab):
    file = file.replace(".pdf","")
    tabpath = os.path.join(file,"Tables",tab)
    df = pd.read_csv(tabpath)

    table_show_window = Toplevel(root)
    table_show_window.geometry("800x600")
    #table_show_frame = LabelFrame(table_show_window, bg = "white",text ="Table Display")
    #table_show_frame.place(relx = 0.5, rely = 0.55, relwidth = 0.45, relheight = 0.3, anchor = "nw")

    #table_text = Text(table_show_frame)
    table_text = Table(table_show_window, dataframe= df, showtoolbar=True, showstatusbar=True)
    table_text.show()
    """
    print(df)
    table_text.model.df = df
    table_text.redraw()
    """

def del_table(file,tab):
    file = file.replace(".pdf","")
    tabpath = os.path.join(file,"Tables",tab)

    table_list.delete(table_list.curselection())

    os.remove(tabpath)

def reset_all():
    figure_list.delete(0,END)
    table_list.delete(0,END)
    label_fig.configure(image="")
    table_info_text.delete("1.0",END)
    files1 = [x for x in os.listdir() if re.search(".pdf",x)]
    file_list.config(values=files1)

# GUI
if __name__=="__main__":
    WIDTH, HEIGHT = 800,700

    # Main window
    root = Tk()
    root.title("extractPDF")
    root.geometry(str(WIDTH) + "x" + str(HEIGHT))
    root.configure(bg = "lightblue")

    varFlav = StringVar(root)
    varFlav.set("Stream")

    varMod = StringVar(root)
    varMod.set("Camelot")

    #canvas = Canvas(root,bg ="lightblue",width = WIDTH, height = HEIGHT)
    #canvas.pack()

    title_frame = Frame(root, bg = "red", bd = 10)
    title_frame.place(relx = 0.05, rely = 0.05, relwidth = 0.9, relheight = 0.1, anchor = "nw")

    title_label = Label(title_frame, text = "PDF Image and Table Extractor", font='Helvetica 18 bold', bg = "red",bd=10)
    title_label.place(relx=0.5,relwidth=1,relheight=1,anchor="n")

    specify_frame = Frame(root, bg = "blue")
    specify_frame.place(relx = 0.05, rely = 0.16, relwidth = 0.9, relheight = 0.05, anchor = "nw")
    file_label = Label(specify_frame,text="Filename")
    file_label.place(relwidth = 0.25, relheight = 1)

    files = [x for x in os.listdir() if re.search(".pdf",x)]
    file_list = ttk.Combobox(specify_frame, values = files)
    file_list.place(relx=0.25,relheight=1,relwidth=0.5)

    reset_button = Button(specify_frame, text = "Reset", command=lambda: reset_all())
    reset_button.place(relx=0.75,relheight=1,relwidth=0.25)

    figure_frame = LabelFrame(root, bg = "white",text ="Figure Extraction")
    figure_frame.place(relx = 0.05, rely = 0.225, relwidth = 0.45, relheight = 0.3, anchor = "nw")

    figure_list = Listbox(figure_frame, selectmode = SINGLE)
    figure_list.place(relwidth = 1,relheight=0.8)

    figure_ex_button = Button(figure_frame, text = "Get Figures", command = lambda: thread_figures(file_list.get(),figure_pages_entry.get()))
    figure_ex_button.place(relx=0,rely=0.8,relwidth = 0.25, relheight=0.1,anchor = "nw")

    plot_button = Button(figure_frame, text = "Plot", command=lambda:plot_figure(file_list.get(),figure_list.get(figure_list.curselection())))
    plot_button.place(relx=0.25, rely= 0.8, relwidth=0.25,relheight=0.1, anchor = "nw")

    figure_delete_button = Button(figure_frame,text="Delete Figure",command=lambda: del_figure(file_list.get(),figure_list.get(figure_list.curselection())))
    figure_delete_button.place(relx=0.5, rely= 0.8, relwidth=0.25,relheight=0.1, anchor = "nw")

    figure_pages_label = Label(figure_frame,text="Pages")
    figure_pages_label.place(relx=0, rely= 0.9, relwidth=0.25,relheight=0.1, anchor = "nw")

    figure_pages_entry = Entry(figure_frame)
    figure_pages_entry.place(relx=0.25, rely= 0.9, relwidth=0.25,relheight=0.1, anchor = "nw")

    figure_plot_frame = LabelFrame(root, bg = "white",text ="Figure Display")
    figure_plot_frame.place(relx = 0.5, rely = 0.225, relwidth = 0.45, relheight = 0.3, anchor = "nw")

    label_fig = Label(figure_plot_frame, bg= "white")
    label_fig.place(relx=0.5, relwidth=1,relheight=1, anchor = "n")

    

    table_frame = LabelFrame(root, bg = "white",text ="Table Extraction")
    table_frame.place(relx = 0.05, rely = 0.55, relwidth = 0.45, relheight = 0.3, anchor = "nw")

    table_list = Listbox(table_frame, selectmode = SINGLE)
    table_list.place(relwidth = 1,relheight=0.9)

    table_pages_label = Label(table_frame, text= "Pages")
    table_pages_label.place(relwidth = 0.25, relheight = 0.1, rely=0.8)

    text_pages_entry = Entry(table_frame)
    text_pages_entry.place(relx=0.25, relwidth = 0.25, relheight = 0.1, rely=0.8)

    table_ex_button = Button(table_frame, text = "Get Tables", command = lambda: thread_tables(file_list.get(),\
        text_pages_entry.get(),varFlav.get(),varMod.get()))
    table_ex_button.place(relx=0.5,rely=0.8,relwidth=0.25,relheight=0.1,anchor = "nw")

    table_show_button = Button(table_frame,text="Show Table",command=lambda: show_table(file_list.get(),table_list.get(table_list.curselection())))
    table_show_button.place(relx=0.75,rely=0.8,relwidth=0.25, relheight=0.1, anchor="nw")

    table_delete_button = Button(table_frame,text="Delete Table",command=lambda: del_table(file_list.get(),table_list.get(table_list.curselection())))
    table_delete_button.place(relx=0,rely=0.9,relwidth=0.25, relheight=0.1, anchor="nw")

    table_flavor_label = Label(table_frame, text = "Options")
    table_flavor_label.place(relx=0.25,rely=0.9,relwidth=0.25, relheight=0.1, anchor="nw")

    table_flavor_entry = OptionMenu(table_frame,varFlav, *["Stream","Lattice"])
    table_flavor_entry.place(relx=0.5,rely=0.9,relwidth=0.25, relheight=0.1, anchor="nw")

    table_module_entry = OptionMenu(table_frame,varMod, *["Camelot","Tabula"])
    table_module_entry.place(relx=0.75,rely=0.9,relwidth=0.25, relheight=0.1, anchor="nw")



    table_info_frame = LabelFrame(root, text= "Table Info", bg = "white")
    table_info_frame.place(relx = 0.5, rely = 0.55, relwidth = 0.45, relheight = 0.3, anchor = "nw")
    table_info_text = Text(table_info_frame, font = "Helvetica 12", wrap=WORD)
    table_info_text.place(relwidth=975,relheight=1)
    scrollb = Scrollbar(table_info_frame, command = table_info_text.yview)
    scrollb.place(relx = 0.975,relwidth = 0.025,relheight=1)
    table_info_text['yscrollcommand'] = scrollb.set

    root.mainloop()


"""
    table_show_window = Toplevel(root)
    table_show_window.geometry("400x400")
    table_show_frame = LabelFrame(table_show_window, bg = "white",text ="Table Display")
    table_show_frame.place(relx = 0.5, rely = 0.55, relwidth = 0.45, relheight = 0.3, anchor = "nw")

    #table_text = Text(table_show_frame)
    table_text = Table(table_show_frame, showtoolbar=True, showstatusbar=True)
    table_text.show()
    #table_text.place(relx=0.5,relwidth=1,relheight=0.9,anchor="n")
"""
