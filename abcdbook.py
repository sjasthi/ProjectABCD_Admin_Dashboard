import sys, io, requests, threading, webbrowser, urllib, os, platform, openpyxl
import tkinter as tk
from tkinter import ttk, messagebox
from pptx import Presentation
import pptx.util
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image
import pandas as pd
from pandas import Series, DataFrame
import textwrap

# pip install googletrans
# pip install googletrans==4.0.0-rc1
import googletrans

root = tk.Tk()
root.title("Project ABCD Book Compiler")
root.geometry("1000x600")
root.minsize(1000,600)

# Set sys.stdout to use utf-8 encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

MAIN_FONT = ("helvetica", 12)
LABEL_FONT = ("helvetica bold", 14)
LANGUAGES = [
    "Telugu",
    "Hindi",
    "Spanish"
]

# creates preferences dictionary from preferences.txt
# creates default preferences.txt file if one does not exist
try:
    with open("preferences.txt", "r", encoding="utf8") as file:
        lines = file.readlines()
        preferences = {}

        for line in lines:
            key, value = line.split('=')
            preferences[key.strip()] = value.strip().replace('“', '').replace('”', '').replace('"', '').replace("'", '')
except FileNotFoundError:
    preferences = {
        "TEXT_SIZE" : "14",
        "TEXT_FONT" : "Times New Roman",
        "TITLE_SIZE" : "32",
        "TITLE_FONT" : "Arial",
        "SUBTITLE_SIZE" : "24",
        "SUBTITLE_FONT" : "Arial",
        "PIC_WIDTH" : "720",
        "PIC_HEIGHT" : "1040"
    }
    tk.messagebox.showwarning(title='Warning', message='No preferences.txt file exists in directory. Default preferences.txt will be created and used.')
    print('No preferences.txt file exists in directory. Default preferences.txt will be created and used.')
    with open('preferences.txt', 'w') as f:
        f.write('TEXT_SIZE = 14\nTEXT_FONT = Times New Roman\nTITLE_SIZE = 32\nTITLE_FONT = Arial\nSUBTITLE_SIZE = 24\nSUBTITLE_FONT = Arial\nPIC_WIDTH = 720\nPIC_HEIGHT = 1040')


'''
Fetch the dress data from API
'''
def getDressInfoFromAPI():

    # get update for dress number input
    update_dress_list = []
    get_text_field = text_field.get("1.0", "end-1c").split(',')

    # add to list
    for number in get_text_field:
        if (number.strip().isnumeric()):
            update_dress_list.append(int(number.strip()))
    
    # remove duplicates
    dress_ids = []
    [dress_ids.append(x) for x in update_dress_list if x not in dress_ids]

    # get dress data from API
    dress_data = []
    for id_number in dress_ids:
        response = requests.get(f'https://abcd2.projectabcd.com/api/getinfo.php?id={id_number}', headers={"User-Agent": "XY"})
        # append dress info to dress data if response status_code == 200
        if response.ok:
            dress_data.append(response.json()['data'])
        else:
            print(f'Request for dress ID: {id_number} failed.')

    return dress_data

'''
Translate text to selected language 
'''
def translateText(text):
    # Build the translator
    if translate.get() == 1:
        if language.get() == "Telugu":
            dest_language = 'te'
        elif language.get() == "Hindi":
            dest_language = 'hi'
        elif language.get() == "Spanish":
            dest_language = 'es'
    else:
        return text # return text if english

    # Create the translator
    translator = googletrans.Translator()
    translated_text = translator.translate(text, dest = dest_language)
    return translated_text.text

'''
Sorts dress data based on user selection
'''
def sortDresses(dress_data):
    if sort_order.get() == 1:
        return sorted(dress_data, key=lambda x : x['name'])
    elif sort_order.get() == 2:
        return sorted(dress_data, key=lambda x : x['id'])
    elif sort_order.get() == 3:
        return dress_data

'''
Performs update once generate button clicked
'''
def generateUpdate():

    # gather all dress data from api
    dress_data = getDressInfoFromAPI()
    # sort dress data
    sorted_dress_data = sortDresses(dress_data)

    # create the pptx presentation
    prs = Presentation()
    
    # creates directory to save images if one does not exist
    if not os.path.exists('./images'):
        os.makedirs('./images')

    # get dress info for items in list & translate
    for index, dress_info in enumerate(sorted_dress_data):
        dress_name = dress_info['name']
        dress_description = dress_info['description']
        dress_did_you_know = dress_info['did_you_know']

        # download images from web if download images check box is selected
        if download_imgs.get() == 1:
            # gets dress image
            img_url = f'http://projectabcd.com/images/dress_images/{dress_info["image_url"]}'
            img_path = f'./images/{dress_info["image_url"]}'
            opener = urllib.request.build_opener()
            opener.addheaders=[('User-Agent', 'XY')]
            urllib.request.install_opener(opener)
            urllib.request.urlretrieve(img_url, img_path)

        if layout.get() == 1: # layout 1 == picture on left page - text on right page
            prs.slide_width = pptx.util.Inches(6.05)
            prs.slide_height = pptx.util.Inches(8.22)
            # Choose a slide layout that has a content placeholder at index 1
            image_slide_layout = prs.slide_layouts[6]  # Index 1 often corresponds to a title and content layout

            # Add a slide with the chosen layout
            image_slide = prs.slides.add_slide(image_slide_layout)

            # Creates another slide for text
            text_slide_layout = prs.slide_layouts[6]  
            text_slide = prs.slides.add_slide(text_slide_layout)

            # title
            left = Inches(0.5)
            top = Inches(0.3)
            width = Inches(4.99)
            height = Inches(1.25)   
            title_box1 = image_slide.shapes.add_textbox(left, top, width, height)
            title_box1f = title_box1.text_frame

            title = title_box1f.add_paragraph()
            title.alignment = PP_ALIGN.CENTER
            title.font.name = title_font_var.get()
            title.font.size = Pt(int(title_size_var.get()))
            title.text = f'{dress_name.upper()}'

            # image
            left = Inches(0.18)
            top = Inches(1.72)
            width = Inches(4)
            height = Inches(5.7) 
            try:
                picture = image_slide.shapes.add_picture(f'./images/{dress_info["image_url"]}', left, top, width, height)
            except FileNotFoundError:
                tk.messagebox.showerror(title="Error", message="When the Download Images check box is not checked make sure you have an images \
                                                                directory in the root of this project that includes the correct images. (example: 1 == Slide1.PNG)")
                print(f'Image {dress_info["image_url"]} Not Found!')
                break

            # text box
            left = Inches(0.5)
            top = Inches(1.32)
            width = Inches(4.73)
            height = Inches(5.28)   
            text_box1 = text_slide.shapes.add_textbox(left, top, width, height)
            text_box1f = text_box1.text_frame
            text_box1f.word_wrap = True

            description_subtitle = text_box1f.add_paragraph()
            description_subtitle.font.name = subtitle_font_var.get()
            description_subtitle.font.bold = True
            description_subtitle.font.size = Pt(int(subtitle_size_var.get()))
            description_subtitle.text = translateText("DESCRIPTION:")

            description_text = text_box1f.add_paragraph()
            description_text.font.name = text_font_var.get()
            description_text.font.size = Pt(int(text_size_var.get()))
            description_text.text = f'{translateText(dress_description)}'

            did_you_know_subtitle = text_box1f.add_paragraph()
            did_you_know_subtitle.font.name = subtitle_font_var.get()
            did_you_know_subtitle.font.bold = True
            did_you_know_subtitle.font.size = Pt(int(subtitle_size_var.get()))
            did_you_know_subtitle.text =  f'\n{translateText("DID YOU KNOW?")}'

            did_you_know_text = text_box1f.add_paragraph()
            did_you_know_text.font.name = text_font_var.get()
            did_you_know_text.font.size = Pt(int(text_size_var.get()))
            did_you_know_text.text = f'{translateText(dress_did_you_know)}'

            if numbering.get() == 1: # show page number and dress id
                #dress id
                left = Inches(2.82)
                top = Inches(7.42)
                width = Inches(1.28)
                height = Inches(0.4)   
                dress_id_box1 = image_slide.shapes.add_textbox(left, top, width, height)
                dress_id_box1f = dress_id_box1.text_frame

                dress_id = dress_id_box1f.add_paragraph()
                dress_id.text = f'Dress ID: {dress_info["id"]}'

                # page number
                left = Inches(4.56)
                top = Inches(7.42)
                width = Inches(1.28)
                height = Inches(0.4)   
                page_number_box1 = image_slide.shapes.add_textbox(left, top, width, height)
                page_number_box1f = page_number_box1.text_frame

                page_number = page_number_box1f.add_paragraph()
                page_number.text = f'Page No. {index+1}'

            elif numbering.get() == 2: # show page number
                # page number
                left = Inches(4.56)
                top = Inches(7.42)
                width = Inches(1.28)
                height = Inches(0.4)   
                page_number_box1 = image_slide.shapes.add_textbox(left, top, width, height)
                page_number_box1f = page_number_box1.text_frame

                page_number = page_number_box1f.add_paragraph()
                page_number.text = f'Page No. {index+1}'
            
            elif numbering.get() == 3: # show dress id
                #dress id
                left = Inches(4.56)
                top = Inches(7.42)
                width = Inches(1.28)
                height = Inches(0.4)   
                dress_id_box1 = image_slide.shapes.add_textbox(left, top, width, height)
                dress_id_box1f = dress_id_box1.text_frame

                dress_id = dress_id_box1f.add_paragraph()
                dress_id.text = f'Dress ID: {dress_info["id"]}'

        elif layout.get() == 2: # layout 2 == picture on right - text on left
            prs.slide_width = pptx.util.Inches(11.69)
            prs.slide_height = pptx.util.Inches(8.27)
            slide_layout = prs.slide_layouts[6]

            # Add a slide with the chosen layout
            slide = prs.slides.add_slide(slide_layout)

            # title
            left = Inches(0.5)
            top = Inches(0.3)
            width = Inches(10.71)
            height = Inches(1.25)   
            title_box1 = slide.shapes.add_textbox(left, top, width, height)
            title_box1f = title_box1.text_frame

            title = title_box1f.add_paragraph()
            title.alignment = PP_ALIGN.CENTER
            title.font.name = title_font_var.get()
            title.font.size = Pt(int(title_size_var.get()))
            title.text = f'{dress_name.upper()}'

            # image
            left = Inches(7.07)
            top = Inches(1.72)
            width = Inches(4.27)
            height = Inches(5.7)
            try:
                picture = slide.shapes.add_picture(f'./images/{dress_info["image_url"]}', left, top, width, height)
            except FileNotFoundError:
                tk.messagebox.showerror(title="Error", message="When the Download Images check box is not checked make sure you have an images \
                                                                directory in the root of this project that includes the correct images. (example: 1 == Slide1.PNG)")
                print(f'Image {dress_info["image_url"]} Not Found!')
                break

            # text box
            left = Inches(0.18)
            top = Inches(1.72)
            width = Inches(6.63)
            height = Inches(4.34)   
            text_box1 = slide.shapes.add_textbox(left, top, width, height)
            text_box1f = text_box1.text_frame
            text_box1f.word_wrap = True

            description_subtitle = text_box1f.add_paragraph()
            description_subtitle.font.name = subtitle_font_var.get()
            description_subtitle.font.bold = True
            description_subtitle.font.size = Pt(int(subtitle_size_var.get()))
            description_subtitle.text = translateText("DESCRIPTION:")

            description_text = text_box1f.add_paragraph()
            description_text.font.name = text_font_var.get()
            description_text.font.size = Pt(int(text_size_var.get()))
            description_text.text = f'{translateText(dress_description)}'

            did_you_know_subtitle = text_box1f.add_paragraph()
            did_you_know_subtitle.font.name = subtitle_font_var.get()
            did_you_know_subtitle.font.bold = True
            did_you_know_subtitle.font.size = Pt(int(subtitle_size_var.get()))
            did_you_know_subtitle.text =  f'\n{translateText("DID YOU KNOW?")}'

            did_you_know_text = text_box1f.add_paragraph()
            did_you_know_text.font.name = text_font_var.get()
            did_you_know_text.font.size = Pt(int(text_size_var.get()))
            did_you_know_text.text = f'{translateText(dress_did_you_know)}'

            if numbering.get() == 1: # show page number and dress id
                #dress id
                left = Inches(8.01)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                dress_id_box1 = slide.shapes.add_textbox(left, top, width, height)
                dress_id_box1f = dress_id_box1.text_frame

                dress_id = dress_id_box1f.add_paragraph()
                dress_id.text = f'Dress ID: {dress_info["id"]}'

                # page number
                left = Inches(9.93)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                page_number_box1 = slide.shapes.add_textbox(left, top, width, height)
                page_number_box1f = page_number_box1.text_frame

                page_number = page_number_box1f.add_paragraph()
                page_number.text = f'Page No. {index+1}'

            elif numbering.get() == 2: # show page number
                # page number
                left = Inches(9.93)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)
                page_number_box1 = slide.shapes.add_textbox(left, top, width, height)
                page_number_box1f = page_number_box1.text_frame

                page_number = page_number_box1f.add_paragraph()
                page_number.text = f'Page No. {index+1}'
            
            elif numbering.get() == 3: # show dress id
                #dress id
                left = Inches(9.42)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                dress_id_box1 = slide.shapes.add_textbox(left, top, width, height)
                dress_id_box1f = dress_id_box1.text_frame

                dress_id = dress_id_box1f.add_paragraph()
                dress_id.text = f'Dress ID: {dress_info["id"]}'

        elif layout.get() == 3: # layout 3 == picture on left - text on right
            prs.slide_width = pptx.util.Inches(11.69)
            prs.slide_height = pptx.util.Inches(8.27)
            # Choose a slide layout that has a content placeholder at index 1
            slide_layout = prs.slide_layouts[6]  # Index 1 often corresponds to a title and content layout

            # Add a slide with the chosen layout
            slide = prs.slides.add_slide(slide_layout)

            # title
            left = Inches(0.5)
            top = Inches(0.3)
            width = Inches(10.71)
            height = Inches(1.25)   
            title_box1 = slide.shapes.add_textbox(left, top, width, height)
            title_box1f = title_box1.text_frame

            title = title_box1f.add_paragraph()
            title.alignment = PP_ALIGN.CENTER
            title.font.name = title_font_var.get()
            title.font.size = Pt(int(title_size_var.get()))
            title.text = f'{dress_name.upper()}'

            # image
            left = Inches(0.18)
            top = Inches(1.72)
            width = Inches(4.27)
            height = Inches(5.7) 
            try:
                picture = slide.shapes.add_picture(f'./images/{dress_info["image_url"]}', left, top, width, height)
            except FileNotFoundError:
                tk.messagebox.showerror(title="Error", message="When the Download Images check box is not checked make sure you have an images \
                                                                directory in the root of this project that includes the correct images. (example: 1 == Slide1.PNG)")
                print(f'Image {dress_info["image_url"]} Not Found!')
                break

            # text box
            left = Inches(4.7)
            top = Inches(1.72)
            width = Inches(6.63)
            height = Inches(4.34)   
            text_box1 = slide.shapes.add_textbox(left, top, width, height)
            text_box1f = text_box1.text_frame
            text_box1f.word_wrap = True

            description_subtitle = text_box1f.add_paragraph()
            description_subtitle.font.name = subtitle_font_var.get()
            description_subtitle.font.bold = True
            description_subtitle.font.size = Pt(int(subtitle_size_var.get()))
            description_subtitle.text = translateText("DESCRIPTION:")

            description_text = text_box1f.add_paragraph()
            description_text.font.name = text_font_var.get()
            description_text.font.size = Pt(int(text_size_var.get()))
            description_text.text = f'{translateText(dress_description)}'

            did_you_know_subtitle = text_box1f.add_paragraph()
            did_you_know_subtitle.font.name = subtitle_font_var.get()
            did_you_know_subtitle.font.bold = True
            did_you_know_subtitle.font.size = Pt(int(subtitle_size_var.get()))
            did_you_know_subtitle.text =  f'\n{translateText("DID YOU KNOW?")}'

            did_you_know_text = text_box1f.add_paragraph()
            did_you_know_text.font.name = text_font_var.get()
            did_you_know_text.font.size = Pt(int(text_size_var.get()))
            did_you_know_text.text = f'{translateText(dress_did_you_know)}'

            if numbering.get() == 1: # show page number and dress id
                #dress id
                left = Inches(8.01)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                dress_id_box1 = slide.shapes.add_textbox(left, top, width, height)
                dress_id_box1f = dress_id_box1.text_frame

                dress_id = dress_id_box1f.add_paragraph()
                dress_id.text = f'Dress ID: {dress_info["id"]}'

                # page number
                left = Inches(9.93)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                page_number_box1 = slide.shapes.add_textbox(left, top, width, height)
                page_number_box1f = page_number_box1.text_frame

                page_number = page_number_box1f.add_paragraph()
                page_number.text = f'Page No. {index+1}'

            elif numbering.get() == 2: # show page number
                # page number
                left = Inches(9.93)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                page_number_box1 = slide.shapes.add_textbox(left, top, width, height)
                page_number_box1f = page_number_box1.text_frame

                page_number = page_number_box1f.add_paragraph()
                page_number.text = f'Page No. {index+1}'
            
            elif numbering.get() == 3: # show dress id
                #dress id
                left = Inches(9.42)
                top = Inches(7.63)
                width = Inches(1.28)
                height = Inches(0.4)   
                dress_id_box1 = slide.shapes.add_textbox(left, top, width, height)
                dress_id_box1f = dress_id_box1.text_frame

                dress_id = dress_id_box1f.add_paragraph()
                dress_id.text = f'Dress ID: {dress_info["id"]}'

    try:
        prs.save('abcdbook.pptx')
    except:
        tk.messagebox.showerror(title="Error", message="Access to abcdbook.pptx denied. Make sure there is not an abcdbook.pptx currently open")
        print("Access to abcdbook.pptx denied. Make sure it is not currently open")
    finally:
        generate_button.config(state="normal")

    # Opens ppt depending on OS
    current_os = platform.system()
    try:
        if current_os == "Windows":
            os.system("start abcdbook.pptx")
        elif current_os == "Darwin":
            os.system("open abcdbook.pptx")
        elif current_os == "Linux":
            os.system("xdg-open abcdbook.pptx")
        else:
            print("Error: Cannot open file " + current_os + " not supported.")
    except Exception as e:
        print("Error:", e)

'''
Opens excel file & gets data
'''
def diffReport():
    # helper for table item selection
    def item_select(_):
        for item in table.selection():
            print(table.item(item)['values'])

    # helper to wrap text in treeview cell
    def wrap(string, length=150):
        return '\n'.join(textwrap.wrap(string, length))
    
    # work in progress!!!!!!
    def row_order(order):
        if order == 'id':
            sorted_diff_data = sorted(diff_dress_data, key=lambda x : x[0])
        elif order == 'name':
            sorted_diff_data = sorted(diff_dress_data, key=lambda x : x[1])
        elif order == 'description':
            pass

        

    file_path = 'APIData.xlsx' # Change to path where file is located

    diff_dress_data = [] # data in spreadsheet that is different from API
    api_dress_data = sorted(getDressInfoFromAPI(), key=lambda x : x['id']) # gets dress data from API and sorts by ID

    try:
        sheet_dress_data = pd.read_excel(file_path) # gets dress data from .xlsx spreadsheet
        sheet_dress_data.dropna(subset=['id'], inplace=True) # drops any rows with na ID

        # cycle through api_dress_data
        for index, api_data in enumerate(api_dress_data):
            # row of data with an ID that matches api_data ID
            row = sheet_dress_data.loc[api_data['id']-1] # row of data in spreadsheet

            # check if name, description, or did_you_know is different from the API data
            if row.loc['name'] != api_data['name']:
                diff_dress_data.append([item for item in row])
                continue
            if row.loc['description'] != api_data['description']:
                diff_dress_data.append([item for item in row])
                continue
            if row.loc['did_you_know'] != api_data['did_you_know']:
                diff_dress_data.append([item for item in row])
                continue

        # find largest description field
        row_size_flag = 0
        for value in diff_dress_data:
            if len(value[2]) > row_size_flag:
                row_size_flag = len(value[2])

        # create window to display table
        table_window = tk.Toplevel(root)
        table_window.title("Difference Report")
        table_window.geometry(f"1000x600")
        # table_window.geometry(f"{table_window.winfo_screenwidth()}x{table_window.winfo_screenheight()}")

        # create frame to hold table
        table_frame = tk.Frame(table_window, height=200)
        table_frame.pack()

        # vertical scrollbar
        table_scrolly = tk.Scrollbar(table_frame)
        table_scrolly.pack(side="right", fill='y')
        # horizontal scrollbar
        table_scrollx = tk.Scrollbar(table_frame, orient='horizontal')
        table_scrollx.pack(side="bottom", fill='x')

        # using style to set row height and heading colors
        style = ttk.Style()
        style.theme_use('clam')
        if row_size_flag <= 500:
            style.configure('Treeview', rowheight=75)
        elif row_size_flag > 500 and row_size_flag < 1000:
            style.configure('Treeview', rowheight=100)
        elif row_size_flag >= 1000:
            style.configure('Treeview', rowheight=150)
        style.configure('Treeview.Heading', background='#848484', foreground='white')

        # use ttk Treeview to create table
        table = ttk.Treeview(table_frame, yscrollcommand=table_scrolly.set, xscrollcommand=table_scrollx.set, columns=('id', 'name', 'description', 'did_you_know'), show='headings')
        # configure the scroll bars with the table
        table_scrolly.config(command=table.yview)
        table_scrollx.config(command=table.xview)
        # create the headers
        table.heading('id', text='id', command=lambda: row_order('id'))
        table.heading('name', text='name', command=lambda: row_order('name'))
        table.heading('description', text='description')
        table.heading('did_you_know', text='did_you_know')
        # table.heading('desc_word_count', text='desc_word_count')
        # table.heading('dyk_word_count', text='dyk_word_count')
        # set column variables
        table.column('id', width=50, stretch='no')
        table.column('name', width=145, stretch='no')
        table.column('description', width=1000, anchor='nw', stretch=False)
        table.column('did_you_know', anchor='nw', stretch=False, width=800)
        # pack table into table_window
        table.pack(fill='both', expand=True)

        # fill table with difference report data
        for index, data in enumerate(diff_dress_data):
            # word wrap text in cell description and did_you_know columns
            # if text in description is greater than 500 words extend the word wrap
            if len(data[2]) > 500:
                data[2] = wrap(data[2], 250)
            else:
                data[2] = wrap(data[2])
            data[3] = wrap(data[3], 100)

            # if even row, set tag to evenrow
            # if odd row, set tag to oddrow
            if index % 2 == 0:
                table.insert(parent='', index=tk.END, values=data, tags=('evenrow',))
            else:
                table.insert(parent='', index=tk.END, values=data, tags=('oddrow',))

        # alternate colors each line
        table.tag_configure('evenrow', background='#e8f3ff')
        table.tag_configure('oddrow', background='#f7f7f7')

        # monitor select event on items
        table.bind('<<TreeviewSelect>>', item_select)

        # create button frame and place one table_window
        btn_frame = tk.Frame(table_frame, bg='white', padx=10)
        btn_frame.place(relx=0.01, rely=0.89, height=50, relwidth=.9)
        # create buttons and pack on button_frame
        btn = tk.Button(btn_frame, text='test', width=25, height=2, bg="#007FFF", fg="#ffffff")
        btn.pack(side='left', padx=25)
        btn2 = tk.Button(btn_frame, text='test2', width=25, height=2, bg="#007FFF", fg="#ffffff")
        btn2.pack(side='left', padx=25)

        #raise scrollbars to the top of frame
        table_scrolly.lift()
        table_scrollx.lift()

        # print ("Changes were made to the following dresses:")
        # print(diff_dress_data)

    except FileNotFoundError:
        print(f"File '{file_path}' not found.")
    finally:
        diff_button.config(state="normal")

'''
Spins up new thread to run generateUpdate function
'''
def startGenerateThread():
    generate_button.config(state="disabled")
    generate_thread = threading.Thread(target=generateUpdate)
    generate_thread.start()

'''
Spins up new thread to run diffReport function
'''
def startDiffReportThread():
    diff_button.config(state='disabled')
    diff_report_thread = threading.Thread(target=diffReport)
    diff_report_thread.start()

'''
Launch help site when user clicks Help button
'''
def launchHelpSite():
    # create help site
    with open('help.html', 'w') as file:
        file.write('<!DOCTYPE html>\n<html>\n<head>\n\t<meta charset="utf8">\n\t<title>abcd Help</title>\n</head>\n<body>\n\t\t<h1 style="text-align: center;">Welcome to the help site</h1>\n</body>\n</html>\n')
    
    # open help site
    webbrowser.open('help.html')


#--------------------------------Main GUI-------------------------------------------------------------------------------------------------
#--------------------------------Main Frame-----------------------------------------------------------------------------------------------
# main frame
main_frame = tk.Frame(root, width=1000, height=600)
main_frame.pack_propagate(False)
main_frame.pack()

#--------------------------------Text Field-----------------------------------------------------------------------------------------------
# text field label
text_field_label = tk.Label(main_frame, text="Dress Numbers:", font=LABEL_FONT)
text_field_label.place(x=25, y=72.5)

# text field
text_field = tk.Text(main_frame)
text_field.place(x=175, y=10, width=800, height=135)

# text field initialization
try:
    with open('slide_numbers.txt', 'r') as file:
        slide_number_content = file.readline().strip()
        text_field.insert("1.0", slide_number_content)
except FileNotFoundError:
    print(FileNotFoundError)

#--------------------------------Layout Radio Buttons-------------------------------------------------------------------------------------
# layout variable
layout = tk.IntVar()
layout.set(1)

# layout frame
layout_frame = tk.Frame(main_frame)
# layout radio buttons
layout_radio1 = tk.Radiobutton(layout_frame, text="Picture on Left Page - Text on Right Page - Two Page Mode - Portrait Mode",
                                font=MAIN_FONT, variable=layout, value=1)
layout_radio2 = tk.Radiobutton(layout_frame, text="Picture on Right - Text on Left - Single Page - Landscape Mode",
                                font=MAIN_FONT, variable=layout, value=2)
layout_radio3 = tk.Radiobutton(layout_frame, text="Picture on Left - Text on Right - Single Page - Landscape Mode",
                                font=MAIN_FONT, variable=layout, value=3)
# pack radio buttons into layout frame
layout_radio1.pack(anchor="nw")
layout_radio2.pack(anchor="nw")
layout_radio3.pack(anchor="nw")
# place layout frame on main frame
layout_frame.place(x=175, y=150, width=800)
# layout radio buttons label
layout_label = tk.Label(main_frame, text="Layout:", font=LABEL_FONT)
layout_label.place(x=25, y=150)

# separator line
separator1 = ttk.Separator(main_frame)
separator1.place(x=175, y=150, relwidth=.8)

#--------------------------------Sort Radio Buttons---------------------------------------------------------------------------------------
# sort variable
sort_order = tk.IntVar()
sort_order.set(1)

# sort frame
sort_frame = tk.Frame(main_frame)
# sort radio buttons
sort_radio1 = tk.Radiobutton(sort_frame, text="By Name", font=MAIN_FONT, variable=sort_order, value=1)
sort_radio2 = tk.Radiobutton(sort_frame, text="By ID", font=MAIN_FONT, variable=sort_order, value=2)
sort_radio3 = tk.Radiobutton(sort_frame, text="By Input Order", font=MAIN_FONT, variable=sort_order, value=3)
# pack radio buttons into sort frame
sort_radio1.pack(side="left")
sort_radio2.pack(side="left")
sort_radio3.pack(side="left")
# place sort frame on main frame
sort_frame.place(x=175, y=250, width=800)

# sort radio buttons label
sort_label = tk.Label(main_frame, text="Sort Order:", font=LABEL_FONT)
sort_label.place(x=25, y=250)

# separator line
separator2 = ttk.Separator(main_frame)
separator2.place(x=175, y=250, relwidth=.8)

#--------------------------------Preferences----------------------------------------------------------------------------------------------
# Initialize variables for storing values
text_size_var = tk.StringVar()
title_size_var = tk.StringVar()
subtitle_size_var = tk.StringVar()
text_font_var = tk.StringVar()
title_font_var = tk.StringVar()
subtitle_font_var = tk.StringVar()
pic_width_var = tk.StringVar()
pic_height_var = tk.StringVar()

# Set initial values for the Entry fields
text_size_var.set(preferences["TEXT_SIZE"])
title_size_var.set(preferences["TITLE_SIZE"])
subtitle_size_var.set(preferences["SUBTITLE_SIZE"])
text_font_var.set(preferences["TEXT_FONT"])
title_font_var.set(preferences["TITLE_FONT"])
subtitle_font_var.set(preferences["SUBTITLE_FONT"])
pic_width_var.set(preferences["PIC_WIDTH"])
pic_height_var.set(preferences["PIC_HEIGHT"])

# preferences frame
preferences_frame = tk.Frame(main_frame)
# preferences labels and entry fields
text_size_label = tk.Label(preferences_frame, text="Text Size:", font=MAIN_FONT)
text_size = tk.Entry(preferences_frame, width=3, textvariable=text_size_var, state="disabled", font=MAIN_FONT)

title_size_label = tk.Label(preferences_frame, text="Title Size:", font=MAIN_FONT)
title_size = tk.Entry(preferences_frame, width=3, textvariable=title_size_var, state="disabled",  font=MAIN_FONT)

subtitle_size_label = tk.Label(preferences_frame, text="Subtitle Size:", font=MAIN_FONT)
subtitle_size = tk.Entry(preferences_frame, width=3, textvariable=subtitle_size_var, state="disabled",  font=MAIN_FONT)

text_font_label = tk.Label(preferences_frame, text="Text Font:", font=MAIN_FONT)
text_font = tk.Entry(preferences_frame, width=25, textvariable=text_font_var, state="disabled",  font=MAIN_FONT)

title_font_label = tk.Label(preferences_frame, text="Title Font:", font=MAIN_FONT)
title_font = tk.Entry(preferences_frame, width=25, textvariable=title_font_var, state="disabled",  font=MAIN_FONT)

subtitle_font_label = tk.Label(preferences_frame, text="Subitle Font:", font=MAIN_FONT)
subtitle_font = tk.Entry(preferences_frame, width=25, textvariable=subtitle_font_var, state="disabled",  font=MAIN_FONT)

pic_width_label = tk.Label(preferences_frame, text="Pic Width:", font=MAIN_FONT)
pic_width = tk.Entry(preferences_frame, width=6, textvariable=pic_width_var, state="disabled",  font=MAIN_FONT)

pic_height_label = tk.Label(preferences_frame, text="Pic Height:", font=MAIN_FONT)
pic_height = tk.Entry(preferences_frame, width=6, textvariable=pic_height_var, state="disabled",  font=MAIN_FONT)

# grid preference labels and entry fields into preferences frame
# column 1 + 2
text_size_label.grid(row=1, column=1, pady=10)
text_size.grid(row=1, column=2)

title_size_label.grid(row=2, column=1, pady=10)
title_size.grid(row=2, column=2)

subtitle_size_label.grid(row=3, column=1, pady=10, padx=15)
subtitle_size.grid(row=3, column=2)

# column 3 + 4
text_font_label.grid(row=1, column=3)
text_font.grid(row=1, column=4)

title_font_label.grid(row=2, column=3)
title_font.grid(row=2, column=4)

subtitle_font_label.grid(row=3, column=3, padx=15)
subtitle_font.grid(row=3, column=4)

# column 5 + 6
pic_width_label.grid(row=1, column=5)
pic_width.grid(row=1, column=6)

pic_height_label.grid(row=2, column=5, padx=15)
pic_height.grid(row=2, column=6)

# place preferences frame on main frame
preferences_frame.place(x=175, y=295, width=800)

# preferences label
preferences_label = tk.Label(main_frame, text="Preferences:", font=LABEL_FONT)
preferences_label.place(x=25, y=295)

# separator line
separator3 = ttk.Separator(main_frame)
separator3.place(x=175, y=295, relwidth=.8)

#--------------------------------Numbering Radio Buttons--------------------------------------------------------------------------------------
# numbering variable
numbering = tk.IntVar()
numbering.set(1)

# numbering frame
numbering_frame = tk.Frame(main_frame)
# numbering radio buttons
numbering_radio1 = tk.Radiobutton(numbering_frame, text="Show both Page Number and Dress ID", font=MAIN_FONT, variable=numbering, value=1)
numbering_radio2 = tk.Radiobutton(numbering_frame, text="Show Page Number", font=MAIN_FONT, variable=numbering, value=2)
numbering_radio3 = tk.Radiobutton(numbering_frame, text="Show Dress ID", font=MAIN_FONT, variable=numbering, value=3)
# pack numbering buttons into sort frame
numbering_radio1.pack(side="left")
numbering_radio2.pack(side="left")
numbering_radio3.pack(side="left")
# place numbering frame on main frame
numbering_frame.place(x=175, y=445, width=800)

# numbering radio buttons label
numbering_label = tk.Label(main_frame, text="Numbering:", font=LABEL_FONT)
numbering_label.place(x=25, y=445)

# separator line
separator4 = ttk.Separator(main_frame)
separator4.place(x=175, y=445, relwidth=.8)

#--------------------------------Translate and Image Check Buttons--------------------------------------------------------------------------------------
# translate check variable
translate = tk.IntVar()
translate.set(0)
# language options variable
language = tk.StringVar()
language.set(LANGUAGES[0])
# Download image variable
download_imgs = tk.IntVar()
download_imgs.set(1)

# translate frame
check_button_frame = tk.Frame(main_frame)
# translate check button
translate_checkbutton = tk.Checkbutton(check_button_frame, text="Translate to:", font=MAIN_FONT, variable=translate, onvalue=1, offvalue=0)
# language options
language_options = tk.OptionMenu(check_button_frame, language, *LANGUAGES)
# download images
download_images = tk.Checkbutton(check_button_frame, text="Download Images", font=MAIN_FONT, variable=download_imgs, onvalue=1, offvalue=0)

# pack translate options into translate frame
translate_checkbutton.pack(side="left")
language_options.pack(side="left")
download_images.pack(side="left")
# place translate frame on main frame
check_button_frame.place(x=175, y=495)

# separator line
separator5 = ttk.Separator(main_frame)
separator5.place(x=175, y=495, relwidth=.8)

#--------------------------------Generate, Help, and Difference Buttons--------------------------------------------------------------------------------------
# button frame
button_frame = tk.Frame(main_frame)
# generate button
generate_button = tk.Button(button_frame, text="Generate", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startGenerateThread)
# help button
help_button = tk.Button(button_frame, text="Help", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=launchHelpSite)
# upload button
diff_button = tk.Button(button_frame, text="Diff Report", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startDiffReportThread)
# pack buttons into button frame
generate_button.pack(side="left", padx=35)
help_button.pack(side="left")
diff_button.pack(side="left", padx=30)

# place button frame on main frame
button_frame.pack(side="bottom", pady=10)

# main gui loop
root.mainloop()