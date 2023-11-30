import sys, io, requests, threading, webbrowser, urllib, os, platform, openpyxl, string, textwrap, re, time, textstat
import tkinter as tk
from tkinter import ttk, messagebox
from pptx import Presentation
import pptx.util
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from PIL import Image
import pandas as pd
from pandas import Series, DataFrame
from concurrent.futures import ThreadPoolExecutor, as_completed
from textblob import TextBlob
import wikipediaapi
# pip install googletrans
# pip install googletrans==4.0.0-rc1
import googletrans

ROOT_WIDTH = 1000 # app window width
ROOT_HEIGHT = 600 # app window height

root = tk.Tk()
root.title("Project ABCD Admin Panel")
sw_placement = int(root.winfo_screenwidth()/2 - ROOT_WIDTH/2) # to place at half width of screen
sh_placement = int(root.winfo_screenheight()/2 - ROOT_HEIGHT/2) # to place at half height of screen
root.geometry(f"{ROOT_WIDTH}x{ROOT_HEIGHT}+{sw_placement}+{sh_placement}")
root.minsize(ROOT_WIDTH, ROOT_HEIGHT)

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
Gathers data from API
'''
def downloadAPIData(url, id_number):
    try:
        response = requests.get(url, headers={"User-Agent": "XY"})
        # append dress info to dress data if response status_code == 200
        if response.ok:
            return response.json()['data']
        else:
            print(f'Request for dress ID: {id_number} failed.')
    except requests.exceptions.RequestException as e:
        print(f'-- DEBUG -- in downloadAPIData: {e}')
    except Exception as e:
        # tk.messagebox.showerror(title="Error", message=f'Could not make connection!\n\nError: {e}')
        print(f'Error: {e}')
    
'''
Sets up and starts threads for gathering API data
'''
def apiRunner():
    # dress ids in entry field
    dress_ids = getSlideNumbers()

    # create list of all urls to send requests to
    url_list = []
    for id_number in dress_ids:
        url_list.append(f'https://abcd2.projectabcd.com/api/getinfo.php?id={id_number}')

    dress_data = [] # dress data from API
    threads= [] # working threads

    # progress bar window for API data retrieval
    progress_window = tk.Toplevel(root)
    progress_window.title('Retrieving API Data')
    sw = int(progress_window.winfo_screenwidth()/2 - 450/2)
    sh = int(progress_window.winfo_screenheight()/2 - 70/2)
    progress_window.geometry(f'450x70+{sw}+{sh}')
    progress_window.resizable(False, False)
    progress_window.attributes('-disable', True)
    progress_window.focus()

    # progress bar custom style
    pb_style = ttk.Style()
    pb_style.theme_use('clam')
    pb_style.configure('green.Horizontal.TProgressbar', foreground='#1ec000', background='#1ec000')

    # frame to hold progress bar
    pb_frame = tk.Frame(progress_window)
    pb_frame.pack()

    # progress bar
    pb = ttk.Progressbar(pb_frame, length=400, style='green.Horizontal.TProgressbar', mode='determinate', maximum=100, value=0)
    pb.pack(pady=10)

    # label for percent complete
    percent_label = tk.Label(pb_frame, text='Retrieving API Data...0%')
    percent_label.pack()

    # spins up 10 threads at a time and stores retrieved data into dress_data upon completion
    with ThreadPoolExecutor(max_workers=10) as exec:
        for index, url in enumerate(url_list):
            threads.append(exec.submit(downloadAPIData, url, dress_ids[index]))
        
        complete = 0 # number of threads that have finished
        for task in as_completed(threads):
            complete += 1
            pb['value'] = (complete/len(dress_ids))*100 # calculate percentage of data retrieved
            percent_label.config(text=f'Retrieving API Data...{int(pb["value"])}%') # update completion percent label

            if task.result() is not None:
                dress_data.append(task.result()) # append retrieved data to dress_data

    progress_window.destroy()
    return dress_data

'''
Gets dress IDs from entry field
'''
def getSlideNumbers():
    # get update for dress number input
    update_dress_list = []
    # get dress numbers from text field
    get_text_field = text_field.get("1.0", "end-1c").split(',')
    
    # add to list
    for number in get_text_field:
        if (number.strip().isnumeric()):
            update_dress_list.append(int(number.strip()))
    
    # remove duplicates
    dress_ids = []
    [dress_ids.append(x) for x in update_dress_list if x not in dress_ids]

    return dress_ids

'''
Downloads dress images
'''
def downloadImages(folder_name, url, img_name):
    try:
        # downloads dress image
        img_url = url
        img_path = f'./{folder_name}/{img_name}'
        opener = urllib.request.build_opener()
        opener.addheaders=[('User-Agent', 'XY')]
        urllib.request.install_opener(opener)
        urllib.request.urlretrieve(img_url, img_path)
    except Exception as e:
        print(f'Error downloading images: {e}')

'''
Sets up and starts threads for downloading dress images
'''
def imageRunner(dress_data):
    # list of all urls to send requests to
    url_list = []
    # list of all image names
    img_name_list = []

    for data in dress_data:
        url_list.append(f'http://projectabcd.com/images/dress_images/{data["image_url"]}')
        img_name_list.append(f'{data["image_url"]}')

    # progress bar window for API data retrieval
    progress_window = tk.Toplevel(root)
    progress_window.title('Downloading Images')
    sw = int(progress_window.winfo_screenwidth()/2 - 450/2)
    sh = int(progress_window.winfo_screenheight()/2 - 70/2)
    progress_window.geometry(f'450x70+{sw}+{sh}')
    progress_window.resizable(False, False)
    progress_window.attributes('-disable', True)
    progress_window.focus()

    # progress bar custom style
    pb_style = ttk.Style()
    pb_style.theme_use('clam')
    pb_style.configure('green.Horizontal.TProgressbar', foreground='#1ec000', background='#1ec000')

    # frame to hold progress bar
    pb_frame = tk.Frame(progress_window)
    pb_frame.pack()

    # progress bar
    pb = ttk.Progressbar(pb_frame, length=400, style='green.Horizontal.TProgressbar', mode='determinate', maximum=100, value=0)
    pb.pack(pady=10)

    # label for percent complete
    percent_label = tk.Label(pb_frame, text='Downloading Images...0%')
    percent_label.pack()

    threads= [] # working threads

    # spins up 10 threads at a time and calls downloadImages with url and image name
    with ThreadPoolExecutor(max_workers=10) as exec:
        for index, url in enumerate(url_list):
            threads.append(exec.submit(downloadImages, "images", url, img_name_list[index]))

        complete = 0 # number of threads that have finished
        for task in as_completed(threads):
            complete += 1
            pb['value'] = (complete/len(url_list))*100 # calculate percentage of images downloaded
            percent_label.config(text=f'Downloading Images...{int(pb["value"])}%') # update completion percent label
    
    progress_window.destroy() # close progress bar window

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
    try:
        # Create the translator
        translator = googletrans.Translator()
        translated_text = translator.translate(text, dest = dest_language)
        return translated_text.text
    except Exception as e:
        print("Error:", e)
        return ""

'''
Sorts dress data based on user selection
'''
def sortDresses(dress_data):
    if sort_order.get() == 1:
        return sorted(dress_data, key=lambda x : str(x['name']).lower())
    elif sort_order.get() == 2:
        return sorted(dress_data, key=lambda x : x['id'])
    elif sort_order.get() == 3:
        return dress_data

'''
Opens file depending on OS
'''
def openFile(file_name):
    current_os = platform.system()
    try:
        if current_os == "Windows":
            print(f"-- DEBUG -- Windows {file_name}")
            os.system(f"start {file_name}")
        elif current_os == "Darwin":
            print(f"-- DEBUG -- Darwin {file_name}")
            os.system(f"open {file_name}")
        elif current_os == "Linux":
            print(f"-- DEBUG -- Linux {file_name}")
            os.system(f"xdg-open {file_name}")
        else:
            print("Error: Cannot open file " + current_os + " not supported.")
    except Exception as e:
        print("Error:", e)

'''
Create the Dress Name Title textbox
'''
def add_title_box(slide, dress_name, left, top, width, height):
    title = slide.shapes.title
    title.left = Inches(left)
    title.top = Inches(top)
    title.width = Inches(width)
    title.height = Inches(height)
    title.text = f'{dress_name.upper()}'

    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0x09, 0x09, 0x82)
    title.text_frame.paragraphs[0].text = title.text_frame.paragraphs[0].text.upper()
    title.text_frame.paragraphs[0].font.name = title_font_var.get()
    title.text_frame.paragraphs[0].font.size = Pt(int(title_size_var.get()))
    title.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

'''
Create the subtitle highlight
'''
def add_subtitle_highlight(slide, left, top, width, height):
    # create the shape
    rectangle1 = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    rectangle1.rotation = 357

    # color fill
    fill1 = rectangle1.fill
    fill1.solid()
    fill1.fore_color.rgb = RGBColor(202, 246, 189)

    # no outline
    line1 = rectangle1.line
    line1.color.rgb = RGBColor(255, 255, 255)

    # no shaddow
    shadow1 = rectangle1.shadow
    shadow1.inherit = False

'''
Create the Description subtitle
'''
def add_description_subtitle(slide, left, top, width, height):
    description_subtitle_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    description_subtitle_frame = description_subtitle_box.text_frame

    description_subtitle = description_subtitle_frame.add_paragraph()
    description_subtitle.font.color.rgb = RGBColor(0x6E, 0xD8, 0xFF)
    description_subtitle.font.bold = True
    description_subtitle.font.name = subtitle_font_var.get()
    description_subtitle.text = translateText("DESCRIPTION:")

    # make text smaller for layout 4
    if layout.get() == 4:
        description_subtitle.font.size = Pt(16)
    else:
        description_subtitle.font.size = Pt(int(subtitle_size_var.get()))

'''
Add the Description text
'''
def add_description_text(slide, dress_description, left, top, width, height):
    # description - text (left, top, width, height)
    description_text_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    description_text_frame = description_text_box.text_frame

    description_text_frame.word_wrap = True
    description_text = description_text_frame.add_paragraph()
    description_text.font.name = text_font_var.get()
    description_text.text = f'{translateText(dress_description)}'

    # make text smaller for layout 4
    if layout.get() == 4:
        description_text.font.size = Pt(11)
    else:
        description_text.font.size = Pt(int(text_size_var.get()))

'''
Create Did You Know subtitle
'''
def add_did_you_know_subtitle(slide, left, top, width, height):
    did_you_know_subtitle_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    did_you_know_subtitle_frame = did_you_know_subtitle_box.text_frame

    did_you_know_subtitle = did_you_know_subtitle_frame.add_paragraph()
    did_you_know_subtitle.font.color.rgb = RGBColor(0x6E, 0xD8, 0xFF)
    did_you_know_subtitle.font.bold = True
    did_you_know_subtitle.font.name = subtitle_font_var.get()
    did_you_know_subtitle.text = translateText("DID YOU KNOW?")

    # make text smaller for layout 4
    if layout.get() == 4:
        did_you_know_subtitle.font.size = Pt(16)
    else:
        did_you_know_subtitle.font.size = Pt(int(subtitle_size_var.get()))

'''
Create Did You Know text
'''
def add_did_you_know_text(slide, dress_did_you_know, left, top, width, height):
    did_you_know_text_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    did_you_know_text_frame = did_you_know_text_box.text_frame

    did_you_know_text_frame.word_wrap = True
    did_you_know_text = did_you_know_text_frame.add_paragraph()
    did_you_know_text.font.name = text_font_var.get()
    did_you_know_text.text = f'{translateText(dress_did_you_know)}'

    # make text smaller for layout 4
    if layout.get() == 4:
        did_you_know_text.font.size = Pt(11)
    else:
        did_you_know_text.font.size = Pt(int(text_size_var.get()))

'''
Add the dress image 
'''
def add_image(slide, dress_info, left, top):
    image_width_px = int(pic_width_var.get())
    image_height_px = int(pic_height_var.get())
    image_width_inch = image_width_px / 96
    image_height_inch = image_height_px / 96

    # change image size for layout 4 
    if layout.get() == 4:
        image_width_inch = 3.71
        image_height_inch = 4.94

    try:
        picture = slide.shapes.add_picture(f'./images/Slide{dress_info["id"]}.png', 0, Inches(0.83), Inches(image_width_inch), Inches(image_height_inch))
    except FileNotFoundError:
        # tk.messagebox.showerror(title="Error", message="When the Download Images check box is not checked make sure you have an images \
        #                                                 directory in the root of this project that includes the correct images. (example: 1 == Slide1.PNG)")
        print(f'Image Slide{dress_info["id"]}.png Not Found!')

'''
Create the Dress Id & Page Number text
'''
def add_numbering(slide, dress_info, index, left, top, width, height, left2, top2, width2, height2):
    # show both page num & dress id
    if numbering.get() == 1:
        # dress id (left, top, width, height)
        dress_id_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
        dress_id_box_frame = dress_id_box.text_frame

        dress_id = dress_id_box_frame.add_paragraph()
        dress_id.font.name = text_font_var.get()
        dress_id.text = translateText(f'Dress ID: {dress_info["id"]}')
        
        # aligns text left & make text smaller
        if layout.get() == 4:
            dress_id.alignment = PP_ALIGN.LEFT
            dress_id.font.size = Pt(11)
        else:
            dress_id.alignment = PP_ALIGN.RIGHT
            dress_id.font.size = Pt(int(text_size_var.get()))

        # page number (left, top, width, height)
        page_number_box = slide.shapes.add_textbox(Inches(left2), Inches(top2), Inches(width2), Inches(height2))
        page_number_box_frame = page_number_box.text_frame

        page_number = page_number_box_frame.add_paragraph()
        page_number.font.name = text_font_var.get()
        page_number.text = translateText(f'Page No. {index+1}')

        # aligns text left & make text smaller
        if layout.get() == 4:
            page_number.alignment = PP_ALIGN.LEFT
            page_number.font.size = Pt(11)
        else:
            page_number.alignment = PP_ALIGN.RIGHT
            page_number.font.size = Pt(int(text_size_var.get()))
    # show page num or dress id
    elif numbering.get() == 2 or numbering.get() == 3:
        number_box = slide.shapes.add_textbox(Inches(left2), Inches(top2), Inches(width2), Inches(height2))
        number_box_frame = number_box.text_frame

        number_box = number_box_frame.add_paragraph()
        number_box.font.name = text_font_var.get()

        # aligns text left & make text smaller
        if layout.get() == 4:
            number_box.alignment = PP_ALIGN.LEFT
            number_box.font.size = Pt(11)
        else:
            number_box.alignment = PP_ALIGN.RIGHT
            number_box.font.size = Pt(int(text_size_var.get()))

        # show page num
        if numbering.get() == 2:
            number_box.text = translateText(f'Page No. {index+1}')
        # show dress id
        elif numbering.get() == 3:
            number_box.text = translateText(f'Dress ID: {dress_info["id"]}')

'''
Closes the pop alert message
'''
def close_popup(popup):
    popup.destroy()

'''
Updates the timer for the alert popup
'''
def update_timer(popup, timer_label, seconds_left):
    timer_label.config(text=f"Auto closes in {seconds_left} sec.")

    if seconds_left > 0:
        # Update the timer every second
        root.after(1000, update_timer, popup, timer_label, seconds_left - 1)
    else:
        close_popup(popup)

'''
Displays the alert message
'''
def show_error_popup(text_message, duration_num):
    popup = tk.Toplevel(root)
    popup.title("Alert Message")
    
    label = tk.Label(popup, text=text_message)
    label.pack(padx=10, pady=10)

    timer_label = tk.Label(popup, text="")
    timer_label.pack(pady=5)

    duration = duration_num
    update_timer(popup, timer_label, duration)

'''
Performs update once generate button clicked
'''
def generateBook():
    # check if Generate from Local is active
    if gen_local.get() == 1:
        # get update for dress number input
        update_dress_list = []
        # get dress numbers from text field
        get_text_field = text_field.get("1.0", "end-1c").split(',')
        
        # add to list
        for number in get_text_field:
            if (number.strip().isnumeric()):
                update_dress_list.append(int(number.strip()))
        
        # remove duplicates
        dress_ids = []
        [dress_ids.append(x) for x in update_dress_list if x not in dress_ids]

        # path to local Excel data
        file_path = "APIData.xlsx"

        # gets and cleans dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)
        sheet_dress_data.dropna(subset=['id'], inplace=True) # drops any rows with na ID
        sheet_dress_data['description'].fillna('', inplace=True) # removes na/nan from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True) # removes na/nan from did_you_know column
        sheet_dress_data['description'] = sheet_dress_data['description'].astype(str).apply(openpyxl.utils.escape.unescape) # convert escaped strings to ASCII
        sheet_dress_data['did_you_know'] = sheet_dress_data['did_you_know'].astype(str).apply(openpyxl.utils.escape.unescape) # Convert escaped strings to ASCII
        
        # holds dress data from local Excel sheet
        dress_data = [] 
        # cycle through dress_ids and append data from excel sheet to dress_data
        for id in dress_ids:
            row = sheet_dress_data.loc[id-1]
            dress_data.append({'id':row.loc['id'], 'name':row.loc['name'], 'description':row.loc['description'], 'did_you_know':row.loc['did_you_know']})   
    else:
        dress_data = apiRunner() # gather all dress data from api
    
    # sort dress data
    sorted_dress_data = sortDresses(dress_data)

    # if there is no image in the local folder then download from web
    if download_imgs.get() == 0:
        if not os.path.exists('./images'):
            os.makedirs('./images')
        if not os.listdir('./images'):
            text_message = "Local folder is empty. Attempting to grab image(s) from web."
            duration_num = 4
            show_error_popup(text_message, duration_num)
            time.sleep(duration_num+1)
            imageRunner(sorted_dress_data)

    # download images from web if download images check box is selected
    if download_imgs.get() == 1:
        # creates directory to save images if one does not exist
        if not os.path.exists('./images'):
            os.makedirs('./images')
        imageRunner(sorted_dress_data) # download images for each dress in list

    # create powerpoint
    prs = Presentation() # create the pptx presentation
    ppt_file_name = "abcdbook.pptx"
    file_name = "abcdbook.pptx"
    count = 0
    while os.path.exists(file_name): # check if file name exist,
        count += 1
        file_name = f"{os.path.splitext(ppt_file_name)[0]}({count}).pptx" # if file name exisit create new filename

    # progress bar window for API data retrieval
    progress_window = tk.Toplevel(root)
    progress_window.title('Creating Book')
    sw = int(progress_window.winfo_screenwidth()/2 - 450/2)
    sh = int(progress_window.winfo_screenheight()/2 - 70/2)
    progress_window.geometry(f'450x70+{sw}+{sh}')
    progress_window.resizable(False, False)
    progress_window.attributes('-disable', True)
    progress_window.focus()

    # progress bar custom style
    pb_style = ttk.Style()
    pb_style.theme_use('clam')
    pb_style.configure('green.Horizontal.TProgressbar', foreground='#1ec000', background='#1ec000')

    # frame to hold progress bar
    pb_frame = tk.Frame(progress_window)
    pb_frame.pack()

    # progress bar
    pb = ttk.Progressbar(pb_frame, length=400, style='green.Horizontal.TProgressbar', mode='determinate', maximum=100, value=0)
    pb.pack(pady=10)

    # label for percent complete
    percent_label = tk.Label(pb_frame, text='Creating Book...0%')
    percent_label.pack()
    complete = 0

    # get dress info for items in list & translate
    for index, dress_info in enumerate(sorted_dress_data):
        left = None
        image_left = None
        
        dress_name = dress_info['name']
        dress_description = dress_info['description']
        dress_did_you_know = dress_info['did_you_know']
        dress_description_len = len(dress_description)

        #--------------------------------Portrait--------------------------------
        # PORTRAIT MODE
        if layout.get() == 1 or layout.get() == 4: 
            prs.slide_width = pptx.util.Inches(7.5) # define slide width
            prs.slide_height = pptx.util.Inches(10.83) # define slide height
            slide_layout = prs.slide_layouts[5] # use slide with only title
            slide_layout2 = prs.slide_layouts[6] # use empty slide

            # LAYOUT 1 == picture on left page - text on right page - two page
            if layout.get() == 1:
                slide_empty = prs.slides.add_slide(slide_layout2) 
                slide_title = prs.slides.add_slide(slide_layout) 

                add_image(slide_empty, dress_info, 0, 0)
                add_title_box(slide_title, dress_name, 0, 0.15, 7.5, 0.91) 
                add_subtitle_highlight(slide_title, 0.37, 1.58, 2.44, 0.3) # description - highlight box
                add_description_subtitle(slide_title, 0.28, 1.07, 6.94, 0.51)
                add_description_text(slide_title, dress_description, 0.28, 1.65, 6.94, 5.99)
                add_subtitle_highlight(slide_title, 0.37, 8.36, 2.78, 0.3) # did you know - highlight box
                add_did_you_know_subtitle(slide_title, 0.28, 7.87, 6.94, 0.51)
                add_did_you_know_text(slide_title, dress_did_you_know, 0.28, 8.46, 6.94, 1.04)
                add_numbering(slide_title, dress_info, index, 4.47, 10.06, 1.28, 0.34, 5.94, 10.06, 1.28, 0.34)

            # LAYOUT 4 == picture on left - text on right - single page
            elif layout.get() == 4: 
                slide_title = prs.slides.add_slide(slide_layout) 

                add_image(slide_title, dress_info, 0, 1.39)
                add_title_box(slide_title, dress_name, 0, 0.09, 7.5, 0.91) 
                add_subtitle_highlight(slide_title, 3.71, 1.21, 1.83, 0.23) # description - highlight box
                add_description_subtitle(slide_title, 3.63, 0.88, 3.71, 0.37)
                add_description_text(slide_title, dress_description, 3.63, 1.35, 3.71, 7.57)

                # adjust the text height based on text length
                if dress_description_len < 600:
                    add_subtitle_highlight(slide_title, 3.72, 5.83, 1.83, 0.23) # did you know - highlight box
                    add_did_you_know_subtitle(slide_title, 3.63, 5.42, 3.71, 0.37)
                    add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 5.94, 3.71, 0.91)

                elif dress_description_len > 600 and dress_description_len < 1300:
                    add_subtitle_highlight(slide_title, 3.72, 7.99, 1.83, 0.23) # did you know - highlight box
                    add_did_you_know_subtitle(slide_title, 3.63, 7.58, 3.71, 0.37)
                    add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 8.1, 3.71, 0.91)
                    
                else:
                    add_subtitle_highlight(slide_title, 3.72, 9.32, 1.83, 0.23) # did you know - highlight box
                    add_did_you_know_subtitle(slide_title, 3.63, 8.91, 3.71, 0.37)
                    add_did_you_know_text(slide_title, dress_did_you_know, 3.63, 9.43, 3.71, 0.91)

                add_numbering(slide_title, dress_info, index, 0.49, 6.46, 1.33, 0.27, 1.95, 6.46, 1.33, 0.27)
 
        #--------------------------------Landscape--------------------------------
        # LANDSCAPE MODE
        elif layout.get() == 2 or layout.get() == 3:
            left = None
            image_left = None
            rectangle_left = None

            # slide size (left, top, width, height)
            prs.slide_width = pptx.util.Inches(16.18) # define slide width
            prs.slide_height = pptx.util.Inches(12.53) # define slide height
            slide_layout = prs.slide_layouts[5] # use empty slide
            slide = prs.slides.add_slide(slide_layout) # add empty slide to pptx

            # LAYOUT 2 == picture on right - text on left
            if layout.get() == 2:
                rectangle_left = 0.4
                left = 0.34
                image_left = 8.45
                image_top = 1.17
                numbering1_left = 1.05
                numbering2_left = 2.53

            # LAYOUT 3 == picture on left - text on right
            elif layout.get() == 3:
                image_left = 0.25
                image_top = 1.17
                rectangle_left = 8.15
                left = 8.09
                numbering1_left = 12.78
                numbering2_left = 14.26

            add_title_box(slide, dress_name, 0, 0.15, 16.18, 0.91) 
            add_subtitle_highlight(slide, rectangle_left, 1.75, 2.44, 0.3) # decription - highlight box
            add_description_subtitle(slide, left, 1.26, 7.81, 0.51)
            add_description_text(slide, dress_description, left, 1.88, 7.81, 5.35)
            add_subtitle_highlight(slide, rectangle_left, 8.04, 2.76, 0.28) # did you know - highlight box
            add_did_you_know_subtitle(slide, left, 7.56, 7.81, 0.51)
            add_did_you_know_text(slide, dress_did_you_know, left, 8.19, 7.81, 1.11)
            add_numbering(slide, dress_info, index, numbering1_left, 11.08, 1.28, 0.34, numbering2_left, 11.08, 1.28, 0.34)
            add_image(slide, dress_info, image_left, image_top)
        complete += 1
        pb['value'] = (complete/len(sorted_dress_data))*100 # calculate percentage of images downloaded
        percent_label.config(text=f'Creating Book...{int(pb["value"])}%') # update completion percent label
    try:
        prs.save(file_name)
    except Exception as e:
        print(f"-- DEBUG -- saving presentation: {e}")
    finally:
        book_gen_generate_button.config(state="normal")
    
    progress_window.destroy() # close progress bar window

    openFile(file_name)

'''
Helper function to wrap text
'''
def wrap(string, length=150):
    return '\n'.join(textwrap.wrap(string, length))

'''
Order sort for treeview table
'''
def rowOrder(order, diff_dress_data, table):
    if order == 'id':
        table.delete(*table.get_children())
        for index, data in enumerate(sorted(diff_dress_data, key=lambda x : x[0])):
            # word wrap text
            for i, cell in enumerate(data): 
                if len(str(data[i])) > 2500:
                    data[i] = wrap(str(cell), 400)
                else:
                    data[i] = wrap(str(cell), 250)
            # if even row, set tag to evenrow
            # if odd row, set tag to oddrow
            if index % 2 == 0:
                table.insert(parent='', index=tk.END, values=data, tags=('evenrow',))
            else:
                table.insert(parent='', index=tk.END, values=data, tags=('oddrow',))

    elif order == 'name':
        table.delete(*table.get_children())
        for index, data in enumerate(sorted(diff_dress_data, key=lambda x : str(x[1]).lower())):
            # word wrap text
            for i, cell in enumerate(data): 
                if len(str(data[i])) > 2500:
                    data[i] = wrap(str(cell), 400)
                else:
                    data[i] = wrap(str(cell), 250)
            # if even row, set tag to evenrow
            # if odd row, set tag to oddrow
            if index % 2 == 0:
                table.insert(parent='', index=tk.END, values=data, tags=('evenrow',))
            else:
                table.insert(parent='', index=tk.END, values=data, tags=('oddrow',))

    # elif order == 'did_you_know':
    #     table.delete(*table.get_children())
    #     for index, data in enumerate(sorted(diff_dress_data, key=lambda x : x[1])):
    #     # word wrap text
    #         for i, cell in enumerate(data): 
    #             data[i] = wrap(str(cell), 250)
    #         # if even row, set tag to evenrow
    #         # if odd row, set tag to oddrow
    #         if index % 2 == 0:
    #             table.insert(parent='', index=tk.END, values=data, tags=('evenrow',))
    #         else:
    #             table.insert(parent='', index=tk.END, values=data, tags=('oddrow',))    
'''
Exports table data to Excel file
'''
def exportExcel(data, excel_columns, sheet_name):
    df = pd.DataFrame(data, columns=excel_columns)
    df.to_excel(f'{sheet_name}.xlsx', index=False)

'''
Exports table data to SQL file
'''
def exportSQL(data, sql_columns, table_name):
    with open(f'{table_name}.sql', 'w') as sql_file:
        # Create SQL CREATE TABLE statement
        create_table_query = f'CREATE TABLE IF NOT EXISTS {table_name} (\n'
        create_table_query += ', '.join(f'{column} TEXT' for column in sql_columns)
        create_table_query += '\n);\n\n'
        sql_file.write(create_table_query)

        # Create SQL INSERT INTO statement
        sql_file.write(f'INSERT INTO {table_name} ({", ".join(sql_columns)}) VALUES\n')

        # Iterate through data and write values
        for row in data:
            values = ', '.join(f"'{str(value)}'" for value in row)
            sql_file.write(f'({values}),\n')

        # Remove the trailing comma from the last line
        sql_file.seek(sql_file.tell() - 2)
        sql_file.truncate()

        # Add a semicolon to the end of the SQL script
        sql_file.write(';')
        
def diffReport():
    # helper for table item selection
    def item_select(_):
        if len(table.selection()) > 0:
            for item in table.selection():
                print(table.item(item)['values'])

    file_path = 'APIData.xlsx' # Change to path where file is located

    dress_ids = sorted(getSlideNumbers()) # gets dress IDs in entry field
    diff_dress_data = [] # data in spreadsheet that is different from API
    changed_or_new = [] # to keep track of what is changed data or new data
    api_dress_data = sorted(apiRunner(), key=lambda x : x['id']) # gets dress data from API and sorts by ID

    try:
        # gets and cleans dress data from Excel file
        sheet_dress_data = pd.read_excel(file_path)
        sheet_dress_data.dropna(subset=['id'], inplace=True) # drops any rows with na ID
        sheet_dress_data['description'].fillna('', inplace=True) # removes na/nan from description column
        sheet_dress_data['did_you_know'].fillna('', inplace=True) # removes na/nan from did_you_know column
        sheet_dress_data['description'] = sheet_dress_data['description'].astype(str).apply(openpyxl.utils.escape.unescape) # convert escaped strings to ASCII
        sheet_dress_data['did_you_know'] = sheet_dress_data['did_you_know'].astype(str).apply(openpyxl.utils.escape.unescape) # Convert escaped strings to ASCII

        # cycle through api_dress_data
        for index, api_data in enumerate(api_dress_data):
            # row of data with an ID that matches api_data ID
            row = sheet_dress_data.loc[api_data['id']-1] # row of data in spreadsheet

            # check if name, description, or did_you_know is different from the API data
            if row.loc['name'] != api_data['name']:
                diff_dress_data.append([item for item in row])
                changed_or_new.append('changed')
                continue
            if row.loc['description'] != api_data['description']:
                diff_dress_data.append([item for item in row])
                changed_or_new.append('changed')
                continue
            if row.loc['did_you_know'] != api_data['did_you_know']:
                diff_dress_data.append([item for item in row])
                changed_or_new.append('changed')
                continue

        # check for new entries in Excel sheet that do not exist in retrieved api data
        for id in dress_ids:
            if not any(data['id'] == id for data in api_dress_data) and (sheet_dress_data['id']==id).any():
                row = sheet_dress_data.loc[sheet_dress_data['id']==id]
                diff_dress_data.append([item for item in row.values[0]])
                changed_or_new.append('new')

        # find largest description field
        row_size_flag = 0
        for value in diff_dress_data:
            if len(str(value[2])) > row_size_flag:
                row_size_flag = len(str(value[2]))

        # create window to display table
        table_window = tk.Toplevel(root)
        table_window.title("Difference Report")
        table_window.geometry(f"1000x600")
        table_window.minsize(1000,600)

        # create frame to hold table
        table_frame = tk.Frame(table_window)
        table_frame.pack_propagate(False)
        table_frame.place(x=0, y=0, relwidth=1, relheight=.89, anchor="nw")

        # using style to set row height and heading colors
        style = ttk.Style()
        style.theme_use('clam')
        if row_size_flag < 500:
            style.configure('Treeview', rowheight=75)
        elif row_size_flag >= 500 and row_size_flag < 1000:
            style.configure('Treeview', rowheight=100)
        elif row_size_flag >= 1000 and row_size_flag < 2000:
            style.configure('Treeview', rowheight=150)
        elif row_size_flag >= 2000:
            style.configure('Treeview', rowheight=200)
        style.configure('Treeview.Heading', background='#848484', foreground='white')

        # vertical scrollbar
        table_scrolly = tk.Scrollbar(table_frame)
        table_scrolly.pack(side="right", fill='y')
        # horizontal scrollbar
        table_scrollx = tk.Scrollbar(table_frame, orient='horizontal')
        table_scrollx.pack(side="bottom", fill='x')

        # use ttk Treeview to create table
        table = ttk.Treeview(table_frame, yscrollcommand=table_scrolly.set, xscrollcommand=table_scrollx.set, columns=('id', 'name', 'description', 'did_you_know'), show='headings')

        # configure the scroll bars with the table
        table_scrolly.config(command=table.yview)
        table_scrollx.config(command=table.xview)

        # create the headers
        table.heading('id', text='id', command=lambda: rowOrder('id', diff_dress_data, table))
        table.heading('name', text='name', command=lambda: rowOrder('name', diff_dress_data, table))
        table.heading('description', text='description', command=lambda: rowOrder('description', diff_dress_data, table))
        table.heading('did_you_know', text='did_you_know', command=lambda: rowOrder('did_you_know', diff_dress_data, table))

        # set column variables
        table.column('id', width=75, stretch=False)
        table.column('name', width=145, stretch=False)
        table.column('description', width=1000, anchor='nw', stretch=False)
        table.column('did_you_know', anchor='nw', stretch=False, width=800)
        
        # pack table into table_frame
        table.pack(fill='both', expand=True)

        # fill table with difference report data
        for index, data in enumerate(diff_dress_data):
            # word wrap text
            for i, cell in enumerate(data): 
                if len(str(data[i])) > 2500:
                    data[i] = wrap(str(cell), 400)
                else:
                    data[i] = wrap(str(cell), 250)

            # if even row, set tag to evenrow
            # if odd row, set tag to oddrow
            if changed_or_new[index] == 'new':
                table.insert(parent='', index=tk.END, values=data, tags=('new',))
            elif changed_or_new[index] == 'changed':
                table.insert(parent='', index=tk.END, values=data, tags=('changed',))

        # alternate colors each line
        table.tag_configure('new', background='#BAFFA4')
        table.tag_configure('changed', background='#FFA5A4')

        # monitor select event on items
        table.bind('<<TreeviewSelect>>', item_select)

        column_headers = ['id', 'name', 'description', 'did_you_know'] # column headers

        # create button frame and place one table_window
        btn_frame = tk.Frame(table_window)
        btn_frame.pack(side='bottom', pady=15)
        # create buttons and pack on button_frame
        btn = tk.Button(btn_frame, text='Export SQL File', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportSQL(diff_dress_data, changed_or_new))
        btn.pack(side='left', padx=25)

        btn = tk.Button(btn_frame, text='Export to HTML', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportHTML(diff_dress_data, column_headers, 'difference_report'))
        btn.pack(side='left', padx=25)

        btn2 = tk.Button(btn_frame, text='Export Excel File', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportExcel(diff_dress_data, column_headers, 'difference_report'))
        btn2.pack(side='left', padx=25)

    except FileNotFoundError:
        tk.messagebox.showerror(title="Error in diffReport", message=f"File '{file_path}' not found.")
        print(f"File '{file_path}' not found.")
    except Exception as e:
        tk.messagebox.showerror(title="Error in diffReport", message=f'Error: {e}')
        print(f'Error: {e}')
    finally:
        diff_report_button.config(state="normal")

'''
Returns 3 google images url and put into excel sheet
'''
def googleImage():
    api_dress_data = sorted(apiRunner(), key=lambda x : x['id'])

    excel_file_name = "googleimages.xlsx"
    file_name = "googleimages.xlsx"
    count = 0

    # Rename file if the file name exists
    while os.path.exists(file_name):
        count += 1
        file_name = f"{os.path.splitext(excel_file_name)[0]}({count}).xlsx" # if file name exisit create new filename

    # Open the Excel file
    workbook = openpyxl.Workbook()
    sheet = workbook.active # active sheet (first sheet)

    # Create header for first row
    header_row = ["ID", "Name", "URL1", "URL2", "URL3"]
    sheet.append(header_row)

    # progress bar window for creating excel file
    excel_progress_window = tk.Toplevel(root)
    excel_progress_window.title('Creating Excel File')
    sw = int(excel_progress_window.winfo_screenwidth()/2 - 450/2)
    sh = int(excel_progress_window.winfo_screenheight()/2 - 70/2)
    excel_progress_window.geometry(f'450x70+{sw}+{sh}')
    excel_progress_window.resizable(False, False)
    excel_progress_window.attributes('-disable', True)
    excel_progress_window.focus()

    # progress bar custom style
    pb_style = ttk.Style()
    pb_style.theme_use('clam')
    pb_style.configure('green.Horizontal.TProgressbar', foreground='#1ec000', background='#1ec000')

    # frame to hold progress bar
    pb_frame = tk.Frame(excel_progress_window)
    pb_frame.pack()

    # progress bar
    pb = ttk.Progressbar(pb_frame, length=400, style='green.Horizontal.TProgressbar', mode='determinate', maximum=100, value=0)
    pb.pack(pady=10)

    # label for percent complete
    percent_label = tk.Label(pb_frame, text='Creating Excel File...0%')
    percent_label.pack()
    complete = 0

    url_list = []
    img_name_list = []

    # Loop through items in list 
    row_count = 2
    for item in api_dress_data:
        # Use google to search for image of each dress and return 3 image url
        word = f"{item['name']} india"
        url = 'https://www.google.com/search?q={0}&tbm=isch'.format(word)
        content = requests.get(url).content
        soup = BeautifulSoup(content, 'html.parser')
        images = soup.findAll('img')

        # Adds the dress number and name
        sheet.cell(row=row_count, column=1, value=item['id'])
        sheet.cell(row=row_count, column=2, value=item['name'])

        # Adds the 3 image urls to the excel sheet
        for j, image in enumerate(images[1:4]):
            img_src = image.get('src')
            try:
                response = requests.get(img_src, stream=True) # Sends a HTTP GET request to the url
                if response.status_code == 200:
                    content_type = response.headers['Content-Type'] # Gets the content type (ex: image/png)
                    img_extension = content_type.split('/')[-1].lower() # Gets the content after the "/" and put into lowercase

                    valid_extensions = {'jpg', 'jpeg', 'gif', 'png'} 
                    if img_extension in valid_extensions: # if url is a valid extension add to sheet
                        sheet.cell(row=row_count, column=j + 3, value=f"{img_src}.{img_extension}")

                        # if user chooses to download google images
                        if download_google_imgs.get() == 1: 
                            image_num = j+1
                            # creates directory to save images if one does not exist
                            if not os.path.exists('./google_images'):
                                os.makedirs('./google_images')

                            dress_name = item["name"].replace(" ", "_") # changes all space in dress name to "_"
                            new_dress_name = re.sub(r'[\\/:*?"<>|]', '', dress_name) # removes special characters
                            new_dress_name = new_dress_name.replace('\n\n', '') # removes "\n\n"

                            image_name = f'{item["id"]}_{new_dress_name}_{image_num}.{img_extension}' # sets the image file name
                            img_name_list.append(image_name)
                            url_list.append(image.get('src'))
            except Exception as e:
                print(f"Error: {e}")
        row_count+=1
        complete += 1
        pb['value'] = (complete/len(api_dress_data))*100 # calculate percentage of data retrieved
        percent_label.config(text=f'Creating Excel File...{int(pb["value"])}%') # update completion percent label

    excel_progress_window.destroy() # close progress bar window
    workbook.save(file_name)

    if download_google_imgs.get() == 1: 
        # progress bar window for downloading google image
        google_image_progress_window = tk.Toplevel(root)
        google_image_progress_window.title('Downloading Images')
        sw = int(google_image_progress_window.winfo_screenwidth()/2 - 450/2)
        sh = int(google_image_progress_window.winfo_screenheight()/2 - 70/2)
        google_image_progress_window.geometry(f'450x70+{sw}+{sh}')
        google_image_progress_window.resizable(False, False)
        google_image_progress_window.attributes('-disable', True)
        google_image_progress_window.focus()

        # progress bar custom style
        pb_style = ttk.Style()
        pb_style.theme_use('clam')
        pb_style.configure('green.Horizontal.TProgressbar', foreground='#1ec000', background='#1ec000')

        # frame to hold progress bar
        pb_frame = tk.Frame(google_image_progress_window)
        pb_frame.pack()

        # progress bar
        pb = ttk.Progressbar(pb_frame, length=400, style='green.Horizontal.TProgressbar', mode='determinate', maximum=100, value=0)
        pb.pack(pady=10)

        # label for percent complete
        percent_label = tk.Label(pb_frame, text='Downloading Images...0%')
        percent_label.pack()
        
        threads= [] # working threads

        # spins up 10 threads at a time and calls downloadImages with url and image name
        with ThreadPoolExecutor(max_workers=10) as exec:
            for index, url in enumerate(url_list):
                threads.append(exec.submit(downloadImages, "google_images", url, img_name_list[index]))

            complete = 0 # number of threads that have finished
            for task in as_completed(threads):
                complete += 1
                pb['value'] = (complete/len(url_list))*100 # calculate percentage of images downloaded
                percent_label.config(text=f'Downloading Images...{int(pb["value"])}%') # update completion percent label
    
        google_image_progress_window.destroy() # close progress bar window

    openFile(file_name)
    google_image_search_button.config(state='normal')

'''
Performs word analysis on given dress IDs
'''
def wordAnalysis():
    # helper for table item selection
    def item_select(_):
        for item in table.selection():
            print(table.item(item)['values'])

    word_analysis_data = [] # data from word analysis
    api_dress_data = sorted(apiRunner(), key=lambda x : x['id']) # gets dress data from API and sorts by ID

    try:
        # cycle through api_dress_data for word analysis
        for dress_data in api_dress_data:
            noun_count = 0 # number of nouns in text
            adjective_count = 0 # number of adjectives in text

            # concatenation of description and did_you_know
            text = f'{str(dress_data["description"])} {str(dress_data["did_you_know"])}'
            # TextBlob word analysis
            blob = TextBlob(text)

            # cycle through key:value pairs of TextBlob analysis to get noun and adjective count
            for k,v in blob.tags:
                if v == 'NN' or v == 'NNS' or v == 'NNP' or v == 'NNPS':
                    noun_count += 1
                elif v == 'JJ' or v == 'JJR' or v == 'JJS':
                    adjective_count += 1

            ease = textstat.flesch_reading_ease(text)
            kincaid = textstat.flesch_kincaid_grade(text)
            readability = textstat.automated_readability_index(text)

            # data to be displayed in table
            word_analysis_data.append([dress_data['id'], dress_data['name'], len(str(dress_data['description']).strip(string.punctuation).split()), 
                                       len(str(dress_data['did_you_know']).strip(string.punctuation).split()), str(noun_count), str(adjective_count),
                                       str(ease), str(kincaid), str(readability)])

        # create window to display table
        table_window = tk.Toplevel(root)
        table_window.title("Word Analysis")
        table_window.geometry(f"1000x600")
        table_window.minsize(1000,600)

        # create frame to hold table
        table_frame = tk.Frame(table_window)
        table_frame.pack_propagate(False)
        table_frame.place(x=0, y=0, relwidth=1, relheight=.89, anchor="nw")

        # using style to set row height and heading colors
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Treeview', rowheight=50)
        style.configure('Treeview.Heading', background='#848484', foreground='white')

        # vertical scrollbar
        table_scrolly = tk.Scrollbar(table_frame)
        table_scrolly.pack(side="right", fill='y')
        # horizontal scrollbar
        table_scrollx = tk.Scrollbar(table_frame, orient='horizontal')
        table_scrollx.pack(side="bottom", fill='x')

        # use ttk Treeview to create table
        table = ttk.Treeview(table_frame, yscrollcommand=table_scrolly.set, xscrollcommand=table_scrollx.set,
                            columns=('id', 'name', 'description_word_count', 'did_you_know_word_count', 'total_noun_count', 'total_adjective_count', 'reading_ease', 'kincaid_grade', 'readability_index'), show='headings')

        # configure the scroll bars with the table
        table_scrolly.config(command=table.yview)
        table_scrollx.config(command=table.xview)

        # create the headers
        table.heading('id', text='id')
        table.heading('name', text='name')
        table.heading('description_word_count', text='description_word_count')
        table.heading('did_you_know_word_count', text='did_you_know_word_count')
        table.heading('total_noun_count', text='total_noun_count')
        table.heading('total_adjective_count', text='total_adjective_count')
        table.heading('reading_ease', text='reading_ease')
        table.heading('kincaid_grade', text='kincaid_grade')
        table.heading('readability_index', text='readability_index')

        # set column variables
        table.column('id', width=75, stretch=False)
        table.column('name', width=145, stretch=False)
        table.column('description_word_count', width=200, anchor='center', stretch=False)
        table.column('did_you_know_word_count', width=200, anchor='center', stretch=False)
        table.column('total_noun_count', width=200, anchor='center', stretch=False)
        table.column('total_adjective_count', width=200, anchor='center', stretch=False)
        table.column('reading_ease', width=200, anchor='center', stretch=False)
        table.column('kincaid_grade', width=200, anchor='center', stretch=False)
        table.column('readability_index', width=200, anchor='center', stretch=False)
        
        # pack table into table_frame
        table.pack(fill='both', expand=True)

        # fill table with difference report data
        for index, data in enumerate(word_analysis_data):
            # word wrap text
            for i, cell in enumerate(data): 
                data[i] = wrap(str(cell), 250)

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

        # column headers
        column_headers = ['id', 'name', 'description_word_count', 'did_you_know_word_count', 'total_noun_count', 'total_adjective_count', 'reading_ease', 'kincaid_grade', 'readability_index']

        # create button frame and place one table_window
        btn_frame = tk.Frame(table_window)
        btn_frame.pack(side='bottom', pady=15)

        # create buttons and pack on button_frame
        btn = tk.Button(btn_frame, text='Export to HTML', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportHTML(word_analysis_data, column_headers, 'word_analysis_report'))
        btn.pack(side='left', padx=25)

        btn2 = tk.Button(btn_frame, text='Export Excel File', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportExcel(word_analysis_data, column_headers, 'word_analysis_report'))
        btn2.pack(side='left', padx=25)

    except Exception as e:
        tk.messagebox.showerror(title="Error in wordAnalysis", message=f'Error: {e}')
        print(f'Error: {e}')
    finally:
        word_analysis_button.config(state='normal')
'''
Generate Wiki Link
'''
def generateWikiLink():
    # set user agent
    USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3"
    wiki_wiki = wikipediaapi.Wikipedia("en",headers={"User-Agent": USER_AGENT})

    # gather all dress data from api
    wiki_link_data = apiRunner()
    wiki_data = []
    for item in wiki_link_data:
        try:
            page = wiki_wiki.page(item["name"])
            if page.exists():
                item["wiki_page_link"] = page.fullurl
                wiki_data.append([item['id'], item['name'],item['wiki_page_link']])
        except Exception as e:
            print(f"Error retrieving Wikipedia data for {item['name']}: {e}")
    row_size_flag = 0

    # create window to display table
    table_window = tk.Toplevel(root)
    table_window.title("Wiki Page Link")
    table_window.geometry(f"1000x600")
    table_window.minsize(1000,600)

    # create frame to hold table
    table_frame = tk.Frame(table_window)
    table_frame.pack_propagate(False)
    table_frame.place(x=0, y=0, relwidth=1, relheight=.89, anchor="nw")

    # using style to set row height and heading colors
    style = ttk.Style()
    style.theme_use('clam')
    if row_size_flag <= 500:
        style.configure('Treeview', rowheight=75)
    elif row_size_flag > 500 and row_size_flag < 1000:
        style.configure('Treeview', rowheight=100)
    elif row_size_flag >= 1000:
        style.configure('Treeview', rowheight=150)
    elif row_size_flag >= 2500:
        style.configure('Treeview', rowheight=200)
    style.configure('Treeview.Heading', background='#848484', foreground='white')
    
    # vertical scrollbar
    table_scrolly = tk.Scrollbar(table_frame)
    table_scrolly.pack(side="right", fill='y')
    
    # horizontal scrollbar
    table_scrollx = tk.Scrollbar(table_frame, orient='horizontal')
    table_scrollx.pack(side="bottom", fill='x')

    # use ttk Treeview to create table
    table = ttk.Treeview(table_frame, yscrollcommand=table_scrolly.set, xscrollcommand=table_scrollx.set, columns=('id', 'name', 'wiki_page_link'), show='headings')

    # configure the scroll bars with the table
    table_scrolly.config(command=table.yview)
    table_scrollx.config(command=table.xview)
    
    # create the headers
    table.heading('id', text='id')
    table.heading('name', text='name')
    table.heading('wiki_page_link', text='wiki_page_link')
    
    # set column variables
    table.column('id', width=75, stretch=False)
    table.column('name', width=145, stretch=False)
    table.column('wiki_page_link', width=850, anchor='nw', stretch=False)
        
    # pack table into table_frame
    table.pack(fill='both', expand=True)

    # fill table with difference report data
    for index, data in enumerate(wiki_data):
    
        # word wrap text
        values_copy = list(data)  # Create a copy of the list

        for i, cell in enumerate(values_copy):
            values_copy[i] = wrap(str(cell), 250)

        # if even row, set tag to evenrow
        if index % 2 == 0:
            table.insert(parent='', index=tk.END, values=tuple(values_copy), tags=('evenrow',))
    
        # if odd row, set tag to oddrow
        else:
            table.insert(parent='', index=tk.END, values=tuple(values_copy), tags=('oddrow',))

    # alternate colors each line
    table.tag_configure('evenrow', background='#e8f3ff')
    table.tag_configure('oddrow', background='#f7f7f7')

    # monitor select event on items
    table.bind('<<TreeviewSelect>>')
    
    # create button frame and place one table_window
    btn_frame = tk.Frame(table_window)
    btn_frame.pack(side='bottom', pady=15)
    
    # create buttons and pack on button_frame
    btn = tk.Button(btn_frame, text='Generate SQL File', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff",command=lambda: exportSQL(wiki_data,excel_columns,'WIKI_LINK'))
    btn.pack(side='left', padx=25)
    
    # column headers for exporting Excel
    excel_columns = ['id', 'name', 'wiki_page_link'] 
    btn2 = tk.Button(btn_frame, text='Generate Excel File', font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: exportExcel(wiki_data, excel_columns, 'WIKI_LINK'))
    btn2.pack(side='left', padx=25)

'''
Spins up new thread to run generateUpdate function
'''
def startGenerateBookThread():
    book_gen_generate_button.config(state="disabled")
    generate_thread = threading.Thread(target=generateBook)
    generate_thread.start()

'''
Spins up new thread to run diffReport function
'''
def startDiffReportThread():
    diff_report_button.config(state='disabled')
    diff_report_thread = threading.Thread(target=diffReport)
    diff_report_thread.start()
'''
Spins up new thread to run wordAnalysis function
'''
def startWordAnalysisThread():
    word_analysis_button.config(state='disabled')
    word_analysis_thread = threading.Thread(target=wordAnalysis)
    word_analysis_thread.start()
def startGenerateWikiLinkThread():
    wiki_link_gen_button.config(state='disabled')
    wiki_link_thread = threading.Thread(target=generateWikiLink)
    wiki_link_thread.start()
'''
Spins up new thread to run googleImage function
'''
def startGoogleImageThread():
    google_image_search_button.config(state='disabled')
    google_image_search_thread = threading.Thread(target=googleImage)
    google_image_search_thread.start()

'''
Launch help site when user clicks Help button
'''
def launchHelpSite():
    # create help site
    with open('help.html', 'w') as file:
        file.write('<!DOCTYPE html>\n<html>\n<head>\n\t<meta charset="utf8">\n\t<title>abcd Help</title>\n</head>\n<body>\n\t\t<h1 style="text-align: center;">Welcome to the help site</h1>\n</body>\n</html>\n')
    
    # open help site
    webbrowser.open('help.html')

'''
Raise selected frame to the top
'''
def raiseFrame(frame):
    if frame == 'main_frame':
        main_frame.tkraise()
        root.title("Project ABCD Admin Panel")
    elif frame == 'book_gen_frame':
        book_gen_frame.tkraise()
        text_field_label.tkraise()
        text_field.tkraise()
        root.title("Project ABCD Book Generation")
    elif frame == 'diff_report_frame':
        diff_report_frame.tkraise()
        text_field_label.tkraise()
        text_field.tkraise()
        root.title("Project ABCD Difference Report")
    elif frame == 'word_analysis_frame':
        word_analysis_frame.tkraise()
        text_field_label.tkraise()
        text_field.tkraise()
        root.title("Project ABCD Word Analysis")
    elif frame == 'wiki_link_frame':
        wiki_link_frame.tkraise()
        text_field_label.tkraise()
        text_field.tkraise()
        root.title("Project ABCD Wiki Link")

#--------------------------------Main Frame-----------------------------------------------------------------------------------------------
#-----------------------------------------------------------------------------------------------------------------------------------------
# configure grid to fill extra space and center
tk.Grid.rowconfigure(root, 0, weight=1)
tk.Grid.columnconfigure(root, 0, weight=1)

# main frame
main_frame = tk.Frame(root, width=1000, height=600)
main_frame.pack_propagate(False)
main_frame.grid(row=0, column=0, sticky='news')

#--------------------------------Main Title-----------------------------------------------------------------------------------------------
# Create a label widget
title_label = tk.Label(main_frame, text="Project ABCD\nMain Menu",  font=('Arial', 20))
title_label.pack(pady=100)

#--------------------------------Main Buttons-----------------------------------------------------------------------------------------------
# Create buttons widget
## Button settings
main_button_frame = tk.Frame(main_frame)
main_button_frame.place(relx=.5, rely=.5, anchor='center')
button_width = 18
button_height = 3
button_bgd_color = "#007FFF"
button_font_color = "#ffffff"

# Generate Book: Gets selected dress from API and import into ppt
generate_book_button = tk.Button(main_button_frame, text="Generate Book", font=LABEL_FONT, width=button_width, height=button_height, bg=button_bgd_color, fg=button_font_color, command=lambda: raiseFrame('book_gen_frame'))
generate_book_button.pack(side="left", padx=10)

# Diff Report: Create a SQL file of dresses that got changed from excel sheet byt comparing to API
diff_report_button = tk.Button(main_button_frame, text="Difference Report", font=LABEL_FONT, width=button_width, height=button_height, bg=button_bgd_color, fg=button_font_color, command=lambda: raiseFrame('diff_report_frame'))
diff_report_button.pack(side="left", padx=10, anchor='center')

# Generate Book: Get selected dress that user input & put into a table (ID, Name, Description Count, DYK Count, Total Nouns Count, Total Adjectives Count)
word_analysis_report_button = tk.Button(main_button_frame, text="Word Analysis Report", font=LABEL_FONT, width=button_width, height=button_height, bg=button_bgd_color, fg=button_font_color, command=lambda: raiseFrame('word_analysis_frame'))
word_analysis_report_button.pack(side="left", padx=10)

## Wiki Link:
wiki_link_button = tk.Button(main_button_frame, text="Wiki Link", font=LABEL_FONT, width=button_width, height=button_height, bg=button_bgd_color, fg=button_font_color, command=lambda: raiseFrame('wiki_link_frame'))
wiki_link_button.pack(side="left", padx=10)

main_button_frame2 = tk.Frame(main_frame)
main_button_frame2.place(relx=.5, rely=.7, anchor='center')

## Google Images: Create an Excel file with 3 image links to the selected dresses
google_image_button = tk.Button(main_button_frame2, text="Google Image", font=LABEL_FONT, width=button_width, height=button_height, bg=button_bgd_color, fg=button_font_color, command=lambda: raiseFrame('google_image_frame'))
google_image_button.pack(side="left", padx=50)


#--------------------------------Book Gen Frame---------------------------------------------------------------------------------------------
#-------------------------------------------------------------------------------------------------------------------------------------------
# book gen frame
book_gen_frame = tk.Frame(root, width=1000, height=600)
book_gen_frame.pack_propagate(False)
book_gen_frame.grid(row=0, column=0, sticky='news')

#--------------------------------Text Field-----------------------------------------------------------------------------------------------
# text field label
text_field_label = tk.Label(root, text="Dress Numbers:", font=LABEL_FONT)
text_field_label.place(x=25, y=72.5)

# text field
text_field = tk.Text(root)
text_field.place(x=175, y=10, relwidth=.8, height=135)

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
layout.set(4)

# layout frame
layout_frame = tk.Frame(book_gen_frame)
# layout radio buttons
layout_radio4 = tk.Radiobutton(layout_frame, text="Picture on Left - Text on right - Single Page - Portrait Mode",
                                font=MAIN_FONT, variable=layout, value=4)
layout_radio1 = tk.Radiobutton(layout_frame, text="Picture on Left Page - Text on Right Page - Two Page Mode - Portrait Mode",
                                font=MAIN_FONT, variable=layout, value=1)
layout_radio2 = tk.Radiobutton(layout_frame, text="Picture on Right - Text on Left - Single Page - Landscape Mode",
                                font=MAIN_FONT, variable=layout, value=2)
layout_radio3 = tk.Radiobutton(layout_frame, text="Picture on Left - Text on Right - Single Page - Landscape Mode",
                                font=MAIN_FONT, variable=layout, value=3)

# pack radio buttons into layout frame
layout_radio4.pack(anchor="nw")
layout_radio1.pack(anchor="nw")
layout_radio2.pack(anchor="nw")
layout_radio3.pack(anchor="nw")

# place layout frame on main frame
layout_frame.place(x=175, y=150, width=800)

# layout radio buttons label
layout_label = tk.Label(book_gen_frame, text="Layout:", font=LABEL_FONT)
layout_label.place(x=25, y=170)

# separator line
separator1 = ttk.Separator(book_gen_frame)
separator1.place(x=175, y=150, relwidth=.8)

#--------------------------------Sort Radio Buttons---------------------------------------------------------------------------------------
# sort variable
sort_order = tk.IntVar()
sort_order.set(1)

# sort frame
sort_frame = tk.Frame(book_gen_frame)

# sort radio buttons
sort_radio1 = tk.Radiobutton(sort_frame, text="By Name", font=MAIN_FONT, variable=sort_order, value=1)
sort_radio2 = tk.Radiobutton(sort_frame, text="By ID", font=MAIN_FONT, variable=sort_order, value=2)
sort_radio3 = tk.Radiobutton(sort_frame, text="By Input Order", font=MAIN_FONT, variable=sort_order, value=3)

# pack radio buttons into sort frame
sort_radio1.pack(side="left")
sort_radio2.pack(side="left")
sort_radio3.pack(side="left")

# place sort frame on main frame
sort_frame.place(x=175, y=265, width=800)

# sort radio buttons label
sort_label = tk.Label(book_gen_frame, text="Sort Order:", font=LABEL_FONT)
sort_label.place(x=25, y=265)

# separator line
separator2 = ttk.Separator(book_gen_frame)
separator2.place(x=175, y=265, relwidth=.8)


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
preferences_frame = tk.Frame(book_gen_frame)

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
preferences_label = tk.Label(book_gen_frame, text="Preferences:", font=LABEL_FONT)
preferences_label.place(x=25, y=350)

# separator line
separator3 = ttk.Separator(book_gen_frame)
separator3.place(x=175, y=295, relwidth=.8)

#--------------------------------Numbering Radio Buttons--------------------------------------------------------------------------------------
# numbering variable
numbering = tk.IntVar()
numbering.set(1)

# numbering frame
numbering_frame = tk.Frame(book_gen_frame)
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
numbering_label = tk.Label(book_gen_frame, text="Numbering:", font=LABEL_FONT)
numbering_label.place(x=25, y=445)

# separator line
separator4 = ttk.Separator(book_gen_frame)
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
download_imgs.set(0)
# Generate book from local Excel sheet
gen_local = tk.IntVar()
gen_local.set(0)

# translate frame
check_button_frame = tk.Frame(book_gen_frame)
# translate check button
translate_checkbutton = tk.Checkbutton(check_button_frame, text="Translate to:", font=MAIN_FONT, variable=translate, onvalue=1, offvalue=0)
# language options
language_options = tk.OptionMenu(check_button_frame, language, *LANGUAGES)
# download images
download_images = tk.Checkbutton(check_button_frame, text="Download Images", font=MAIN_FONT, variable=download_imgs, onvalue=1, offvalue=0, command=lambda: gen_local.set(0))
# generate book from local Excel sheet
generate_from_local = tk.Checkbutton(check_button_frame, text="Generate from Local", font=MAIN_FONT, variable=gen_local, onvalue=1, offvalue=0, command=lambda: download_imgs.set(0))

# pack translate options into translate frame
translate_checkbutton.pack(side="left")
language_options.pack(side="left")
download_images.pack(side="left")
generate_from_local.pack(side="left")
# place translate frame on main frame
check_button_frame.place(x=175, y=495)

# separator line
separator5 = ttk.Separator(book_gen_frame)
separator5.place(x=175, y=495, relwidth=.8)

#--------------------------------Book Gen Buttons--------------------------------------------------------------------------------------
# button frame
book_gen_button_frame = tk.Frame(book_gen_frame)
# generate button
book_gen_generate_button = tk.Button(book_gen_button_frame, text="Generate", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startGenerateBookThread)
# help button
book_gen_help_button = tk.Button(book_gen_button_frame, text="Help", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=launchHelpSite)
# upload button
book_gen_back_button = tk.Button(book_gen_button_frame, text="Back", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: raiseFrame('main_frame'))
# pack buttons into button frame
book_gen_generate_button.pack(side="left", padx=35)
book_gen_help_button.pack(side="left")
book_gen_back_button.pack(side="left", padx=30)

# place button frame on main frame
book_gen_button_frame.pack(side="bottom", pady=10)


#--------------------------------Diff Report Frame-----------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
diff_report_frame = tk.Frame(root, width=1000, height=600)
diff_report_frame.pack_propagate(False)
diff_report_frame.grid(row=0, column=0, sticky='news')

#--------------------------------Diff Report Buttons-----------------------------------------------------------------------------------------------
# button frame
diff_report_button_frame = tk.Frame(diff_report_frame)
# difference report button
diff_report_button = tk.Button(diff_report_button_frame, text="Diff Report", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startDiffReportThread)
# back button
diff_back_button = tk.Button(diff_report_button_frame, text="Back", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: raiseFrame('main_frame'))
# pack buttons into button frame
diff_report_button.pack(side="left", padx=35)
diff_back_button.pack(side="left", padx=30)

# place button frame on diff report frame
diff_report_button_frame.pack(side="bottom", pady=10)


#--------------------------------Word Analysis Frame-----------------------------------------------------------------------------------------------
#--------------------------------------------------------------------------------------------------------------------------------------------------
word_analysis_frame = tk.Frame(root, width=1000, height=600)
word_analysis_frame.pack_propagate(False)
word_analysis_frame.grid(row=0, column=0, sticky='news')

#--------------------------------Word Analysis Buttons-----------------------------------------------------------------------------------------------
# button frame
word_analysis_button_frame = tk.Frame(word_analysis_frame)
# word analysis button
word_analysis_button = tk.Button(word_analysis_button_frame, text="Word Analysis", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startWordAnalysisThread)
# back button
word_analysis_back_button = tk.Button(word_analysis_button_frame, text="Back", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: raiseFrame('main_frame'))
# pack buttons into button frame
word_analysis_button.pack(side="left", padx=35)
word_analysis_back_button.pack(side="left", padx=30)

# place button frame on word analysis frame
word_analysis_button_frame.pack(side="bottom", pady=10)

#--------------------------------Wiki Link Frame-----------------------------------------------------------------------------------------------
#------------------------------------------------------------------------------------------------------------------------------------------------
wiki_link_frame = tk.Frame(root, width=1000, height=600)
wiki_link_frame.pack_propagate(False)
wiki_link_frame.grid(row=0, column=0, sticky='news')
 
#--------------------------------Wiki Link Buttons--------------------------------------------------------------------------------------
# button frame
wiki_link_gen_button_frame = tk.Frame(wiki_link_frame)

wiki_link_gen_button = tk.Button(wiki_link_gen_button_frame, text="Generate", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startGenerateWikiLinkThread)
wiki_link_help_button = tk.Button(wiki_link_gen_button_frame, text="Help", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=launchHelpSite)
wiki_link_back_button = tk.Button(wiki_link_gen_button_frame, text="Back", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=lambda: raiseFrame('main_frame'))

# pack buttons into button frame
wiki_link_gen_button.pack(side="left", padx=35)
wiki_link_help_button.pack(side="left")
wiki_link_back_button.pack(side="left", padx=30)

# place button frame on word analysis frame
wiki_link_gen_button_frame.pack(side="bottom", pady=10)

# raise main_frame to start
main_frame.tkraise()

# main gui loop
root.mainloop()
# End