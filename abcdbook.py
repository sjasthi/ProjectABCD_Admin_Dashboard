import sys, io, requests, threading, webbrowser, urllib, os, platform
import tkinter as tk
from tkinter import ttk
from pptx import Presentation
import pptx.util
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from PIL import Image

# pip install googletrans
# pip install googletrans==4.0.0-rc1
import googletrans

# Set sys.stdout to use utf-8 encoding
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

MAIN_FONT = ("helvetica", 12)
LABEL_FONT = ("helvetica bold", 14)
LANGUAGES = [
    "Telugu",
    "Hindi",
    "Spanish"
]

# Creates preferences dictionary from preferences.txt
try:
    with open("preferences.txt", "r", encoding="utf8") as file:
        lines = file.readlines()
        preferences = {}

        for line in lines:
            key, value = line.split('=')
            preferences[key.strip()] = value.strip().replace('“', '').replace('”', '').replace('"', '').replace("'", '')
except FileNotFoundError:
    print("No preferences.txt file exists")

'''
Fetch the dress data from API
'''
def getDressInfoFromAPI(sorted_dress_ids):
    # get dress data from API
    dress_data = []
    for id_number in sorted_dress_ids:
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
    # if translate.get() != 1:
    #     return text
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

    # gather all dress data from api
    dress_data = getDressInfoFromAPI(dress_ids)
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
            picture = image_slide.shapes.add_picture(f'./images/{dress_info["image_url"]}', left, top, width, height)

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
            picture = slide.shapes.add_picture(f'./images/{dress_info["image_url"]}', left, top, width, height)

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
            picture = slide.shapes.add_picture(f'./images/{dress_info["image_url"]}', left, top, width, height)

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
Spins up new thread to run generateUpdate method
'''
def startThread():
    generate_button.config(state="disabled")
    generate_thread = threading.Thread(target=generateUpdate)
    generate_thread.start()

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
Main method  
'''
def main():
    # global variables to be used
    global layout, sort_order, numbering, translate, language, text_field, generate_button, text_size_var
    global title_size_var, subtitle_size_var, text_font_var, title_font_var, subtitle_font_var, pic_width_var, pic_height_var
    

    #--------------------------------Main Frame-----------------------------------------------------------------------------------------------
    # main frame
    main_frame = tk.Frame(root, width=1000, height=600)
    main_frame.pack_propagate(False)
    main_frame.pack()


    #--------------------------------Text Field-----------------------------------------------------------------------------------------------
    # text field
    text_field = tk.Text(main_frame)
    text_field.place(x=175, y=10, width=800, height=135)

    # text field label
    text_field_label = tk.Label(main_frame, text="Dress Numbers:", font=LABEL_FONT)
    text_field_label.place(x=25, y=72.5)

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


    #--------------------------------Translate Check Button--------------------------------------------------------------------------------------
    # translate check variable
    translate = tk.IntVar()
    translate.set(0)
    # language options variable
    language = tk.StringVar()
    language.set(LANGUAGES[0])

    # translate frame
    translate_frame = tk.Frame(main_frame)
    # translate check button
    translate_checkbutton = tk.Checkbutton(translate_frame, text="Translate to:", font=MAIN_FONT, variable=translate, onvalue=1, offvalue=0)
    # language options
    language_options = tk.OptionMenu(translate_frame, language, *LANGUAGES)
    # pack translate options into translate frame
    translate_checkbutton.pack(side="left")
    language_options.pack(side="left")
    # place translate frame on main frame
    translate_frame.place(x=175, y=495)

    # separator line
    separator5 = ttk.Separator(main_frame)
    separator5.place(x=175, y=495, relwidth=.8)


    #--------------------------------Generate and Help Buttons--------------------------------------------------------------------------------------
    # button frame
    button_frame = tk.Frame(main_frame)
    # generate button
    generate_button = tk.Button(button_frame, text="Generate", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=startThread)
    # help button
    help_button = tk.Button(button_frame, text="Help", font=LABEL_FONT, width=25, height=1, bg="#007FFF", fg="#ffffff", command=launchHelpSite)
    # pack buttons into button frame
    generate_button.pack(side="left", padx=35)
    help_button.pack(side="left")
    # place button frame on main frame
    button_frame.pack(side="bottom", pady=10)


    # main gui loop
    root.mainloop()

if __name__ == '__main__':
    root = tk.Tk()
    root.title("Project ABCD Book Compiler")
    root.geometry("1000x600")
    root.minsize(1000,600)
    main()
