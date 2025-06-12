import os
import streamlit as st
import json
from pathlib import Path
from PIL import Image
from deep_translator import GoogleTranslator

import img2pdf
from pypdf import PdfReader, PdfWriter, Transformation, PageObject, PaperSize
from pypdf.generic import RectangleObject
from pypdf.annotations import FreeText

# This is a generalized application to present PowerPoint slides and notes as slideshow through Streamlit.
# You can adapt the script with header and path to a specific presentation. To do this, just replace the initial informations below.
# If the presentation part of an multipage app, you need to remove two parts that are clearly marked in the following script (around lines 205 and 313)

#####################
# ADAPT FROM HERE ###
#####################

# Folder to your presentation
presentation_folder = "slides/SlideJet_Overview"

# Header and subheader
header_text = 'SlideJet Presentation'
subheader_text = 'Overview and Demo'

######################
# ADAPT UNTIL HERE ###
######################


# --- Header content Copyright ---
year = 2025 
authors = {
    "Thomas Reimann": [1],  # Author 1 belongs to Institution 1
    "Nils Wallenberg": [2],  # Author 2 also belongs to Institution 1
}
institutions = {
    1: "TU Dresden",
    2: "University of Gothenburg"
}
author_list = [f"{name}{''.join(f'<sup>{i}</sup>' for i in idxs)}" for name, idxs in authors.items()]
institution_text = " | ".join([f"<sup>{i}</sup> {inst}" for i, inst in institutions.items()])


# Functions

def patch_translations_if_missing(slide_data, target_lang, translate_func):
    """Translate only slides that do not yet have translated notes."""
    if not target_lang:
        return slide_data

    for slide in slide_data:
        if "translated_notes" not in slide:
            try:
                slide["translated_notes"] = translate_func(slide["notes"], target_lang)
            except Exception as e:
                slide["translated_notes"] = f"[Translation failed: {e}]"
    return slide_data


def generate_pdf(slides, img_folder, pres_folder, trans_lan, with_notes=False, text='Download pdf (no notes)'):
    # Convert all files ending in .png inside a directory
    imgs = []
    for i in range(len(slides)):
        fname = slides[i]['image'].split('/')[1]        
        path = os.path.join(img_folder, fname)
        imgs.append(path)

    with open(pres_folder + '/presentation.pdf','wb') as f:
        f.write(img2pdf.convert(imgs))

    # Rotate from landscape to portrait
    pdf_as_portrait(pres_folder + '/presentation.pdf', len(slides))

    # Add notes
    if with_notes:
        slide_data = patch_translations_if_missing(slides, trans_lan, translate_notes)
        add_notes(pres_folder + '/presentation.pdf', slides, trans_lan)


    # Open file in memory to be able to save with download button
    with open(pres_folder + '/presentation.pdf', 'rb') as pdf_file:
        PDFbyte = pdf_file.read()

    # Add download button
    with pcol3:
        # Save as presentation.pdf
        st.download_button(
            label=text,
            data=PDFbyte,
            file_name='presentation.pdf',
            mime='application/octet-stream',
            icon=':material/download:',
            type='primary'
        )

def add_notes(pdf_path, slide_data, target_lang):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)

    for i in range(len(slide_data)):
        notes_to_add = slide_data[i]['notes']
        
        if target_lang and "translated_notes" in slide_data[i]:
            notes_to_add += "\n\n" + slide_data[i]["translated_notes"]
        elif target_lang:
            trans_notes = translate_notes(notes_to_add, target_lang)
            notes_to_add += "\n\n" + trans_notes

        annotation = FreeText(
            text=notes_to_add,
            rect=(60, 25, 0.9 * reader.pages[i].mediabox.width, reader.pages[i].mediabox.height * 0.5),
            font="Arial",
            bold=False,
            italic=False,
            font_size="11pt",
            font_color="#000000",
            border_color="#FFFFFF",
            background_color="#FFFFFF",
        )
        annotation.flags = 4
        writer.add_annotation(page_number=i, annotation=annotation)

    with open(pdf_path, 'wb') as fp:
        writer.write(fp)

def pdf_as_portrait(pdf_path, nr_pages):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for i in range(nr_pages):
        page = reader.pages[i]

        # A4 size in points
        A4_w = PaperSize.A4.width
        A4_h = PaperSize.A4.height

        # Original slide dimensions
        h = float(page.mediabox.height)
        w = float(page.mediabox.width)

        # Define margin (2.5 cm = ~71 pt)
        margin_cm = 2.5
        margin_pt = margin_cm * 72 / 2.54

        # Compute usable width and height
        usable_width = A4_w - 2 * margin_pt
        usable_height = A4_h - margin_pt  # Only top margin; allow flexible bottom

        # Compute scale to fit slide within usable area
        scale_factor = min(usable_width / w, usable_height / h)

        # Compute offsets for positioning
        x_offset = (A4_w - w * scale_factor) / 2
        y_offset = A4_h - h * scale_factor - margin_pt  # From bottom up

        # Apply transformation
        transform = Transformation().scale(scale_factor, scale_factor).translate(x_offset, y_offset)
        page.add_transformation(transform)

        # Create A4 blank page
        page_A4 = PageObject.create_blank_page(width=A4_w, height=A4_h)

        # Ensure cropbox/mediabox match A4
        page.cropbox = RectangleObject((0, 0, A4_w, A4_h))
        page.mediabox = RectangleObject((0, 0, A4_w, A4_h))

        # Merge original onto A4 page
        page_A4.merge_page(page)

        writer.add_page(page_A4)

    # Save updated PDF
    with open(pdf_path, 'wb') as fp:
        writer.write(fp)

def scale_pdf(pdf_path, nr_pages):
    # Read the input
    reader = PdfReader(pdf_path)
    writer = PdfWriter()
    for i in range(nr_pages):
        page = reader.pages[i]

        # Scale
        # page.scale_by(0.5)
        op = Transformation().scale(sx=0.75, sy=0.75)
        page.add_transformation(op)

        # Write the result to a file        
        writer.add_page(page)
    with open(pdf_path, 'wb')as fp:
        writer.write(fp)

@st.cache_data(show_spinner=False)
def translate_notes(text, target_lang):
    try:
        return GoogleTranslator(source='auto', target=target_lang).translate(text)
    except Exception as e:
        return f"[Translation failed: {e}]"

##########################################
# PART OF MULTIPAGEAPP? REMOVE START 1 ###
##########################################

# !! IF THE PRESENTATION IS PART OF A MULTIPAGE APP - THIS NEEDS TO BE REMOVED !!
# --- MUST be first: layout setup ---
if "layout_choice" in st.session_state:
    st.session_state.layout_choice_SJ = st.session_state.layout_choice  # use app-wide layout
elif "layout_choice_SJ" not in st.session_state:
    st.session_state.layout_choice_SJ = "centered"  # fallback

st.set_page_config(
    page_title="SlideJet - Present",
    page_icon="🚀",
    layout=st.session_state.layout_choice_SJ
    )
########################################
# PART OF MULTIPAGEAPP? REMOVE END 1 ###
########################################
    
# --- Streamlit App Content ---
st.title("Presentation Slides")
st.header(f':red-background[{header_text}]')
st.subheader(subheader_text, divider='red')

st.markdown(""" 
    **About the SlideJet presentation:** _You can navigate the slides using the +/- buttons or by entering a slide number. Speaker notes can be translated into your preferred language; the original text remains visible. Use the layout toggle to switch to a horizontal view for better device compatibility._
        """)

# Define the default folder - the structure is fix and provided by the convert_ppt_slides.py application
JSON_file = os.path.join(presentation_folder, "slide_data.json")
images_folder = os.path.join(presentation_folder, "images")

# Language selection
# --- Language options with flags ---
languages = {
    "🇬🇧 English": "en",
    "🇪🇸 Spanish": "es",
    "🇫🇷 French": "fr",
    "🇩🇪 German": "de",
    "🇮🇹 Italian": "it",
    "🇸🇪 Swedish": "sv",
    "🇩🇰 Danish": "da",
    "🇳🇴 Norwegian": "no",
    "🇷🇺 Russian": "ru",
    "🇨🇳 Chinese (Simplified)": "zh-CN",
    "🇮🇳 Hindi": "hi",
    "🇧🇩 Bengali": "bn",
    "🇺🇾 Urdu": "ur",    
    "🇦🇪 Arabic": "ar",
    "🇯🇵 Japanese": "ja",
    "🇰🇷 Korean": "ko",
    "🇻🇳 Vietnamese": "vi",
    "🇹🇷 Turkish": "tr",
    "🇵🇹 Portuguese": "pt",
    "🇵🇱 Polish": "pl",
    "🇳🇱 Dutch": "nl", 
    "🇮🇩 Indonesian": "id",
    "🇹🇭 Thai": "th",
}

# Set 'None' as the default for original, assuming no fixed source language
language_names = ["🌐 Original Notes"] + list(languages.keys())

# Initialize
slide_data = None  # Initialize

# Open files - first, check if default JSON exists
if os.path.exists(JSON_file):
    #st.success(f"Found slide data in `{JSON_file}`. Loading automatically.")
    with open(JSON_file, "r") as f:
        slide_data = json.load(f)
else:
    # Upload any file inside the desired folder
    uploaded_file = st.file_uploader("Select any file inside your target folder", type=None)

    if uploaded_file is not None:
        # Get folder path from uploaded file (only possible locally)
        uploaded_path = Path(uploaded_file.name)
        presentation_folder = str(uploaded_path.parent)

        JSON_file = os.path.join(presentation_folder, "slide_data.json")
        images_folder = os.path.join(presentation_folder, "images")

        st.write(f"Detected folder: `{presentation_folder}`")
        st.write(JSON_file)

        # Try to load JSON from the newly detected folder
        if os.path.exists(JSON_file):
            with open(JSON_file, "r") as f:
                slide_data = json.load(f)
        else:
            st.error(f"No `slide_data.json` found in `{presentation_folder}`.")

# Only continue if slide data was loaded
if slide_data:
    
    # Start with the first slide
    if "slide_index" not in st.session_state:
        st.session_state["slide_index"] = 1

    num_slides = len(slide_data)
    
    # Layout and Language selection    
    lc, cc, rc = st.columns((1,1,1))
    with lc:
        # --- Layout toggle switch ---
        
        ##########################################
        # PART OF MULTIPAGEAPP? REMOVE START 2 ###
        ##########################################
        wide_mode = st.toggle("Use wide layout", value=(st.session_state.layout_choice_SJ == "wide"))
        new_layout = "wide" if wide_mode else "centered"
        
        if new_layout != st.session_state.layout_choice_SJ:
            st.session_state.layout_choice_SJ = new_layout
            st.rerun()
        ########################################
        # PART OF MULTIPAGEAPP? REMOVE END 2 ###
        ########################################
        
        horizontal = st.toggle('Toggle to show notes beside slides')
    with cc:
        selected_lang_display = st.selectbox("Language for speaker notes", options=language_names)
    with rc:
        st.write('Number of slides in the presentation: %3i' %num_slides)
        
    # None = no translation
    target_lang = None if selected_lang_display == "🌐 Original Notes" else languages[selected_lang_display]
    
    # Navigation buttons
    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
            st.session_state["slide_index"] = st.number_input('Slide number to show', 1, num_slides)

    # Display slides
    selected_slide = slide_data[st.session_state["slide_index"] - 1]
    image_path = os.path.join(images_folder, os.path.basename(selected_slide["image"]))

    note_text = selected_slide["notes"]

    if target_lang:
        if "translated_notes" not in selected_slide:
            selected_slide["translated_notes"] = translate_notes(note_text, target_lang)
        translated_text = selected_slide["translated_notes"]

    if horizontal:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.image(image_path)
        with col2:
            if target_lang:
                st.write(f"**Translated Notes ({selected_lang_display})**\n\n{translated_text}")
                with st.expander("Show original notes"):
                    st.write(note_text)
            else:
                st.write(f"**Notes:**\n\n{note_text}")
    else:
        st.image(image_path)
        if target_lang:
            st.write(f"**Translated Notes ({selected_lang_display})**\n\n{translated_text}")
            with st.expander("Show original notes"):
                st.write(note_text)
        else:
            st.write(f"**Notes:**\n\n{note_text}")

    ### Download presentation as pdf ###
    # Navigation buttons
    pcol1, pcol2, pcol3 = st.columns([3, 1, 3])
    with pcol1:    
        if st.button('Prepare pdf (no notes) for download'):
            generate_pdf(slide_data, images_folder, presentation_folder, target_lang)
        
        if st.button('Prepare pdf (with notes) for download'):
            generate_pdf(slide_data, images_folder, presentation_folder, target_lang, True, 'Download pdf (with notes)')

else:
    st.warning("No slide data loaded yet.")

'---'
# Render footer with authors, institutions, and license logo in a single line
columns_lic = st.columns((2,1))
with columns_lic[0]:
    st.markdown(f'**SlideJet developed by** <br> {", ".join(author_list)} ({year}). <br> {institution_text}', unsafe_allow_html=True)
with columns_lic[1]:
    st.markdown(f'**Open-source license for SlideJet:**', unsafe_allow_html=True)
    try:
        st.image(Image.open("FIGS/CC_BY-SA_icon.png"))
    except FileNotFoundError:
        st.image("https://raw.githubusercontent.com/gw-inux/SlideJet/main/FIGS/CC_BY-SA_icon.png")
