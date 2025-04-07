import os
import streamlit as st
import json
from pathlib import Path
from PIL import Image

# This is a generalized application to present Powerpoint slides and notes as slideshow through Streamlit
# You can adapt the script with header and path to a specific presentation. To do this, just replace the initial informations.

# --- Header content Copyright ---
year = 2025 
authors = {
    "Thomas Reimann": [1],  # Author 1 belongs to Institution 1
   #"Colleague Name": [2],  # Author 2 also belongs to Institution 1
}
institutions = {
    1: "TU Dresden",
   #2: "Second Institution / Organization"
}
author_list = [f"{name}{''.join(f'<sup>{i}</sup>' for i in idxs)}" for name, idxs in authors.items()]
institution_text = " | ".join([f"<sup>{i}</sup> {inst}" for i, inst in institutions.items()])

### SlideJet - Present

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
    
################
# ADAPT HERE ###
################

# Folder to your presentation

presentation_folder = "slides/SlideJet_Overview"

# Presettings - Here you can adjust the application to your specific application

header_text = 'SlideJet Presentation'
subheader_text = 'Overview and Demo'

# --- Streamlit App Content ---
st.title("Presentation Slides")
st.header(f':red-background[{header_text}]')
st.subheader(subheader_text, divider='red')

st.markdown(""" 
    You can move through the slides with the left/right button. Alternatively, you can switch through the slides with the slider. Finally, you can use the toggle to switch to an vertical layout to eventually adapt the app to your device. 
        """)

# Define the default folder - the structure is fix and provided by the convert_ppt_slides.py application
JSON_file = os.path.join(presentation_folder, "slide_data.json")
images_folder = os.path.join(presentation_folder, "images")

# Initialize
slide_data = None  # Initialize

# TODO: THis is not fully working / Fix

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
    
    # Toggle for vertical content organization and streamlit layout
    
    lc, cc, rc = st.columns((1,1,1))
    with lc:
        vertical = st.toggle('Click here for vertical layout')
    with cc:
# --- Layout toggle switch ---
        wide_mode = st.toggle("Use wide layout", value=(st.session_state.layout_choice_SJ == "wide"))
        new_layout = "wide" if wide_mode else "centered"
        
        if new_layout != st.session_state.layout_choice_SJ:
            st.session_state.layout_choice_SJ = new_layout
            st.rerun()
    with rc:
        st.write('Number of slides in the presentation: %3i' %num_slides)
        
    # Navigation buttons
    col1, col2, col3 = st.columns([1, 3, 1])
    with col2:
            st.session_state["slide_index"] = st.number_input('Slide number to show', 1, num_slides)

    # Display slides
    selected_slide = slide_data[st.session_state["slide_index"] - 1]
    image_path = os.path.join(images_folder, os.path.basename(selected_slide["image"]))

    if vertical:
        st.image(image_path)
        st.write(f"**Notes:**\n\n{selected_slide['notes']}")
    else:
        col1, col2 = st.columns([3, 1])
        with col1:
            st.image(image_path)
        with col2:
            st.write(f"**Notes:**\n\n{selected_slide['notes']}")
else:
    st.warning("No slide data loaded yet.")

'---'
# Render footer with authors, institutions, and license logo in a single line
columns_lic = st.columns((4,1))
with columns_lic[0]:
    st.markdown(f'Developed by {", ".join(author_list)} ({year}). <br> {institution_text}', unsafe_allow_html=True)
with columns_lic[1]:
    try:
        st.image(Image.open("FIGS/CC_BY-SA_icon.png"))
    except FileNotFoundError:
        st.image("https://raw.githubusercontent.com/gw-inux/SlideJet/main/FIGS/CC_BY-SA_icon.png")
