import os
import tempfile
import shutil
import json
import yaml
import streamlit as st
import pythoncom
import win32com.client
from PIL import Image

###
# Authors, institutions, and year
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

def save_yaml_config(yaml_output_path, slides_subfolder, header_text, subheader_text, mode, yaml_repo_path=None):
    if mode == "Local use":
        presentation_folder = slides_subfolder
    else:  # Online use (Streamlit Cloud)
        if yaml_repo_path is None:
            raise ValueError("For online use, 'yaml_repo_path' must be provided.")
        # Construct path inside the repo
        presentation_folder = os.path.join(yaml_repo_path, slides_subfolder).replace("\\", "/")

    config = {
        "presentation_folder": presentation_folder,
        "header_text": header_text,
        "subheader_text": subheader_text
    }

    with open(yaml_output_path, "w") as f:
        yaml.dump(config, f, default_flow_style=False, sort_keys=False)


# Header and title
st.set_page_config(
    page_title="SlideJet - Convert",
    page_icon="🚀",
)

# Streamlit App Title
st.title("🚀 SlideJet - Convert")
st.header("PowerPoint to Streamlit-Ready Slideshow", divider= "green")

st.markdown(""" 
    SlideJet-convert allows you to transfer a Powerpoint presentation with notes into *.png graphics and a JSON file that contains the slide notes. You can upload any Powerpoint file. Subsequently, you can define the name of the folder where the tools save the JSON file and the 'images' folder.
    
    SlideJet works with PowerPoint in the background to convert your slides. **Now** it's a good moment to safe your open presentations in case of unexpected troubles with PowerPoint.
        """)

# File upload section
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])

if uploaded_file:
    
    # Extract filename (without extension)
    pptx_filename = os.path.splitext(uploaded_file.name)[0]

    # SlideJet_present folder (YAML will be saved here)
    present_folder = st.text_input(
        "Enter the folder where SlideJet_present is located (YAML will be saved here)",
        value=os.getcwd()
    )
    deployment_mode = st.radio("Deployment Mode",["Local use", "Online use (Streamlit Cloud)"], index=0)
    if deployment_mode == "Online use (Streamlit Cloud)":
        yaml_repo_path = st.text_input(
            "Enter the repository-relative path to the YAML file (e.g., `project\presentations`)",
            value=""
        )
    
    # Slides folder (relative to SlideJet_present folder)
    slides_subfolder = st.text_input(
        "Enter the relative path where the slides will be saved (from YAML location)",
        value=os.path.join("slides", pptx_filename).replace("\\", "/")
    )
    # Compute the absolute output directory
    slides_absolute_path = os.path.join(present_folder, slides_subfolder)
    
    default_header = f"{pptx_filename}"
    default_subheader = "Interactive Slideshow"
    header_text = st.text_input("Enter header text (main title)", value=default_header)
    subheader_text = st.text_input("Enter subheader text (subtitle or description)", value=default_subheader)


    # Define output folders dynamically
    OUTPUT_DIR = slides_absolute_path
    IMAGE_DIR = os.path.join(OUTPUT_DIR, "images")
    JSON_FILE = os.path.join(OUTPUT_DIR, "slide_data.json")


    def clear_old_files(folder_path):
        """Deletes all files in the folder before writing new ones."""
        if os.path.exists(folder_path):
            shutil.rmtree(folder_path)  # Delete everything inside the folder
        os.makedirs(folder_path)  # Recreate the folder

    def convert_ppt_to_images_using_powerpoint(ppt_path, image_dir):
        """Uses PowerPoint COM automation to export full slides as images"""
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # Run PowerPoint in the background

        # Open the presentation
        presentation = powerpoint.Presentations.Open(ppt_path, WithWindow=False)

        # Ensure output directory exists and clear old images
        clear_old_files(image_dir)

        slide_data = []
        slide_count = presentation.Slides.Count

        for i in range(1, slide_count + 1):
            slide_filename = f"slide_{i}.png"
            slide_path = os.path.join(image_dir, slide_filename)

            try:
                presentation.Slides(i).Export(slide_path, "PNG")

                # Extract notes
                notes = "No notes"
                if presentation.Slides(i).NotesPage.Shapes.Count > 1:
                    notes_shape = presentation.Slides(i).NotesPage.Shapes.Placeholders(2)
                    if notes_shape.TextFrame.HasText:
                        notes = notes_shape.TextFrame.TextRange.Text.strip()
                
                # Store only relative path for the JSON
                slide_data.append({"image": f"images/{slide_filename}", "notes": notes})

            except Exception as e:
                st.error(f"Error exporting slide {i}: {e}")

        presentation.Close()
        #powerpoint.Quit()
        pythoncom.CoUninitialize()

        return slide_data

    def save_slide_data_json(slide_data, json_file):
        """Saves slide images and notes in a structured JSON format for Streamlit slideshow."""
        with open(json_file, "w") as f:
            json.dump(slide_data, f, indent=4)

    if st.button("Convert to Slideshow Data"):
        # Save uploaded file to a unique temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp_file:
            tmp_file.write(uploaded_file.getbuffer())
            temp_ppt_path = tmp_file.name

        # Convert PPT slides to images using PowerPoint automation and extract notes
        slide_data = convert_ppt_to_images_using_powerpoint(temp_ppt_path, IMAGE_DIR)

        # Save slide data JSON for Streamlit slideshow
        save_slide_data_json(slide_data, JSON_FILE)

        if slide_data:
            st.success(f"Slides and notes successfully saved in `{OUTPUT_DIR}`.")
            st.success(f"Slide data JSON saved for Streamlit slideshow in `{JSON_FILE}`.")
            
            # Write YAML file in parent folder
#           yaml_file = os.path.join(script_dir, f"{pptx_filename}_slidejet_config.yaml")
            yaml_file = os.path.join(present_folder, f"{pptx_filename}_slidejet_config.yaml")
            # Compute correct relative path from script dir to output dir
            #slides_relative_path = os.path.relpath(OUTPUT_DIR, script_dir).replace("\\", "/")
#           save_yaml_config(slides_relative_path, yaml_file, header_text, subheader_text)
            if deployment_mode == "Online use (Streamlit Cloud)":
                save_yaml_config(
                    yaml_output_path=yaml_file,
                    slides_subfolder=slides_subfolder,
                    header_text=header_text,
                    subheader_text=subheader_text,
                    mode=deployment_mode,
                    yaml_repo_path=yaml_repo_path
                )
            else:
                save_yaml_config(
                    yaml_output_path=yaml_file,
                    slides_subfolder=slides_subfolder,
                    header_text=header_text,
                    subheader_text=subheader_text,
                    mode=deployment_mode
                )


            st.success(f"YAML config for SlideJet Present saved as `{yaml_file}`.")
            st.markdown("""
            #### Next steps
            Now you will find the slides, the speaker notes (as *.json file), and the YAML file in the generated folders - see messages above. The see and present the slides, use the **SlideJet_present**.py application with the generated YAML file. You have two options:
            (1) Directly refer to the YAML file from the SlideJet_present.py application, or
            (2) Select and load the YAML file from the SlideJet_present.py application.
            """)
        
        # Delete temporary file
        os.remove(temp_ppt_path)

'---'
# Render footer with authors, institutions, and license logo in a single line
columns_lic = st.columns((4,1,1))
with columns_lic[0]:
    st.markdown(f'Developed by {", ".join(author_list)} ({year}). <br> {institution_text}', unsafe_allow_html=True)
with columns_lic[1]:
    try:
        img_sj = Image.open("FIGS/SlideJet_Logo_Wide_small.png")
        st.image(img_sj)
    except FileNotFoundError:
        st.image("https://raw.githubusercontent.com/gw-inux/SlideJet/main/FIGS/SlideJet_Logo_Wide_small.png")
    
with columns_lic[2]:
    try:
        st.image(Image.open("FIGS/CC_BY-SA_icon.png"))
    except FileNotFoundError:
        st.image("https://raw.githubusercontent.com/gw-inux/SlideJet/main/FIGS/CC_BY-SA_icon.png")