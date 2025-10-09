import os
import tempfile
import shutil
import json
import yaml
import streamlit as st
import pythoncom
import win32com.client
from PIL import Image
from pathlib import Path
import re

# SlideJet_convert is a tool to turn PowerPoint presentations in Streamlit Slideshows
#
# Execute this script with Streamlit on your local computer
#
# SlideJet_convert guides you through the process of converting your presentations


# --- Functions ---------------------------------------------------------------

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


def emit_present_script(
    yaml_file: str | Path,
    template_source: str | Path | None = None,
    template_name: str = "SlideJet_present_template.py",
    multipage: bool = False,
    app_id: str = "app_01",
) -> Path:
    """
    Simple version with two options:
      - multipage=True  -> comment out st.set_page_config(...) in the presenter
      - app_id          -> injected into the presenter (via __APP_ID__ or regex)
    """
    yaml_path = Path(yaml_file).resolve()
    yaml_dir  = yaml_path.parent
    yaml_stem = yaml_path.stem

    # Derive a clean base name WITHOUT the trailing 'SJconfig' (case-insensitive, with optional -/_/.)
    base_name = re.sub(r'(?i)[\-\_\.]?SJconfig$', '', yaml_stem).strip()
    if not base_name:  # safety fallback
        base_name = yaml_stem

    # presenter goes next to the YAML
    target_dir = yaml_dir
    target_dir.mkdir(parents=True, exist_ok=True)

    # load template text
    if template_source:
        tpl_text = Path(template_source).read_text(encoding="utf-8")
    else:
        tpl_text = (Path(__file__).parent / template_name).read_text(encoding="utf-8")

    # compute relative YAML path from presenter location
    rel_yaml = Path(os.path.relpath(yaml_path, start=target_dir)).as_posix()

    # --- token replacements (lightweight)
    out_text = tpl_text
    out_text = out_text.replace("__SLIDEJET_YAML__", rel_yaml)
    out_text = out_text.replace("__IN_MULTIPAGE__", "True" if multipage else "False")
    out_text = out_text.replace("__PAGE_TITLE__", base_name.replace("_", " "))
    out_text = out_text.replace("__APP_ID__", app_id)

    # robust fallbacks: patch literal constants if template lacks placeholders
    # YAML_PATH = "..."
    out_text = re.sub(
        r'(YAML_PATH\s*=\s*)(["\']).*?\2',
        rf'\1"{rel_yaml}"',
        out_text
    )
    # IN_MULTIPAGE = True/False
    out_text = re.sub(
        r'(IN_MULTIPAGE\s*=\s*)(True|False)',
        rf'\1{"True" if multipage else "False"}',
        out_text
    )
    # APP_ID = "..."
    out_text = re.sub(
        r'(APP_ID\s*=\s*)(["\']).*?\2',
        rf'\1"{app_id}"',
        out_text
    )

    # --- if multipage, comment out st.set_page_config(...) line
    if multipage:
        # comment out any line that starts with st.set_page_config(…)
        out_text = re.sub(
            r'^\s*st\.set_page_config\([^\n]*\)\s*$',
            lambda m: f'# {m.group(0)}',
            out_text,
            flags=re.MULTILINE
        )

    # write presenter: <BASE>_SJpresent.py
    present_name = f"{base_name}_SJpresent.py"
    out_file = target_dir / present_name
    out_file.write_text(out_text, encoding="utf-8")
    return out_file

# --- Application -------------------------------------------------------------

# Header and title
st.set_page_config(
    page_title="SlideJet - Convert",
    page_icon="🚀",
)

st.title("🚀 SlideJet - Convert")
st.header("PowerPoint to Streamlit-Ready Slideshow", divider= "green")

st.markdown(""" 
    SlideJet-convert allows you to transfer a Powerpoint presentation with notes into *.png graphics and a JSON file that contains the slide notes. You can upload any Powerpoint file. Subsequently, you can define the name of the folder where the tools save the JSON file and the 'images' folder.
    
    SlideJet works with PowerPoint in the background to convert your slides. **Now** it's a good moment to safe your open presentations in case of unexpected troubles with PowerPoint.
        """)

st.subheader('First step: Upload the presentation file', divider = 'green')
# File upload section
uploaded_file = st.file_uploader("Upload a PowerPoint file", type=["pptx"])


if uploaded_file:
    
    # --- Step: Define Paths 
    st.subheader('Next steps: Define paths', divider = 'green')
    st.markdown("""
    #### Local path
    :green[**SlideJet-Convert**] will generate slideshow data that can be seen by the :orange[**SlideJet-Present**] app. The YAML-file, which contains some settings for :orange[**SlideJet-Present**], is saved in the same folder. Define the path of this folder (on your computer) below:
    """
    )
    
    # Extract filename (without extension)
    pptx_filename = os.path.splitext(uploaded_file.name)[0].replace(" ", "_")

    # SlideJet_present folder (YAML will be saved here)
    present_folder = st.text_input(
        "Enter the folder where SlideJet_present is located (YAML will be saved here)",
        value=os.getcwd()
    )
    st.markdown("""
    #### Information for online deployment
    :orange[**SlideJet-Present**] is typically placed in an online repository (e.g., GitHub) from where the Streamlit-app is deployed. Accordingly, the YAML file needs specific informations to find the presentation for online use. Define the path from the repository to the YAML file below:
    """
    )
    
    deployment_mode = st.radio("Deployment Mode",["Local use", "Online use (Streamlit Cloud)"], index=0)
    
    with st.expander("Click here to show an example and further explanation"):
        st.markdown("""
        ***Further information for the 'Online use (Streamlit Cloud)' option***
        
        Consider that an USER_X operates an REPOSITORY_A on GitHub.com. Within this repository, the user defines a folder where all SlideJet presentations will be saved. This folder is named 'SlideJet_presentations'. Inside of this folder, all SlideJet_present files, the YAML files, and folders with presentation data (see next step) will be saved.
        
        The structure would look like following:
        
        :blue[**GitHub.com/USER_X/Repository_A/**]:orange[SlideJet_presentations/]
        
        The SlideJet_present Python script and the YAML file will be placed in :orange[SlideJet_presentations/]. The relative  path from the repository would be :orange[SlideJet_presentations/].
        """)
    
    if deployment_mode == "Online use (Streamlit Cloud)":
        yaml_repo_path = st.text_input(
            "Enter the relative path from the repository to the YAML file (e.g., `SlideJet_presentations`). The name of the repository is usually not part of the path",
            value="SlideJet_presentations"
        )
    
    st.markdown("""
    #### Relative path to the presentation data
    :orange[**SlideJet-Present**] access the data from a folder that is typically named :blue[**SJ_Data**]. Within :blue[**SJ_Data**], each presentation can be saved in a subfolder. Subsequently, you can define the relative path, whereas the presetting is likely suitable for most users:
    """
    )
    
    with st.expander("Click here to show an example and further explanation"):
        st.markdown("""
        ***Further information for the relative path***
        
        Consider your SlideJet_present script and the YAML file are placed in a folder on your local computer or online repository (e.g., 'SlideJet_presentations'). Within this folder, the slides (as *.png graphics) and the JSON-file with the speaker notes can be saved in subfolders. Typically, the main subfolder will be named 'SJ_Data' and contains subfolders for the different presentations like 'Presentation01', 'Presentation02', and so on.
        
        The relative path is used in your YAML file to allow SLideJet_present to identify the data that are required to generate the Streamlit slideshow.
        """)
    
    # Slides folder (relative to SlideJet_present folder)
    slides_subfolder = st.text_input(
        "Relative path where the slides will be saved (from YAML location). Usually it is 'SJ_Data/[NAME_OF_PRESENTATION]'. Eventually modify this.",
        value=os.path.join("SJ_DATA", pptx_filename).replace("\\", "/")
    )
    
    # Compute the absolute output directory
    slides_absolute_path = os.path.join(present_folder, slides_subfolder)
    
    # --- Step: PRESENTATION HEADERS
    st.subheader('SlideJet-Present header information', divider = 'green')
    st.markdown("""
    :orange[**SlideJet-Present**] is an interactive Streamlit app that shows your presentation with notes as a slideshow. Subsequently, you can define the header and subheader for your specific :orange[**SlideJet-Present**] Streamlit app. This information will be safed in the YAML-file.
    """
    )
    
    default_header = f"{pptx_filename}"
    default_subheader = "Interactive Slideshow"
    header_text = st.text_input("Enter header text (main title)", value=default_header)
    subheader_text = st.text_input("Enter subheader text (subtitle or description)", value=default_subheader)

    # Define output folders dynamically
    OUTPUT_DIR = slides_absolute_path
    IMAGE_DIR = os.path.join(OUTPUT_DIR, "images")
    JSON_FILE = os.path.join(OUTPUT_DIR, "slide_data.json")

    # --- Final Step: Convert
    st.subheader('Final step: Convert the slideshow', divider = 'green')
    st.markdown("""
    SlideJet can already create the :orange[**SlideJet_Present**] file that represents the Streamlit SlideShow. Also, please check if your presentation will be part of an multipage app - in that case, the SlideJet presentation will be without a separate page title.
    
    Use the subsequent checkboxes to proceed. 
    """)
    make_presenter = st.checkbox("Create SlideJet_present file next to the YAML", value=True)
    multipage_true = st.checkbox("The SlideJet presentation will be part of an multipage app", value=False)

    app_id = "app_01"
    if multipage_true:
        app_id = st.text_input("Multipage app_id (used for namespacing state etc.)", value="app_01")
    
    col1, col2, col3 = st.columns((1,1,1))
    with col2:
        start_convert = st.button(":rainbow[**Convert to Slideshow Data**]")
        
    if start_convert:
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
            st.success(f"Slide data JSON saved in `{JSON_FILE}`.")
            
            # Write YAML file in parent folder
            yaml_file = os.path.join(present_folder, f"{pptx_filename}_SJconfig.yaml")
            # Compute correct relative path from script dir to output dir
            #slides_relative_path = os.path.relpath(OUTPUT_DIR, script_dir).replace("\\", "/")
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

            st.success(f"YAML config for SlideJet_present saved as `{yaml_file}`.")
            
            
            
            try:
                if make_presenter:
                    # Option A: template file lives next to this converter script
                    template_path = Path(__file__).parent / "SlideJet_present_template.py"

                    out_present = emit_present_script(
                        yaml_file=yaml_file,
                        template_source=template_path,
                        multipage=multipage_true,
                        app_id=app_id,
                    )
                    st.success(f"SlideJet_present script created in: `{out_present}`")

            except Exception as e:
                st.warning(f"Could not create presenter script automatically: {e}")

            
            
            
            st.markdown("""
            #### Next steps
            Now you will find the slides, the speaker notes (as *.json file), the SlideJet_presentation, and the YAML file in the generated folders - see messages above. The see and present the slides, use the generated :orange[**SlideJet_present.py**] application (contains *_SJpresent.py* in the filename). Run this file on your local computer from the command prompt (CMD) with 'streamlit run ... YOUR_PRESENTATION_SJpresent.py'.
            """)
        
        # Delete temporary file
        os.remove(temp_ppt_path)

'---'
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