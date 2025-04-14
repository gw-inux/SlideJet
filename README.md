## 🚀 SlideJet

**SlideJet** is a lightweight tool that transforms PowerPoint presentations into beautiful, interactive web apps using [Streamlit](https://streamlit.io). It converts your slides into scrollable or slideshow-style views — complete with speaker notes — making it easy to share presentations online.

SlideJet consists of two components:

1. **SlideJet-Convert**: A Streamlit application that runs locally and converts PowerPoint presentations (including speaker notes) into web-ready graphics and note files.
2. **SlideJet-Present**: A Streamlit application to present the converted slides in a user-friendly web format. Presentations can be deployed online via platforms like Streamlit Community Cloud.

[**See an working example by clicking here**](https://slidejet-overview.streamlit.app/)

---

### ✨ Features

- Convert PowerPoint slides and notes into web-optimized figures and data.
- Display slides directly in a Streamlit app.
- Show speaker notes alongside each slide.
- Designed for future multi-language support (automatic translation planned).
- Output format is optimized for deployment on GitHub Pages or other static hosts.
- Ideal for teaching, workshops, training materials, and documentation.

---

### 💻 Requirements

To run **SlideJet-Convert**, you need the following installed locally (on **Windows**):

- [Python](https://www.python.org/downloads/)
- [Microsoft PowerPoint](https://www.microsoft.com/microsoft-365/powerpoint)
- Python packages: `streamlit`, `pywin32`, `Pillow`, and `deep-translator`
  
---

### ⚙️ Installation and running the app

If not already on your computer: Please install Python (e.g., through a distribution like Anaconda, or by downloading from www.python.org)

Download the Python files **SlideJet_convert.py** and eventually **SlideJet_present.py** on a folder on your local computer

Open a **Command Prompt** and run the following command:

```bash
pip install streamlit
```

Then you add the required Python libraries with the command: pip install …
 (https://docs.streamlit.io/get-started/installation)
 
 ```bash
pip install pywin32
```

(this should include pythoncom and win32com)

 ```bash
pip install Pillow
```

(this should include PIL)

 ```bash
pip install deep-translator
```

Then start the app with the command window from the folder where SlideJet_convert.py is saved with:

```bash
streamlit run SlideJet_convert.py
```

### 📺 Getting Started

A short tutorial and video guide (link coming soon) will walk you through the full process.
With a decent internet connection, the entire setup takes just a few minutes.

<img src="figs/SlideJet_Logo_Wide_small.png" alt="SlideJet Logo" width="300">
