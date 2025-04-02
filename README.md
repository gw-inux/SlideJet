## 🚀 SlideJet

**SlideJet** is a lightweight tool that transforms PowerPoint presentations into beautiful, interactive web apps using [Streamlit](https://streamlit.io). It converts your slides into scrollable or slideshow-style views — complete with speaker notes — making it easy to share presentations online.

SlideJet consists of two components:

1. **SlideJet-Convert**: A Streamlit application that runs locally and converts PowerPoint presentations (including speaker notes) into web-ready graphics and note files.
2. **SlideJet-Present**: A Streamlit application to present the converted slides in a user-friendly web format. Presentations can be deployed online via platforms like Streamlit Community Cloud.

---

### ✨ Features

- Convert PowerPoint slides and notes into web-optimized figures and structured data.
- Display slides directly in a Streamlit app — share your presentation via link or QR code.
- Show speaker notes alongside each slide.
- Designed for future multi-language support (automatic translation planned).
- Output format optimized for deployment on GitHub Pages or other static hosting platforms.
- Ideal for teaching, conferences, workshops, training materials, and documentation.
- Provide your audience with the most recent version of your slides and update content quickly and efficiently.

---

### 💻 Requirements

To run **SlideJet-Convert**, you need the following installed locally (on **Windows**):

- [Python](https://www.python.org/downloads/)
- [Microsoft PowerPoint](https://www.microsoft.com/microsoft-365/powerpoint)
- Python packages: `streamlit`, `pywin32`, and `pillow`

---

### ⚙️ Installation

Open a **Command Prompt** and run the following command:

```bash
pip install streamlit pywin32 pillow
```

Then start the app with:

```bash
streamlit run app.py
```

### 📺 Getting Started

A short tutorial and video guide (link coming soon) will walk you through the full process.
With a decent internet connection, the entire setup takes just a few minutes.

<img src="figs/SlideJet_Logo_Wide_small.png" alt="SlideJet Logo" width="300">
