# 🧠 Sprint Review PowerPoint Generator

This script automates the generation of a Sprint Review PowerPoint presentation by pulling **User Stories** and their related **Features** from Azure DevOps, and inserting them into a slide deck using a PowerPoint template.

---

## 🚀 Features

- Fetches current sprint data from Azure DevOps
- Filters work items to include only **User Stories**
- Maps User Stories to their associated **Features**
- Inserts the information into a specific slide in a PowerPoint template
- Saves the output with a timestamped filename

---

## 🛠 Requirements

- Python 3.7+
- PowerPoint `.pptx` template with appropriate placeholders
- Access to your Azure DevOps organization and a valid Personal Access Token (PAT)

---

## 📦 Setup

1. **Clone the repo**  
   ```bash
   git clone https://your.repo.url/ppt-generator.git
   cd ppt-generator
2. **Install Dependencies**
   - pip install -r requirements.txt

3. **Configure the config.json**
   - project: The name of the Azure DevOps project.
   - team: The team within the Azure DevOps project.
   - pat: Your Personal Access Token (PAT) from Azure DevOps.
   - template_path: The path to your PowerPoint template file (.pptx).

4. **Get this running**
   ```bash
   python sprint_review_generator.py
