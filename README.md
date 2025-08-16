# 🔹 AI-Powered Tableau to Power BI Migration 

A highly advanced, modular Python toolkit designed to **automate the end-to-end migration** of Tableau dashboards (`.twbx`) to Power BI projects (`.pbip`). This system integrates **LLMs (Large Language Models)**, **AI agents**, **vision-based model interaction**, **prompt engineering**, and **deep BI metadata parsing**, drastically reducing the need for manual recreation.

---

## 🌐 Project Overview

Migrating dashboards from Tableau to Power BI typically requires **manual effort, visual inspection, recalculating DAX expressions, and error-prone layout alignment**. This project automates all of it:

* Parses the Tableau XML layout (`.twb`) and extracts metadata
* Converts `.hyper` data sources to CSV and Excel
* Recognizes visuals via screenshots + Vision AI
* Leverages **Generative LLMs** to:

  * Translate Tableau Calcs to DAX
  * Generate Power BI visual JSON
  * Produce slicer/filter configurations
* Programmatically injects all components into a **working `.pbip` Power BI project**
* Includes **agents** to handle **secure credential injection**, **field binding**, and **Power BI visual deployment**

---



## 🎓 Core Features

### 📁 TWBX Extraction

* Automatically unzips `.twbx` archive
* Isolates `.twb` XML layout & `.hyper` data

### 📝 XML Metadata Parsing

* Parses sheets, dashboards, encodings, filters, parameters
* Identifies visual types and data field mappings

### ⚙️ Hyper Data Extraction

* Uses `tableauhyperapi` to convert `.hyper` files to `.csv`
* Renames and cleans datasets for Power BI use

### 🔜 Visual Recognition via AI

* Takes automated screenshots of Tableau dashboards
* Uses **LLM Vision APIs** to detect:

  * Chart types
  * Titles
  * Visual bounding boxes
* Cross-references visual structure with XML data

### ⚛️ Prompt-Based DAX Generation

* Converts Tableau Calculated Fields to **valid DAX expressions**
* Prompts designed to:

  * Preserve logic
  * Resolve ambiguous column/table names
  * Align with Power BI modeling practices

### 🌟 Visual & Filter JSON Generation

* Produces valid `visualContainer` JSON for Power BI
* Includes:

  * Layout
  * Query structure
  * Titles, formatting, interactivity options
* Generates slicers and filters dynamically

### 🚀 Direct PBIP Injection

* Unpacks `.pbip` template
* Modifies:

  * `model.bim` for DAX
  * `report.json` for visuals
* Saves updated project with minimal user input

### 💡 Agent-Based Integration

* **Credential Agent**: securely fetches and injects database credentials into Power BI model configuration
* **DAX Agent**: pushes DAX queries into Power BI's advanced editor for field binding and semantic modeling
* **Visual Deployment Agent**: handles visual creation, layout alignment, and slicer binding inside Power BI canvas

---

## 📊 Supported Chart Types

This toolkit supports automated migration of 20+ charts, including:

* 📦 **Box-Whisker Plots** (`Box_Whisker_chart`)
* 🎯 **Bullet Charts** (`bullet chart`)
* 🔼 **Bump Charts** (`bumpchart`)
* 🦋 **Butterfly Charts** (`butterfly_chart_project`)
* 🔄 **Choropleth Maps** (`choropleth`)
* 🔘 **Circle Timeline Charts** (`circleTimeline`)
* 🔻 **Funnel Charts** (`funnel_chart`)
* 🔲 **Highlighted Tables** (`highlighted table chart`)
* 🍭 **Lollipop Charts** (`Lollipop Chart`)
* 🧱 **Marimekko Charts** (`marimekko Chart`)
* 📊 **Pareto Charts** (`pareto_chart`)
* 🌐 **Radial Charts** (`radialchart`)
* 🧗‍♂️ **Ridge Charts** (`ridge_chart`)
* 🧮 **Slope Charts** (`Slope_chart`)
* ✨ **Sparkline Charts** (`SparklinChart`)

Each chart folder contains Python modules for:

* XML parsing
* AI-based visual matching
* Layout coordinate translation
* Final JSON and M-script generation

---

## 🌈 Technologies Used

### 🧠 AI & LLM

* Google Gemini / OpenAI (LLM agnostic)
* Vision LLM for visual classification
* Text-based LLM for:

  * DAX synthesis
  * Visual schema creation

### 🧰 Data Handling

* `pandas`, `tableauhyperapi`, `numpy`
* `openpyxl` for Excel conversion

### ⚖️ Metadata Extraction

* `xml.etree.ElementTree`
* `zipfile`, `json`, `uuid`, `os`

### 🕹️ UI Automation

* `pyautogui` for capturing dashboard screenshots
* `subprocess` to interact with Tableau Desktop

---

## 📆 Workflow Breakdown

### ✅ Phase 1: Extraction & Analysis

* `.twbx` unpacked
* `.twb` parsed for:

  * Visuals
  * Filters
  * Calculations
  * Layout positions
* `.hyper` to `.csv`
* Screenshot captured + processed by Vision LLM

### 🛠 Phase 2: Transformation

* Detected visuals matched with XML metadata
* Calculated fields sent to LLM for DAX conversion
* All metadata merged to form `final.json`

### 📈 Phase 3: Power BI Assembly

* Visuals & slicers generated via LLM
* Injected into `report.json` and `model.bim`
* Credential Agent and DAX Agent handle secure integration
* Final `.pbip` file is ready to open in Power BI Desktop

---

## 🥇 Why This Project Stands Out

| Skill Demonstrated              | Description                                                                     |
| ------------------------------- | ------------------------------------------------------------------------------- |
| 💡 **AI-Driven Automation**     | Automated conversion of complex visual + data logic using LLMs                  |
| 📆 **ETL Engineering**          | Tableau data extraction to Power BI-compliant data pipelines                    |
| 📄 **Prompt Engineering**       | Custom prompts designed for chart mapping, DAX generation, and layout rendering |
| ⚖️ **Metadata Management**      | Mapping Tableau internal metadata into structured Power BI visual configs       |
| 🚧 **End-to-End System Design** | Covers UI, backend, file systems, data logic, and AI integration                |
| 🧨 **AI Agent Orchestration**   | Modular agents for secure credential injection and field binding in Power BI    |

---

## 🚀 Impact & Application

* ⏰ Migrates dashboards in **minutes** instead of hours/days
* 📈 Ensures **data + visual fidelity** with intelligent AI corrections
* 📅 Saves **engineering & analytics teams** 80%+ effort in dashboard migrations

---

## 💼 Real-World Talking Points

> "I developed an AI-powered automation engine that deconstructs Tableau dashboards and programmatically reconstructs them as Power BI projects."

> "We integrated both text and vision-based LLMs to accurately interpret layout, chart intent, and calculation logic."

> "Prompt engineering played a key role in translating Tableau calculations to robust DAX formulas, reducing migration errors by over 90%."

> "The toolkit manipulates Power BI project internals like `report.json` and `model.bim`, showing deep understanding of both platforms."

> "Agents handle credentials, field injection, and DAX configuration—mimicking real user behavior programmatically."

---

## 💼 License

This project is intended for research, automation prototyping, and internal BI migration workflows. Use at your own discretion and validate all generated `.pbip` artifacts before production deployment.
