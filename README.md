# UmlagerungsTool

An internal Python project for generating relocation documents, managing article templates, integrating barcodes, and automating Excel-based warehouse workflows.

The software was developed to simplify real business processes in a German warehouse and e-commerce environment and is actively used in daily operations.

---

# Background

This project originated from a practical business problem.

Originally, warehouse relocation processes were documented using handwritten notes. Later, these were replaced by printed Word templates that still had to be filled out manually.

As the number of products and Amazon FBA shipments increased, several problems became apparent:

* Manual data entry
* SKU transcription mistakes
* Inconsistent documentation
* Missing barcode integration
* Time-consuming repetitive tasks
* Difficult maintenance of templates

The goal of the project was to standardize and automate these workflows.

---

# Project Structure

The project consists of two connected applications.

## UmlagerungsTool

The main application used during daily warehouse operations.

Features:

* SKU-based article selection
* Relocation document generation
* Amazon FBA shipment processing
* Barcode generation
* Printing of relocation documents
* Support for various warehouse workflows
* Excel template integration

## Artikelverwaltung

Administration tool for article data and templates.

Features:

* Create new articles
* Edit existing articles
* Delete articles
* Export articles
* Import articles
* Manage SKU information
* Maintain template-related data

The article management application is launched directly from UmlagerungsTool and serves as the central interface for maintaining the data used by the main application.

---

# Supported Workflows

Generated documents can contain information for various warehouse operations, including:

* Relabeling
* Repackaging
* Removal of bands or packaging materials
* Dangerous goods processing
* Internal documentation
* Amazon FBA preparation

Barcodes are automatically generated and embedded into documents to reduce manual data entry and improve accuracy.

---

# Screenshots

## UmlagerungsTool

<img width="602" height="287" alt="grafik" src="https://github.com/user-attachments/assets/094f4ba6-0813-4456-a2e0-86d2408102a9" />


## Artikelverwaltung

<img width="716" height="547" alt="grafik" src="https://github.com/user-attachments/assets/a6d8fb00-0745-40c1-b51d-cff005658816" />


## New Article Creation

<img width="412" height="381" alt="grafik" src="https://github.com/user-attachments/assets/96df5154-7f93-4a61-a591-a5111618c96c" />


---

# Development History

This project was not originally developed using Git or GitHub.

The software started as a small internal tool and evolved through many local versions before being uploaded to GitHub.

Earlier development versions have been preserved inside the `archive` directory for historical reference.

These archived versions document the progression from the first working prototype to the current release and provide insight into how the software evolved over time.

The repository also contains source code archives and screenshots of older releases that illustrate the development process and the gradual expansion of functionality.

## Evolution of the User Interface

The screenshots below show selected milestones during the development of the project.

Version 1.0 represents the first working release.

Version 1.6 introduced additional usability improvements and user feedback features.

Version 1.13 represents the final major release before the transition to version 2.0.

Version 2.0 expanded the project by introducing the separate article management application and transforming the software from a single utility into a larger workflow system.

### Version 1.0

<img width="506" height="232" alt="grafik" src="https://github.com/user-attachments/assets/45d97c60-00c8-46a0-8351-e370e95349a4" />


### Version 1.6

<img width="583" height="259" alt="grafik" src="https://github.com/user-attachments/assets/db4989bd-2dd4-4dac-8916-394f44501af8" />


### Version 1.13

<img width="603" height="289" alt="grafik" src="https://github.com/user-attachments/assets/d1fb85eb-0d3e-46d4-9bd0-20baf6d88d15" />


### Version 2.0

<img width="602" height="287" alt="grafik" src="https://github.com/user-attachments/assets/be493c22-bf2e-4e0c-bd25-222a956d3fd4" />


---

# Technologies Used

* Python
* Tkinter
* OpenPyXL
* Pillow
* PyInstaller
* Inno Setup

---

# Repository Structure

```text
assets/
├─ C1A.ico
├─ arial.ttf
└─ icon.jpg

archive/
├─ source/
└─ screenshots/

build/
└─ build_commands.txt

installer/
└─ setup.iss

src/
├─ Artikelverwaltung.pyw
└─ Umlagerungs-Menü.pyw

templates/
└─ %Vorlage%.xlsx

README.md
```

---

# Setup

Expected directory structure:

```text
Application/
├─ UmlagerungsTool.exe
├─ Artikelverwaltung.exe
└─ Umlagerung/
   └─ %Vorlage%.xlsx
```

UmlagerungsTool serves as the main application.

When the user opens the article management module, UmlagerungsTool searches for:

```text
Artikelverwaltung.exe
```

within the same directory and launches it automatically.

Both applications share the same `Umlagerung` directory and use the same template files.

---

# Build

Build commands used for PyInstaller are stored in:

```text
build/build_commands.txt
```

The installer configuration is located in:

```text
installer/setup.iss
```

and can be compiled using Inno Setup.

---

# Status

Version: 2.0.0

Status: Active Development

Type: Internal Business Tool, Portfolio Project and Process Automation Tool

The software is actively used in a real warehouse and e-commerce environment and continues to evolve based on operational requirements.

---

# Note

This project was not created as a tutorial or learning exercise.

It was developed as a practical solution to a real business problem and remains in active use within daily operations.

This project was not created as a tutorial, coding exercise, or portfolio demonstration.

It was developed as a practical solution to a real-world warehouse and e-commerce workflow problem and remains actively used in daily operations.

Platform Support

The application was developed and tested exclusively for Microsoft Windows environments.

Supported operating systems:

Windows 7
Windows 8 / 8.1
Windows 10
Windows 11

Linux and macOS are not officially supported and have not been tested.

License

This project is released under the PolyForm Noncommercial License 1.0.0.

You are welcome to view, study, modify, and use the source code for personal, educational, or internal non-commercial purposes.

Commercial resale, commercial redistribution, or selling this project or substantially identical derivatives is not permitted without prior written permission from the author.

