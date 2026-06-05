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

<img width="602" height="287" alt="grafik" src="https://github.com/user-attachments/assets/8afab66b-554a-46a2-aa99-83076fef9821" />


## Artikelverwaltung

<img width="716" height="547" alt="grafik" src="https://github.com/user-attachments/assets/2c8148db-9bde-4e75-beb9-b420dea05c1a" />


## New Article Creation

<img width="412" height="381" alt="grafik" src="https://github.com/user-attachments/assets/b20f676f-e4d9-4ab4-ac82-ea1b17b9eb3b" />


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

Type: Internal Business Tool / Portfolio Project

The software is actively used in a real warehouse and e-commerce environment and continues to evolve based on operational requirements.

---

# Note

This project was not created as a tutorial or learning exercise.

It was developed as a practical solution to a real business problem and remains in active use within daily operations.
