# UmlagerungsTool

A Python-based warehouse and inventory management tool developed to support article administration, relocation workflows, barcode generation, and Excel-based processes.

The project was created to solve practical workflow challenges in a German warehouse and e-commerce environment. It evolved through multiple major versions and is actively used as an internal productivity tool.

## Features

### Umlagerungs-Menü

* Management of relocation templates
* Barcode generation and printing
* Excel integration
* Template processing
* Relocation document creation
* Print support

### Artikelverwaltung

* Create new articles
* Edit existing articles
* Delete articles
* Import and export article data
* Automatic template generation
* Barcode support
* Centralized article management

## Technologies

* Python
* Tkinter
* OpenPyXL
* Pillow
* PyInstaller
* Inno Setup

## Project Structure

```text
assets/
├─ C1A.ico
├─ icon.jpg
└─ arial.ttf

build/
└─ pyinstaller_commands.txt

installer/
└─ setup2.iss

src/
├─ Artikelverwaltung.pyw
└─ Umlagerungs-Menü.pyw

templates/
└─ %Vorlage%.xlsx

README.md
```

## Setup

The application requires a folder named:

```text
Umlagerung
```

to exist in the same directory as the executable.

Copy the template file:

```text
%Vorlage%.xlsx
```

from the `templates` directory into the `Umlagerung` folder before running the application.

The software expects this structure and may not function correctly if the folder or template file is missing.

The project also includes the required font file (`arial.ttf`) and build instructions used to generate the executable versions with PyInstaller.

## Project Background

Unlike many learning projects, this software was developed to solve real operational problems rather than as a programming exercise.

The project combines inventory management, document generation, barcode handling, Excel automation, and installation packaging into a single workflow.

The user interface is currently available only in German because the software was designed for use in a German-speaking business environment.

## Status

Current Version: 2.0.0

Development Status: Active

Repository Type: Internal Business Tool / Portfolio Project
