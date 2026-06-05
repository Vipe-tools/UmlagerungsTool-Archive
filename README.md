# UmlagerungsTool

A Python-based internal tool developed to simplify inventory relocation management and article administration.

This project was created to solve real-world workflow problems in a German e-commerce and warehouse environment and evolved through multiple major versions during active use.

## Features

### Umlagerungs-Menü

* Management of relocation templates
* Barcode generation
* Excel integration
* Print support
* Template selection and processing

### Artikelverwaltung

* Create new articles
* Edit existing articles
* Delete articles
* Import and export article data
* Automatic template creation
* Barcode support

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
└─ icon.jpg

src/
├─ Artikelverwaltung.pyw
└─ Umlagerungs-Menü.pyw

templates/
└─ %Vorlage%.xlsx
```

## Setup

Before running the application, create a folder named:

```text
Umlagerung
```

in the same directory as the application.

Copy the provided template file:

```text
%Vorlage%.xlsx
```

from the `templates` directory into the `Umlagerung` folder.

The software expects this structure and may not function correctly if the required folder or template file is missing.

## Project Background

This project was developed as a practical business tool rather than a learning exercise.

The software is actively used to support warehouse and administrative workflows. The user interface is currently available only in German because it was designed for a German-speaking work environment.

## Status

Current Version: 2.0.0

Development Status: Active
