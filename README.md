# UmlagerungsTool-Archive

[![Python](https://img.shields.io/badge/Python-3.11+-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/windows)
[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Tkinter](https://img.shields.io/badge/GUI-Tkinter-FFCC00)](https://docs.python.org/3/library/tkinter.html)
[![OpenPyXL](https://img.shields.io/badge/Excel-openpyxl-107C41)](https://openpyxl.readthedocs.io/)
[![PyInstaller](https://img.shields.io/badge/Build-PyInstaller-orange)](https://pyinstaller.org/)
[![Inno Setup](https://img.shields.io/badge/Installer-Inno%20Setup-red)](https://jrsoftware.org/isinfo.php)

> [!IMPORTANT]
>
> **This repository is archived and is no longer under active development.**
>
> Development of the project has continued in the completely rewritten **UmlagerungsTool 3.0** repository:
>
> **[UmlagerungsTool 3.0](https://github.com/Vipe-tools/UmlagerungsTool)**
>
> This repository remains available as a historical archive documenting the evolution of the original project from its first prototype to the final Version 2.0 release.

---

## Overview

UmlagerungsTool-Archive contains the original implementation of the warehouse automation software developed before the complete architectural redesign that resulted in UmlagerungsTool 3.0.

The software was created to automate repetitive warehouse workflows such as relocation document generation, article management, barcode creation, Excel automation, and Amazon FBA preparation.

Originally developed to solve a real operational problem inside a German warehouse environment, the application evolved over several years into a larger workflow system consisting of two connected desktop applications.

Although this repository is no longer actively maintained, it documents the complete evolution of the original code base and remains available for educational and historical purposes.

---

# Background

The project originated from a practical business problem.

Originally, warehouse relocation documents were completed entirely by hand.

Later, handwritten paperwork was replaced by Microsoft Word templates, but employees still had to enter all information manually for every shipment.

As warehouse operations continued to grow, several recurring issues became increasingly obvious:

* Manual SKU entry
* Typing mistakes
* Missing barcode support
* Repetitive data entry
* Difficult maintenance of templates
* Inconsistent documentation
* Time-consuming daily workflows

The first versions automated only a small portion of these tasks.

Over time the project expanded into a much larger application capable of generating relocation documents, managing article templates, creating Code128 barcodes, automating Excel templates, and significantly reducing manual work.

Version 2.0 represents the final release of the original architecture before the complete rewrite that became UmlagerungsTool 3.0.

---

# Project Structure

The project consists of two connected desktop applications.

## UmlagerungsTool

The primary application used during daily warehouse operations.

Main features include:

* SKU-based article selection
* Automatic relocation document generation
* Barcode generation
* Excel template automation
* Amazon FBA shipment support
* Printing workflow
* Multiple warehouse relocation workflows

## Artikelverwaltung

The integrated article management application serves as the central interface for maintaining article templates used by the main application.

Features include:

* Create articles
* Edit articles
* Delete articles
* Import article database
* Export article database
* Maintain SKU information
* Manage template-related settings

Both applications share the same article database and work together as one integrated warehouse solution.

---

# Supported Workflows

The generated documents support a variety of warehouse operations, including:

* Relabeling
* Repacking
* Removal of packaging bands
* Dangerous goods processing
* Internal warehouse documentation
* Amazon FBA preparation

Automatically generated Code128 barcodes are embedded directly into the Excel templates to improve accuracy while reducing manual work.

---

# Screenshots

## UmlagerungsTool

<img width="602" height="287" alt="grafik" src="https://github.com/user-attachments/assets/094f4ba6-0813-4456-a2e0-86d2408102a9" />

---

## Artikelverwaltung

<img width="716" height="547" alt="grafik" src="https://github.com/user-attachments/assets/a6d8fb00-0745-40c1-b51d-cff005658816" />

---

## New Article Creation

<img width="412" height="381" alt="grafik" src="https://github.com/user-attachments/assets/96df5154-7f93-4a61-a591-a5111618c96c" />

---

# Development History

Unlike UmlagerungsTool 3.0, this repository was not originally developed using Git.

Development started as a small internal automation tool and evolved through numerous local versions before eventually being published on GitHub.

Earlier source code snapshots have intentionally been preserved inside the `archive` directory.

These archived versions document the progression from the first working prototype to the final Version 2.0 release and provide insight into how both the software architecture and user interface evolved over time.

The repository also contains historical screenshots and archived source code illustrating the continuous expansion of the project.

Following Version 2.0, development moved to a completely new repository.

Instead of continuing to expand the original architecture, the project was rewritten almost entirely from scratch.

The successor introduces a modular architecture, improved maintainability, cleaner separation between GUI and business logic, and a significantly improved project structure.

The actively maintained successor is available here:

**[UmlagerungsTool 3.0](https://github.com/Vipe-tools/UmlagerungsTool)**

---

# Evolution of the User Interface

The screenshots below document several milestones throughout the development of the original project.

Version **1.0** represents the very first working release.

Version **1.6** introduced numerous usability improvements.

Version **1.13** became the final major release before the transition toward Version 2.0.

Version **2.0** introduced the dedicated article management application and transformed the software from a simple utility into a considerably larger warehouse workflow system.

## Version 1.0

<img width="506" height="232" alt="grafik" src="https://github.com/user-attachments/assets/45d97c60-00c8-46a0-8351-e370e95349a4" />

---

## Version 1.6

<img width="583" height="259" alt="grafik" src="https://github.com/user-attachments/assets/db4989bd-2dd4-4dac-8916-394f44501af8" />

---

## Version 1.13

<img width="603" height="289" alt="grafik" src="https://github.com/user-attachments/assets/d1fb85eb-0d3e-46d4-9bd0-20baf6d88d15" />

---

## Version 2.0

<img width="602" height="287" alt="grafik" src="https://github.com/user-attachments/assets/be493c22-bf2e-4e0c-bd25-222a956d3fd4" />

---

# Technologies

* Python
* Tkinter
* Pillow
* OpenPyXL
* PyInstaller
* Inno Setup

---

# Build

Build commands used for PyInstaller are stored inside:

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

**Version:** 2.0.0

**Repository Status:** Archived

**Development:** Discontinued

**Successor:** UmlagerungsTool 3.0

**Current Development Repository:**

https://github.com/Vipe-tools/UmlagerungsTool

This repository remains available as historical documentation of the original implementation.

---

# Platform Support

The application was developed and tested exclusively for Microsoft Windows.

Supported operating systems include:

* Windows 7
* Windows 8 / 8.1
* Windows 10
* Windows 11

Linux and macOS have never been officially supported.

---

# License

This project is licensed under the **PolyForm Noncommercial License 1.0.0**.

You are free to study, modify and use the source code for personal, educational or internal non-commercial purposes.

Commercial redistribution, resale or commercial use of this project or substantially identical derivatives is not permitted without prior written permission from the author.

See the **LICENSE** file for the complete license text.

---

# Author

Developed by **Vipe**.

Originally created as an internal warehouse automation solution before being released publicly for educational and portfolio purposes.

Although development of this repository has ended, it remains available to document the evolution of the original software.

The actively maintained successor can be found here:

**UmlagerungsTool 3.0**

[UmlagerungsTool 3.0](https://github.com/Vipe-tools/UmlagerungsTool)
