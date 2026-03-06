# Počítače24.cz - Hardware Inventory Script (Python)

This project is a Python-based hardware inventory tool designed to automatically collect detailed information about a computer system. The script gathers data about the machine's hardware and operating system and generates a strcutured report.

The goal of this tool is to help technicians quickly identify and document the hardware configuration of a computer during diagnostics, maintenance, or refurbishment processes.

The script can run on both Windows and Linux environments, making it suitable for cross-platform hardware analysis.

---

## Features

The script collects and generates reports containing the following information :

**System information**

-   Manufacturer
-   Model
-   Serial number

**CPU**

-   Processor model
-   Base clock frequency
-   Number of physical cores
-   Number of logical processors

**RAM**

-   Total installed memory
-   Memory type (DDR3, DDR4, DDR5)
-   Individual memory modules information

**Storage**

-   Disk model
-   Disk size
-   Interface type (SATA / NVMe)

**Graphics**

-   Primary GPU
-   Secondary GPU (if available)

**Display**

-   Screen resolution
-   Estimated screen size

**Network**

-   Wi-Fi adapter detection

**Battery**

-   Battery percentage
-   Estimated battery health

**Operating System**

-   OS name
-   Version

**Additional hardware detection**

-   Bluetooth
-   Webcam
-   Fingerprint reader
-   Keyboard backlight

---

## Output

The script automatically generates two files :
-   machine_report.json – structured JSON report
-   machine_report.txt – human-readable summary

### Example output :

=====================================================================================
SYSTEM MATERIAL REPORT
Generated the : 2026-03-06 10:41:06
=====================================================================================

Model: ASUSTeK COMPUTER INC. VivoBook_ASUSLaptop X521EA_S533EA S/N: M3N0CV11P745118
CPU: 11th Gen Intel(R) Core(TM) i5-1135G7 @ 2.40GHz
RAM: 16GB DDR4
Disk: INTEL SSDPEKNW512G8 512GB NVMe
Mechanika: NE
GPU: Intel(R) Iris(R) Xe Graphics
GPU2: NE
LCD: 15.3 1920x1080
Wifi: Intel(R) Wi-Fi 6 AX201 160MHz [Wifi AX]
BAT: 92.10% [45.00/48.90 Wh]
OS: Microsoft Windows 11 Famille 64 bits (25H2)
BT:Ano WC:Ano FACE:Check B-KBD:Check

---

## Technologies Used

-   Python3
-   PowerShell (Windows hardware queries)
-   Linux system utilities (lscpu, lsblk, lspci)
-   JSON data processing
-   Cross-platform system detection

---

## Use Case

This tool was developed as part of an internship project to assist in the hardware testing and inventory process of refurbished computers

It can be used in environments where technicians need to quickly collect system specifications and generate standardized reports

---

## Author

Developed as part of a software development internship project
