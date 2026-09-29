# 📊 Daily Orders AI Processor

[![Platform](https://img.shields.io/badge/Platform-Windows-blue.svg)](https://microsoft.com)
[![Python Version](https://img.shields.io/badge/Python-3.12.10%2B-green.svg)](https://python.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

An enterprise-grade Python desktop application designed to automate daily financial remittance processing, unstructured order text parsing, Excel ledger generation, intelligent transaction splitting, and automated monthly ledger rollovers.

The application communicates with a dedicated FastAPI backend powered by Google Gemini AI, hosted separately on Render.
For the backend implementation and API source code, see the **[Daily Orders AI Backend Repository](https://github.com/AmmarAlhashemi/Daily-Orders-AI-Backend)**.

---

## 🎯 Business Context & Customization

> **Note:** This project is a custom-tailored software solution engineered specifically for a dedicated operational workflow. Unlike generic order-processing software, this application was built around specific business rules, Excel ledger structures, transaction classification logic, balancing requirements, and monthly rollover procedures defined for the target workflow.
>
> The application generates standardized and balanced Excel ledgers specifically structured for downstream ingestion by the [Orders Management System](https://github.com/AmmarAlhashemi/Orders-Management-System).
>
> While the architecture demonstrates a reusable client-server approach for Python automation and AI-assisted data processing, the underlying business rules and Excel structures are tightly coupled to the intended workflow.

---

## ⚠️ Compatibility Notice & Requirements

> **Important:** The Desktop application requires a Windows Operating System with Microsoft Excel installed locally.

- **Operating System:** Windows 10 / 11 (64-bit).
- **Microsoft Office:** A local installation of Microsoft Excel is required. Excel COM Automation through `pywin32` is used for Windows-based Excel operations.
- **Python:** Python 3.12.10 or higher.
- **Internet Connection:** Required for communication with the remote FastAPI backend.
- **Cross-Platform Limitations:** Linux and macOS are not supported out of the box for the Desktop application due to Windows-specific Excel COM automation.

---

## 🌟 Key Features & Business Value

- **AI-Powered Order Extraction:** Sends unstructured daily remittance text to the FastAPI backend, which processes it using Google Gemini and returns structured transaction records.

- **Standardized Excel Ledger Generation:** Produces clean, structured Excel ledgers containing transaction details, recipient information, account numbers, amounts, remarks, designated entities, and AI warnings when applicable.

- **Smart Transaction Splitting:** Distributes daily transactions between `ENTITY_A` and `ENTITY_B` using configurable processing methods:
  - Sequential / One-by-One
  - Bank-Based Routing
  - AI-Balanced Distribution

- **Automated Monthly Ledger Rollover:** Automatically detects month transitions when the application starts, archives previous daily order files into structured monthly directories, and rolls active ledgers into the new month with the required opening balance carry-over.

- **Primary & Fallback AI Models:** Uses a primary Gemini model with a fallback model when the primary model is unavailable.

- **Secure API Architecture:** Gemini API credentials are isolated from the Desktop application and handled by the remote backend through environment variables.

- **Responsive Desktop UI:** Built with CustomTkinter with background processing, live status updates, and user feedback mechanisms to prevent long-running AI and Excel operations from freezing the interface.

---

## 🏗️ Project Architecture

The project follows a **Modular Client-Server Architecture**. The Desktop application handles the user interface, local Excel processing, file management, and monthly rollover operations, while the remote FastAPI backend handles Gemini AI processing.

```text
                         ┌──────────────────────────┐
                         │  Desktop Application     │
                         │                          │
                         │  CustomTkinter           │
                         │  Excel Processing        │
                         │  File Management         │
                         │  Monthly Rollover        │
                         └────────────┬─────────────┘
                                      │
                                      │ HTTPS / HTTP
                                      ▼
                         ┌──────────────────────────┐
                         │     FastAPI Backend      │
                         │                          │
                         │  Gemini AI Processing    │
                         │  Pydantic Schemas        │
                         │  API Endpoints           │
                         └────────────┬─────────────┘
                                      │
                                      ▼
                              Google Gemini API
```

---

## 🛠️ Tech Stack

### Desktop Application

- Python 3.12.10
- CustomTkinter
- Pandas
- OpenPyXL
- PyWin32
- Requests
- **Primary Model:** Gemini 3.6 Flash
- **Fallback Model:** Gemini 3.5 Flash

---

## 🧪 Demo & Testing

You can test the application using the sample data below:

```text
Ahmed Ali 123456789 
15,000
Mohammed Saleh 
987654321 urgent
25,000
Khaled Hassan 456789123 bank_2 10,000
Omar Abdullah 741852963 
18?000
Yasser Ibrahim 
369258147 
22,0000
Samir Ahmed 852963741 12,500
```

> **Note:** This is fake data for testing only. No real customer data is used.

---

## 🚀 Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/AmmarAlhashemi/Daily-Orders-AI-Processor
cd Daily-Orders-AI-Processor
```

### 2. Create a Virtual Environment

```bash
python -m venv .venv
.venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Run the Application

```bash
python main.py
```

The Desktop application communicates with the FastAPI backend through HTTP. Gemini API credentials are not stored in the Desktop application.

> **Note:** No Gemini API key is required on the Desktop client as the API credentials are securely configured on the Backend server.

---

## 🛡️ Data Privacy & Dummy Data Notice

This repository does not contain real client financial records or API credentials.

The included Excel structures and demonstration data are anonymized / dummy data intended solely for demonstrating the application's processing workflow.

---

## 📝 License

This project is licensed under the MIT License.

See the [`LICENSE`](LICENSE) file for details.
