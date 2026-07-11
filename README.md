# Контроль пациентов с инфекционными заболеваниями
Скрипт анализирует базу данных пациентов и используется заведующим отделением пластической хирургии и главным врачом для контроля пациентов с инфекционными заболеваниями при планировании операций

# Patient Disease Screening

> Python application for identifying surgical patients with infectious diseases based on medical records.

## Description

This project analyzes surgical schedules together with patient medical notes to identify patients with infectious diseases before planned operations.

The application automatically classifies diseases using keyword recognition, prepares detailed reports and generates summary statistics for medical management.

Important:

The project requires access to the hospital scheduling database.

Without the original database structure the application cannot be executed.

## Business Goal

The application helps medical staff:

- identify patients with infectious diseases before surgery
- improve operating room safety
- prepare infection control reports
- monitor disease statistics over time

The tool is used by the Head of Plastic Surgery Department and the Chief Medical Officer.

## Features

- SQL database integration
- Medical schedule analysis
- Patient record processing
- Disease classification using text recognition
- Hepatitis detection
- HIV detection
- Syphilis detection
- Treatment status detection
- Monthly disease statistics
- Excel report generation
- Conditional formatting
- Telegram notifications

## Tech Stack

- Python
- pandas
- NumPy
- SQLAlchemy
- Firebird SQL
- xlsxwriter
- python-dotenv
- pyTelegramBotAPI

## How It Works

1. Loads surgical schedule
2. Loads patient medical notes
3. Filters surgical procedures
4. Searches medical comments for infectious diseases
5. Determines treatment status
6. Generates patient list
7. Creates monthly summary statistics
8. Exports formatted Excel reports

## Example / Demo

### Input

Hospital database:

- Surgery schedule
- Patient records
- Medical comments

### Output

Excel workbook containing:

- All surgeries
- Patients with detected diseases
- Monthly disease statistics

This project can be used for:

- healthcare analytics
- medical reporting
- patient safety
- clinical data analysis
- hospital automation
