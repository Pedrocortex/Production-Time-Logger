# Production Time Logger

A lightweight Python desktop application for recording production operation times and organizing the data in an Excel spreadsheet for later analysis.

## Features

- Desktop interface built with Tkinter
- Record operation time
- Select work shift:
  - Morning
  - Afternoon
  - Night
- Select production pace:
  - Normal
  - Intense
  - Critical
- Automatically creates a new daily record
- Updates existing daily records
- Stores all data in an Excel workbook
- Color-coded shifts for better visualization

## Technologies

- Python 3
- Tkinter
- OpenPyXL

## Project Structure

```
.
├── interfacEst.py      # User interface
├── guuar.py            # Excel data management
├── Dados1.xlsx         # Data storage
└── README.md
```

## Installation

Clone the repository:

```bash
git clone https://github.com/yourusername/production-time-logger.git
cd production-time-logger
```

Install the required dependency:

```bash
pip install openpyxl
```

## Usage

Run the application:

```bash
python interfacEst.py
```

Fill in:

- Operation Time
- Work Shift
- Production Pace

Click **Save Data** to record the information into the Excel spreadsheet.

## Data Organization

Each day is automatically organized in the spreadsheet with:

- Date
- Work Shift
- Operation Time
- Production Pace
- Record Number

This structure makes it easy to create dashboards, charts, and productivity reports using Excel or Power BI.

## Future Improvements

- Export to CSV
- Statistics dashboard
- Average operation time
- Daily and monthly reports
- Data visualization charts
- SQLite database support
- PDF report generation

## Requirements

- Python 3.10+
- OpenPyXL

## Demo 

https://github.com/user-attachments/assets/4e408adc-b953-473f-9fec-f10bde14109d



## License

This project is licensed under the MIT License.
