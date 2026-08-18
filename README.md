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

## Performance Optimization

The data handling logic was optimized by eliminating unnecessary linear searches when determining the insertion position.

### Before

The program previously relied on `max_row` and iterated through columns to find the next available position:

* Determining the insertion point required additional calculations.
* Finding the next column could require traversing existing columns, resulting in **O(n)** complexity.

### After

The current implementation stores the last relevant row in cell `M1` and retrieves it directly:

```python
ultima_linha = aba["M1"].value
```

The next position is then calculated directly from the stored value:

```python
prog = aba.cell(row=ultima_linha, column=1).value + 1
```

This changes the complexity of **finding the insertion position** from **O(n)** to **O(1)**.

### Impact

| Operation               |                      Before |     After |
| ----------------------- | --------------------------: | --------: |
| Determine last position |                        O(n) |  **O(1)** |
| Find next column        |                    **O(n)** |  **O(1)** |
| Insert/update data      |                   Unchanged | Unchanged |

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
git clone https://github.com/yourusername/Production-Time-Logger.git
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
