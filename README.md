# Money Management System

A Python-based money management system that helps you track and manage your expenses and income. This project uses the `openpyxl` library to work with Excel files for recording and analyzing financial data.

## Features

- Add income and expenses
- Categorize transactions (e.g., Food, Rent, Entertainment, etc.)
- Automatically updates balance
- Export and save records in an Excel spreadsheet using Openpyxl
- View financial summary (monthly, weekly, or custom time periods)
- Simple and user-friendly interface

## Prerequisites

Before you begin, make sure you have Python installed on your system. You will also need the following Python packages:

- `openpyxl`

You can install the package using pip:

```bash
pip install openpyxl
```

## Installation

1. Clone this repository:

```bash
git clone https://github.com/Sudhanshu611/Money-Management-System-using-Ms-Excel.git
```

2. Navigate to the project directory

3. Install the dependencies:

```bash
pip3 install openpyxl
```

## Usage

1. Run the main.py file to start the program:

```bash
python main.py
```

2. Follow the on-screen instructions to add income or expenses.

3. The system will automatically update the balance and save the data in an Excel file.

4. You can view your financial summary anytime.

## File Structure

```bash
.
├── main.py               # Main script to run the program
├── README.md             # Project documentation
└── INSTRUCTION.md      # INSTRUCTION for the project
```

## How It Works

1. The system prompts the user to enter income or expense amounts and assigns them to specific categories.
2. Data is stored in an Excel file (finances.xlsx), which is managed using the openpyxl library.
3. The program keeps a running balance of total income and expenses and allows you to view summaries.

## Future Improvements

1. Add data visualization (graphs and charts) to analyze spending habits.
2. Implement a budgeting feature to set limits on specific categories.
3. Create a user authentication system for multiple users.

## Contributing

If you'd like to contribute, feel free to fork the repository and submit a pull request. For major changes, please open an issue first to discuss what you'd like to change.
