# Digital Receipt Extraction and Analysis Script

## Overview

This script is designed to help you extract digital receipts from stores like **Albertsons**, **Safeway**, and others, then save the data into a database for analysis. It parses emails containing receipts, extracts relevant information, and stores the data in an **Access Database**.

**Fun Note**:  
I joke with my spouse that the only reason I made this script was to prove how much I spend on bags from the store. But the real story is:  
1. **I was bored** and thought this would be fun.  
2. **I spend nearly a thousand dollars a month at these stores** and wanted to understand where my money goes.

With this script, I can analyze spending trends such as:
- **Price over time**: Track how the price of items increases over time to create my own inflation statistics.
- **Spending patterns**: Find out what times or days I shop most frequently.
- **Purchasing frequency**: Determine what items I buy most often, so I can buy in bulk and save money.
- **Inventory management**: The ultimate goal is to build a web app that not only tracks inventory but also helps me make informed decisions. The app would track expiration dates, so if I'm at the store and need cheese but can't remember how old the cheese I already have is, I can easily check. Additionally, the app would learn usage trends, so if I’m making soup and need 3 cups of broth, it could estimate how much I have left based on my consumption patterns. For example, if it detects that I typically use a certain amount of broth, it could tell me I have about 80% left, helping me know whether I need to buy more.

## Features
- Extract data such as **transaction number**, **authorization time**, **total amount**, **items purchased**, and more.
- Automatically detect duplicate receipts to prevent multiple entries.
- Save extracted receipt data into a **Microsoft Access Database** (`Receipts.accdb`).
- Analyze receipt history to understand spending habits.

## Setup

1. **Create a Gmail account**:
    - Gmail is NOT requited for this script it can be any email as long as its connected to Outlook.
3. **Update the email account name in the script**:
   - Go to **line 28** and replace `"Account name"` with your actual Gmail account (e.g., `"spending@gmail.com"`).
4. **Set output location**:
   - By default, the script saves data to an Access database file called `Receipts.accdb` in your `MyDocuments` folder. You can change the path if you prefer a different location.

5. **Run the script** to start extracting receipts from your email.

## Script Details

The script works by connecting to your **Gmail account** using Outlook and parsing emails from a designated folder (usually the "Inbox"). It extracts receipt data using regular expressions and stores it in an Access database. 

### Key Features of the Script:
- **Transaction Information**: Extracts key details like **Transaction Number**, **Authorization Time**, **Amount**, and **Card Ending**.
- **Item Details**: Extracts details for each item, such as **Product Name**, **Product Price**, and **Quantity**.
- **Data Validation**: Checks for duplicate receipts to ensure data integrity.
- **Database**: The script stores all extracted data in a Microsoft Access database (`Receipts.accdb`) for easy querying and analysis.

### How It Works:
1. The script connects to your Gmail account via Outlook.
2. It retrieves emails from the **specified folder** (Inbox).
3. It processes each email to extract receipt details using regular expressions.
4. Extracted data is stored in the Access database.
5. Duplicate checks are performed to ensure the same receipt is not processed multiple times.

### Extracted Data Includes:
- Email address
- Transaction number
- Authorization date and time
- Reference number
- Amount
- Item details (name, price, quantity)

## Example Use Case

1. After processing your digital receipts, you can:
   - Track how much you’re spending over time.
   - Analyze spending patterns for individual items.
   - Create your own **inflation statistics** based on your purchases.

2. You can even envision using this data to automate home inventory management or track the **expiration dates** of food items. A possible future feature could integrate with a web app to suggest when to buy items based on your spending history.

## Requirements

- **Outlook** installed on your device for email parsing.
- **Microsoft Access** to create and manage the database.

## License
This github project is distributed "as is" under the GPL License, WITHOUT WARRANTY OF ANY KIND. See LICENSE for details.
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License.
This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.


## Disclaimer

This script is provided "as is" with no warranty or guarantee. Use at your own risk.
