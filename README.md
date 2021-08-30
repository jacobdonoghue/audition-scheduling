# Audition Scheduling Scripts
## Author: Jacob Donoghue

## Background
To minimize the time it takes for both groups and auditionees to hold solo auditions, I wrote a two scripts.

`create-slots.py` creates optimal slots and writes the data to an excel file. This allows for editing if anyone has time conflicts during a certain period. 

`send-emails.py` reads the excel file, validates the entries to ensure no repeat auditions or time conflicts, and, if valid, sends emails out to all auditionees! 

## Dependencies
This program relies on:
1. Python v3.8+
2. [xlsxwriter](https://xlsxwriter.readthedocs.io/) to write to excel
3. [pandas](https://pandas.pydata.org/) to read from excel
4. [decouple](https://pypi.org/project/python-decouple/) to access .env environment variables

## Before running

1. Install all dependencies
2. Set up a gmail [account for development](https://realpython.com/python-send-email/) and store the corresponding username and password in a .env file
3. Store all auditions locations in `locations.txt` -- 1 group per line, formatted: `group:location` with no spaces
4. Store all auditionee emails in `emails.txt`, each email separated by a colon

Note: To test, run main() in send-emails.py with 'False' as a parameter, so you don't end up sending out emails prematurely!
