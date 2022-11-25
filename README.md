# What is the Auto Valuer?

Auto-valuation is a program for the Truman Besief Program that takes financial statements for a company from Morningstar and valuates the information to recieve financial information (debt, profit, etc) about certain companies such as AAPL and MSFT.

https://www.youtube.com/watch?v=-uG-n2eKCCA

# How does it work?

Its a python built program that writes date into excel sheets. It uses chromedriver to access Chrome and travel to Morningstar.com to download the financial sheets. If chromedriver fails then the user must manually place the sheets into the program to evaulate.
.....


# Updates

yahoo-fin has been producing errors that causes the get_live_stock_info to crash. Yahoo_fin has recently been updated to fix it.


# To Do
Webdriver_manager fails when attempting to retrieve ticker. Will be fixed


The program fails to write to cell due to a missing arguement. Will be fixed



