This program takes one argument and takes data from the given website which it stores in an excel file and on mongodb.

Install Python, pymongo, xlsxwriter, selenium and download the chrome web driver.
Run program by entering "python reviewscrapper.py {company website}" eg. https://www.cylex-uk.co.uk/

Change "s" to where you have your chrome driver installed.
Change "book" to change excel file name. It will need to be changed each time the program is run as xlsxwriter cannot open excel files, only modify them
while they are open as far as I've found out so far.