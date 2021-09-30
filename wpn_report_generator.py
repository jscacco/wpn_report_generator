import openpyxl
import datetime
import pickle
import sys
import getopt
import math
import time

LG_ORG_ID = 40658
DG_ORG_ID = 35657
CURRENCY = "USD"

DICT_FILENAME = "files/wotc_sku_dict.txt"
NEW_SKUS = {}

DESC_COL_NUM = 3
QTY_COL_NUM = 4

class Transaction:
    def __init__(self, store):
        self.store = store
        self.wpn_org_id = None
        self.date = None
        self.transaction_id = None
        self.quantity_sold = None
        self.unit_price = None
        self.total_sale_price = None
        self.wotc_sku = None
        self.fg_product_desc = None
        self.currency = CURRENCY
        self.customer = None


    def __str__(self):
        output = ""
        output += "\n----------------------------------------\n"
        output += "Transaction " + str(self.transaction_id) + ": "
        output += "\n\n    WPN Org. ID: " +  str(self.wpn_org_id)
        output += "\n    wpn_org_id type: " +  str(type(self.wpn_org_id))
        output += "\n\n    Date: " + str(self.date)
        output += "\n    self.date type: " + str(type(self.date))
        output += "\n\n    Transaction ID: " + str(self.transaction_id)
        output += "\n    self.transaction_id type: " + str(type(self.transaction_id))
        output += "\n\n    Quantity Sold: " + str(self.quantity_sold)
        output += "\n    self.quantity_sold type: " + str(type(self.quantity_sold))
        output += "\n\n    Unit Price: " + str(self.unit_price)
        output += "\n    self.unit_price type: " + str(type(self.unit_price))
        output += "\n\n    Total Sale Price: " + str(self.total_sale_price)
        output += "\n    self.total_sale_price type: " +  str(type(self.total_sale_price))
        output += "\n\n    FG Product Desc: " + str(self.fg_product_desc)
        output += "\n    self.fg_product_desc type: " + str(type(self.fg_product_desc))
        output += "\n\n    WPN SKU: " + str(self.wotc_sku)
        output += "\n    self.wpn type: " + str(type(self.wotc_sku))
        output += "\n----------------------------------------\n"
        
        return output
    
    
    def getInfo(self, line_report_filename, r):
        """Given a line report filename and a row, read in the attributes.
           Note: This will not fill wotc SKU."""
        
        # Set wpn_org_id
        # This can be done whenever, but we do it here
        self.wpn_org_id = LG_ORG_ID if self.store == "LG" else DG_ORG_ID
        
        # Open the line report
        wb = openpyxl.load_workbook(line_report_filename)
        sheet = wb[wb.sheetnames[0]]
        
        for c in range(1, sheet.max_column + 1):
            this_value = sheet.cell(row=r, column=c).value
            
            if c == 1:
                # ID
                self.transaction_id = this_value
            elif c == 2:
                # Date
                self.date = this_value
            elif c == 3:
                # Description
                self.fg_product_desc = this_value
            elif c == 4:
                # Qty
                self.quantity_sold = this_value
            elif c == 5:
                # Retail (unit price)
                self.unit_price = this_value
            elif c == 6:
                # Subtotal
                self.total_sale_price = this_value
            elif c == 10:
                self.customer = this_value
                
                
    def adjustFormatDate(self):
        """If self.date is a string, convert it to a datetime object."""
        if type(self.date) == type('str'):
            split_str = self.date.split('-')
            self.date = datetime.datetime(int(split_str[0]), \
                                          int(split_str[1]), \
                                          int(split_str[2]))
            
            
    def adjustFormatPrices(self):
        """If self.unit_price is a string, convert it to a number.
           Same for self.total_sale_price."""
        
        if type(self.unit_price) == type('str'):
            self.unit_price = float(remove_commas(self.unit_price[1:]))
            
        if type(self.total_sale_price) == type('str'):
            self.total_sale_price = float(remove_commas(self.total_sale_price[1:]))
            

    def getInfoAndFormat(self, line_report_filename, r):
        """ Consolidates the previous three methods."""
        self.getInfo(line_report_filename, r)
        self.adjustFormatDate()
        self.adjustFormatPrices()

        
    """
    def isValidTransaction(self):
        #Returns True if the transaction should be added to the report.

        # Basically, this catches all the transactions we don't want to add into
        # the report. This includes trade ins, admission fees, singles,
        # transaction with a total sale value not greater than zero,
        # and transactions made using the house account.

        # NOTE: this will also filter out any product which has the below keywords in
        # its name. I don't think there are any WotC products which contain those
        # keywords, but it's worth mentioning in case that changes.
        
        desc = self.fg_product_desc.lower()
        return not (desc.__contains__("trade in") or \
                    desc.__contains__("admission") or \
                    desc.__contains__("draft fnm") or \
                    desc.__contains__("booster league") or \
                    desc.__contains__("venue") or \
                    desc.__contains__("single") or \
                    self.quantity_sold < 0)
    """                
    
    def enterIntoWpnReport(self, wpn_report_filename, r):
        """Enter the values of the transaction into the WPN filename."""

        # Open the report
        wb = openpyxl.load_workbook(wpn_report_filename)
        sheet = wb[wb.sheetnames[0]]

        for c in range(1, sheet.max_column + 1):
            this_cell = sheet.cell(row=r, column=c)
            if c == 1:
                # WPN id
                this_cell.value = self.wpn_org_id
            elif c == 2:
                # Date
                this_cell.value = self.date
            elif c == 4:
                # Transaction ID
                this_cell.value = self.transaction_id
            elif c == 6:
                # WotC SKU
                this_cell.value = self.wotc_sku
            elif c == 9:
                # Product Desc.
                this_cell.value = self.fg_product_desc
            elif c == 10:
                # Quantity Sold
                this_cell.value = self.quantity_sold
            elif c == 11:
                # Unit Price
                this_cell.value = self.unit_price
            elif c == 12:
                # Total Sale Price
                this_cell.value = self.total_sale_price
            elif c == 13:
                # Currency
                this_cell.value = self.currency

        wb.save(wpn_report_filename)

        
        
def get_line_report_col_descs(line_report_filename):
    """Given a line report, return a list of the column descriptions.
       Each description's column is equal to its index in the returned list.
       (Note that this means the list will always have filler info at its head)
       IMPORTANT: This assumes that the column descriptions will be in the first row.
       Update 7/29: I don't think this function is necessary."""

    # Open the line report
    col_descs = [None]
    wb = openpyxl.load_workbook(line_report_filename)
    sheet = wb[wb.sheetnames[0]]

    # Iterate through the first row, adding the column descriptions to the list
    for c in range(1, sheet.max_column + 1):
        col_descs.append(sheet.cell(row=1, column=c).value)

    # Return the list of descs
    return col_descs
    

"""
def fill_wpn_report(store, line_report_filename, wpn_report_filename, wotc_skus):
    #Put it all together.
    wb = openpyxl.load_workbook(line_report_filename)
    sheet = wb[wb.sheetnames[0]]
    num_entries = sheet.max_row - 1

    current_line_row = 2
    current_wpn_row = 5

    while current_line_row <= num_entries:
        this_transaction = Transaction(store)
        this_transaction.getInfoAndFormat(line_report_filename, current_line_row)
        current_line_row += 1
        if this_transaction.isValidTransaction():
            set_wotc_sku(this_transaction, wotc_skus)
            this_transaction.enterIntoWpnReport(wpn_report_filename,current_wpn_row)
            current_wpn_row += 1
"""

def isValidTransaction(line_report_filename, line_row):
    """Given a row in the line report, return true if that row contains a valid
       transaction."""

    wb = openpyxl.load_workbook(line_report_filename)
    sheet = wb[wb.sheetnames[0]]
    
    desc = sheet.cell(row=line_row, column=DESC_COL_NUM).value.lower()
    quantity_sold = sheet.cell(row=line_row, column=QTY_COL_NUM).value
    
    return not (desc.__contains__("trade in") or \
                desc.__contains__("admission") or \
                desc.__contains__("draft fnm") or \
                desc.__contains__("booster league") or \
                desc.__contains__("venue") or \
                desc.__contains__("single") or \
                quantity_sold < 0)
    

def fill_wpn_report(store, line_report_filename, wpn_report_filename, wotc_skus):
    """Put it all together."""
    wb = openpyxl.load_workbook(line_report_filename)
    sheet = wb[wb.sheetnames[0]]
    num_entries = sheet.max_row - 1

    current_line_row = 2
    current_wpn_row = 5

    while current_line_row <= num_entries:
        if isValidTransaction(line_report_filename, current_line_row):
            this_transaction = Transaction(store)
            this_transaction.getInfoAndFormat(line_report_filename, current_line_row)
            set_wotc_sku(this_transaction, wotc_skus)
            this_transaction.enterIntoWpnReport(wpn_report_filename, current_wpn_row)
            current_wpn_row += 1
        current_line_row += 1


def pickled_dict_setup(filenames):
    """Setup for the wpn SKU dict given a list of previous WPN reports.
       Should only be needed once."""

    this_dict = {}
    
    for report in filenames:
        print(report)
        wb = openpyxl.load_workbook(report)
        sheet = wb[wb.sheetnames[0]]

        # fg_desc = col 9
        # wpu_sku = col 6

        for r in range(5, sheet.max_row + 1):
            this_fg_desc = sheet.cell(row=r, column=9).value
            if this_fg_desc not in this_dict:
                this_dict[this_fg_desc] = sheet.cell(row=r, column=6).value

    file = open(DICT_FILENAME, 'wb')
    pickle.dump(this_dict, file)
    file.close()
    
    
def set_wotc_sku(transaction, wotc_skus):
    """Gets and sets the wotc sku for the item."""
    
    if transaction.fg_product_desc == None:
        # If there is not description...
        # (I don't think this would ever trigger, but adding it just in case.)
        print("\n\n\n!Error! Transaction has no product description.\n\n")
    elif transaction.fg_product_desc in wotc_skus:
        # If the SKU is already in our dictionary...
        transaction.wotc_sku = wotc_skus[transaction.fg_product_desc]
    elif transaction.fg_product_desc in NEW_SKUS:
        # If the SKU is new, but we've already seen it this session...
        transaction.wotc_sku = NEW_SKUS[transaction.fg_product_desc]
    else:
        # We need to input the SKU and add it in.
        confirmation = "n"
        while confirmation.lower() not in ['', 's']:
            new_sku = input("\nNo SKU found for " + transaction.fg_product_desc + \
                            ". Please enter it now: ")
            confirmation = input("Confirm: is " + new_sku + " the correct SKU for " + \
                                 transaction.fg_product_desc + \
                                 "? (Enter nothing if correct, S to skip, and any other button to re-enter the SKU)")

        if confirmation.lower() == 's':
            new_sku = "SKIPPED"
        NEW_SKUS[transaction.fg_product_desc] = new_sku
        transaction.wotc_sku = new_sku
        print("Working...")
        

def add_new_skus(wotc_skus):
    """Adds the SKUS in NEW_SKUS to the dict file."""
    updated_skus = wotc_skus
    updated_skus.update(NEW_SKUS)
    file = open(DICT_FILENAME, 'wb')
    pickle.dump(updated_skus, file)
    file.close()

    
def remove_commas(this_str):
    """Removes the commas from a given string."""
    new_str = ""
    for char in this_str:
        if char != ',':
            new_str += char
    return new_str
        

def read_wotc_skus():
    """Reads in the dict of wpn skus."""
    
    # https://www.geeksforgeeks.org/how-to-read-dictionary-from-file-in-python/
    with open(DICT_FILENAME, 'rb') as handle:
        data = handle.read()
    return pickle.loads(data)


def seconds_to_formatted_time(seconds):
    """Given a number representing a length of time in seconds, return a string which
       formats the time into a more readable format."""

    seconds = int(seconds)
    
    h = math.floor(seconds / 3600)
    m = math.floor((seconds / 60) % 60)
    s = math.floor(seconds % 60)
    
    return "{h}h, {m}min, {s}sec".format(h=h, m=m, s=s)

    
def generate_report(line_report_filename, wpn_report_filename, store):
    """The bulk of the program, generate a new report."""
    
    # line_report_filename = "files/lg_0321_reports_sales_listings_transaction_line.xlsx"
    # wpn_report_filename = "files/40658_FairGameLaGrange_POSData_0321.xlsx"
    # store = "LG"

    # let's keep track of how long this takes
    start = time.time()
    wotc_skus = read_wotc_skus()
    
    print("Working... (This will take a few minutes)")
    fill_wpn_report(store, line_report_filename, wpn_report_filename, wotc_skus)
    add_new_skus(wotc_skus)

    end = time.time()
    
    print("Finished! Time elapsed: " + seconds_to_formatted_time(end - start))


def print_help_info():
    """Show how to use the program."""
    
    print("\nGeneral usage (for report generation):")
    print("\twpn_report_generator.py -l <line report filename> -w <wpn report filename> -s <store (DG or LG)>")
    print("\nAll parameters:")
    print("\t--line=<line report filename>")
    print("\t--wpn=<wpn report filename>")
    print("\t--store=<store (DG or LG)>")
    print("\t-u (if replacing a WotC SKU)")
    print("\t-o (if looking up a WotC SKU)")
    

def lookup_sku():
    """Looks up a sku."""
    # Get the name of the item to be updated from the user and display the current SKU
    wotc_skus = read_wotc_skus()
        
    desc = input("Please enter the name of the item whose SKU you would like to lookup:\n")
    try:
        current_sku = wotc_skus[desc]
    except KeyError:
        print("No item found with that description")
        return

    print("The current SKU for " + desc + " is: " + str(current_sku) + ".")

    
def update_sku():
    """Update an item's SKU."""
    wotc_skus = read_wotc_skus()
        
    desc = input("Please enter the name of the item whose SKU you would like to update:\n")
    try:
        current_sku = wotc_skus[desc]
    except KeyError:
        print("No item found with that description")
        return

    print("The current SKU for " + desc + " is: " + str(current_sku) + ".")
    
    # Get the new SKU from the user
    confirm = 'n'
    while confirm not in ['y', '', 'yes']:
        new_sku = input("Please enter the new SKU:\n")
        confirm = input("Is " + str(new_sku) + " correct? (Enter 'y', 'yes', or nothing to confirm. Enter any other key to resubmit.")

    # Update the SKU dict
    wotc_skus[desc] = new_sku
    file = open(DICT_FILENAME, 'wb')
    pickle.dump(wotc_skus, file)
    file.close()
    

def main():
    line_report_filename = ''
    wpn_report_filename = ''
    store = ''
    
    try:
        opts, args = getopt.getopt(sys.argv[1:], "hl:w:s:uo", ["line=", "wpn=", "store="])
    except:
        print_help_info()
        sys.exit(2)

    # If there are no arguments, print the help and exit
    if len(opts) == 0:
        print("HERE")
        print(opts)
        print_help_info()
        sys.exit(2)

    # Otherwise, determine functionality based on arguments provided
    for opt, arg in opts:
        if opt == '-h':
            print_help_info()
            sys.exit()
        elif opt == '-u':
            update_sku()
            sys.exit()
        elif opt == '-o':
            lookup_sku()
            sys.exit()
        elif opt in ['-l', '--line']:
            line_report_filename = arg
        elif opt in ['-w', '--wpn']:
            wpn_report_filename = arg
        elif opt in ['-s', '--store']:
            store = arg

    # If we are generating a report and don't have the three arguments we need,
    # inform the user that they need to provide them and exit.
    if line_report_filename == '' or wpn_report_filename == '' or store == '':
        print("Please include a line report filename, wpn report filename, and store.")
        sys.exit(2)

    generate_report(line_report_filename, wpn_report_filename, store)

    
if __name__ == "__main__":
    main()
