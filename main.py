import sys
import openpyxl as xl
import datetime

# PTD/YTD report
path = sys.argv[1]
# Date
arg_date = sys.argv[2]
edit_date = arg_date
# Revenue Summary spreadsheet
rs = sys.argv[3]


# Load the PTD/YTD
ptdytd = xl.load_workbook(path)
ptd_sheet = ptdytd.active

# Identify the cells with the CC Allowances
visa = ptd_sheet['F36'].value
mastercard = ptd_sheet['F35'].value
discover = ptd_sheet['F34'].value
amex = ptd_sheet['F33'].value

####
# Load the RS spreadsheet and make the CC Summary the active workbook
#
ccs_wb = xl.load_workbook(rs)
ccs_sheet = ccs_wb['Credit Card Summary']
rs_sheet = ccs_wb['Revenue Summary']

# Add 4 to the date variable, since the data starts on row 5
edit_date = int(edit_date) + 4


####
# Set variables for pasting locations
#
def ccsPaste():
    visa_paste = ccs_sheet.cell(row = edit_date, column = 3)
    visa_paste.value = visa
    mastercard_paste = ccs_sheet.cell(row = edit_date, column = 4)
    mastercard_paste.value = mastercard
    discover_paste = ccs_sheet.cell(row = edit_date, column = 5)
    discover_paste.value = discover
    amex_paste = ccs_sheet.cell(row = edit_date, column = 10)
    amex_paste.value = amex



#####################################
#                                   #
# Revenue summary workbook section  #
#                                   #
#####################################

# Variables for revenue summary
revenue = ptd_sheet['F181'].value
tax = ptd_sheet['F192'].value
misc = ptd_sheet['F121'].value
outlet = ptd_sheet['F130'].value
phone = ptd_sheet['F157'].value
cash = ptd_sheet['F86'].value
overshort = ptd_sheet['F139'].value

def rsPaste():
    revenue_paste = rs_sheet.cell(row = edit_date, column = 3)
    revenue_paste.value = revenue
    tax_paste = rs_sheet.cell(row = edit_date, column = 4)
    tax_paste.value = tax
    misc_paste = rs_sheet.cell(row = edit_date, column = 5)
    misc_paste.value = misc
    outlet_paste = rs_sheet.cell(row = edit_date, column = 6)
    outlet_paste.value = outlet
    phone_paste = rs_sheet.cell(row = edit_date, column = 7)
    phone_paste.value = phone
    cash_paste = rs_sheet.cell(row = edit_date, column = 8)
    cash_paste.value = cash
    overshort_paste = rs_sheet.cell(row = edit_date, column = 10)
    overshort_paste.value = overshort


#
# See if it's the end of the month, and print a little reminder if so
#
def eom():
    months = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]
    days = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    this_month = datetime.date.today().month
    eom = days[this_month - 1]
    if eom == int(arg_date):
        print('\nIt looks like it\'s the end of the month. Don\'t forget to clean up your files!\n')


#
# Send the info to the RS spreadsheet
#
def wbSave():
    ccs_wb.save(filename=rs)

#
# Allow the user to cancel the script in case of an error
#
def cancel():
    choice = input('\n** ** Would you like to continue? Y/N ** **\n>> ')
    if choice.lower == 'y' or 'yes':
        return choice
    else:
        sys.exit(0)

#
# A little message that prints out to let you know the script is complete.
# Will also let you know if there may have been an error with the date.
def complete():
    if str(arg_date) == str(rs_sheet.cell(row = edit_date, column = 2).value) and str(ccs_sheet.cell(row = edit_date, column = 2).value):
        print('Done!')
    else:
        print('Something didn\'t add up, better check things out manually.')

def doubleCheck():
    print('\n ** ** Before we continue, here are the files you\'re working with: ** **\n')
    print(f'PTDYTD File: {path}\n')
    print(f'Destination File: {rs}\n')
    print(f'\nDate: {arg_date}\n')
    print(f'\nVisa: {visa}\n')
    print(f'Mastercard: {mastercard}\n')
    print(f'Discover: {discover}\n')
    print(f'Amex: {amex}\n\n')
    print(f'Revenue: {revenue}\n')
    print(f'Tax: {tax}\n')
    print(f'Misc: {misc}\n')
    print(f'Phone: {phone}\n')
    print(f'Cash: {cash}\n')
    print(f'Over/Short: {overshort}\n')


if __name__ == '__main__':
    ccsPaste()
    rsPaste()
    doubleCheck()
    wbSave()
    eom()
    complete()
