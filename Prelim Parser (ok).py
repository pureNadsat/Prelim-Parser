"""
// Prelim Parser \\

Searches prelim for relevant information:

Property address, vesting, # of liens and lender, type of property, check for tax/mechanics liens 

and solar endorsements

"""

import PyPDF2
import re
import win32com.client as win32

# PDF will be dragged from email onto desktop

# Open prelim
with open(r'?', 'rb') as pdf:
    pdf_reader = PyPDF2.PdfFileReader(pdf)
    
    # Compile pages of prelim into a string
    prelim = ''
    for page in range(pdf_reader.numPages):
        prelim += pdf_reader.getPage(page).extractText()

pdf.close()

# Remove line breaks and extra spaces
prelim = prelim.replace('   ', ' ').replace('  ', ' ')

# Function for finding a subset string within a string
def parser(first, last, string):
    words = [first, '(.*?)', last]
    try:
        word_search = re.search(''.join(words), string)
        if word_search != None:
            return(word_search.group(1))
        else:
            return('None')
    except ValueError:
        return("Unable to read information")

## Check the title company and choose substrings accordingly ##

     
# For Old Republic Title syntax
if 'old republic title' in prelim[:100].lower():
    
    title_company = 'Old Republic Title'
    
    # Assign variables using find_between function
    prop_address = parser('Property Address: ', 
                             'In response to the above referenced application',
                             prelim)

    vesting = parser('interest at the date hereof is vested in:',
                        'The land referred to in this Report',
                        prelim)

    lender = ['Beneficiary/Lender:', 'Dated:']
    
    lien_amount = ['under the terms thereof,Amount:', 'Trustor/Borrower:']
                          
   
# For First American syntax
elif 'first american' in prelim[:100].lower():
    
    title_company = 'First American Title'
    
    # Variable assignment
    prop_address = parser('Property: ', 
                             'PRELIMINARY REPORT', 
                             prelim)
    
    vesting = parser('interest at the date hereof is vested in: ',
                        'The estate or interest in the land',
                        prelim)[1:]
    
    lender = ['title insurance company beneficiary: ', 
              'Order Number:']
   
# for Stewart Title    
elif 'stewart title' in prelim[:100].lower():
    
    title_company = 'Stewart Title Company'
    
    prop_address = parser('Property Address: ',
                             'In response to the above referenced application',
                             prelim)
    
    vesting = parser('interest at the date hereof is vested in:',
                        'Order No.:',
                        prelim)
    
    lender = ['Beneficiary : ', 'Recorded']
    
    

# Determine the number of liens and derive the name of the lien holder   
po_liens = parser(lender[0], lender[1], prelim)
no_liens = prelim.count(lender[0])

if po_liens == 'None':
    liens = 'Property is Owned Free and Clear'

elif no_liens > 1:
    counter = no_liens
    liens = []
    lien_amounts = []
    splitter = re.split(lender[0], prelim)
    
    while counter > 1:
        
        for banks in range(no_liens):
            if banks % 2 == 0:
                liens.append(parser('', lender[1], splitter[banks + 1]))
                lien_amounts.append(parser('', lien_amount[1], splitter[banks + 1]))
                counter -= 1  
                
            elif banks % 2 != 0:
                liens.append(parser('', lender[1], splitter[banks + 1]))
                lien_amounts.append(parser('', lien_amount[1], splitter[banks + 1]))
                counter -= 1
                
else:  
    liens = po_liens
    lien_amount = parser(lien_amount[0], lien_amount[1], prelim)
    
# Combine list of liens for email msg if necessary    
if type(liens) == list:
    liens_list = []
    for i in range(len(liens)):
        liens_list.append(str(i+1) + '. '  + liens[i] + ': ' + lien_amounts[i])

    liens = '\n'.join(liens_list)    

      
# List of property types
property_list = ['Planned Urban Development',
                 'Planned Unit Development',
                 'PUD',
                 'Single Family Residence',
                 'Condominium',
                 'Multi-Family',
                 'Multiple Family Residence',
                 'Commercial Building']

# Check if property is listed in the prelim
property_type = ''
for prop in range(len(property_list)):
    if property_list[prop].lower() in prelim.lower():
        property_type += property_list[prop]
        break
    
space = '\n\n'

# Email msg to be sent
msg = ('Prelim Information:' + space
         + 'Title Company:' '\n' + title_company
         + space + 'Property Address:' + '\n' + prop_address
         + space + 'Vesting:' + '\n' + vesting
         + space + 'Liens:' + '\n' + str(no_liens) 
         + space + 'Lien Holder(s) and Amount(s):' + '\n' + liens + ': ' + lien_amount
         + space + 'Property Type:' + '\n' + property_type)


# Append email msg if title has solar
if 'solar' in prelim.lower():
    solar = space + 'THERE IS A SOLAR LIEN ON TITLE'
    msg += (solar)
    
print(msg + '\n')

# Ask user if they would look to send the msg via email
response = input('Would you like to send the prelim as an email? (Y/N)' + '\n')

if response.lower() in ['y', 'yes']:

    # Send prelim findings via email
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = '?'
    mail.Subject = 'Prelim: ' + prop_address
    mail.Body = msg
            
    # Attach prelim to the email and send
    attachment  = r'?'
    mail.Attachments.Add(attachment)
    print('\n' + 'Correspondence sent to ' + mail.To)
    mail.Send()
    
else:
    print('\n' + 'ok.')
