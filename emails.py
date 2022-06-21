# Module for email class

class Email: 
    def __init__(self, date="", issueSummary="", product="", name="", customerEmail="", comment="", ipAddress="", browser="", cookies="", followup=False):
        self.date = date
        self.issueSummary = issueSummary
        self.product = product
        self.name = name
        self.customerEmail = customerEmail
        self.comment = comment
        self.ipAddress = ipAddress
        self.browser = browser
        self.cookies = cookies
        self.followup = followup  
    
# Loops through all Outlook folders to find given folder_name
def get_folder_by_name(folder_name, root_folder):
    found_folder = None
    
    for folder in root_folder.Folders:
        if folder.Name == folder_name:
            found_folder = folder

    return found_folder

# Function to create new Email given info
def emailCreator(info):
    date = info['date']
    name = info['First Name'] + ' ' + info['Last Name']
    name = name.title()
    issueSummary = info['Issue Summary']
    customerEmail = info['E-mail']
    comment = info['Comment Value']
    ipAddress = info['Extracted IP Address']
    browser = info['Extracted Browser/OS']
    cookies = info['Cookies']
    product = info['Product']

    newEmail = Email(date, issueSummary, product, name, customerEmail, comment, ipAddress, browser, cookies, False)
    return newEmail


# Function to replace annoying punctuation characters
def replaceCharacters(text):
    result = (text.
    		replace('\\xe2\\x80\\x99', "\'").
            replace('\\xc3\\xa9', 'e').
            replace('\\xe2\\x80\\x90', '-').
            replace('\\xe2\\x80\\x91', '-').
            replace('\\xe2\\x80\\x92', '-').
            replace('\\xe2\\x80\\x93', '-').
            replace('\\xe2\\x80\\x94', '-').
            replace('\\xe2\\x80\\x94', '-').
            replace('\\xe2\\x80\\x98', "\'").
            replace('\\xe2\\x80\\x9b', "\'").
            replace('\\xe2\\x80\\x9c', '\"').
            replace('\\xe2\\x80\\x9c', '\"').
            replace('\\xe2\\x80\\x9d', '\"').
            replace('\\xe2\\x80\\x9e', '\"').
            replace('\\xe2\\x80\\x9f', '\"').
            replace('\\xe2\\x80\\xa6', '...').
            replace('\\xe2\\x80\\xb2', "\'").
            replace('\\xe2\\x80\\xb3', "\'").
            replace('\\xe2\\x80\\xb4', "\'").
            replace('\\xe2\\x80\\xb5', "\'").
            replace('\\xe2\\x80\\xb6', "\'").
            replace('\\xe2\\x80\\xb7', "\'").
            replace('\\xe2\\x81\\xba', "+").
            replace('\\xe2\\x81\\xbb', "-").
            replace('\\xe2\\x81\\xbc', "=").
            replace('\\xe2\\x81\\xbd', "(").
            replace('\\xe2\\x81\\xbe', ")").
            replace('\\xc2\\xa0', " ").
            replace('\\xa0', ' ')
            )
    return result
