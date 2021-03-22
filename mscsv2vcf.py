#!/usr/bin/env python
# -*- coding: utf-8 -*-

# This program converts outlook contacts exported in CSV format to a giant 
# vcard 3.0 file that can be imported into wherever.


#### LIBRARIES ####
import sys
import argparse
import vobject
import csv

#### CONSTANTS ####

# These are the columns for Office365 CSV format.
# This is insane how many columns there are.

TITLE                       = 0
FIRST_NAME                  = 1
MIDDLE_NAME                 = 2
LAST_NAME                   = 3
SUFFIX                      = 4
COMPANY                     = 5
DEPARTMENT                  = 6
JOB_TITLE                   = 7
BUSINESS_STREET             = 8
BUSINESS_STREET2            = 9
BUSINESS_STREET3            = 10
BUSINESS_CITY               = 11
BUSINESS_STATE              = 12
BUSINESS_POSTAL_CODE        = 13
BUSINESS_COUNTRY            = 14
HOME_STREET                 = 15
HOME_STREET2                = 16
HOME_STREET3                = 17
HOME_CITY                   = 18
HOME_STATE                  = 19
HOME_POSTAL_CODE            = 20
HOME_COUNTRY                = 21
OTHER_STREET                = 22
OTHER_STREET2               = 23
OTHER_STREET3               = 24
OTHER_CITY                  = 25
OTHER_STATE                 = 26
OTHER_POSTAL_CODE           = 27
OTHER_COUNTRY               = 28
ASSISTANT_PHONE             = 29
BUSINESS_FAX                = 30
BUSINESS_PHONE              = 31
BUSINESS_PHONE2             = 32
CALLBACK                    = 33
CAR_PHONE                   = 34
COMPANY_MAIN_PHONE          = 35
HOME_FAX                    = 36
HOME_PHONE                  = 37
HOME_PHONE2                 = 38
ISDN                        = 39
MOBILE_PHONE                = 40
OTHER_FAX                   = 41
OTHER_PHONE                 = 42
PAGER                       = 43
PRIMARY_PHONE               = 44
RADIO_PHONE                 = 45
TTY_PHONE                   = 46
TELEX                       = 47
ACCOUNT                     = 48
ANNIVERSARY                 = 49
ASSISTANT_NAME              = 50
BILLING_INFORMATION         = 51
BIRTHDAY                    = 52
BUSINESS_ADDRESS_PO_BOX     = 53
CATEGORIES                  = 54
CHILDREN                    = 55
DIRECTORY_SERVER            = 56
EMAIL_ADDRESS               = 57
EMAIL_TYPE                  = 58
EMAIL_DISPLAY_NAME          = 59
EMAIL2_ADDRESS              = 60
EMAIL2_TYPE                 = 61
EMAIL2_DISPLAY_NAME         = 62
EMAIL3_ADDRESS              = 63
EMAIL3_TYPE                 = 64
EMAIL3_DISPLAY_NAME         = 65
GENDER                      = 66
GOVERNMENT_ID_NUMBER        = 67
HOBBY                       = 68
HOME_ADDRESS_PO_BOX         = 69
INITIALS                    = 70
INTERNET_FREE_BUSY          = 71
KEYWORDS                    = 72
LANGUAGE                    = 73
LOCATION                    = 74
MANAGERS_NAME               = 75
MILEAGE                     = 76
NOTES                       = 77
OFFICE_LOCATION             = 78
ORGANIZATIONAL_ID_NUMBER    = 79
OTHER_ADDRESS_PO_BOX        = 80
PRIORITY                    = 81
PRIVATE                     = 82
PROFESSION                  = 83
REFERRED_BY                 = 84
SENSITIVITY                 = 85
SPOUSE                      = 86
USER1                       = 87
USER2                       = 88
USER3                       = 89
USER4                       = 90
WEB_PAGE                    = 91


def CSVrow2vobject(row):
    '''This function converts a single CSV row to a vcard v3.0 object.'''

    # Create a vobject to return.
    vcard = vobject.vCard()

    # Contact name
    vcard.add('fn')
    if row[MIDDLE_NAME]:
        vcard.fn.value = ' '.join(
            (row[FIRST_NAME], row[MIDDLE_NAME], row[LAST_NAME]))
    else:
        vcard.fn.value = ' '.join((row[FIRST_NAME], row[LAST_NAME]))
    
    vcard.add('n')
    vcard.n.value = vobject.vcard.Name()
    vcard.n.value.family = row[LAST_NAME]
    vcard.n.value.given = row[FIRST_NAME]
    vcard.n.value.additional = row[MIDDLE_NAME]
    vcard.n.value.prefix = row[TITLE]
    vcard.n.value.suffix = row[SUFFIX]


    # Company name/job
    if row[COMPANY]:
        vcard.add('org')
        if row[DEPARTMENT]:
            vcard.org.value = [row[COMPANY], row[DEPARTMENT]]
        else:
            vcard.org.value = [row[COMPANY]]
    if row[JOB_TITLE]:
        vcard.add('title')
        vcard.title.value = row[JOB_TITLE]


    # Company address
    if row[BUSINESS_STREET]:
        workAddr = vcard.add('adr')
        workAddr.type_param = ["WORK", "POSTAL"]
        workAddr.value = vobject.vcard.Address()
        workAddr.value.street = row[BUSINESS_STREET]
        if row[BUSINESS_STREET2]:
            workAddr.value.street += '\n' + row[BUSINESS_STREET2]
        if row[BUSINESS_STREET3]:
            workAddr.value.street += '\n' + row[BUSINESS_STREET3]
        workAddr.value.city = row[BUSINESS_CITY]    
        workAddr.value.region = row[BUSINESS_STATE]
        workAddr.value.code = row[BUSINESS_POSTAL_CODE]
        workAddr.value.country = row[BUSINESS_COUNTRY]

    
    # Home address
    if row[HOME_STREET]:
        homeAddr = vcard.add('adr')
        homeAddr.type_param = ["HOME", "POSTAL"]
        homeAddr.value = vobject.vcard.Address()
        homeAddr.value.street = row[HOME_STREET]
        if row[HOME_STREET2]:
            homeAddr.value.street += '\n' + row[HOME_STREET2]
        if row[HOME_STREET3]:
            homeAddr.value.street += '\n' + row[HOME_STREET3]
        homeAddr.value.city = row[HOME_CITY]    
        homeAddr.value.region = row[HOME_STATE]
        homeAddr.value.code = row[HOME_POSTAL_CODE]
        homeAddr.value.country = row[HOME_COUNTRY]


    # Company phone
    if row[COMPANY_MAIN_PHONE]:
        workPhone = vcard.add('tel')
        workPhone.type_param = ["WORK", "VOICE", "PREF"]
        workPhone.value = row[COMPANY_MAIN_PHONE]
    if row[BUSINESS_PHONE]:
        workPhone = vcard.add('tel')
        if row[COMPANY_MAIN_PHONE]:
            workPhone.type_param = ["WORK", "VOICE"]
        else:
            workPhone.type_param = ["WORK", "VOICE", "PREF"]
        workPhone.value = row[BUSINESS_PHONE]
    if row[BUSINESS_PHONE2]:
        workPhone = vcard.add('tel')
        workPhone.type_param = ["WORK", "VOICE"]
        workPhone.value = row[BUSINESS_PHONE2]
    if row[BUSINESS_FAX]:
        workPhone = vcard.add('tel')
        workPhone.type_param = ["WORK", "FAX"]
        workPhone.value = row[BUSINESS_FAX]


    # Home phone
    if row[HOME_PHONE]:
        homePhone = vcard.add('tel')
        homePhone.type_param = ["HOME", "VOICE", "PREF"]
        homePhone.value = row[HOME_PHONE]
    if row[HOME_PHONE2]:
        homePhone = vcard.add('tel')
        homePhone.type_param = ["HOME", "VOICE"]
        homePhone.value = row[HOME_PHONE2]
    if row[HOME_FAX]:
        homePhone = vcard.add('tel')
        homePhone.type_param = ["HOME", "FAX"]
        homePhone.value = row[HOME_FAX]


    # Other phones
    if row[MOBILE_PHONE]:
        otherPhone = vcard.add('tel')
        otherPhone.type_param = ["CELL"]
        otherPhone.value = row[MOBILE_PHONE]
    if row[OTHER_PHONE]:
        otherPhone = vcard.add('tel')
        otherPhone.type_param = ["VOICE"]
        otherPhone.value = row[OTHER_PHONE]
    if row[OTHER_FAX]:
        otherPhone = vcard.add('tel')
        otherPhone.type_param = ["FAX"]
        otherPhone.value = row[OTHER_FAX]
    if row[PAGER]:
        otherPhone = vcard.add('tel')
        otherPhone.type_param = ["PAGER"]
        otherPhone.value = row[PAGER]
    

    # Email addresses
    if row[EMAIL_ADDRESS]:
        email = vcard.add('email')
        if row[EMAIL_TYPE] == "SMTP":
            email.type_param = ["internet", "PREF"]
        else:
            email.type_param = ["PREF"]
        email.value = row[EMAIL_ADDRESS]
    if row[EMAIL2_ADDRESS]:
        email2 = vcard.add('email')
        if row[EMAIL2_TYPE] == "SMTP":
            email2.type_param = ["internet"]
        email2.value = row[EMAIL2_ADDRESS]


    # Miscellaneous info
    if row[ASSISTANT_NAME]:
        assistant = vcard.add('x-assistant')
        assistant.value = row[ASSISTANT_NAME]
    if row[CATEGORIES]:
        cat = vcard.add('categories')
        cat.value = row[CATEGORIES].split(',')
    if row[NOTES]:
        notes = vcard.add('note')
        notes.value = row[NOTES]
    if row[WEB_PAGE]:
        url = vcard.add('url')
        url.value = row[WEB_PAGE]


    return vcard




#### MAIN ####

def main(argv=sys.argv[1:]):
    '''Import a csv file and dump out the corresponding VCARD to stdout.'''

    #
    # Set up command line arguments.
    #
    parser = argparse.ArgumentParser(description = '''
This program imports a CSV file of contacts from Outlook v16 (Office365
or office 2019) and converts it to a single iCal file which can be 
imported.''')

    parser.add_argument("infile", nargs='?', 
                        help="Specify the outlook CSV filename.",
                        type=argparse.FileType('r'), default=sys.stdin)
    parser.add_argument("outfile", nargs='?', 
                        help="Specify the iCAL filename.",
                        type=argparse.FileType('w'), default=sys.stdout)
    args = parser.parse_args(args=argv)


    # Get the CSV headers.
    CSVreader = csv.reader(args.infile)
    headers = next(CSVreader)

    # Run through each CSV line to create a list of vcard objects.
    linecount = 0
    vcards = []
    for line in CSVreader:
        vcard = CSVrow2vobject(line)
        vcards.append(vcard)

    # Dump the vcards to a file.
    for vcard in vcards:
        args.outfile.write(vcard.serialize())
    args.outfile.close()

        

    

    return 0 # ok status


if __name__=='__main__':
    sys.exit(main())
