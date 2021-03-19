#!/usr/bin/python

# This program converts outlook contacts exported in CSV format to a giant 
# vcard file that can be imported into wherever.

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
WEB_PAGE                    = 100


function CSVrow2vobject():

   # Here's the columns I need to convert:
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
   BUSINESS_FAX                = 30
   BUSINESS_PHONE              = 31
   BUSINESS_PHONE2             = 32
   COMPANY_MAIN_PHONE          = 35
   HOME_FAX                    = 36
   HOME_PHONE                  = 37
   HOME_PHONE2                 = 38
   MOBILE_PHONE                = 40
   OTHER_FAX                   = 41
   OTHER_PHONE                 = 42
   PAGER                       = 43
   ASSISTANT_NAME              = 50
   CATEGORIES                  = 54
   EMAIL_ADDRESS               = 57
   EMAIL_TYPE                  = 58
   EMAIL_DISPLAY_NAME          = 59
   EMAIL2_ADDRESS              = 60
   EMAIL2_TYPE                 = 61
   EMAIL2_DISPLAY_NAME         = 62
   INITIALS                    = 70
   NOTES                       = 77
   OFFICE_LOCATION             = 78
   WEB_PAGE                    = 100

