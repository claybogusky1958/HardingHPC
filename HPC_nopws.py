import json
from os import system
import xlrd
import time
import requests
import base64
import _datetime
import smtplib

#
# Make API calls to RapidAPI to find properties for sale in Harding.
#
#url = "https://realtor.p.rapidapi.com/locations/auto-complete" # used to search for specific properties
url = "https://realtor.p.rapidapi.com/properties/list-for-sale"
headers = {
    'x-rapidapi-host': "realtor.p.rapidapi.com",
    'x-rapidapi-key': "229e80e323msh0697f51cb0c7a00p1d0eacjsn18e02be76e4a"
    }
    
print( "** Starting ...'" ) 

fo = open( '/Users/Clay/Documents/Python/HistoricProperties/HardingListings.json', 'w+' )


print( "** Retrieving remote Real Estate data  ...'" ) 
# retrieve 1 chunk of records.
querystring = {"sort":"relevance","radius":"1","city":"Harding","offset":"0","limit":"200","state_code":"NJ"}
response = requests.request("GET", url, headers=headers, params=querystring)
if response.status_code  != 200:
    print( "Abnormal Status Code from requests.request(GET): ", response.status_code )
    fo.close()
    system.exit( "Processing Terminated" )

hdata = json.dumps( response.text )

print( "** Writing MLS data to file [HardingListings.json] ...'" )
json.dump( response.text, fo )

#debugging stuff
#pretty_json = json.loads( response.text )
#print (json.dumps(pretty_json, indent=2))

fo.close()

print( "** Opening JSON data file [HardingListings.json] ... ")

#
# End of section taken from HardingTwp0Radius.py
#

wp_url = "https://clayurl.com/wp-json/wp/v2"
wp_user = "XXX"   #admin 
wp_passwd = "YYY"

wp_credentials = wp_user + ':' + wp_passwd
wp_token = base64.b64encode( wp_credentials.encode())
wp_header = { 'Authorization': 'Basic ' + wp_token.decode( 'utf-8') }

# load Harding for sale listing data from previous process
json_file = open( '/Users/Clay/Documents/Python/HistoricProperties/HardingListings.json', 'r')
json_object = json.load(json_file)
json_string = json.loads( json_object )

# print( type( json_string[ 'listings' ] ) ) "LIST"
print( "** Total rows found in JSON file ... ", json_string["returned_rows"], " Matching rows: ",  json_string[ "matching_rows"] )

# Print all properties in JSON file 
#for property in json_string[ 'listings' ]:
#    print( property[ 'address' ] + " URL-> " +  property[ 'rdc_web_url' ] ) 

# Give the location of the XL file with Harding Historical Houses
print( "** Opening Excel file with Historic properties ..." )
loc = ("/Users/Clay/Documents/Python/HistoricProperties/HistoricHardingHouses.xlsx") 
  
# Open Historic Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0) 

print ("" )
print ("====== Summary of Historic Property for Sale in Harding Township ======" )
print( "                     ", time.strftime("%c") )
print ("" )
# Extracting number of rows in Spreadsheet
# print(sheet.nrows)

# Mostly Artifacts now..
# Example HTML so that a link can be sent about the property to the Post
# wp_content = '</span><strong><span style="text-decoration: underline;"> ' \
# '<a href="https://www.realtor.com/realestateandhomes-detail/607-Spring-Valley-Rd_Morristown_NJ_07960_M63809-33099">' \
# '607 Spring Valley Rd</a></span></strong></span></em></p>'

wp_part1 = '</span><strong><span style="text-decoration: underline;"> '
wp_part2 = '<a href='
wp_part3 = '</a></span></strong></span></em></p>'

wp_content = "<h4><strong>Please click on the address below for more information.</strong></h4>"
num_houses = 0

print( "** Search for historic properties in MLS listing data ..." )
# Search for historic house in listings data
for i in range(sheet.nrows):
    house_addr = sheet.cell_value(i, 2)
    # print( house_addr )
    for property in json_string[ 'listings' ]:
        # print( house_addr, property[ 'address' ] + " URL-> " +  property[ 'rdc_web_url' ] )
        s = property[ 'address' ]
        if ( s.find( house_addr ) != -1 ):
            print( "Found: " + house_addr + " URL: " + property[ 'rdc_web_url' ] )
            wp_tmp_content = wp_part1 + wp_part2 + '"' + property[ 'rdc_web_url' ] + '">' + house_addr + wp_part3
            wp_content = wp_content + wp_tmp_content
            num_houses += 1
            
print ("" )
print ( "*** Total number of historic properties for sale =", num_houses, "."  )

if ( num_houses == 0 ): # nothing more to do!!
    print( "No more processing required..." )
    exit()

print( "** Sending Data to WordPress POST ..." )
# deal with Current date and time using isoformat:" )
now = _datetime.datetime.now()
wp_date = now.isoformat()
wp_title = 'Historic Property For Sale in Harding on ' + time.strftime("%c")
wp_post = {
    'date':    wp_date,                         # '2020-04-07T07:00:00', YYYY-MM-DD
    'title':   wp_title,
    'content': wp_content,
    'status':  'publish',
    'categories' : [19], # Historic Category
    'slug' :   'historic'
}

wp_status = requests.post( wp_url + '/posts', headers=wp_header, json=wp_post )
print( "** WordPress POST returned status: ", wp_status )

# Readable display format... if needed
#print( json.dumps( json_object, indent=2) )

# Now send email to the members of the HPC using my hotmail address ...

print( "** Sending email to HPC ... " )

hotmail_user = 'XXXX'
hotmail_password = 'PASSWD'
# Email addresses
HPC_email_list =  [  'ltaglairino@hardingnj.org', 'cbogusky@hotmail.com', \
                    'cyates@hardingnj.org' , 'donato@brandesmaselli.com', \
                    'dev.modi@lglmlaw.com', 'dinsmored@hotmail.com', \
                    'gajc@me.com', 'karenalfieri@yahoo.com', 'mjwils65@gmail.com', \
                    'mcooney1973@gmail.com', 'skao1a@yahoo.com', 'tdepoortere@bccbelle.com' ]                  
to = ['claytest1927@gmail.com', 'claybogusky@gmail.com']
to = HPC_email_list

sent_from = hotmail_user
subject = 'List of Historic Properties for sale in Harding ' + time.strftime("%c")
body = "Hello Everyone, \n\nThis is the weekly list of Historic Properties for sale in Harding.\n " \
"Please see this page for more information https://clayurl.com/category/historic/" \
"\n\nThanks! \nClay"

email_text = """\
From: %s
To: %s
Subject: %s

%s
""" % (sent_from, ", ".join(to), subject, body)

try:
# This MS SMTP server & port work well 19 April 2020
    server = smtplib.SMTP( 'smtp-mail.outlook.com', 587 )
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(hotmail_user, hotmail_password )
    server.sendmail(sent_from, to, email_text)
    server.close()

    print ( '*** Email sent successfully ...' )
except:
    print ( 'Error sending Email to HPC ...something went wrong...' )

print( "** Processing finshed Normally ..." )

