#########################################################################################################################################################################################################################################
# Name: AGOL Active Directory Comparison
# Python Version: 3.6
# Author: Zand Bakhtiari w\Todd Grissom
# Date: Sometime between February and April of 2020
#
# Purpose: This script creates a list of AD “UserNames” then compares it against AD to determine if the AGOL users has
# an AD account.  It does so by running a Windows PowerShell Script.  All the verbiage for the PowerShell script is
# created within this script using User information pulled form AGOL. The script assumes that anything before the ‘@’ sign in the email is the AD Username.
#
# Only three inputs are required when running the script;
# 1. A file path to a working Directory,
# 2. Username (The AGOL user must have the credentials to view this information).
# 3. Password
# The final output is a .CSV document that lists the Users Full Name, AGOL Username, Email, Created Date, Last Login, Department, & Division.
# If a record has a blank under Department that means the user does not exist in AD.
#
# NOTE: This scrip was created for the purpose of comparing ArcGIS Online (AGOL) users to active directory users (AD).
# While the script here in may work in my case, I am by no means an AD expert and some portions of this script might
# need to be altered before used by others. Before scripting can begin you must first install the the AD PowerShell Module.
# Link - https://theitbros.com/install-and-import-powershell-active-directory-module/
#
# You will also need to edit line 27 of the code to add your AGOL Organizational URL.
#
##########################################################################################################################################################################################################################################
# Import Modules
from subprocess import check_output
from arcgis.gis import GIS
import csv, time

# This variable will prompt the user to paste in their preferred working directory.
Folder = (input("What directory would you like outputs to be placed? >>>"))

# Sets a variable that will access the Enterprise AGOL Organization and passes in the users Username & Password.
# UN and PW Variables are passed through the Function Arguments
OrgURL = #!!!!!AGOL ORG URL GOES HERE!!!!! <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
UN = (input("Please type in your Username (all lower case)>>>"))
PW = (input("Please type in your Password >>>"))
orgAccess = GIS(OrgURL,UN,PW)

# Creates a list of ALL user emails in the Enterprise AGOL Organization and sets the max results to 1000
orgUsers = orgAccess.users.search(query="*", max_users=1000)
print ("There are",len(orgUsers),"users")
# Creates list
emails = [] # For storing the users email
created = [] # For storing the date the user's account was created
lastlogin = [] # For storing the date the user last logged in
AGOLUN =[] # For storing the users AGOL username
FullName = []  # For Storing the user’s full name
print ("Creating necessary lists ....")

# Iterates through the OrgUsers list and appends data into the their respective lists
for user in orgUsers:
    emails.append(user.email)
    AGOLUN.append(user.username)
    FullName.append(user.fullName)
# Prints the count of the emails list to make sure it is the same count as the OrgUsers list.
print("Done creating Username List.","Final Count is",len(emails))

# Iterates through the orgUsers list and finds the Created date and converts it to a user-friendly format.
# Then appends it to the created list.
for l in orgUsers:
    CreatedOn = time.localtime(l.created / 1000)
    create = ("{}/{}/{}".format(CreatedOn[0], CreatedOn[1], CreatedOn[2]))
    created.append(create)

# Iterates through the orgUsers list and finds the last login date and converts it to a user-friendly format.
# Then appends it to the created list. If the user has never logged in, the list will be appended to state such.
for l in orgUsers:
    if l.lastLogin == -1:
        lastlogin.append("Has never logged in")

    else:
        last_accessed = time.localtime(l.lastLogin / 1000)
        mll = ("{}/{}/{}".format(last_accessed[0], last_accessed[1], last_accessed[2]))
        lastlogin.append(mll)


# Creates and empty list called ADUserName.  This is different than AGOL Username.
# Here we strip off everything before the @ sign in their emails.
# Weather or not the AD Username created here is valid does not matter as that will be validated later on in the script
ADUserName = []
print("Creating a list of AD Usernames.....")
# Iterates through the OrgUsers list, strips off the end of the email(everything after @), then appends the Username to the UserName list.
for user in orgUsers:
    split = '@'
    ADUserName.append(user.email.split(split, 1)[0])
# Prints the count of the UserName list to make sure it is the same count as the OrgUsers list.
print("Done creating Username List.","Final Count is",len(ADUserName))


# Variable setting the location of the PowerShell document based on user input.
PowerShellPath = Folder + '\\AD.ps1'

# Creating empty lists for Department and Division to be appended too.
Department = []
Division = []

# This for-loop iterates over the list of AGOL Usernames from the Usernames variable created above.
for user in ADUserName:

    # Opens the PowerShell file and creates the AD commands using the List of AD Usernames
    f = open(PowerShellPath, "w+")
    f.write('Import-Module ActiveDirectory'); f.write('\n')
    f.write('$User = ' + '\'' + user + '\'')
    f.write('\n')
    f.write('Get-ADUser -Identity $User -Properties Department, Office')
    f.close()

    # If the Username is in AD, then it will find Department & Division & append them to their respective lists.
    try:
        p = check_output(["powershell.exe", PowerShellPath])
        p = p.decode('utf-8')
        x = p.strip()
        colonposition = x.find(':')  # Find the position of the first colon
        carriagepossition = x.find('\r', colonposition) # Find the position of the first carriage return
        Depart = x[colonposition+2 : carriagepossition]
        Department.append(Depart)
        y = x.replace('Office            : ', '@')
        atpos = y.find('@')
        sppos = y.find('\r', atpos)
        Division.append(y[atpos+1 : sppos])
    # If the Username does not exist in AD it will except the IndexError and print, "User cannot be found in AD."
    except:
        IndexError()
        print (user, 'cannot be found in AD')

#  The merge function takes the 7 lists and zips them into a tuple.
def merge(list1, list2, list3, list4, list5, list6, list7):
    merged_list = tuple(zip(list1, list2, list3, list4, list5, list6, list7))
    return merged_list
merged = (merge(FullName, AGOLUN, emails, created, lastlogin, Department, Division))
print(len(merged))

# Creates a csv file in the user defined directory, creates field headers, and writes the data into the respective columns.
with open(Folder+'\\ADUsers.csv', 'w+', newline='') as csvfile:
    fieldnames = ['Full Name','AGOL Username','Email','Create Date', 'Last Login', 'Department', 'Division']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()

    for fn, un, em, cd, ll, dep, div in merged:
        writer.writerow({'Full Name':fn,'AGOL Username': un, 'Email': em,'Create Date':cd, 'Last Login': ll,'Department':dep, 'Division': div})
