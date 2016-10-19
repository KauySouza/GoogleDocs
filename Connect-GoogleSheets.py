# VMware PSO - v0.1
#
# This script syncs data between .csv and Google Sheets. That makes the Residency Dashboard on SharePoint more accurate and easy to maintain. The Google API version when this script was develop is 4.0.
#
#
# You will need:
# 	Internet access to connect on Google Sheet
#	A Google Account to authenticate.
#	Python 2.6 or later - Comes with your MAC.
#	A python package manager called "pip" - Usually comes with python 2.7 or later. Keep in mind that some bugs were found with El Captain Release. Try updating your python (brew install python).
#	Download the packages for API connection in your python repository - sudo pip install --upgrade google-api-python-client
#	Enable API connection on your google account - Follow only the Step 1 of this link https://developers.google.com/sheets/quickstart/python.
#
#


#
# Importing packages and methods.
#

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

# Autheticating and storing credencials for further use.

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

