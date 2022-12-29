import sys
import requests
from bs4 import BeautifulSoup
from shareplum.site import Version
from shareplum import Site, Office365
import pandas as pd

SHAREPOINT_URL = 'https://alphaonline.sharepoint.com'
SHAREPOINT_SITE = 'https://alphaonline.sharepoint.com/sites/Carret'
SHAREPOINT_LIST = 'Staff'
USERNAME = 'msoftdeveloper@alphacalibration.com'
PASSWORD = 'Ram87006'

def authenticate(sp_url, sp_site, user_name, password):
    """
    Takes a SharePoint url, site url, username and password to access the SharePoint site.
    Returns a SharePoint Site instance if passing the authentication, returns None otherwise.
    """
    site = None
    try:
        authcookie = Office365(SHAREPOINT_URL, username=USERNAME, password=PASSWORD).GetCookies()
        site = Site(SHAREPOINT_SITE, version=Version.v365, authcookie=authcookie)
    except:
        # We should log the specific type of error occurred.
        print('Failed to connect to SP site: {}'.format(sys.exc_info()[1]))
    return site

def get_sp_list(sp_site, sp_list_name):
    """
    Takes a SharePoint Site instance and invoke the "List" method of the instance.
    Returns a SharePoint List instance.
    """
    sp_list = None
    try:
        sp_list = sp_site.List(sp_list_name)
    except:
        # We should log the specific type of error occurred.
        print('Failed to connect to SP list: {}'.format(sys.exc_info()[1]))
    return sp_list

def download_list_items(sp_list, view_name=None, fields=None, query=None, row_limit=0):
    """
    Takes a SharePoint List instance, view_name, fields, query, and row_limit.
    The rowlimit defaulted to 0 (unlimited)
    Returns a list of dictionaries if the call succeeds; return a None object otherwise.
    """
    sp_list_items = None
    try:
        sp_list_items = sp_list.GetListItems(view_name=view_name, fields=fields, query=query, row_limit=row_limit)
    except:
        # We should log the specific type of error occurred.
        print('Failed to download list items {}'.format(sys.exc_info()[1]))
        raise SystemExit('Failed to download list items {}'.format(sys.exc_info()[1]))
    return sp_list_items

# Authenciate Sharepoint Connection
sp_site = authenticate(SHAREPOINT_URL, SHAREPOINT_SITE, USERNAME, PASSWORD)
sp_lists = ctx.web.lists

#Get Staff List Details
sp_list = get_sp_list(sp_site, "staff")
list_Staff = download_list_items(sp_list)
df_staff = pd.DataFrame(list_Staff)
df_staff.to_csv("staff.csv")

#Get Client Grid v2.0 List Details
sp_list = get_sp_list(sp_site, "Client Grid v2.0")
list_Client = download_list_items(sp_list)
df_Client = pd.DataFrame(list_Client)
df_Client.to_csv("Client_Grid_v2.0.csv")

#Get AUM & Retro Statement Tracker List Details
sp_list = get_sp_list(sp_site, "AUM & Retro Statement Tracker")
list_AUM = download_list_items(sp_list)
df_AUM = pd.DataFrame(list_AUM)
df_AUM.to_csv("AUM_Retro_Statement_Tracker.csv")

#Get Securities Master DB List Details
sp_list = get_sp_list(sp_site, "Securities Master DB")
list_Security = download_list_items(sp_list)
df_Security = pd.DataFrame(list_Security)
df_Security.to_csv("Securities_Master_DB.csv")

#Get Portfolio Report List Details
sp_site = authenticate(SHAREPOINT_URL, SHAREPOINT_SITE, USERNAME, PASSWORD)
sp_list = get_sp_list(sp_site, "portfolio report")
list_portfolio = download_list_items(sp_list)
df_Portfolio = pd.DataFrame(list_portfolio)
df_Portfolio.to_csv("Portfolio_Report.csv")
