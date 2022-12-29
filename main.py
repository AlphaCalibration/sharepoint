import sys
from shareplum.site import Version
from shareplum import Site, Office365
import pandas as pd
from sqlalchemy import create_engine
from sqlalchemy.engine import URL

import psycopg2

SHAREPOINT_URL = 'https://alphaonline.sharepoint.com'
SHAREPOINT_SITE = 'https://alphaonline.sharepoint.com/sites/Carret'
SHAREPOINT_LIST = 'Staff'
USERNAME = 'msoftdeveloper@alphacalibration.com'
PASSWORD = 'Ram87006'

POSTGRES_SERVER ="carretdb.postgres.database.azure.com"
POSTGRES_DB = "carretdb_admin"
POSTGRES_USERNAME = "carretadmin@carretdb"
POSTGRES_PASSWORD = "Carret@1234"


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
#pd.DataFrame(sp_site.GetListCollection()).to_csv("test.csv")

#Get Staff List Details
#sp_list = get_sp_list(sp_site, "staff")
#list_Staff = download_list_items(sp_list)
#df_staff = pd.DataFrame(list_Staff)
#df_staff.to_csv("staff.csv")

#Get Client Grid v2.0 List Details
sp_list = get_sp_list(sp_site, "Client Grid v2.0")
list_Client = download_list_items(sp_list)
df_Client = pd.DataFrame(list_Client)
df_Client.reset_index(drop=True, inplace=True)
df_Client.rename(columns = {'ID' : 'ID'
            ,'Account Number' : 'Account_Number'
            ,'Mandate Status' : 'Mandate_Status'
            ,'Mandate Termination Date' : 'Mandate_Termination_Date'
            ,'Approval Date' : 'Approval_Date'
            ,'RM1' : 'RM1'
            ,'Nationality/Place of Incorporation' : 'Nationality_Place_of_Incorporation'
            ,'Custodian Bank' : 'Custodian_Bank'
            ,'Corresponding Address Location of Client (according to latest address record in client folder)' : 'Corresponding_Address_Location_of_Client'
            ,'Geographical Location (Correspond to column K/ Based on SFC Form 12 definition)' : 'Geographical_Location'
            ,'Mandate Type (Carret-Client)' : 'Mandate_Type_Carret_Client'
            ,'Mandate Type (Carret -Bank)' : 'Mandate_Type_Carret_Bank'
            ,'Entity Type' : 'Entity_Type'
            ,'PI Type' : 'PI_Type'
            ,'Client Type(AM survey)' : 'Client_Type_AM_survey'
            ,'LPOA' : 'LPOA'
            ,'Private Bank: ID' : 'Private_Bank_ID'
            ,'Account Name' : 'Account_Name'
            ,'Staff Account/Staff Related Account?' : 'Staff_Account'
            ,'Has the client been met Face to Face ?' : 'Has_the_client_been_met_Face_to_Face'
            ,'Method of Sourcing' : 'Method_of_Sourcing'
            ,'Return Rebate to client' : 'Return_Rebate_to_client'
            ,'Whether the client has granted discretionary authourity to CPCL' : 'Whether_the_client_has_granted_discretionary_authourity_to_CPCL'
            ,'Client Risk Rating (CRR)' : 'Client_Risk_Rating_CRR'
            ,'ML/TF Risk' : 'ML_TF_Risk'
            ,'Professional Investor' : 'Professional_Investor'
            ,'PEP' : 'PEP'
            ,'Vulnerable Client' : 'Vulnerable_Client'
            ,'Quantity Review Status' : 'Quantity_Review_Status'
            ,'Quality Review Status' : 'Quality_Review_Status'
            ,'Necessary docs scanned to shared drive?' : 'Necessary_docs_scanned_to_shared_drive'
            ,'Created' : 'Created'
            ,'Fee Category' : 'Fee_Category'
            ,'Modified' : 'Modified'
            ,'Client NameLAST Name, First Name / Corp Name' : 'Client_Name'
            ,'LPOA with Lumen or Carret, and are Co-managed' : 'LPOA_with_Lumen_or_Carret_and_are_Co_managed'
            ,'Level' : 'Level'
            ,'owshiddenversion' : 'owshiddenversion'
            ,'Title' : 'Title'
            ,'URL Path' : 'URL_Path'
            ,'Approval Status' : 'Approval_Status'
            ,'Property Bag' : 'Property_Bag'
            ,'Effective Permissions Mask' : 'Effective_Permissions_Mask'
            ,'Unique Id' : 'Unique_Id'
            ,'ScopeId' : 'ScopeId'
            ,'Name' : 'Name'
            ,'Item Type' : 'Item_Type'
            ,'Address' : 'Address'
            ,'Email # 1' : 'Email_1'
            ,'Client Agreement Date' : 'Client_Agreement_Date'
            ,'Advisory Fees' : 'Advisory_Fees'
            ,'Performance DM & Adv Fees' : 'Performance_DM_Adv_Fees'
            ,'Format of the forms' : 'Format_of_the_forms'
            ,'Questions (consent or not/opportunistic model)' : 'Questions'
            ,'Format of the forms2' : 'Format_of_the_forms2'
            ,'KMM' : 'KMM'
            ,'ExpMM' : 'ExpMM'
            ,'KFI' : 'KFI'
            ,'ExpFI' : 'ExpFI'
            ,'KEQ' : 'KEQ'
            ,'ExpEQ' : 'ExpEQ'
            ,'KMF' : 'KMF'
            ,'ExpMF' : 'ExpMF'
            ,'KFX' : 'KFX'
            ,'ExpFX' : 'ExpFX'
            ,'KComm' : 'KComm'
            ,'ExpComm' : 'ExpComm'
            ,'KHF' : 'KHF'
            ,'ExpHF' : 'ExpHF'
            ,'KPE' : 'KPE'
            ,'ExpPE' : 'ExpPE'
            ,'KSP' : 'KSP'
            ,'ExpSP' : 'ExpSP'
            ,'KDer' : 'KDer'
            ,'ExpDer' : 'ExpDer'
            ,'KHK/ SIN' : 'KHK_SIN'
            ,'ExpHK/ SIN' : 'ExpHK_SIN'
            ,'KAP' : 'KAP'
            ,'ExpAP' : 'ExpAP'
            ,'KEU/ NA' : 'KEU_NA'
            ,'ExpEU/ NA' : 'ExpEU_NA'
            ,'KWorld' : 'KWorld'
            ,'ExpWorld' : 'ExpWorld'
            ,'Objective' : 'Objective'
            ,'Strategy' : 'Strategy'
            ,'Horizon (years)' : 'Horizon_years'
            ,'Vol' : 'Vol'
            ,'Cash' : 'Cash'
            ,'IG' : 'IG'
            ,'HY' : 'HY'
            ,'EQ/ REIT' : 'EQ_REIT'
            ,'Comm.' : 'Comm'
            ,'AI' : 'AI'
            ,'CCY' : 'CCY'
            ,'Max Lev' : 'Max_Lev'
            ,'Last Client Review Completion Date' : 'Last_Client_Review_Completion_Date'
            ,'Next Client Review Start Date (Generated)' : 'Next_Client_Review_Start_Date_Generated'
            ,'New PI / Last PI Confirmation Letter Sent Date (may use account opening date)' : 'New_PI_Last_PI_Confirmation_Letter_Sent_Date'
            ,'Next PI Confirmation Letter Date' : 'Next_PI_Confirmation_Letter_Date'
            ,'Single FI / EQ' : 'Single_FI_EQ'
            ,'Single Funds' : 'Single_Funds'
            ,'Single G10+' : 'Single_G10_plus'
            ,'Sector Other' : 'Sector_Other'
            ,'Single Others' : 'Single_Others'
            ,'Single China' : 'Single_China'
            ,'CCY G10+' : 'CCY_G10_plus'
            ,'CCY Others' : 'CCY_Others'
            ,'CCY China' : 'CCY_China'
            ,'Sector Gov' : 'Sector_Gov'
            ,'Email # 2' : 'Email_2'
            ,'Other Alternative Investments' : 'Other_Alternative_Investments'
            ,'RM2' : 'RM2'
            ,'RM3' : 'RM3'
            ,'Threshold for each trade (USD)' : 'Threshold_for_each_trade_USD'
            ,'DM Fees' : 'DM_Fees'
            ,'Fees Remarks' : 'Fees_Remarks'
            ,'VC Assessment Alert' : 'VC_Assessment_Alert'
            ,'Last VC Assessment Completion Date' : 'Last_VC_Assessment_Completion_Date'
            ,'Next VC Assessment Date' : 'Next_VC_Assessment_Date'
            ,'Anticipated Level of Trading Activity (per annum)(USD)' : 'Anticipated_Level_of_Trading_Activity_per_annum_USD'
            ,'Next Client Review Status' : 'Next_Client_Review_Status'
            ,'Tax Residency' : 'Tax_Residency'
            }, inplace = True)
#df_Client.to_csv("Client_Grid_v2.0.csv")

#Get AUM & Retro Statement Tracker List Details
sp_list = get_sp_list(sp_site, "AUM & Retro Statement Tracker")
list_AUM = download_list_items(sp_list)
df_AUM = pd.DataFrame(list_AUM)
#df_AUM.to_csv("AUM_Retro_Statement_Tracker.csv")
df_AUM.reset_index(drop=True, inplace=True)
df_AUM.rename(columns = {'Account' : 'Account'
                                ,'ID' : 'ID'
                                ,'ClientName' : 'ClientName'
                                ,'Month' : 'Month'
                                ,'Year' : 'Year'
                                ,'Status' : 'Status'
                                ,'RM' : 'RM'
                                ,'Mandate' : 'Mandate'
                                ,'Bank' : 'Bank'
                                ,'Total AUM' : 'Total_AUM'
                                ,'AUM for Management Fee' : 'AUM_for_Management_Fee'
                                ,'Net Asset' : 'Net_Asset'
                                ,'Fee Amount' : 'Fee_Amount'
                                ,'Fees Remarks' : 'Fees_Remarks'
                                ,'Advisory Fees' : 'Advisory_Fees'
                                ,'DM Fees' : 'DM_Fees'
                                ,'Performance DM & Adv Fees' : 'Performance_DM_Adv_Fees'
                                ,'Created' : 'Created'
                                ,'Created By' : 'Created_By'
                                ,'For Action' : 'For_Action'
                                ,'Client Grid: ID:Entity Type' : 'Client_Grid_ID_Entity_Type'
                                ,'Client Grid: ID:PI Type' : 'Client_Grid_ID_PI_Type'
                                ,'Return Rebate to client' : 'Return_Rebate_to_client'
                                ,'ID_val' : 'ID_val'
                                ,'Account ID' : 'Account_ID'
                                ,'Modified By' : 'Modified_By'
                                ,'Modified' : 'Modified'
                                ,'Retrocession Amount' : 'Retrocession_Amount'
                                ,'LPOA with Lumen or Carret, and are Co-managed' : 'LPOA_with_Lumen_or_Carret_Co_managed '
                                ,'Level' : 'Level'
                                ,'Property Bag' : 'Property_Bag'
                                ,'owshiddenversion' : 'owshiddenversion'
                                ,'Effective Permissions Mask' : 'Effective_Permissions_Mask'
                                ,'ScopeId' : 'ScopeId'
                                ,'URL Path' : 'URL_Path'
                                ,'Approval Status' : 'Approval_Status'
                                ,'Unique Id' : 'Unique_Id'
                                ,'Item Type' : 'Item_Type'
                                ,'Value Statement (USD)' : 'Value_Statement_USD'
                                ,'Cash (USD)' : 'Cash_USD'
                                ,'Loan (USD)' : 'Loan_USD'
                                ,'Received Retro Report from Banks' : 'Received_Retro_Report_from_Banks'
                                ,'Retro Amount Confirmed by RM' : 'Retro_Amount_Confirmed_by_RM'
                                ,'Instructed Bank to Pay' : 'Instructed_Bank_to_Pay'
                                ,'Bank paid Carret' : 'Bank_paid_Carret'
                                ,'Management/Advisory Fee Invoice Created' : 'Management_Advisory_Fee_Invoice_Created'
                                ,'Payment Instruction Sent to Bank' : 'Payment_Instruction_Sent_to_Bank'
                                ,'Management/Advisory Fee Amount Confirmed by RM' : 'Management_Advisory_Fee_Amount_Confirmed_by_RM'
                                ,'Payment Instruction Created' : 'Payment_Instruction_Created'
                                ,'Invoice Sent to Client' : 'Invoice_Sent_to_Client'
                                ,'Retro Amount Currency' : 'Retro_Amount_Currency'
                                ,'GenerateFeeAmount' : 'GenerateFeeAmount'
                                ,'Retrocession Amount (USD)' : 'Retrocession_Amount_USD_'
                                ,'Retro Remarks' : 'Retro_Remarks'
                                ,'Remarks' : 'Remarks'
                                ,'Remarks on Management Fee' : 'Remarks_on_Management_Fee'
                                ,'Total Direct Funds' : 'Total_Direct_Funds'
                                ,'Value Statement (Original Currency)' : 'Value_Statement_Original_Currency'
                                ,'Cash (Original Currency)' : 'Cash_Original_Currency'
                                ,'Loan (Original Currency)' : 'Loan_Original_Currency'
                                ,'EstimatedAUMforManagementFee' : 'EstimatedAUMforManagementFee'
                                ,'EstimatedFeeAmount' : 'EstimatedFeeAmount'
                                ,'FeeDifference' : 'FeeDifference'
                                ,'EstimatedTotalDirectFunds' : 'EstimatedTotalDirectFunds'
                                ,'CurrencyType' : 'CurrencyType'
                                ,'Rebate Notice Prepared' : 'Rebate_Notice_Prepared'
                                ,'AUM Local Currency' : 'AUM_Local_Currency'
                              }, inplace=True)

#Get Securities Master DB List Details
sp_list = get_sp_list(sp_site, "Securities Master DB")
list_Security = download_list_items(sp_list)
df_Security = pd.DataFrame(list_Security)
#df_Security.to_csv("Securities_Master_DB.csv")
df_Security.reset_index(drop=True, inplace=True)
df_Security.rename(columns = {'ID' : 'ID'
                    ,'Asset' : 'Asset'
                    ,'Asset Class' : 'Asset_Class'
                    ,'Description' : 'Description'
                    ,'Country' : 'Country'
                    ,'Currency' : 'Currency'
                    ,'Region' : 'Region'
                    ,'Sector' : 'Sector'
                    ,'Derivatives' : 'Derivatives'
                    ,'Country of Risk' : 'Country_of_Risk'
                    ,'S&P Bond Rating' : 'S_P_Bond_Rating'
                    ,'Product Risk Rating' : 'Product_Risk_Rating'
                    ,'Product Type' : 'Product_Type'
                    ,'Created' : 'Created'
                    ,'Cntry_of_risk' : 'Cntry_of_risk'
                    ,'ID NUMBER' : 'ID_NUMBER'
                    ,'Created By' : 'Created_By'
                    ,'Modified By' : 'Modified_By'
                    ,'Level' : 'Level'
                    ,'Property Bag' : 'Property_Bag'
                    ,'owshiddenversion' : 'owshiddenversion'
                    ,'Title' : 'Title'
                    ,'Effective Permissions Mask' : 'Effective_Permissions_Mask'
                    ,'ScopeId' : 'ScopeId'
                    ,'URL Path' : 'URL_Path'
                    ,'Approval Status' : 'Approval_Status'
                    ,'Unique Id' : 'Unique_Id'
                    ,'Item Type' : 'Item_Type'
                    ,'Modified' : 'Modified'
                    ,'Name' : 'Name'
                    ,'Issuer' : 'Issuer'
                    ,'six_valorNumber' : 'six_valorNumber'
                    ,'six_exchangeCode' : 'six_exchangeCode'
                    ,'six_currencyCode' : 'six_currencyCode'
                    ,'six_uniqueCombination' : 'six_uniqueCombination'
                    ,'six_IsColumnUpdated' : 'six_IsColumnUpdated'
                    ,'Six_DataStatus' : 'Six_DataStatus'
                    ,'Ticker Symbol' : 'Ticker_Symbol'
                    ,'High Yield Bond?' : 'High_Yield_Bond'
                    ,'CIS (SFC survey)' : 'CIS_SFC_survey'
                    ,'Complex Product' : 'Complex_Product'
                    ,'Issuer rating' : 'Issuer_rating'
                    ,'Structured Product types(product Survey)' : 'Structured_Product_types'
                    ,'For PI only?' : 'For_PI_only'
                    ,'Experienced liquidity/solvency problems during the reporting period?' : 'Experienced_liquidity_solvency_problems'
                    ,'Leveraged?' : 'Leveraged'
                    }, inplace = True)


#Get Portfolio Report List Details
#sp_site = authenticate(SHAREPOINT_URL, SHAREPOINT_SITE, USERNAME, PASSWORD)
#sp_list = get_sp_list(sp_site, "Portfolio Report1")
#list_portfolio = download_list_items(sp_list,row_limit=100)
#df_Portfolio = pd.DataFrame(list_portfolio)
#df_Portfolio.to_csv("Portfolio_Report.csv")

#engine = create_engine('postgresql://' + POSTGRES_USERNAME + '@' + POSTGRES_PASSWORD + '@' + POSTGRES_SERVER +'/' + POSTGRES_DB)
#                       carretadmin@Carret@1234@carretdb.postgres.database.azure.com:5432/carretdb_admin')
#df_Client.to_sql('ClientGrid_20', engine)

url_object = URL.create(
    "postgresql",
    username=POSTGRES_USERNAME,
    password=POSTGRES_PASSWORD,
    host=POSTGRES_SERVER,
    database=POSTGRES_DB,
)
engine = create_engine(url_object)
conn = engine.connect()
df_Client.to_sql('ClientGrid_20', con=conn,if_exists="replace")
df_AUM.to_sql('AUM_Retro_Statement_Tracker', con=conn,if_exists="replace")
df_Security.to_sql('SecurityMasterDB', con=conn,if_exists="replace")



