import cx_Oracle
import pandas as pd
import datetime as dt
import configparser
import win32com.client
import os
import pythoncom
import shutil
from datetime import datetime, timedelta
import time


#Connect to outlook and transfer below files into it's respective path in shared network directory allowing for dynamic creation of file path.
pythoncom.CoInitialize()
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

DEMYST = outlook.GetDefaultFolder(6).Folders["REPORTING"].Folders["MYST"]

ADOBE = outlook.GetDefaultFolder(6).Folders["REPORTING"].Folders["ADOBE"]

broker_files_path = r"\\msp\Tribe\Biz\reporting\Reporting\Data\Data\2024"

daily_file_path = r"\\asp\Tribe\Biz\reporting\Data Reporting\data"

adobe_file_path = r"\\dsp\Tribe\Biz\reporting\Reporting\Data"

# Map month numbers to month names
month_names = {

    "04": "April",
    "05": "May",
    "06": "June",
    "07": "July",
    "08": "August",
    "09": "September",
    "10": "October",
    "11": "November",
    "12": "December"
}

# Transfer the daily demyst sessions report, broker referral file and adobe file to its respective folders

def transfer_file(inbox, file_start, file_path):
    file_repository = []
    for (current_folder, folders_within_current, files) in os.walk(file_path):
        for file in files:
            file_repository.append(file)

    for email in inbox.items:
        for attachment in email.Attachments:
            if attachment.FileName.startswith(file_start):
                filename = attachment.FileName
                if filename not in file_repository and not email.Subject.startswith("SIT"):
                    month = filename[-9:-7]
                    if month in month_names:
                        month_name = month_names[month]
                        destination_folder = os.path.join(file_path, f"{month} {month_name} 2024")
                        attachment_path = os.path.join(destination_folder, filename)
                        attachment.SaveAsFile(attachment_path)

for emails in ADOBE.items:
    for attachment in emails.Attachments:
        if attachment.FileName.startswith("Daily Reporting Metrics (NEW)"):
            email_date_time = emails.ReceivedTime
            email_month = str(email_date_time - timedelta(days=1))[5:7]
            email_date = str(email_date_time - timedelta(days=1))[8:10]
            new_file_name = f"Daily Reporting Metrics {email_date}.csv"
            month_name = month_names[email_month]
            destination_folder = os.path.join(adobe_file_path, f"{email_month} {month_name} 2024")
            # Check if the file already exists in the destination folder
            file_path = os.path.join(destination_folder, new_file_name)
            if not os.path.isfile(file_path):
                # Save the attachment with the new file name
                attachment.SaveAsFile(file_path)

transfer_file(DEMYST, "Referral", broker_files_path)

transfer_file(DEMYST, "Daily Summary", daily_file_path)


# Transfer LH file to respective Daily reports folder and rename the file

lh_report_path = r"\\Tribe\Biz\Reporting\Reporting"
save_report_path = r"\\Tribe\Biz\Reporting\Reporting\amlreports"


all_files = os.listdir(lh_report_path)
for file in all_files:
    if file.startswith("Report -"):
        month = file[11:13]
        month_name = month_names[month]
        year = file[-9:-5]
        destination_folder = os.path.join(save_lh_report_path, f"{month} {month_name} {year}")

        original_file_path = os.path.join(lh_report_path, file)
        destination_path = os.path.join(destination_folder, file)
        if not os.path.exists(destination_path):
            shutil.copy(original_file_path, destination_path)
        else:
            print("File already exists")
        day = file[9:11]
        original_date = datetime.strptime(f"{day}{month}{year}", "%d%m%Y")
        new_date = original_date + timedelta(days=1)
        new_day = new_date.strftime("%d")
        new_month = new_date.strftime("%m")
        new_year = new_date.strftime("%Y")
        new_file_name = f"Report - {new_day}{new_month}{new_year}.xlsx"
        new_file_path = os.path.join(lh_report_path, new_file_name)
        time.sleep(5)
        os.rename(original_file_path, new_file_path)


# Run downgrader tool

path=r'\\Tribe\Reporting\Reporting\data\D.xlsm'
def run_downgrader_tool(file_path, macro_name):
    excel_app=win32com.client.Dispatch("Excel.Application")
    try:
        workbook=excel_app.Workbooks.Open(file_path)
        excel_app.Application.Run(macro_name)
        workbook.Save()
    except:
        print("An error occured")

    finally:
        workbook.Close(SaveChanges=True)
        excel_app.Quit()

run_downgrader_tool(path, "SetFolder")



config = configparser.ConfigParser()
config.read("config.ini")

# create connection to the database parsing credentials from config file

username = config.get("PO-PROD", "username")
password = config.get("PO-PROD", "password")
host = config.get("PO-PROD", "host")
port = config.getint("PO-PROD", "port")
service_name = config.get("PO-PROD", "service_name")

dsn = cx_Oracle.makedsn(host, port, service_name=service_name)

connection = cx_Oracle.connect(user=username, password=password, dsn=dsn)

print()

print("Database connection successful")

print()

# Find WIP file from directory

# Directory where Lighthouse report is stored
directory_path = r'\\csm\Tribe\Reporting'

files_in_directory = os.listdir(directory_path)

for file in files_in_directory:
    if file.startswith("Report -"):
        file_path = os.path.join(directory_path, file)
df_WIP = pd.read_excel(file_path, sheet_name="WS - LH - WIP")
WIP_apps = df_WIP.iloc[:, 5]
WIP_apps = pd.DataFrame(WIP_apps)


# Delete previous day total_apps file

old_total_apps_path = r"\\Reporting\Reporting\analysis\Data"
old_all_total_apps = os.listdir(old_total_apps_path)
for old_total_apps in old_all_total_apps:
    if old_total_apps.startswith("total"):
        old_total_apps_filepath = os.path.join(old_total_apps_path, old_total_apps)
        os.remove(old_total_apps_filepath)

# Get new apps
cur = connection.cursor()

# Calculate yesterday date
yesterday = dt.datetime.now() - dt.timedelta(days=1)
yesterday_str = yesterday.strftime("%d-%b-%Y")  # Format date as "DD-MON-YYYY"

# Fetching & storing new apps
new_apps = f"""
select distinct po.ORDERNUMBER
from PROCESS.ORDER PO JOIN PROCESS.STATE OS ON PO.id=OS.orderid,
JSON_TABLE(PO.DATA, '$' COLUMNS (id NUMBER PATH '$.Application[*].externalSystemApplicationId',
submissioncode VARCHAR2 PATH '$.Application[*].SubmissionReasonCode')) x
WHERE trunc(po.CREATEDDATE)=to_date('{yesterday_str}','DD-MON-YYYY') AND po.APPLICATIONTYPE='OLA' and x.olasubmissioncode in ('RS1','RS2','RS3','RS4') and  po.VERSION=1
"""



cur.execute(new_apps)

new_rows = cur.fetchall()
new_df = pd.DataFrame(new_rows, columns=["ORDERNUMBER"])
new_df.rename(columns={"ORDERNUMBER": "BBD APP #"}, inplace=True)
total_apps = pd.concat([WIP_apps, new_df])
total_apps.to_excel(
    r"\\Reporting\Reporting\analysis\apps.xlsx",
    index=False)

# Close first cursor object
cur.close()

# Convert list of BBD Apps in total_apps file into a Python List object
total_apps_path = r'\\Reporting\analysis'

all_files = os.listdir(total_apps_path)

for total_apps_file in all_files:
    if total_apps_file.startswith("total"):
        total_apps_file_path = os.path.join(total_apps_path, total_apps_file)
df_total_apps = pd.read_excel(total_apps_file_path)
total_final_apps = df_total_apps["BBD APP #"].tolist()
# print(total_final_apps)

cur2 = connection.cursor()

# Calculate yesterday's date
yesterday_date = dt.datetime.now() - dt.timedelta(days=1)
yesterday_date_str = yesterday_date.strftime('%d-%b-%Y')

# # Your SQL query with placeholders
total_apps_status = """Select
id,
Ordernumber,
Createddate,
Createdby,
Entityname,
Applicationtype,
Orderid,
State,
Amount,submissioncode,
State_Createddate,
Trunc(State_Createddate) As State_Createddate_Fmt,
Status,
Lh_Comment,
Appwithdrawreason,
Decision,
Decisiondatetime,
Case When Rownumber=1 Then 1 Else Null End As Latest_Flag
From
(

Select  Po.Ordernumber,
        X.id,
        Po.Createddate,
        Po.Createdby,
        Po.Entityname,
        Po.Applicationtype,
        Os.Orderid,
        Os.State,
        X.Amount,
        X.submissioncode,
        Os.Createddate As State_Createddate,
        Case When Os.State In ('FINALISED') Then 'Completed'
             When Os.State In ('DRAWDOWN_COMPLETED') Then 'Awaiting Finalisation'
             When Os.State In ('PENDING_WITHDRAWN') Then 'Withdrawn'
             When Os.State In ('ASSESSMENT_DECLINED') Then 'Declined'
             When Os.State In ('AUTO_DECLINED') Then 'Declined'
             When Os.State In ('PENDING_DOC_RETURN') Then 'Awaiting Document Return'
             When Os.State In ('AWAITING_WELCOME_PACK') Then 'Awaiting Document Return'
             When Os.State In ('AWAITING_REWORK') Then 'Unsubmitted for Rework'
             When Os.State In ('PENDING_ASSESSMENT') Then 'Assigned to Assessor'
             When Os.State In ('AUTO_APPROVED') Then 'Awaiting Assessment'
             When Os.State In ('NEW') Then 'Unsubmitted'
             When Os.State In ('PENDING_CUSTOMER_LINK') Then 'Pre-processing'
             Else Null
             End As Status,
        Case When X.submissioncode In ('RS1') Then 'RS1'
             When X.submissioncode In ('RS2') Then 'RS2'
             When Os.State In ('FINALISED') Then 'Drawn'
             When Os.State In ('DRAWDOWN_COMPLETED') Then 'Drawn'
             When Os.State In ('PENDING_WITHDRAWN') And Regexp_Like(Substr(X.Appwithdrawreason, 3, 1), '^[0-9]$') Then 'Failed validation'
             When Os.State In ('PENDING_WITHDRAWN') And Not Regexp_Like(Substr(X.Appwithdrawreason, 3, 1), '^[0-9]$') Then 'Withdrawn'
             When Os.State In ('ASSESSMENT_DECLINED') And Cd.Decision='DC' Then 'Declined assessment - CTS'
             When Os.State In ('ASSESSMENT_DECLINED') And Cd.Decision='DB' Then 'Declined assessment - bureau'
             When Os.State In ('ASSESSMENT_DECLINED') And Cd.Decision In ('DO','DEC') Then 'Declined assessment - scorecard'
             When Os.State In ('ASSESSMENT_DECLINED') And Cd.Decision Is Not Null Then 'Declined assessment - other'
             When Os.State In ('AUTO_DECLINED') And Cd.Decision='DC' Then 'Declined assessment - CTS'
             When Os.State In ('AUTO_DECLINED') And Cd.Decision='DB' Then 'Declined assessment - bureau'
             When Os.State In ('AUTO_DECLINED') And Cd.Decision In ('DO','DEC') Then 'Declined assessment - scorecard'
             When Os.State In ('AUTO_DECLINED') And Cd.Decision Is Not Null Then 'Declined assessment - other'
             When Os.State In ('PENDING_DOC_RETURN') Then 'Docs out'
             When Os.State In ('AWAITING_WELCOME_PACK') Then 'Docs out'
             When Os.State In ('AWAITING_REWORK') Then 'Awaiting full assessment'
             When Os.State In ('PENDING_ASSESSMENT') Then 'Awaiting full assessment'
             When Os.State In ('AUTO_APPROVED') Then 'Awaiting full assessment'
             When Os.State In ('NEW') Then 'Awaiting validation'
             When Os.State In ('PENDING_CUSTOMER_LINK') Then 'Awaiting validation'
             When Os.State In ('PENDING_FULFILLMENT') Then 'Docs out'
             When Os.State In ('AIP_AWAITING_ASSET_DETAILS') Then 'Awaiting full assessment'
             Else Null
             End As Lh_Comment,
            X.Appwithdrawreason,
            Cd.Decision,
            Cd.Decisiondatetime,
            Row_Number() Over (Partition By Po.Ordernumber Order By Os.Createddate Desc) As Rownumber


From
    Process.Po_Order Po
Join
    Json_Table(Po.Data, '$' Columns (Olaid Number Path '$.Application[*].externalSystemApplicationId',
                                     Olasubmissioncode Varchar2 Path '$.Application[*].OLASubmissionReasonCode',
                                     Appwithdrawreason Varchar2 Path '$.Application[*].appWithdrawReason',
                                     Amount Number Path '$.Application[*].totalLendingIncrease')) X
On 1 = 1

Inner Join
    Process.Order_State Os
On
    Po.Id=Os.Order
Left Join
  Process.Decision Cd
On
    Os.Orderid=Cd.Orderid And Os.Date=Cd.Decisiondate
Where
     Os.State In ('FINALISED','DRAWDOWN_COMPLETED','PENDING_WITHDRAWN','ASSESSMENT_DECLINED','AUTO_DECLINED','PENDING_DOC_RETURN','AWAITING_WELCOME_PACK','PENDING_ASSESSMENT','AUTO_APPROVED','NEW','AWAITING_REWORK','PENDING_CUSTOMER_LINK','PENDING_FULFILLMENT','AIP_AWAITING_ASSET_DETAILS')
     And Trunc(Os.Createddate)<=To_Date('{yesterday_date_str}','DD-MON-YYYY') And
     Po.Ordernumber In ({})
) Rankeddata
Order By
Ordernumber, State_Createddate
""".format(','.join(str(ordernumber) for ordernumber in total_final_apps), yesterday_date_str=yesterday_date_str)
#


#
cur2.execute(total_apps_status)

column_names = [col[0] for col in cur2.description]
total_apps_status_rows = cur2.fetchall()
total_apps_status_rows_df = pd.DataFrame(total_apps_status_rows, columns=column_names)
total_apps_status_rows_df.to_excel(r"\\Reporting\analysis\Data\History.xlsx",index=False, header=True)