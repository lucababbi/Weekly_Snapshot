import os
from datetime import date
from dateutil.relativedelta import relativedelta
from Functions.Index_Series import Index_List, Get_TimeSeries
from Functions.FX_Series import FX_List, Get_TimeSeriesFX
from Functions.Dynamic_Date import Last_Friday
from Functions.Annual_Returns import Annual_Returns
from Functions.Annual_Returns_FX import Annual_Returns_FX
from Functions.Update_Excel import Update_Excel
from Functions.Sharepoint_Upload import SharePointUpload
from Functions.PDF_Exporter import ExportWeeklySnapshot, SharePointUploadPDF
from Functions.Outlook_Sender import OutlookEmail
from constants import SECTOR_INDICES
loc = os.getcwd()+ "/Dashboard"

# Sharepoint locations
Archive_path = "/Users/yukesun/Library/CloudStorage/OneDrive-ISS/02_Tasks/10_Weekly_Benchmark_Update/Archive"
PDF_Path = "/Users/yukesun/Library/CloudStorage/OneDrive-ISS/02_Tasks/10_Weekly_Benchmark_Update/Snapshots"
if_update = True

# Users 
#!! please change your refinitiv credentials in lseg-data.config.json !!
USER = "Yuke_Sun"

# Reference Dates
today = date.today()
EDate = Last_Friday(today)
EDate = date(2026, 4, 10) # For Testing Purposes
SDate = EDate - relativedelta(months=36)

# Parameters for Time_Series
Parameters = dict(
    SDate=SDate.strftime("%Y-%m-%d"),
    EDate=EDate.strftime("%Y-%m-%d"),
    Frq="D"
)



# Update Excel File with the new data
if __name__ == "__main__":

    if if_update == True:

        # Function to retrieve the Time_Series of all the Indices
        Time_Series = Get_TimeSeries(
            universe=Index_List["RIC"].to_list(),
            fields=["TR.ClosePrice.date", "TR.ClosePrice"],
            parameters=Parameters
            )

        # Calculate Annual Returns of the Indices
        Returns_Data = Annual_Returns(
            Time_Series=Time_Series,
            Date_Column="Date",
            EDate=EDate,
            Index_Frame=Index_List
            )

        # Get FX Rates Time_Series
        FX_Series = Get_TimeSeriesFX(
            universe=FX_List["RIC"].to_list(),
            fields=["TR.MIDPRICE.date", "TR.MIDPRICE"],
            parameters=Parameters
            )

        # Calculate Annual Returns of the FX Rates
        FX_Returns = Annual_Returns_FX(
            Time_Series=FX_Series,
            Date_Column="Date",
            EDate=EDate,
            Index_Frame=FX_List
            )

        Update_Excel(loc, Time_Series, FX_Series, Returns_Data, FX_Returns, Index_List, EDate, SECTOR_INDICES)
        print("Dashboard Updated Successfully!")
        # SharePoint Upload
        SharePointUpload(Excel_Path = loc + "/Excel_Dashboard.xlsx", 
                        Sharepoint_Folder = Archive_path)

        # PDF Export
        PDF_Export = ExportWeeklySnapshot(loc + "/Excel_Dashboard.xlsx", loc + "/Snapshots", USER)
        SharePointUploadPDF(PDF_Path = loc + "/Snapshots/" + date.today().strftime("%Y%m%d") + "_Weekly_Benchmarks_Snapshot_" + USER + ".pdf", 
                            Sharepoint_Folder = PDF_Path,
                            username=USER)

    # Outlook Email Sender
    OutlookEmail(
        pdf_path=loc + "/Snapshots/" + date.today().strftime("%Y%m%d") + "_Weekly_Benchmarks_Snapshot_" + USER + ".pdf",
        to_emails=[
            "stoxx-Index-Business@iss-stoxx.com"
        ],
        cc_emails=[
            "stoxxstrategy@iss-stoxx.com",
            "stoxx-DAXStrategy@iss-stoxx.com"
        ],
        subject=f"Weekly Benchmarks Snapshot as of {date.today().strftime('%d/%m/%Y')}",
        body_text="STOXX & DAX Benchmarks Teams",
        dpi=300,
        max_pages=1,
        send_automatically=False
    )
