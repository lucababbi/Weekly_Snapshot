import polars as pl
from datetime import date
from dateutil.relativedelta import relativedelta
from Index_Series import Index_List, Get_TimeSeries
from FX_Series import FX_List, Get_TimeSeriesFX
from Dynamic_Date import Last_Friday
from Annual_Returns import Annual_Returns
from Annual_Returns_FX import Annual_Returns_FX
from Update_Excel import openwb, writewb, savewb, Update_Excel

# Change to your Path
loc = "/Users/luccababbi/Documents/GitHub/Weekly_Snapshot/Dashboard"

# Reference Dates
today = date.today()
EDate = Last_Friday(today)
EDate = date(2026, 4, 3) # For Testing Purposes
SDate = EDate - relativedelta(months=36)

# Parameters for Time_Series
Parameters = dict(
    SDate=SDate.strftime("%Y-%m-%d"),
    EDate=EDate.strftime("%Y-%m-%d"),
    Frq="D"
)

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

# Array for Sectors
Sector =[
    ".SXAP",
    ".SX7P",
    ".SXPP",
    ".SX4P",
    ".SXOP",
    ".SXFP",
    ".SXDP",
    ".SXNP",
    ".SXIP",
    ".SXMP",
    ".SXEP",
    ".SX8P",
    ".SXKP",
    ".SX6P",
    ".SXRP",
    ".SXTP",
    ".SX86P",
    ".S600CPP",
    ".S600ENP",
    ".S600FOP",
    ".S600PDP",
]

if __name__ == "__main__":
    Update_Excel(loc, Time_Series, FX_Series, Returns_Data, FX_Returns, Index_List, EDate, Sector)

    print("Update complete!")