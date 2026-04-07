import polars as pl
from datetime import timedelta, date
from dateutil.relativedelta import relativedelta

def Annual_Returns(
    Time_Series,
    Date_Column="Date",
    EDate=None,
    Index_Frame=None
):
    # Reference Date for Annual Returns
    WoW_Change = EDate - relativedelta(days=7)
    One_Month = EDate - relativedelta(months=1)
    YTD = Time_Series.filter(pl.col("Date").dt.year() == pl.lit(EDate.year)).sort("Date").head(1)["Date"][0]
    OneYear = EDate - relativedelta(years=1)
    ThreeYear = EDate - relativedelta(years=3)

    # Latest Price
    Last_Price = Time_Series.tail(1).select(pl.exclude(Date_Column))

    # Periods
    Periods = [
        ("WoW", WoW_Change),
        ("1 Month", One_Month),
        ("YTD", YTD),
        ("1 Year", OneYear),
        ("3 Year", ThreeYear)
    ]

    PriceChanges = []

    for label, RefDate in Periods:
        ref_row = Time_Series.filter(pl.col(Date_Column) == RefDate).select(pl.exclude(Date_Column))

        if ref_row.height > 0:
            PriceChange = ((Last_Price / ref_row) - 1)
            PriceChange = PriceChange.with_columns(pl.lit(label).alias("Period"))
        else:
            # Missing information
            PriceChange = pl.DataFrame({"Period": [label]})

        PriceChanges.append(PriceChange)

    # Concatenate all PriceChanges
    Annual_Returns = pl.concat(PriceChanges)

    # Reorder Period first, then Instruments
    Melted_Frame = Annual_Returns.melt(
                                        id_vars=["Period"],
                                        value_vars=[col for col in Annual_Returns.columns if col != "Period"],
                                        variable_name="Instrument",
                                        value_name="Return"
                                    )
    
    # Replace RIC with Full Name
    Melted_Frame = Melted_Frame.join(
        Index_Frame.select(pl.col("RIC"), pl.col("Full_Name")),
        left_on="Instrument",
        right_on="RIC",
        how="left"
    )
    
    return Melted_Frame.pivot(index="Instrument", columns="Period", values="Return")