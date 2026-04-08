import polars as pl
import lseg.data as ld

# Index_List
Index_List = pl.DataFrame(
    {
        "Symbol": [
            "DAX",
            "MDAX",
            "SDAX",
            "DAXMS",
            "DAXAC",
            "SX5E",
            "SX5P",
            "SXXP",
            "SX50UP",
            "SXUSX1P",
            "SXP1E",
            "STXWAP",
            "V1X",
            "V2TX",
            "SXAP",
            "SX7P",
            "SXPP",
            "SX4P",
            "SXOP",
            "SXFP",
            "SXDP",
            "SXNP",
            "SXIP",
            "SXMP",
            "SXEP",
            "SX8P",
            "SXKP",
            "SX6P",
            "SXRP",
            "SXTP",
            "SX86P",
            "S600CPP",
            "S600ENP",
            "S600FOP",
            "S600PDP",
            "SX5K",
            "SX5L",
            "SXXL",
            "SX50UL",
            "SXUSX1L"
        ],
        "Full_Name": [
            "DAX",
            "MDAX",
            "SDAX",
            "DAX MidSmall Cap",
            "DAX All Cap",
            "EURO STOXX 50",
            "STOXX Europe 50",
            "STOXX Europe 600",
            "STOXX USA 500",
            "STOXX US Nexus 100",
            "STOXX Asia/Pacific 600",
            "STOXX World AC All Cap",
            "VDAX",
            "EURO STOXX 50 Volatility (VSTOXX)",
            "STOXX Europe 600 Automobiles & Parts",
            "STOXX Europe 600 Banks",
            "STOXX Europe 600 Basic Resources",
            "STOXX Europe 600 Chemicals",
            "STOXX Europe 600 Construction & Materials",
            "STOXX Europe 600 Financial Services",
            "STOXX Europe 600 Health Care",
            "STOXX Europe 600 Industrial Goods & Services",
            "STOXX Europe 600 Insurance",
            "STOXX Europe 600 Media",
            "STOXX Europe 600 Oil & Gas",
            "STOXX Europe 600 Technology",
            "STOXX Europe 600 Telecommunications",
            "STOXX Europe 600 Utilities",
            "STOXX Europe 600 Retail",
            "STOXX Europe 600 Travel & Leisure",
            "STOXX Europe 600 Real Estate",
            "STOXX Europe 600 Consumer Products and Services",
            "STOXX Europe 600 Energy",
            "STOXX Europe 600 Food Beverage and Tobacco",
            "STOXX Europe 600 Personal Care Drug and Grocery Stores",
            "EURO STOXX 50",
            "STOXX Europe 50",
            "STOXX Europe 600",
            "STOXX USA 500",
            "STOXX US Nexus 100"
        ],
        "RIC": [
            ".GDAXI",
            ".MDAXI",
            ".SDAXI",
            ".DAXMS",
            ".DAXAC",
            ".STOXX50E",
            ".STOXX50",
            ".STOXX",
            ".SX50UP",
            ".SXUSX1P",
            ".SXP1E",
            ".STXWAP",
            ".V1XI",
            ".V2TX",
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
            ".STOXX50ED",
            ".STOXX50D",
            ".STOXXD",
            ".SX50UL",
            ".SXUSX1L"
        ]
    }
)

# Function to get Time Series data
def Get_TimeSeries(
    universe,
    fields,
    parameters,
    session_already_open=False
):
    opened_here = False

    try:
        if not session_already_open:
            ld.open_session(config_name="lseg-data.config.json")
            opened_here = True

        df_raw = ld.get_data(universe=universe, fields=fields, parameters=parameters)
        df = pl.DataFrame(df_raw)

        # Implement NULL:SKIP: drop rows where Close Price is null
        df = df.drop_nulls("Close Price")

        # Aggregate and afjust in case of duplicates
        Output = (
                    df
                    .unique(subset=["Date", "Instrument"], keep="last")
                    .pivot(index="Date", columns="Instrument", values="Close Price")
                    .sort("Date")
                    .with_columns([
                        pl.col("Date").dt.date(),  # convert DateTime → Date
                        pl.exclude("Date").forward_fill()  # forward fill price columns
                    ])
                )

        return Output

    finally:
        if opened_here:
            ld.close_session()