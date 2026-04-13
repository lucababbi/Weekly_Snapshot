import polars as pl
import lseg.data as ld

# FX_List
FX_List = pl.DataFrame(
    {
        "Symbol": [
            "EURUSD",
            "EURCHF",
            "EURGBP"
                ],
        "RIC": [
            "EUR17H=",
            "EURCHF17H=",
            "EURGBP17H="]
    }
)

# Function to retrieve the Time_Series of all the FX Rates
def Get_TimeSeriesFX(
    universe,
    fields,
    parameters,
    session_already_open=False
):
    opened_here = False

    try:
        if not session_already_open:
            ld.get_config()["http.request-timeout"] = 240.0
            ld.open_session()
            opened_here = True

        df_raw = ld.get_data(universe=universe, fields=fields, parameters=parameters)
        df = pl.DataFrame(df_raw)

        # Implement NULL:SKIP: drop rows where Mid Price is null
        df = df.drop_nulls("Mid Price")

        # Aggregate and afjust in case of duplicates
        Output = (
                    df
                    .unique(subset=["Date", "Instrument"], keep="last")
                    .pivot(index="Date", columns="Instrument", values="Mid Price")
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

    return df