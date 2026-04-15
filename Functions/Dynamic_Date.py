from datetime import date, timedelta

def Last_Friday(ref_date=None):
    if ref_date is None:
        ref_date = date.today()
    days_back = (ref_date.weekday() - 4) % 7  # Friday = 4
    return ref_date - timedelta(days=days_back)

