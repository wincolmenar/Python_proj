from datetime import datetime

def dynamic_yymm(start_year, end_year):

    # empty list to store the mm.yy values
    mm_yy_series = []

    # year loop.  Python generates sequence of numbers up to, but not including, the stop. True number is until 2040
    for year in range(start_year, end_year):
        # month loop.  Python generates sequence of numbers up to, but not including, the stop. True number is until 12
        for month in range(1, 13):
            indiv_date_start = datetime(year, month, 1)
            mm_yy_series.append(indiv_date_start.strftime('%m.%y'))
    
    # return the list of mm.yy values
    return mm_yy_series

mm_yy = dynamic_yymm(2022, 2041)

print(mm_yy)