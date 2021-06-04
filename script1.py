import pandas
from datetime import datetime, timedelta

df = pandas.read_excel("Incidents.xlsx")

days = 9
to_date = datetime(2021,5,10)

date_list =[]
average_list=[]
average_dict ={}
from_date = to_date - timedelta(days=days)
delta = timedelta(days=1)
while from_date < to_date:
    total_age = 0
    count = 0

    for nu,cr,cl in zip(df["Number"],df["Created On"],df["Closed On"]):
        if (from_date < cl  or pandas.isna(cl)) and from_date > cr:
            total_age += int((from_date-cr).days)
            count += 1
    from_date += delta
    if count == 0:
        average_age = 0
    else:
        average_age = total_age/count
    date_list.append(from_date)
    average_list.append(average_age)
    average_dict['Date'] = from_date
    average_dict['Average Age'] = average_age
dump_list = [date_list,average_list]  

dump_frame = pandas.DataFrame(dump_list).transpose()
dump_frame.columns=['Date','Average Age']   

dump_frame.to_excel("Average Age.xlsx")
