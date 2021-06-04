import pandas
from datetime import datetime, timedelta
from bokeh.plotting import figure,output_file,show

df = pandas.read_excel("Incidents.xlsx")

days = 9
#to_date = datetime(2021,5,10)
to_date = datetime.today() - timedelta(days=1)

date_list =[]
average_list=[]
average_dict ={}
from_date = to_date - timedelta(days=days)
delta = timedelta(days=1)
while from_date < to_date:
    total_age = 0
    count = 0

    for nu,cr,cl in zip(df["Number"],df["Created"],df["Closed"]):
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

output_file("Average Age.html")
f = figure(plot_width=500, plot_height=250, x_axis_type='datetime')
f.title.text = "Trend of Average Age of Open Incidents"
f.title.text_color = "Brown"
f.title.text_font = "times"
f.title.text_font_style = "bold"
f.yaxis.minor_tick_line_color = "Yellow"
f.xaxis.axis_label = "Date"
f.yaxis.axis_label = "Average Age"

excel_frame = pandas.read_excel("Average Age.xlsx",parse_dates=["Date"])
f.line(excel_frame["Date"],excel_frame["Average Age"])

show(f)