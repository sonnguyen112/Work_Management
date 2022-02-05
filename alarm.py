import datetime
import json
import os
import sys

try:
    #Get data
    if getattr(sys, 'frozen', False):
        f = open(f"{os.path.dirname(sys.executable)}/data.json")
            # or a script file (e.g. `.py` / `.pyw`)
    elif __file__:
        f = open(f"{os.path.dirname(__file__)}/data.json")
    fdata = json.load(f)
    f.close()
    date_now = datetime.datetime.now()
    date_now_str = date_now.strftime("%d/%m/%Y")
    date_now_arr = date_now_str.split("/")
    day_now = int(date_now_arr[0])
    month_now = int(date_now_arr[1])
    year_now = int(date_now_arr[2])
    weekday_now = datetime.datetime.now().strftime("%A")
    is_alarm = False

    #Check alarm
    for work in fdata:
        date_data = work["time"]
        date_data_arr = date_data.split("/")
        day_data = int(date_data_arr[0])
        month_data = int(date_data_arr[1])
        year_data = int(date_data_arr[2])
        weekday_data = datetime.datetime(
            year_data, month_data, day_data
        ).strftime("%A")
        if work["loop"] == "Một lần":
            if (
                day_now == day_data
                and month_now == month_data
                and year_now == year_data
            ):
                is_alarm = True
        elif work["loop"] == "Mỗi ngày":
            is_alarm = True
            if not(day_now == day_data and month_now == month_data and year_now == year_data):
                work["state"] = "Chưa hoàn thành"
                work["time"] = f"{day_now}/{month_now}/{year_now}"
        elif work["loop"] == "Mỗi tuần":
            if weekday_now == weekday_data:
                is_alarm = True
            
            reset_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days = 1)
            next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(weeks = 1)
            if datetime.datetime.strptime(date_now_str, "%d/%m/%Y") >= reset_date:
                work["state"] = "Chưa hoàn thành"
                work["time"] = f"{next_date.strftime('%d')}/{next_date.strftime('%m')}/{next_date.strftime('%Y')}"
                    
            
        elif work["loop"] == "Mỗi tháng":
            if day_now == day_data:
                is_alarm = True
            
            reset_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days = 1)
            if (month_data == 1 or month_data == 3 or month_data == 5 or month_data == 7 or month_data == 8 or month_data == 10 or month_data == 12):
                next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days=31)
            elif (month_data == 4 or month_data == 6 or month_data == 9 or month_data == 11):
                next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days=30)
            elif (month_data == 2):
                if ((year_data % 4 == 0) and (year_data % 100 != 0) or (year_data % 400 == 0)):
                    next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days=29)
                else:
                    next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days=28)
            if datetime.datetime.strptime(date_now_str, "%d/%m/%Y") >= reset_date:
                work["state"] = "Chưa hoàn thành"
                work["time"] = f"{next_date.strftime('%d')}/{next_date.strftime('%m')}/{next_date.strftime('%Y')}"

        elif work["loop"] == "Mỗi quý":
            if day_now == date_data and (month_now - month_now) % 3 == 0:
                is_alarm = True
            
            reset_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days = 1)
            if datetime.datetime.strptime(date_now_str, "%d/%m/%Y") >= reset_date:
                work["state"] = "Chưa hoàn thành"
                work["time"] = f"{day_data}/{month_data + 4}/{year_data}"

        else:
            if day_now == day_data and month_now == month_data:
                is_alarm = True
            
            reset_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days = 1)
            if ((year_data % 4 == 0) and (year_data % 100 != 0) or (year_data % 400 == 0)):
                next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days=366)
            else:
                next_date = datetime.datetime(year_data, month_data, day_data) + datetime.timedelta(days=365)

            if datetime.datetime.strptime(date_now_str, "%d/%m/%Y") >= reset_date:
                work["state"] = "Chưa hoàn thành"
                work["time"] = f"{next_date.strftime('%d')}/{next_date.strftime('%m')}/{next_date.strftime('%Y')}"

    if getattr(sys, "frozen", False):
        with open(f"{os.path.dirname(sys.executable)}/data.json", "w") as f:
            json.dump(fdata, f)
    elif __file__:
        with open(f"{os.path.dirname(__file__)}/data.json", "w") as f:
            json.dump(fdata, f)

    #Alarm:
    if (is_alarm):
        if getattr(sys, 'frozen', False):
            f = open(f"{os.path.dirname(sys.executable)}/alarm.json")
            # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            f = open(f"{os.path.dirname(__file__)}/alarm.json")
        alarm = json.load(f)
        f.close()
        alarm["alarm"] = 1
        
        if getattr(sys, 'frozen', False):
            with open(f"{os.path.dirname(sys.executable)}\\alarm.json", "w") as f:
                json.dump(alarm, f)
        elif __file__:
            with open(f"{os.path.dirname(__file__)}/alarm.json", "w") as f:
                print("TH2")
                json.dump(alarm, f)

        if getattr(sys, 'frozen', False):
            os.system(f'start "file" "{os.path.dirname(sys.executable)}/WorkManage.exe"')
            # or a script file (e.g. `.py` / `.pyw`)
        elif __file__:
            os.system(f'start "file" "{os.path.dirname(__file__)}/WorkManage.exe"')
except Exception as Ex:
    print(f"Ex")
    a = input()