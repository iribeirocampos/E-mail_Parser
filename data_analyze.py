import pandas as pd
import datetime
import numpy as np


WORKDAY = datetime.timedelta(hours=9, minutes=0, seconds=0)
# 8 hours of work +1 lunch time
WORKDAY_WEEKAND = datetime.timedelta(hours=4)
# 4 hours of work on weekend


def data_anal():
    data = pd.read_csv("sent_emails.csv", ";")
    df_data = pd.DataFrame()

    df_data._reindex_columns = ["Date", "Week_day", "Worked_hours", "Overtime"]

    print(data)
    data["Date"] = data.Date.str.split(".", expand=True)
    data["Date"] = data.Date.str.split("+", expand=True)
    data[["Date", "Hour"]] = data.Date.str.split(" ", expand=True)

    data["Date"] = pd.to_datetime(data["Date"], format="%Y-%m-%d")

    def make_timedelta(time):
        print(time)
        hour, minute, seconds = time.split(":")
        hour, minute, seconds = int(hour), int(minute), int(seconds)
        hora = datetime.timedelta(hours=hour, minutes=minute, seconds=seconds)
        return hora

    data["Hour"] = data["Hour"].apply(make_timedelta)
    # converts string to time delta

    hour_bound = datetime.timedelta(hours=8)
    filt_hour = data["Hour"] >= hour_bound
    data = data[filt_hour]
    # removes all e-mails sent between as 24 an 8 am.

    df_data["Date"] = data["Date"].drop_duplicates()

    df_data["Week_day"] = df_data["Date"].dt.dayofweek

    df_data = df_data.reset_index(drop=True)

    inicial = []
    final = []
    number_mails = []
    for date in df_data["Date"]:
        hour = data.loc[data["Date"] == date, "Hour"]
        inicial.append(hour.min())
        final.append(hour.max())
        number_mails.append(len(hour))

    df_data.insert(2, column="Last_mail", value=final, allow_duplicates=False)
    # Column with hour of last sent e-mail
    df_data.insert(2, column="First_mail", value=inicial, allow_duplicates=False)
    # Column with first sent e-mail
    df_data["Worked_hours"] = df_data["Last_mail"] - df_data["First_mail"]
    # Column with worked hours
    df_data.insert(5, column="Num_Mails", value=number_mails, allow_duplicates=False)
    # Column with the number of mails sent

    filt_sat = df_data["Week_day"] == 5
    filt_sun = df_data["Week_day"] == 6

    df_results_fds = df_data.loc[filt_sat | filt_sun]

    df_results_sem = df_data.loc[~filt_sat | ~filt_sun]

    filt = df_results_sem["Worked_hours"] > WORKDAY
    df_results_sem = df_results_sem.loc[filt]

    df_results_sem["Overtime"] = df_results_sem["Worked_hours"] - WORKDAY

    filt_fds = df_results_fds["Worked_hours"] > WORKDAY_WEEKAND
    df_results_fds = df_results_fds[filt_fds]
    df_results_fds["Overtime"] = df_results_fds["Worked_hours"] - WORKDAY_WEEKAND

    df_final = df_results_sem.append(df_results_fds)

    df_final.insert(
        0, column="Day", value=df_final["Date"].dt.day, allow_duplicates=False
    )
    df_final.insert(
        0, column="Month", value=df_final["Date"].dt.month, allow_duplicates=False
    )
    df_final.insert(
        0, column="Year", value=df_final["Date"].dt.year, allow_duplicates=False
    )

    table = pd.pivot_table(
        df_final,
        values="Overtime",
        index=["Year", "Month"],
        columns=["Week_day"],
        aggfunc=(np.sum),
        fill_value="0",
    )
    table["Total"] = table.sum(axis=1) / np.timedelta64(1, "h")
    table["Total"] = table["Total"].round(decimals=2)
    table = table.rename(
        columns={
            0: "Monday",
            1: "Tuesday",
            2: "Wednesday",
            3: "Thursday",
            4: "Friday",
            5: "Saturday",
            6: "Sunday",
        }
    )
    print(f"You have {table['Total'].sum().round(decimals=2)} hours of overtime.")
    writer = pd.ExcelWriter("Overtime.xlsx", engine="xlsxwriter")

    # store your dataframes in a  dict, where the key is the sheet name you want
    frames = {"Overtime": table, "E_mails": df_data}

    for sheet, frame in frames.items():
        frame.to_excel(writer, sheet_name=sheet)

    writer.save()


if __name__ == "__main__":
    data_anal()
