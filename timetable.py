import pyodbc
import pandas as pd
import matplotlib.pyplot as plt
import os
from string import digits
import re
from datetime import datetime


def connect_to_access(path):
    # Try connecting to Microsoft access
    print("Connecting to Microsoft Access File...")
    try:
        conn_str = r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};" fr"DBQ={path};"
        conn = pyodbc.connect(conn_str)
        crsr = conn.cursor()
        for table_info in crsr.tables(tableType="TABLE"):
            print(table_info.table_name)
        return conn

    except pyodbc.Error as e:
        print("Error in Connection", e)


def generate_timetable(dfs):
    fig, axs = plt.subplots(len(dfs), sharex=True, sharey=True)
    fig.set_size_inches(18.5, 10.5, forward=True)

    # hide the axes
    fig.patch.set_visible(False)

    for index, cls_df in enumerate(dfs):
        cls_name, df = cls_df

        axs[index].axis("off")
        axs[index].set_title(cls_name)

        axs[index].table(
            cellText=df.values,
            colLabels=df.columns,
            rowLabels=df.index,
            cellLoc="center",
            rowColours=["palegreen"] * len(df.values),
            colColours=["palegreen"] * len(df.columns),
            loc="upper left",
        )

    now = datetime.now()
    plt.savefig(f'output-{now.strftime("%Y%m%d")}.png')


def day_of_week(day: str) -> int:
    dayofweek = {
        "Monday": 1,
        "Tuesday": 2,
        "Wednesday": 3,
        "Thursday": 4,
        "Friday": 5,
        "Saturday": 6,
        "Sunday": 7,
    }

    return dayofweek[day]


def modify_data_to_timetable(df: pd.DataFrame) -> list[pd.DataFrame]:

    classes = df["class_name"].unique()
    dfs: list[pd.DataFrame] = []

    for cls in classes:  # Create new table for each class
        cls_df = df[df["class_name"] == cls].copy(deep=True)

        # Create the new table
        cls_df["timeslot"] = (
            cls_df["start_time"].dt.strftime("%H:%M %p")
            + " - "
            + cls_df["end_time"].dt.strftime("%H:%M %p")
        )

        cls_df["subject"] = cls_df["subject_id"].str.translate(
            str.maketrans("", "", digits)
        )

        cls_df = cls_df[["timeslot", "Day", "subject"]]
        cls_df = pd.pivot(cls_df, index="Day", columns="timeslot", values="subject")

        for index in cls_df.index:
            cls_df.loc[index, "day_num"] = day_of_week(index)

        cls_df.sort_values(["day_num"], inplace=True)
        cls_df.drop("day_num", axis=1, inplace=True)

        cls_df.fillna(" ", inplace=True)

        dfs.append((cls, cls_df))

    return dfs


def load_sql_to_dataframe(query: str, file_name: str, testing=False) -> pd.DataFrame:

    path = os.path.join(os.getcwd(), file_name)

    if not os.path.exists("data.pkl") or not testing:
        conn = connect_to_access(path)

        # Using pandas to execute the query
        df = pd.read_sql_query(query, conn)
        df.to_pickle("./data.pkl")

    else:
        print("Load already existing data...")
        df = pd.read_pickle("./data.pkl")

    return df


def check_conflict(df: pd.Series):
    teacher_time = set()
    pattern = re.compile(r"(\d+)T(\d+)[I,S,R,K,J,A](\d+)")

    for index, value in df.iteritems():
        timetable_id = pattern.findall(value)

        t_time = (timetable_id[0][1], timetable_id[0][2])
        if t_time not in teacher_time:
            teacher_time.add(t_time)

        elif t_time in teacher_time:
            raise ValueError(f"Conflict found at index {index}")


def main():
    # SQL Query for joining all the tables
    query = "SELECT * FROM classTimetableQ"
    access_file_name = "database.accdb"

    df = load_sql_to_dataframe(query, access_file_name)

    df["start_time"] = pd.to_datetime(df["start_time"])
    df["end_time"] = pd.to_datetime(df["end_time"])

    print(df)

    check: str = input("Would you like to check for conflicts? (Y/N): ")
    if check.lower() in ["y", "yes"]:
        check_conflict(df["timetable_id"])

    classes_dfs = modify_data_to_timetable(df)

    generate_timetable(classes_dfs)

    print("Successfully created a new timetable.")


if __name__ == "__main__":
    main()
