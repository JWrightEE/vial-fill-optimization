import pandas as pd
from datetime import datetime


def calculate_schedule(df, sort=False):
    if sort:
        df = df.sort_values(["batch_type", "fill_time"], ascending=[True, False]).copy()

    schedule = []
    time = 0
    last_clean_time = 0
    last_batch_type = None

    # Start with a cleaning
    schedule.append(("Cleaning", 24, 0))
    time += 24
    last_clean_time = time

    while not df.empty:
        for idx, row in df.iterrows():
            batch_id = row["batch_id"]
            batch_type = row["batch_type"]
            fill_time = row["fill_time"]

            if last_batch_type is None:
                changeover_time = 0
            elif batch_type == last_batch_type:
                changeover_time = 4
            else:
                changeover_time = 8

            # Check if this batch would trigger a cleaning
            if time - last_clean_time + fill_time + changeover_time > 120:
                if sort:
                    # Don't fill this batch now, try with the next one
                    continue
                else:
                    # Perform a cleaning
                    schedule.append(("Cleaning", 24, 0))
                    time += 24
                    last_clean_time = time
                    last_batch_type = None
                    changeover_time = 0

            # If we've reached this point, we can fill the current batch without triggering a cleaning
            schedule.append((batch_id, fill_time, changeover_time))
            time += fill_time + changeover_time
            last_batch_type = batch_type
            df = df.drop(idx)
            break

        else:
            if sort:
                # If we've gone through all unfilled batches and none of them can fit in the current window, then clean
                schedule.append(("Cleaning", 24, 0))
                time += 24
                last_clean_time = time
                last_batch_type = None
                changeover_time = 0

    schedule_df = pd.DataFrame(schedule, columns=["Event", "Fill Time", "Changeover Time"])
    schedule_df["Total Time"] = schedule_df["Fill Time"] + schedule_df["Changeover Time"]
    schedule_df["Cumulative Time"] = schedule_df["Total Time"].cumsum()

    return schedule_df


# Load the data
df = pd.read_excel("input_data.xlsx")

# Calculate the fill time for each batch
df["fill_time"] = df["vial_count"] / (332 * 60)  # convert minutes to hours

# Calculate the schedule for the original input data
original_schedule_df = calculate_schedule(df.copy(), sort=False)

# Calculate the optimized schedule
optimized_schedule_df = calculate_schedule(df.copy(), sort=True)

# Calculate the total times
original_total_time = original_schedule_df["Cumulative Time"].iloc[-1]
optimized_total_time = optimized_schedule_df["Cumulative Time"].iloc[-1]

# Calculate the percent improvement
percent_improvement = ((original_total_time - optimized_total_time) / original_total_time) * 100

# Define the path for the output file
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
output_file_path = f"filling_schedule_final_{timestamp}.xlsx"

# Create a Pandas Excel writer using XlsxWriter as the engine
writer = pd.ExcelWriter(output_file_path, engine='xlsxwriter')

# Write each DataFrame to a different worksheet
original_schedule_df.to_excel(writer, sheet_name='Original Schedule', index=False)
optimized_schedule_df.to_excel(writer, sheet_name='Optimized Schedule', index=False)

# Write the total times and percent improvement to a new worksheet
total_times_df = pd.DataFrame({
    "Schedule": ["Original", "Optimized"],
    "Total Time": [original_total_time, optimized_total_time],
    "Improvement": ["-", percent_improvement]
})
total_times_df.to_excel(writer, sheet_name='Summary', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets['Summary']

# Create a chart object
chart = workbook.add_chart({'type': 'column'})

# Configure the series of the chart from the dataframe data
chart.add_series({
    'categories': ['Summary', 1, 0, 2, 0],
    'values':     ['Summary', 1, 1, 2, 1],
    'name':       ['Summary', 0, 1],
})

# Insert the chart into the worksheet
worksheet.insert_chart('F2', chart)

# Close the Pandas Excel writer and output the Excel file
writer.close()
