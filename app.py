# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
from flask import Flask, request, send_file, flash, redirect, render_template
import pandas as pd
from io import BytesIO
import tempfile
import os

app = Flask(__name__)
app.secret_key = 'your_very_secret_key'

def process_excel_file(file_stream):
    # 示例：读取上传的Excel文件
    from io import StringIO
    file_stream_string = StringIO(file_stream.read().decode('utf-8'))
    data_df = pd.read_csv(file_stream_string)

    #No 1

    # Define the filtering criteria based on the actual column names and values
    month_criteria = "Mar'24"
    lead_sub_status_criteria = 'Procured'
    cities_criteria = ['Bangalore', 'Chennai', 'Coimbatore']
    
    # Apply the filter criteria
    filtered_data = data_df[
        (data_df['Month (Dashboard Purpose)'] == month_criteria) &
        (data_df['Lead Sub Status'] == lead_sub_status_criteria) &
        (data_df['City'].isin(cities_criteria))
    ]
    
    # Count the number of 'Procured' entries for each city
    procured_counts = filtered_data['City'].value_counts().reindex(cities_criteria, fill_value=0).reset_index()
    procured_counts.columns = ['City', 'Procured']
    
    procured_counts.insert(1, 'Plan C2B', [None]*len(procured_counts), True)
    procured_counts.insert(2, 'MTD Plan-C2B', [None]*len(procured_counts), True)
    procured_counts.insert(4, 'MTD Achievement-C2B', [None]*len(procured_counts), True)
    
    total_row = pd.DataFrame({
        'City': ['Total'],
        'Plan C2B': [None],
        'MTD Plan-C2B': [None],
        'Procured': [procured_counts['Procured'].sum()],
        'MTD Achievement-C2B': [None]
    })
    
    procured_counts = pd.concat([procured_counts, total_row], ignore_index=True)

    #No 2

    
    # Trim whitespace from the 'City' column to ensure consistent matching
    data_df['City'] = data_df['City'].str.strip()
    
    # Define the filtering criteria based on the actual column names and values
    month_criteria = "Mar'24"
    lead_sub_status_criteria = 'Procured'
    inspected_criteria = 'Yes'
    cities_criteria = ['Bangalore', 'Coimbatore']
    
    # Apply the filters for each criteria
    scheduled_filter = data_df['Month (Dashboard Purpose)'] == month_criteria
    completed_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Inspected'] == inspected_criteria)
    procured_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Lead Sub Status'] == lead_sub_status_criteria)
    
    
    # Apply the filters for each criteria and count the occurrences for each city
    # For the 'Inspections Scheduled' column (simply filter by month as all leads are scheduled by default)
    scheduled_counts = data_df[scheduled_filter]['City'].value_counts().reindex(cities_criteria, fill_value=0)
    
    # For the 'Inspections Completed' column (filter by month and 'Inspected' status)
    completed_counts = data_df[completed_filter]['City'].value_counts().reindex(cities_criteria, fill_value=0)
    
    # For the 'Procurement Digital Leads' column (filter by month and 'Lead Sub Status')
    procured_counts1 = data_df[procured_filter]['City'].value_counts().reindex(cities_criteria, fill_value=0)
    
    # Combine the counts into one DataFrame for output
    combined_counts = pd.DataFrame({
        'City': cities_criteria,
        'Inspections Scheduled': scheduled_counts.values,
        'Inspections Completed': completed_counts.values,
        'Procurement Digital Leads': procured_counts1.values
    })
    total_row_for_combined_counts = pd.DataFrame({
        'City': ['Total'],
        'Inspections Scheduled': [combined_counts['Inspections Scheduled'].sum()],
        'Inspections Completed': [combined_counts['Inspections Completed'].sum()],
        'Procurement Digital Leads': [combined_counts['Procurement Digital Leads'].sum()],
    })
    
    combined_counts = pd.concat([combined_counts, total_row_for_combined_counts], ignore_index=True)
    
    combined_counts.insert(1, 'Raw Leads', [None]*len(combined_counts), True)
    combined_counts.insert(5, 'Raw Leads Vs Procurement Conversion %', [None]*len(combined_counts), True)
    
    
    
    combined_counts['Inspections Scheduled Vs Procurement'] = combined_counts.apply(
        lambda row: "{:.2%}".format(row['Procurement Digital Leads'] / row['Inspections Scheduled'])
        if row['Inspections Scheduled'] > 0 else None, axis=1
    )
    
    
    combined_counts['Inspections Completed Vs Conversion'] = combined_counts.apply(
        lambda row: "{:.2%}".format(row['Procurement Digital Leads'] / row['Inspections Completed'])
        if row['Inspections Completed'] > 0  else None, axis=1
    )
    combined_counts.insert(8, 'Raw Leads Vs Scheduled', [None]*len(combined_counts), True)
    
    combined_counts['Inspections Scheduled Vs Inspections Completed'] = combined_counts.apply(
        lambda row: "{:.2%}".format(row['Inspections Completed'] / row['Inspections Scheduled'])
        if row['Inspections Scheduled'] > 0  else None, axis=1
    )

    #No 3
    # Trim whitespace from the 'City' column to ensure consistent matching
    data_df['City'] = data_df['City'].str.strip()
    
    # Define the filtering criteria based on the actual column names and values
    month_criteria = "Mar'24"
    lead_sub_status_criteria = 'Procured'
    inspected_criteria = 'Yes'
    cities_criteria = ['Bangalore', 'Chennai', 'Coimbatore']
    
    # Apply the filters for each criteria
    scheduled_filter = data_df['Month (Dashboard Purpose)'] == month_criteria
    completed_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Inspected'] == inspected_criteria)
    procured_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Lead Sub Status'] == lead_sub_status_criteria)
    
    # Apply the filters for each criteria and count the occurrences for each city
    # For the 'Inspections Scheduled' column (simply filter by month as all leads are scheduled by default)
    scheduled_counts = data_df[scheduled_filter]['City'].value_counts().reindex(cities_criteria, fill_value=0)
    
    # For the 'Inspections Completed' column (filter by month and 'Inspected' status)
    completed_counts = data_df[completed_filter]['City'].value_counts().reindex(cities_criteria, fill_value=0)
    
    # For the 'Procurement Digital Leads' column (filter by month and 'Lead Sub Status')
    procured_counts1 = data_df[procured_filter]['City'].value_counts().reindex(cities_criteria, fill_value=0)
    
    # Combine the counts into one DataFrame for output
    combined_counts2 = pd.DataFrame({
        'City': cities_criteria,
        'Inspections Scheduled': scheduled_counts.values,
        'Inspections Completed': completed_counts.values,
        'Procurement Digital Leads': procured_counts1.values
    })
    
    total_row_for_combined_counts2 = pd.DataFrame({
        'City': ['Total'],
        'Inspections Scheduled': [combined_counts2['Inspections Scheduled'].sum()],
        'Inspections Completed': [combined_counts2['Inspections Completed'].sum()],
        'Procurement Digital Leads': [combined_counts2['Procurement Digital Leads'].sum()],
    })
    
    combined_counts2 = pd.concat([combined_counts2, total_row_for_combined_counts2], ignore_index=True)
    
    combined_counts2.insert(1, 'Raw Leads', [None]*len(combined_counts2), True)
    combined_counts2.insert(5, 'Raw Leads Vs Procurement Conversion %', [None]*len(combined_counts2), True)
    
    combined_counts2['Inspections Scheduled Vs Procurement'] = combined_counts2.apply(
        lambda row: "{:.2%}".format(row['Procurement Digital Leads'] / row['Inspections Scheduled'])
        if row['Inspections Scheduled'] > 0 else None, axis=1
    )
    
    
    combined_counts2['Inspections Completed Vs Conversion'] = combined_counts2.apply(
        lambda row: "{:.2%}".format(row['Procurement Digital Leads'] / row['Inspections Completed'])
        if row['Inspections Completed'] > 0  else None, axis=1
    )
    combined_counts2.insert(8, 'Raw Leads Vs Scheduled', [None]*len(combined_counts2), True)
    
    combined_counts2['Inspections Scheduled Vs Inspections Completed'] = combined_counts2.apply(
        lambda row: "{:.2%}".format(row['Inspections Completed'] / row['Inspections Scheduled'])
        if row['Inspections Scheduled'] > 0  else None, axis=1
    )

    #No.4
    
    # Define the filtering criteria based on the actual column names and values
    month_criteria = "Mar'24"
    cities_criteria = ['Bangalore', 'Chennai', 'Coimbatore']
    lead_sub_status_criteria = 'Procured'
    status_columns = [
        'Already Sold', 'Bad Vehicle Quality', 'Improper Documents',
        'Inspected, but Sold Outside', 'Inspected but, Not Interested',
        'Inspection Rescheduled', 'Not Interested', 'Out of Purchase Criteria',
        'Out of Station', 'Pricing Gap', 'RNR', 'Under Discussion', 'Vehicle at Service'
    ]
    final_results = pd.DataFrame(columns=['City'] + status_columns)
    final_results['City'] = cities_criteria
    month_filter = data_df['Month (Dashboard Purpose)'] == month_criteria
    procured_filter= (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Lead Sub Status'] == lead_sub_status_criteria)
    
    
    # Use 'groupby' and 'size' to count occurrences instead of 'value_counts' to avoid index issues
    qualified_leads_counts = data_df[month_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    procured_counts1 = data_df[procured_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    # Assign the counts to 'final_results'
    final_results['Total Qualified Leads Received'] = qualified_leads_counts.values
    final_results['MTD Procurement'] = procured_counts1.values
    
    
    
    # Loop through each status column and get the counts for each city
    for status in status_columns:
        status_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Lead Sub Status'] == status)
        status_counts = data_df[status_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
        final_results[status] = status_counts.values
    
        
    # To calculate 'MTD LOST', we need to sum all the relevant columns except the four specified
    # First, let's add a placeholder for 'MTD LOST' in the correct position
    final_results.insert(final_results.columns.get_loc('Already Sold'), 'MTD LOST', 0)
    
    # Define the columns to exclude from the sum for 'MTD LOST'
    exclude_columns = ['Inspection Rescheduled', 'Out of Station', 'Under Discussion', 'Vehicle at Service']
    
    # Calculate the 'MTD LOST' for each city by summing all the relevant status columns
    final_results['MTD LOST'] = final_results[status_columns].drop(columns=exclude_columns).sum(axis=1)
    
    cols = final_results.columns.tolist()
    new_order = cols[:1] + cols[-2:] + cols[1:-2]
    final_results = final_results[new_order]
    
    totals_data = {
        'City': ['Total'],
        'Total Qualified Leads Received': [final_results['Total Qualified Leads Received'].sum()],
        'MTD Procurement': [final_results['MTD Procurement'].sum()],
        'MTD LOST': [final_results['MTD LOST'].sum()]
    }
    
    # Add each status column's total to the dictionary
    for status in status_columns:
        totals_data[status] = [final_results[status].sum()]
        
    totals_data = pd.DataFrame(totals_data)
    
    final_results = pd.concat([final_results, totals_data], ignore_index=True)


    #No.5

    # Standardize the 'Source' column by stripping whitespace and converting to a consistent case
    data_df['Source'] = data_df['Source'].str.strip().str.lower()
    
    # Filter the data for March 2024
    march_data = data_df[data_df['Month (Dashboard Purpose)'] == "Mar'24"]
    
    # Now we group by the cleaned 'Source' and count the 'Scheduled/Qualified Leads'
    grouped_source = march_data.groupby('Source').size().reset_index(name='Scheduled/Qualified Leads')
    
    # Ensure the source names are in title case for presentation
    grouped_source['Source'] = grouped_source['Source'].str.title()
    
    grouped_source.insert(1, 'Raw Leads', ['WIP']*len(grouped_source), True)
    
    group_data = {
        'Source': ['Total'],
        'Raw Leads':'WIP',
        'Scheduled/Qualified Leads': [grouped_source['Scheduled/Qualified Leads'].sum()]
    }
    
    group_data=pd.DataFrame(group_data)
    
    grouped_source = pd.concat([grouped_source, group_data], ignore_index=True)

    # No.6-1 Bikes

    # Define the criteria for the different columns
    month_criteria = "Mar'24"
    type_criteria = 'Bike'
    procured_criteria = 'Procured'
    inspected_criteria='Yes'
    
    
    # Filter the data based on the given criteria for 'Bikes Scheduled'
    bikes_scheduled_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Type'] == type_criteria)
    bikes_scheduled_data = data_df[bikes_scheduled_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    # Filter the data based on the given criteria for 'Bikes Procured'
    bikes_procured_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & \
                            (data_df['Type'] == type_criteria) & (data_df['Lead Sub Status'] == procured_criteria)
    bikes_procured_data = data_df[bikes_procured_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    # Create the final DataFrame with the required columns
    final_table_six = pd.DataFrame({
        'City': cities_criteria,
        'Bikes Scheduled': bikes_scheduled_data,
        'Bikes Procured': bikes_procured_data
    })
    
    bikes_completed_filter=(data_df['Month (Dashboard Purpose)'] == month_criteria) &\
                          (data_df['Type'] == type_criteria) & (data_df['Inspected']==inspected_criteria)
    bikes_completed_data = data_df[bikes_completed_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    final_table_six['Bike Inspections Completed']=bikes_completed_data
    
    total_final_data1 = pd.DataFrame({
        'City': ['Total'],
        'Bikes Scheduled': [final_table_six['Bikes Scheduled'].sum()],
        'Bikes Procured': [final_table_six['Bikes Procured'].sum()],
        'Bike Inspections Completed': [final_table_six['Bike Inspections Completed'].sum()],
    })
    
    final_table_six = pd.concat([final_table_six, total_final_data1], ignore_index=True)
    
    
    
    final_table_six['% Bikes Procured Vs Total Inspections Scheduled'] = final_table_six.apply(
        lambda row: "{:.2%}".format(row['Bikes Procured'] / row['Bikes Scheduled'])
        if row['Bikes Scheduled'] > 0 else None, axis=1
    )
    
    
    final_table_six['% Bikes Procured Vs Total Inspections Completed'] = final_table_six.apply(
        lambda row: "{:.2%}".format(row['Bikes Procured'] / row['Bike Inspections Completed'])
        if row['Bike Inspections Completed'] > 0 else None, axis=1
    )
    
    # Assuming final_table_six is your DataFrame
    columns_list = final_table_six.columns.tolist()
    
    # Swap the positions of the 4th and 5th columns
    columns_list[3], columns_list[4] = columns_list[4], columns_list[3]
    
    # Reassign the columns to the DataFrame based on the new order
    final_table_six = final_table_six[columns_list]

   # No.6-2 Scooter

    # Define the criteria for the different columns
    month_criteria = "Mar'24"
    type_criteria = 'Scooter'
    procured_criteria = 'Procured'
    
    # Filter the data based on the given criteria for 'Bikes Scheduled'
    scooters_scheduled_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Type'] == type_criteria)
    scooters_scheduled_data = data_df[scooters_scheduled_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    # Filter the data based on the given criteria for 'Bikes Procured'
    scooters_procured_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & \
                            (data_df['Type'] == type_criteria) & (data_df['Lead Sub Status'] == procured_criteria)
    scooters_procured_data = data_df[scooters_procured_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    # Create the final DataFrame with the required columns
    final_table_six2 = pd.DataFrame({
        'City': cities_criteria,
        'Scooters Scheduled': scooters_scheduled_data,
        'Scooters Procured': scooters_procured_data
    })
    
    
    scooters_completed_filter=(data_df['Month (Dashboard Purpose)'] == month_criteria) &\
                          (data_df['Type'] == type_criteria) & (data_df['Inspected']==inspected_criteria)
    scooters_completed_data = data_df[scooters_completed_filter].groupby('City').size().reindex(cities_criteria, fill_value=0)
    
    final_table_six2['Scooter Inspections Completed']=scooters_completed_data
    
    total_final_data2 = pd.DataFrame({
        'City': ['Total'],
        'Scooters Scheduled': [final_table_six2['Scooters Scheduled'].sum()],
        'Scooters Procured': [final_table_six2['Scooters Procured'].sum()],
        'Scooter Inspections Completed': [final_table_six2['Scooter Inspections Completed'].sum()],
    })
    
    final_table_six2 = pd.concat([final_table_six2, total_final_data2], ignore_index=True)
    
    
    
    final_table_six2['% Scooters Procured Vs Total Inspections Scheduled'] = final_table_six2.apply(
        lambda row: "{:.2%}".format(row['Scooters Procured'] / row['Scooters Scheduled'])
        if row['Scooters Scheduled'] > 0 else None, axis=1
    )
    
    
    final_table_six2['% Scooters Procured Vs Total Inspections Completed'] = final_table_six2.apply(
        lambda row: "{:.2%}".format(row['Scooters Procured'] / row['Scooter Inspections Completed'])
        if row['Scooter Inspections Completed'] > 0 else None, axis=1
    )
    
    # Assuming final_table_six is your DataFrame
    columns_list_2 = final_table_six2.columns.tolist()
    
    # Swap the positions of the 4th and 5th columns
    columns_list_2[3], columns_list_2[4] = columns_list_2[4], columns_list_2[3]
    
    # Reassign the columns to the DataFrame based on the new order
    final_table_six2 = final_table_six2[columns_list_2]

    #No.7
    # Define the criteria for filtering
    month_criteria = "Mar'24"
    inspected_criteria = 'Yes'
    procured_criteria = 'Procured'
    
    # Filter the data based on the criteria for 'Inspections Scheduled Till Date'
    inspections_scheduled_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria)
    inspections_scheduled_data = data_df[inspections_scheduled_filter].groupby(['City', 'Evaluator Name']).size().reset_index(name='Inspections Scheduled Till Date')
    
    # Filter the data based on the criteria for 'Inspections Completed Till Date'
    inspections_completed_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Inspected'] == inspected_criteria)
    inspections_completed_data = data_df[inspections_completed_filter].groupby(['City', 'Evaluator Name']).size().reset_index(name='Inspections Completed Till Date')
    
    # Filter data for 'MTD C2B Procurement'
    procured_filter = (data_df['Month (Dashboard Purpose)'] == month_criteria) & (data_df['Lead Sub Status'] == procured_criteria)
    procured_data = data_df[procured_filter].groupby(['City', 'Evaluator Name']).size().reset_index(name='MTD C2B Procurement')
    
    
    # Merge the dataframes
    merged_data = pd.merge(inspections_scheduled_data, inspections_completed_data, on=['City', 'Evaluator Name'], how='left')
    merged_data = pd.merge(merged_data, procured_data, on=['City', 'Evaluator Name'], how='left')
    
    # Fill NaN values with zeros since they represent counts
    merged_data.fillna(0, inplace=True)
    
    # Convert float columns to int, because after merging, counts may become float if there are NaNs
    int_cols = ['Inspections Scheduled Till Date', 'Inspections Completed Till Date', 'MTD C2B Procurement']
    merged_data[int_cols] = merged_data[int_cols].astype(int)
    
    merged_data.insert(4, 'Target C2B', [None]*len(merged_data), True)
    
    merged_data['C2B Conversion % Lead Vs Achievement'] = merged_data.apply(
        lambda row: "{:.2%}".format(row['MTD C2B Procurement'] / row['Inspections Scheduled Till Date'])
        if row['Inspections Scheduled Till Date'] > 0 else None, axis=1
    )
    
    
    # Step 1: Group by 'City' and calculate the sum for numeric columns
    city_totals = merged_data.groupby('City')[int_cols].sum().reset_index()
    
    # Step 2: Create a separate overall total row
    overall_total = city_totals[int_cols].sum()
    overall_total['City'] = 'Grand Total'
    overall_total = overall_total.to_frame().transpose()
    
    # Step 3: Insert total rows under each city in the original dataframe
    final_data_with_city_totals = pd.DataFrame()
    for city in merged_data['City'].unique():
        city_data = merged_data[merged_data['City'] == city]
        total_row = city_totals[city_totals['City'] == city]
        final_data_with_city_totals = pd.concat([final_data_with_city_totals, city_data, total_row], ignore_index=True)
    
    # Step 4: Append the overall total row at the end of the DataFrame
    final_data_with_totals = pd.concat([final_data_with_city_totals, overall_total], ignore_index=True)
    
    # Re-add the 'Target C2B' and 'C2B Conversion % Lead Vs Achievement' columns
    final_data_with_totals['Target C2B'] = None
    final_data_with_totals['C2B Conversion % Lead Vs Achievement'] = final_data_with_totals.apply(
        lambda row: "{:.2%}".format(row['MTD C2B Procurement'] / row['Inspections Scheduled Till Date'])
        if row['Inspections Scheduled Till Date'] > 0 else None, axis=1
    )
    
    # Assuming 'final_data_with_totals' is your DataFrame
    
    # Get all the unique city names including 'Grand Total'
    city_names = final_data_with_totals['City'].unique()
    
    # Update 'City' for city total rows and empty the 'Evaluator Name' for both city totals and 'Grand Total'
    for city in city_names:
        if city != "Grand Total":
            # City totals
            city_total_filter = (final_data_with_totals['City'] == city) & final_data_with_totals['Evaluator Name'].isna()
            final_data_with_totals.loc[city_total_filter, 'City'] = f'Total {city}'
            final_data_with_totals.loc[city_total_filter, 'Evaluator Name'] = ''
        else:
            # Grand Total
            grand_total_filter = final_data_with_totals['City'] == 'Grand Total'
            final_data_with_totals.loc[grand_total_filter, 'Evaluator Name'] = ''
        
    
    
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
       startrow = 0
       dataframes = {
           'procured_counts': procured_counts,
           'combined_counts': combined_counts,
           'combined_counts2': combined_counts2,
           'final_results': final_results,
           'grouped_source': grouped_source,
           'final_table_six': final_table_six,
           'final_table_six2': final_table_six2,
           'final_data_with_totals': final_data_with_totals
       }
       for df_name, df in dataframes.items():
           df.to_excel(writer, sheet_name='Sheet1', startrow=startrow, index=False)
           startrow += len(df.index) + 2  # 2 rows of padding between DataFrames
    output.seek(0)
    return output


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['file']
        if file:
            # Process the CSV file
            output = process_excel_file(file)
            output.seek(0)
            # Make sure to specify the correct MIME type for an Excel file
            return send_file(
                output, 
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                as_attachment=True, 
                download_name="processed_file.xlsx"
            )
        else:
            flash('No file selected')
            return redirect(request.url)

    return render_template("index.html")

if __name__ == '__main__':
    app.run(debug=True, port=8080)
