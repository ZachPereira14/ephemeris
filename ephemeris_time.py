import pandas as pd 
import plotly.express as px
import pytz
from datetime import datetime, timedelta
from openpyxl.drawing.image import Image
import plotly.io as pio
import os
import json

def convert_utc_to_est(utc_dt):
    """
    Convert a UTC datetime object to Eastern Standard Time (EST).

    Parameters:
    utc_dt (datetime): A naive or aware datetime object in UTC.

    Returns:
    datetime: The converted datetime object in EST.
    """
    utc_zone = pytz.utc
    est_zone = pytz.timezone('America/New_York')
    utc_dt = utc_zone.localize(utc_dt)
    est_dt = utc_dt.astimezone(est_zone)
    return est_dt

def save_configuration(settings, file_path='config.json'):
    """
    Save user configuration settings to a file for future use.

    Parameters:
    settings (dict): A dictionary of settings to save.
    file_path (str): The path to the configuration file. Defaults to 'config.json'.

    Returns:
    None: The function saves the settings to a configuration file.
    
    Example usage: 
        user_settings = {
        'magnitude_limit': (0, 14.5),
        'air_mass_lim': True,
        'transit_depth_limit': (0, 0.5),
        'max_airmass': (2, 2),
        'setup_time': False,
    }

    save_configuration(user_settings)
    """
    try:
        with open(file_path, 'w') as config_file:
            json.dump(settings, config_file, indent=4)
        print(f"Configuration saved successfully to {file_path}.")
    except Exception as e:
        print(f"An error occurred while saving configuration: {e}")

def calculate_transit_times(midpoint, duration, setup_time):
    """
    Calculate the transit start and end times based on a midpoint and duration.

    Parameters:
    midpoint (str or datetime): The midpoint of the transit.
    duration (float): The duration of the transit in hours.
    setup_time (bool): Indicates whether setup time is included in the calculation.

    Returns:
    tuple: A tuple containing the transit start time and transit end time as datetime objects.
    """
    if setup_time:
        transit_start_time = pd.to_datetime(midpoint) - pd.Timedelta(minutes=30) - pd.Timedelta(hours=duration / 2)
        transit_end_time = pd.to_datetime(midpoint) + pd.Timedelta(minutes=30) + pd.Timedelta(hours=duration / 2)
    else:
        transit_start_time = pd.to_datetime(midpoint) - pd.Timedelta(hours=duration / 2)
        transit_end_time = pd.to_datetime(midpoint) + pd.Timedelta(hours=duration / 2)
    return transit_start_time, transit_end_time

def format_datetime(dt):
    """
    Format a datetime object as a string in the format 'YYYY-MM-DD HH:MM'.

    Parameters:
    dt (datetime): The datetime object to format.

    Returns:
    str: The formatted datetime string.
    """
    return dt.strftime("%Y-%m-%d %H:%M")

def analyze_schedules(all_schedules):
    """
    Analyze schedules for optimal target selection based on given criteria.

    Parameters:
    all_schedules (list): A list of schedules containing various parameters for analysis.

    Returns:
    None: The function modifies the state but does not return any value.
    """
    # Convert all_schedules to a DataFrame for easier manipulation
    all_schedules_df = pd.DataFrame(all_schedules, columns=[ 
        'Name', 'Duration (hours)', 'Midpoint', 'Transit Start Time', 'Transit End Time', 
        'RA', 'Dec', 'Period', 'Transit Depth', 'Air Mass', 'Magnitude K'
    ])
    
    # Implement your criteria for "optimal" here
    # For example, find schedules with minimal gaps, maximize number of scheduled targets, etc.

    # Example analysis (modify based on your criteria):
    num_targets = all_schedules_df['Name'].nunique()
    total_time = (all_schedules_df['Transit End Time'].max() - all_schedules_df['Transit Start Time'].min()).total_seconds() / 3600


def plot_schedule_with_plotly(schedule_df, setup_time):
    """
    Plot the optimized schedule using Plotly and save it as an image.

    Parameters:
    schedule_df (DataFrame): A DataFrame containing the schedule to plot.

    Returns:
    None: The function modifies the state but does not return any value.
    """
    if setup_time: 
        plot_df = pd.DataFrame({
            'Task': schedule_df['Name'],
            'Start': schedule_df['Schedule Start Time'],
            'Finish': schedule_df['Schedule End Time'],
            'Resource': schedule_df['Name']
        })
    
    else: 
        plot_df = pd.DataFrame({
            'Task': schedule_df['Name'],
            'Start': schedule_df['Transit Start Time'],
            'Finish': schedule_df['Transit End Time'],
            'Resource': schedule_df['Name']
        })

    fig = px.timeline(plot_df, x_start='Start', x_end='Finish', y='Resource', color='Resource')
    fig.update_layout(title='Optimized Exoplanet Observation Schedule',
                      xaxis_title='Time (UTC)',
                      yaxis_title='Exoplanet Names',
                      xaxis=dict(tickformat="%m/%d %H:%M"),
                      yaxis=dict(title='Exoplanets', autorange="reversed"))
    
    # Save the figure as a PNG file
    image_file = "optimized_exoplanet_observation_schedule.png"
    pio.write_image(fig, image_file)

    # Show the figure in the notebook (or browser)
    fig.show()

    #print(f"Figure saved as {image_file}")

def optimize_schedule(df, magnitude_limit=(0, 14.5), air_mass_lim=True, setup_time=False, 
                      period_limit=False, transit_depth_limit=(0, 0.5), 
                      max_airmass=(2, 2), nanignore_airmass=False, 
                      time_window=(None, None)):
    """
    Optimize the observation schedule based on various criteria and limits.

    Parameters:
    df (DataFrame): A DataFrame containing the data to optimize.
    magnitude_limit (tuple): A tuple defining the acceptable magnitude range (min, max).
    air_mass_lim (bool): Whether to apply air mass limits.
    setup_time (bool): Whether to include setup time in the calculation.
    period_limit (bool): Whether to apply limits on the period of the planets.
    transit_depth_limit (tuple): A tuple defining the acceptable transit depth range (min, max).
    max_airmass (tuple): A tuple defining the maximum acceptable air mass (ingress, egress).
    nanignore_airmass (bool): Whether to ignore NaN values in air mass calculations.
    time_window (tuple): A tuple defining the time window as (start_time, end_time) in UTC. 
                         If either value is None, that limit is ignored.

    Returns:
    tuple: A tuple containing two DataFrames: the optimized schedule and the cut list of rejected candidates.
    """
    maxingressairmass, maxegressairmass = max_airmass
    valid_candidates = []
    cut_list = []
    all_schedules = []  # Store all valid schedules for comparison

    # Check if ingress/egress airmass columns exist in the dataframe
    ingress_column_exists = 'ingressairmass' in df.columns
    egress_column_exists = 'egressairmass' in df.columns
    
    # Time window limits
    window_start, window_end = time_window
    
    for index, row in df.iterrows():
        name = row['planetname']
        duration = row['transitduration']
        midpoint = row['midpointcalendar']
        ra = row['ra']
        dec = row['dec']
        period = row['period']
        transitdepth = row['transitdepthcalc']
        air_mass = row['midpointairmass']

        ingress_air_mass = row.get('ingressairmass', None) if ingress_column_exists else None
        egress_air_mass = row.get('egressairmass', None) if egress_column_exists else None

        mag_k = row['magnitude_k']

        # Initialize current_row_info to None
        current_row_info = None

        if pd.isna(duration):
            cause = "Duration is NaN"
            cut_list.append([name, duration, midpoint, None, None, ra, dec, period, transitdepth, air_mass, mag_k, cause])
            continue

        if period_limit and (period < period_limit[0] or period > period_limit[1]):
            cause = "Period limit exceeded"
            cut_list.append([name, duration, midpoint, None, None, ra, dec, period, transitdepth, air_mass, mag_k, cause])
            continue

        transit_start_time, transit_end_time = calculate_transit_times(midpoint, duration, setup_time)

        # Check against the time window
        cause = None
        if window_start is not None and transit_start_time < window_start:
            cause = "Transit start time is before the time window"
        elif window_end is not None and transit_end_time > window_end:
            cause = "Transit end time is after the time window"

        if cause:
            # Append current_row_info if defined; otherwise, append a new entry
            cut_list.append(current_row_info + [cause] if current_row_info else [name, duration, midpoint, None, None, ra, dec, period, transitdepth, air_mass, mag_k, cause])
            continue

        # Only assign after checking conditions
        current_row_info = [name, duration, midpoint, transit_start_time, transit_end_time, 
                            ra, dec, period, transitdepth, air_mass, mag_k]

        cause = None
        if magnitude_limit and (mag_k < magnitude_limit[0] or mag_k > magnitude_limit[1]):
            cause = "Magnitude limit exceeded"
            cut_list.append(current_row_info + [cause])
        elif air_mass_lim and air_mass > 2:
            cause = "Air mass limit exceeded"
            cut_list.append(current_row_info + [cause])
        elif transit_depth_limit and (transitdepth < transit_depth_limit[0] or transitdepth > transit_depth_limit[1]):
            cause = "Transit depth limit exceeded"
            cut_list.append(current_row_info + [cause])
        elif not nanignore_airmass and (pd.isna(ingress_air_mass) or pd.isna(egress_air_mass)):
            if ingress_column_exists and egress_column_exists:
                cause = "NaN in ingress/egress air mass"
            else:
                cause = "Ingress/Egress air mass data missing"
            cut_list.append(current_row_info + [cause])
        elif (not pd.isna(ingress_air_mass) and ingress_air_mass > maxingressairmass) or \
             (not pd.isna(egress_air_mass) and egress_air_mass > maxegressairmass):
            cause = "Ingress/Egress air mass limit exceeded"
            cut_list.append(current_row_info + [cause])
        else:
            valid_candidates.append(current_row_info)
            all_schedules.append(current_row_info)

    valid_candidates_df = pd.DataFrame(valid_candidates, columns=[ 
        'Name', 'Duration (hours)', 'Midpoint', 'Transit Start Time', 'Transit End Time', 
        'RA', 'Dec', 'Period', 'Transit Depth', 'Air Mass', 'Magnitude K'
    ])

    valid_candidates_df['Transit Start Time'] = pd.to_datetime(valid_candidates_df['Transit Start Time'], format='%m/%d/%Y %H:%M')
    valid_candidates_df['Transit End Time'] = pd.to_datetime(valid_candidates_df['Transit End Time'], format='%m/%d/%Y %H:%M')
    valid_candidates_df = valid_candidates_df.sort_values(by='Transit End Time').reset_index(drop=True)

    optimized_schedule = []
    last_end_time = None

    for index, row in valid_candidates_df.iterrows():
        # Determine the effective start and end times based on setup_time
        if setup_time:
            schedule_start_time = row['Transit Start Time'] - timedelta(minutes=30)
            schedule_end_time = row['Transit End Time'] + timedelta(minutes=30)
        else:
            schedule_start_time = row['Transit Start Time']
            schedule_end_time = row['Transit End Time']

        # Check for overlap based on the determined schedule times
        if last_end_time is None or schedule_start_time >= last_end_time:
            optimized_schedule.append(row)
            last_end_time = schedule_end_time
        else:
            cut_list.append(row.tolist() + ["Overlapping with another target"])

    optimized_schedule_df = pd.DataFrame(optimized_schedule)
    cut_list_df = pd.DataFrame(cut_list, columns=[ 
        'Name', 'Duration (hours)', 'Midpoint', 'Transit Start Time', 'Transit End Time', 
        'RA', 'Dec', 'Period', 'Transit Depth', 'Air Mass', 'Magnitude K', 'Cause'
    ])
    
    # Calculate the schedule time range
    if not optimized_schedule_df.empty:
        first_target_start = optimized_schedule_df['Transit Start Time'].min()
        last_target_end = optimized_schedule_df['Transit End Time'].max()

        # Calculate adjusted schedule start and end times for each target
        if setup_time:
            optimized_schedule_df['Schedule Start Time'] = optimized_schedule_df['Transit Start Time'] - timedelta(minutes=30)
            optimized_schedule_df['Schedule End Time'] = optimized_schedule_df['Transit End Time'] + timedelta(minutes=30)
        else:
            optimized_schedule_df['Schedule Start Time'] = optimized_schedule_df['Transit Start Time']
            optimized_schedule_df['Schedule End Time'] = optimized_schedule_df['Transit End Time']

    plot_schedule_with_plotly(optimized_schedule_df, setup_time)

    print(f"\nTotal number of valid schedules: {len(valid_candidates_df)}")
    print(f"Maximum number of transits that can fit in a night: {len(optimized_schedule_df)}")

    return optimized_schedule_df, cut_list_df

def count_max_schedules(all_schedules, optimized_schedule_df):
    """
    Count the number of schedules that have the maximum number of non-overlapping transits 
    matching the optimized schedule.

    Parameters:
    all_schedules (list of list): A list of candidate schedules, where each candidate schedule 
                                   is a list of events. Each event is expected to be a list 
                                   containing details of the transit, with the start and end 
                                   times at indices 3 and 4, respectively.
    optimized_schedule_df (pd.DataFrame): A DataFrame containing the optimized schedule, 
                                           with each row representing a transit.

    Returns:
    int: The count of candidate schedules that have the same maximum number of transits 
         as the optimized schedule.
    """
    # Get the maximum number of transits in the optimized schedule
    max_transit_count = len(optimized_schedule_df)
    
    # Track the number of schedules with the same maximum transit count
    max_schedules_count = 0

    # Compare every potential schedule with the optimized schedule based on non-overlapping criteria
    for candidate_schedule in all_schedules:
        # Generate a temporary schedule from the list of candidate transits and check if it would match the optimized schedule
        temp_schedule = []
        last_end_time = None

        for event in candidate_schedule:
            # event[3] = transit start time, event[4] = transit end time
            event_start_time, event_end_time = event[3], event[4]
            if last_end_time is None or event_start_time >= last_end_time:
                temp_schedule.append(event)
                last_end_time = event_end_time

        if len(temp_schedule) == max_transit_count:
            max_schedules_count += 1

    return max_schedules_count



def gen_schedule(file_path, magnitude_limit=(0, 14.5), air_mass_lim=True, save=True, 
                 setup_time=False, period_limit=False, transit_depth_limit=(0, 0.5), 
                 max_airmass=(2, 2), nanignore_airmass=False, time_window=(None, None)):
    """
    Generate an optimized schedule of astronomical transits based on specified criteria from a CSV file. 
    The function also provides options to save the results to an Excel file and prints messages for missing data.

    Parameters:
    file_path (str): The path to the CSV file containing transit data.
    magnitude_limit (tuple): A tuple specifying the lower and upper limits for the magnitude 
                             of the transits (default: (0, 14.5)).
    air_mass_lim (bool): A flag indicating whether to apply air mass limits (default: True).
    save (bool): A flag indicating whether to save the optimized schedule to an Excel file (default: True).
    setup_time (bool): A flag indicating whether to consider setup time in the scheduling (default: False).
    period_limit (bool): A flag indicating whether to apply period limits in the scheduling (default: False).
    transit_depth_limit (tuple): A tuple specifying the lower and upper limits for the transit depth 
                                  (default: (0, 0.5)).
    max_airmass (tuple): A tuple specifying the maximum allowable air mass (default: (2, 2)).
    nanignore_airmass (bool): A flag indicating whether to ignore NaN values in air mass calculations (default: False).
    time_window (tuple): A tuple defining the time window as (start_time, end_time) in UTC. 
                         If either value is None, that limit is ignored.

    Returns:
    None: The function prints the optimized schedule and cut list, and optionally saves them to an Excel file.
    """
    try:
        df = pd.read_csv(file_path, comment='#')
        print("DataFrame loaded successfully.")
        
        # Initialize flags for printing messages
        air_mass_found = True
        egress_air_mass_found = True
        ingress_air_mass_found = True

        # Check for required columns
        if 'midpointairmass' not in df.columns:
            air_mass_found = False

        if 'egressairmass' not in df.columns:
            egress_air_mass_found = False

        if 'ingressairmass' not in df.columns:
            ingress_air_mass_found = False

        # Print messages if columns are not found
        if not air_mass_found:
            print("Mid Point air mass not found in CSV.")
        if not egress_air_mass_found:
            print("Egress air mass not found in CSV.")
        if not ingress_air_mass_found:
            print("Ingress air mass not found in CSV.")

        final_schedule, cut_list = optimize_schedule(df, magnitude_limit, air_mass_lim, 
                                                     setup_time, period_limit, transit_depth_limit,
                                                     max_airmass=max_airmass, nanignore_airmass=nanignore_airmass,
                                                     time_window=time_window)
        
        print("\nOptimized Schedule")
        print(final_schedule)

        print("\nCut List")
        print(cut_list)

        if save:
            output_file = f"optimized_schedule_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                final_schedule.to_excel(writer, sheet_name='Optimized Schedule', index=False)
                cut_list.to_excel(writer, sheet_name='Cut List', index=False)

                # Insert the plot image into the 'Schedule' sheet
                # Create a new sheet for the plot
                workbook = writer.book
                worksheet = workbook.create_sheet(title='Schedule')
                worksheet.append(['Optimized Schedule Plot'])  # Add a title or description

                # Add the image to the worksheet
                img = Image('optimized_exoplanet_observation_schedule.png')
                worksheet.add_image(img, 'A3')  # Adjust the position as needed

            print(f"\nOutput saved to {output_file}")

        # Delete the plot file after saving to Excel

        if os.path.exists('optimized_exoplanet_observation_schedule.png'):
            os.remove('optimized_exoplanet_observation_schedule.png')  # Delete the saved plot image file

    except Exception as e:
        print(f"An error occurred: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
        
# Example usage with a time window
start_time_window = pd.to_datetime("2024-10-09 20:00:00")  # Example start time
end_time_window = pd.to_datetime("2024-10-10 06:00:00")    # Example end time
gen_schedule(r"C:\Users\Zachary\Downloads\transits_1009071212_2024.10.09_07.15.45.csv", 
             magnitude_limit=(0, 14.5), 
             air_mass_lim=True, 
             save=True, 
             setup_time=True, 
             period_limit=False, 
             transit_depth_limit=(0, 0.5), 
             max_airmass=(2, 2), 
             nanignore_airmass=False,
             time_window=(start_time_window, end_time_window))
