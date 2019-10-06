"""Stores all the functions needed for displaying the menu and the progress"""

import os
import inquirer
from processing import *
import matplotlib.pyplot as plt

from matplotlib import dates as mpl_dates

plt.style.use('seaborn-deep')


def print_progressbar(iteration, total, prefix='', suffix='', decimals=1, length=100, fill='█'):
    """
    Call in a loop to create terminal progress bar
    @params:
        iteration   - Required  : current iteration (Int)
        total       - Required  : total iterations (Int)
        prefix      - Optional  : prefix string (Str)
        suffix      - Optional  : suffix string (Str)
        decimals    - Optional  : positive number of decimals in percent complete (Int)
        length      - Optional  : character length of bar (Int)
        fill        - Optional  : bar fill character (Str)
    """
    percentage = (iteration / float(total) * 100)
    percent = ("{0:." + str(decimals) + "f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    print('\r%s |%s| %s%% %s' % (prefix, bar, percent, suffix), end='\r')
    # Print New Line on Complete
    if iteration == total:
        print()


def show_progress_graph(sheet_name):
    """Ask the user if they want to see the progress graphs"""
    question = {
        inquirer.List('show_progress_graph',
                      message='Which graph do you want to display ',
                      choices=['Progess over time', 'Hours this week', "Don't show any graph"],
                      default="Don't show any graph",
                      ),
    }
    answer = inquirer.prompt(question)
    choice = answer['show_progress_graph']
    if choice == 'Progess over time':
        plot_progress_over_all_time(sheet_name)
    elif choice == "Hours this week":
        plot_hours_per_day(sheet_name)
    elif choice == "Don't show any graph":
        pass


def plot_hours_per_day(sheet_name):
    """Plots the hours invested per day"""
    date_strings, hours = read_hours_from_worksheet(sheet_name=sheet_name)
    date_objs = convert_datetime_objs(date_strings)

    plt.plot_date(date_objs, hours, linestyle='solid')

    plt.gcf().autofmt_xdate()

    date_format = mpl_dates.DateFormatter('%d.%m')
    plt.gca().xaxis.set_major_formatter(date_format)

    plt.ylabel('Hours invested')
    plt.title('Invested hours per day')

    plt.grid(True)
    plt.tight_layout()

    plt.show()


def plot_progress_over_all_time(sheet_name):
    """Plots the entire progress over the time of the game."""
    total_hours = read_total_hours_from_worksheet(sheet_name)
    date_strings, hours = read_hours_from_worksheet(sheet_name)

    starting_hours = 0
    # for the case that the user has hours before the logging with the program
    for hour in hours:
        if hour is None:
            continue
        else:
            starting_hours += hour
    starting_hours = total_hours - starting_hours

    hours_progressing = []
    for hour in hours:
        if hour is None:
            continue
        else:
            starting_hours += hour
            hours_progressing.append(starting_hours)

    date_objs = convert_datetime_objs(date_strings)

    plt.plot_date(date_objs, hours_progressing, linestyle='solid')

    plt.gcf().autofmt_xdate()

    date_format = mpl_dates.DateFormatter('%d.%m')
    plt.gca().xaxis.set_major_formatter(date_format)

    plt.ylabel('Total hours invested')
    plt.title(f'Progress in {sheet_name}')
    plt.grid(True)
    plt.tight_layout()

    plt.show()
