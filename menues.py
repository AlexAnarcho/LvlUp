"""Stores all the functions needed for displaying the menu and the progress"""

import inquirer
import openpyxl
from graphs import *
from essentials import *


def load_menu_options():
    """
    Loading the sheet names for the menu in the selector
    :return:
    """
    workbook_path = get_workbook_path()
    wb = openpyxl.load_workbook(workbook_path)
    sheet_names = wb.sheetnames
    return sheet_names


def skill_menu():
    """
    Make a selection from the skill menu. The skill selected will be loaded from the spreadsheet.
    :return:
    """
    while True:
        menu_options = load_menu_options()
        menu_complete = menu_options[:]
        menu_complete.append('Add new skill')
        questions = [
            inquirer.List('skill',
                          message="What skill are you improving?",
                          choices=menu_complete,
                          carousel=True
                          ),
        ]
        selections = inquirer.prompt(questions)
        if selections['skill'] == menu_complete[-1]:
            entering_active = True
            while entering_active:
                skill_name = input("\nGive this new skill a name: ").title()
                if len(skill_name) > 3:
                    initiate_new_sheet(skill_name)
                    entering_active = False
                else:
                    print('ERROR: The skill name must be longer than three characters.')
        elif selections['skill'] in menu_options:
            # Get the sheet object of Python
            sheet_name = selections['skill']
            break
    return sheet_name


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
