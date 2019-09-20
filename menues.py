"""Stores all the functions needed for displaying the menu and the progress"""

import inquirer
import openpyxl
from graphs import *

workbook_path = '/Users/alex/Documents/LevelUp Progress/lvlupProgress.xlsx'


def load_menu_options():
    """
    Loading the sheet names for the menu in the selector
    :return:
    """
    wb = openpyxl.load_workbook(workbook_path)
    sheet_names = wb.sheetnames
    return sheet_names


def skill_menu():
    """
    Make a selection from the skill menu. The skill selected will be loaded from the spreadsheet.
    :return:
    """
    # skill_name = ""
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


def initiate_new_sheet(sheet_name):
    """
    A function that kicks in when the user wants to add a new skill. Then a new Worksheet will be initiated.
    :return:
    """
    wb = openpyxl.load_workbook(workbook_path)
    wb.create_sheet(sheet_name)
    sheet_obj = wb.get_sheet_by_name(sheet_name)
    sheet_obj['A1'] = 'Date'
    sheet_obj['B1'] = 'Hours'
    sheet_obj['C1'] = 'Total Hours'
    sheet_obj['D1'] = 'Current Level'
    sheet_obj['E1'] = 'XP Accumulated'
    sheet_obj['C3'] = 'Total XP for LevelUp'
    sheet_obj['D3'] = 'Next Level'
    sheet_obj['E3'] = 'XP needed for LevelUp'
    sheet_obj['D5'] = 'LevelUp Flag'

    sheet_obj['C2'] = int(0)
    sheet_obj['D2'] = 0
    sheet_obj['E2'] = 0
    sheet_obj['C4'] = 0
    sheet_obj['D4'] = sheet_obj.cell(row=2, column=4).value + 1
    sheet_obj['E4'] = 0
    sheet_obj['D6'] = False
    wb.save(workbook_path)


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
                      default='No',
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
