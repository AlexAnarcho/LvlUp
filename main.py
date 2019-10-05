#! /usr/local/bin python3
# Describe your level in python

from menues import *
from processing import *
from graphs import *
import os
from essentials import create_config_file


def main_output(sheet_obj, sheet_name):
    """
    Getting information from the updated spreadsheet and giving output in the terminal
    :return:
    """
    # Getting the information from the just updated xlsx sheet
    total_hours, total_xp_for_lvlup, current_level, next_level, xp_accumulated, xp_delta, daily_hours, flag_levelup = \
        get_progress_sheet(sheet_obj)

    # The messages formatted
    display_length = 80  # the number that changes how long the terminal window is
    str_greeting = 'Level Chart'
    str_ending = 'Keep Going!'
    str_hours_daily = 'Hours invested today:'
    str_hours_total = 'Total hours:'
    str_current_level = 'Current Level: '
    str_1_next_level = f'{str(xp_delta)} more hours required to level up to {str(next_level)}.'
    str_2_next_level = f"until Level {str(next_level)}"
    str_congrats = f'Congratulations! You have reached Level {str(current_level)}!'

    # Centering the strings
    str_displ_daily_hours = f'{str(daily_hours).rjust(display_length - len(str_hours_daily), ".")}'
    str_displ_total_hours = f'{str(total_hours).rjust(display_length - len(str_hours_total), ".")}'
    str_displ_current_level = f'{str_current_level}{current_level}'.center(display_length)
    str_displ_next_level = f'{str_1_next_level}'.center(display_length)
    str_displ_congrats = f'{str_congrats}'.center(display_length)
    str_displ_ending = str_ending.center(display_length, "+")

    # Starting the Output
    print(f'\n{"-" * display_length}')
    print(f'\n{str_greeting.center(display_length, "+")}\n')  # printing the greeting
    print(f'{str_hours_daily}{str_displ_daily_hours}')  # printing the hours invested today
    print(f'{str_hours_total}{str_displ_total_hours}')  # printing the hours invested today
    if flag_levelup:
        print(f'\n\n{str_displ_congrats}')
    else:
        print(f'\n{str_displ_current_level}')
        print(f'\n{str_displ_next_level}')
        print_progressbar(iteration=xp_accumulated, total=total_xp_for_lvlup, suffix=str_2_next_level,
                          length=(display_length - (len(str_2_next_level) + 10)), fill='█')

    # prompt user if they want to see the graph of the progress
    print(f'\n\n{str_displ_ending}\n')  # printing the ending
    print(f'\n{"-" * display_length}')
    print()
    show_progress_graph(sheet_name=sheet_name)


def main_program():
    """
    Processing the data and giving output.
    Main Program.
    :return:
    """
    workbook_path = get_workbook_path()
    sheet_obj, sheet_name = main_processing()
    main_output(sheet_obj, sheet_name)


if __name__ == '__main__':
    main_program()
