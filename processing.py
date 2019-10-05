"""Stores the logic for calculating levels and the processing of information"""

import datetime
import openpyxl
import os
from menues import *
from essentials import *


def input_pomos():
    """
    Asks the user to specify the hours spent, two pomos are equal to one hour
    Only allows ints in the range between 1 and 20
    :return: the number of hours to be registered with the program
    """
    active = True
    daily_pomos = ""
    while active:
        daily_pomos = input('\nEnter the number of completed Pomodoros: ')
        try:  # check if the input is an integer
            if float(daily_pomos) in range(0, 20):
                active = False
        except ValueError:
            print('ERROR: Please enter an integer as the number of Pomodoros.')
    daily_hours = float(daily_pomos) / 2  # two pomodoros are equal to one hour
    return daily_hours


def load_sheet(sheet_name):
    """Load the specified worksheet"""
    workbook_path = get_workbook_path()
    wb = openpyxl.load_workbook(workbook_path)
    sheet_obj = wb[sheet_name]
    return sheet_obj, wb


def lvl_algo(next_level):
    """
    Calculating the xp_needed for the next level
    :return:
    """
    total_xp_needed = (next_level * next_level)
    return total_xp_needed


def checking_lvl(xp_accumulated, total_xp_needed):
    """
    checking to see if the next level has been reached
    :param total_xp_needed: Required    : The amount of XP to LevelUp
    :param xp_accumulated:  Required    : XP already gathered by the player
    :return:
    """
    if xp_accumulated >= total_xp_needed:  # check if total ours satisfy next level
        return True
    else:  # if no new level has been reached
        return False


def get_progress_sheet(sheet_obj):
    """
    gather all the important information out of the spreadsheet
    :return:
    """
    sheet = sheet_obj
    total_hours = sheet['A2'].value
    total_xp_for_lvlup = sheet['A4'].value
    current_level = sheet['B2'].value
    next_level = sheet['B4'].value
    xp_accumulated = sheet['C2'].value
    xp_delta = sheet['C4'].value
    flag_levelup = sheet['B6'].value
    daily_hours = sheet.cell(row=sheet.max_row, column=2).value

    return total_hours, total_xp_for_lvlup, current_level, next_level, xp_accumulated, xp_delta, daily_hours, \
           flag_levelup


def write_daily_hours(daily_hours, sheet_obj, wb_obj):
    """
    Updating the progress sheet with the data from the main_processing function.
    :param daily_hours:     Required:   Writes the daily hours that were input at the starting
    prompt, adds to existing daily hours if availiable
    :param sheet_obj:       Required:   Needs the sheet object to write the hours to
    :param wb_obj:          Required:   the workbook object, needed for saving the progress
    :return:
    """
    # Load the workbook for the progress tracking
    if daily_hours > 0:
        workbook_path = get_workbook_path()
        sheet = sheet_obj
        wb = wb_obj
        # getting the date for today
        datetime_obj = datetime.datetime.now()
        today_formatted = datetime_obj.strftime('%d.%m.%Y')
        # Fill in the progress made, if date already exists, hours are added, else new line with date + hours created
        new_daily_hours_cell = sheet.cell(row=sheet.max_row, column=2)
        if sheet.cell(row=sheet.max_row, column=1).value == today_formatted:
            hours = new_daily_hours_cell.value
            new_daily_hours_cell.value = daily_hours + float(hours)
        else:
            sheet.cell(row=(sheet.max_row + 1), column=1).value = today_formatted
            new_daily_hours_cell = sheet.cell(row=sheet.max_row, column=2)
            new_daily_hours_cell.value = daily_hours
        wb.save(workbook_path)


def write_progress_sheet(total_hours, current_level, next_level, xp_accumulated, total_xp_for_lvlup,
                         xp_delta, flag_levelup, sheet_obj, wb_obj):
    """
    :param total_hours:             Required: Writes the total hours in the specified field
    :param current_level:           Required: Writes current level
    :param next_level:              Required: Writes next level, needed later for the terminal output
    :param xp_accumulated:          Required: Writes the xp gathered in the current level
    :param total_xp_for_lvlup:      Required: Writes the total xp needed to reach the next level
    :param xp_delta:                Required: Writes the delta between total xp and accumulated xp
    :param flag_levelup:            Required: Tells if the player reached a new level
    :param sheet_obj:               Required: The needed sheet to write the stats in
    :param wb_obj:                  Required: Workbook object needed to save progress
    :return:
    """

    # Load the workbook for the progress tracking
    workbook_path = get_workbook_path()
    wb = wb_obj
    sheet = sheet_obj

    # Getting the appropriate cells in variables
    cell_total_hours = sheet['A2']
    cell_total_xp_for_lvlup = sheet['A4']
    cell_current_level = sheet['B2']
    cell_next_level = sheet['B4']
    cell_xp_accumulated = sheet['C2']
    cell_xp_delta = sheet['C4']
    cell_flag_levelup = sheet['B6']

    # Update the other fields
    cell_total_hours.value = total_hours
    cell_current_level.value = current_level
    cell_next_level.value = next_level
    cell_xp_accumulated.value = xp_accumulated
    cell_total_xp_for_lvlup.value = total_xp_for_lvlup
    cell_xp_delta.value = xp_delta
    cell_flag_levelup.value = flag_levelup

    # Save the made progress to the spreadsheet
    wb.save(workbook_path)  # Placeholder sheet


def main_processing():
    """
    Does the bulk of the work with processing input, and updating the spreadsheet
    :return:
    """
    # start the menu
    sheet_name = skill_menu()
    sheet_obj, wb = load_sheet(sheet_name)

    # ask user for pomos when program starts
    daily_hours = input_pomos()

    # Getting the needed values from the xlsx sheet
    total_hours, total_xp_for_lvlup, current_level, next_level, xp_accumulated, xp_delta, hours_, flag_levelup = \
        get_progress_sheet(sheet_obj=sheet_obj)

    # Processing the progress made
    # Adding the daily_hours
    if total_hours is None:
        total_hours = daily_hours
        xp_accumulated = daily_hours
    else:
        total_hours += daily_hours  # add daily hours to total hours
        xp_accumulated += daily_hours  # add the daily hours to the eXP
    next_level = current_level + 1

    # Checking for LvlUp
    total_xp_needed = lvl_algo(next_level)  # calculating xp_needed based on the next level
    if xp_accumulated >= total_xp_needed:
        current_level += 1
        next_level += 1
        xp_accumulated = xp_accumulated - total_xp_needed
        xp_delta = total_xp_needed - xp_accumulated
        total_xp_needed = lvl_algo(next_level)
        flag_levelup = True
    else:
        xp_delta = total_xp_needed - xp_accumulated
        flag_levelup = False
    # Updating the xlsx sheet with the progress made
    write_daily_hours(daily_hours, sheet_obj=sheet_obj, wb_obj=wb)
    write_progress_sheet(total_hours=total_hours, current_level=current_level,
                         next_level=next_level, xp_accumulated=xp_accumulated, total_xp_for_lvlup=total_xp_needed,
                         xp_delta=xp_delta, flag_levelup=flag_levelup, sheet_obj=sheet_obj, wb_obj=wb)
    wb.close()
    return sheet_obj, sheet_name
