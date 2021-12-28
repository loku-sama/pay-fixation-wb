"""
    This module covers the important functions for the application.
"""
# Necessary Imports
import pandas as pd
import app_layout as al
import requests

# Excel pass = 'unlockme' # Password for unlocking excel file
# Information of current version
curr_x = 1
curr_y = 0
curr_z = 1
icon = r'assets/new-icon.ico'  # Application icon file


def check_update():
    """  Function to check the app Updates
    :return: True or False
    :rtype: bool
    """
    try:
        response = requests.get("https://api.github.com/repos/loku-sama/pay-fixation-wb/releases/latest")
        # Getting the latest release version from GitHub API
        github_version = str(response.json()["tag_name"])  # Processing the JSON file
        x = github_version[0:1]  # Slicing the string
        y = github_version[2:3]
        z = github_version[4:5]
        al.sg.popup_no_buttons("Checking for Updates......\n\n", title='Processing', auto_close=True, icon=icon)
        if int(x) > curr_x or int(y) > curr_y or int(z) > curr_z:
            update = al.sg.popup_yes_no("Good News! An Updated version is available.\n"
                                        "Do you want to update now to use the latest features?\n"
                                        "(You can delete the old version after the update)", title="Update Now?",
                                        icon=icon)
            if update == 'Yes':
                al.sg.webbrowser.open(url='https://github.com/loku-sama/pay-fixation-wb/releases/latest', new=2)
            return True
        else:
            al.sg.popup("You already have the latest release. Please check again for updates.", title='Update', icon=icon)
            return False
    except:
        al.sg.popup("Please check your internet connection and try again.", title='Error!', icon=icon)


def downloadExcelFile():
    """ Function for downloading Excel pay matrix file if not found.

    :return: True if succeeded
    :rtype: bool
    """
    url = 'https://github.com/loku-sama/pay-fixation-wb/raw/main/assets/pay%20matrix.xlsx'
    try:
        notFound = al.sg.popup_yes_no(
            "Pay Matrix Excel file not found in the App Directory.\nDo You like to download the Excel "
            "file from our GitHub Page?\n(Please place the excel file in the assets folder"
            " in the main App directory.)", title='Error', icon=icon)
        if notFound == 'Yes':
            r = requests.get(url, allow_redirects=True, stream=True)
            with open(f'./assets/pay matrix.xlsx', 'wb') as f:
                for chunk in r.iter_content(chunk_size=None):
                    al.sg.popup_no_buttons("Your file is Downloading....\n\n", title='Download', auto_close=True, icon=icon)
                    f.write(chunk)
            al.sg.popup("Downloaded. Please Reload the Pay Matrix File from Main Menu or Restart the Application.",
                        title='Success', icon=icon)
            return True
    except:
        al.sg.popup("Something went wrong. Please try later.\n(Please check your internet connection)", title='Error',
                    icon=icon)


def loadPayMatrix():
    """ Load the Excel pay matrix file.

    :return: Panda dataframe
    :rtype: any
    """
    try:
        payMat = pd.read_excel('assets/pay matrix.xlsx', header=0)  # Reads the excel file
        return payMat
    except Exception as e:
        downloadExcelFile()
        return e


payMatrix = loadPayMatrix()  # Loads Pay matrix dataframe


def reLoadPayMatrix():
    """ Reloads the Excel pay matrix file.

    :return: Panda dataframe
    :rtype: any
    """
    try:
        payMat = pd.read_excel('assets/pay matrix.xlsx', header=0)
        al.sg.popup("Excel Pay Matrix File Reloaded Successfully.", title='Success', icon=icon)
        return True, payMat
    except Exception as e:
        return False, e


def getPayLvl(payLvl):
    """ Function for returning the Pay Level entered by User

    :param payLvl: Pay Level entered ny the User
    :type payLvl: str
    :return: Pay Level in String or integer format as applicable
    :rtype: str or int
    """
    try:
        return int(payLvl)
    except:
        return str(payLvl)


# def getPayList(payLevel):  # (Not ready Yet)
#     """ Function for getting a list of Pay Levels from Excel file
#
#     :param payLevel:
#     :type payLevel:
#     :return:
#     :rtype:
#     """
#     pay_list = payMatrix[payLevel].iloc[4:]
#     return list(pay_list)


def getPayRowNo(payLvl, curBasic):
    """ Function for returning row no from excel sheet of the current pay level

    :param payLvl: Pay Level entered ny the User
    :type payLvl: str
    :param curBasic: Current Basic Pay entered ny the User
    :type curBasic: int
    :return: Row no. from Excel sheet
    :rtype: int
    """
    try:
        rowNo = payMatrix[payMatrix[payLvl] == curBasic].index[0]
        return rowNo
    except:
        al.sg.popup(f"Basic {curBasic} was not found in Pay Level {payLvl}.", title='Error', icon=icon)
        return None


def getNormalInc(currRowNo, payLevel, noOfInc):
    """ Function for Calculating Increments

    :param currRowNo: Row no. from the excel sheet of the Current Basic Pay.
    :type currRowNo: int
    :param payLevel: Pay Level entered ny the User
    :type payLevel: str or int
    :param noOfInc: No. of Increments.
    :type noOfInc: int
    :return: New Basic and corresponding Row no.
    :rtype: int or str
    """
    try:
        newRowNo = currRowNo + int(noOfInc)
        newBasic = payMatrix[payLevel][newRowNo]
        return round(newBasic), newRowNo
    except:
        return round(0), currRowNo
