"""
    Application User Interface layout design module
"""
# Necessary Imports
import PySimpleGUI as sg
import datetime
from jinja2 import Environment, FileSystemLoader
import os
# Experimental Features (Not ready yet)
    # GTK_Folder = r"C:\Program Files\GTK3-Runtime Win64\bin"
    #os.environ['PATH'] = GTK_Folder + os.pathsep + os.environ.get('PATH', '')
    # from weasyprint import HTML, CSS

# Some Important Variable Declarations
payFixationTypeList = ['Select Type', 'Appointment', 'Non Functional Promotion (MCAS)', 'Functional Promotion',
                       'Confirmation', 'ROPA Fixation', 'Pay Protection', 'Others']  # List of Pay Fixation Types
payLevels = ['Select One', 1, 2, 3, 4, 5, 6, '6A', 7, 8, 9, '9A', 10, '10A', '10B', '10C', 11, 12, '12A', '12B', 13, 14,
             15, '15A', '16A', 16, 17, 18, 19, '19A', 20, 21, 22, 23, 24]  # List of Pay levels under ROPA 2019
incrementTypes = ['Select One', 'Normal Increment', 'Promotional Increment', 'Additional Increment', 'Others']  # List of Increment Types
current_directory = os.getcwd()  # Gets the current working directory
time_stamp = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')  # Generates date time in specific format
sg.theme('TanBlue')  # Application theme
icon = r'assets/new-icon.ico'  # Application icon file
try:
    env = Environment(loader=FileSystemLoader('./assets'))  # Loads assets folder
    report_template1 = env.get_template('ropa19reporttemp.html')  # Loads report format 1
    report_template2 = env.get_template('ropa09reporttemp.html')  # Loads report format 2
except Exception as g:
    sg.popup(f"""The following error is detected :\nTemplate File "{g}" was not found in the main app directory.
\nPlease report this bug to our GitHub.""",
                   title="Error", icon=icon)  # Return errors if template not found. (If any)
    pass

# Start of Layout Functions


def headMenu():
    """ Application Main Menu List

    :return: Menus as a List
    :rtype: list
    """
    menu_def = [
        ['Help and Support   ', [
             'Reload Excel Pay Matrix File',
             '---',
             'About the App',
             '---',
             'Contact Us',
             '---',
             'Fork Me on Github',
             '---',
             'Check for Updates',
             '---',
             'End-User Rights',
             '---',
             'Exit']]]
    return menu_def


def getNextIncrementDate(actual_date):
    """ Function to get the date of next Increment.

    :param actual_date: Date of Actual Promotion
    :type actual_date: str
    :return: First day of July (DD/MM/YYYY) of same year if Actual date is before july or returns the first day of July
    of next year.
    :rtype: str
    """
    date = datetime.datetime.strptime(actual_date, '%d/%m/%Y')
    month = date.month
    year = date.year
    if month < 7:
        return f'01/07/{year}'
    elif month >= 7:
        return f'01/07/{year + 1}'


def header_layout():
    """ Application Header Layout

    :return: Layout as a List
    :rtype: list
    """
    head = [
        [sg.Text("AUTOMATIC PAY FIXATION CALCULATOR", size=(100, 1), font=('arial', 15), justification='center')],
        [sg.Text("Pay Fixation Calculator for Employees of Govt. of WB", size=(50, 1),
                 font=('arial', 10, 'underline'), justification='center')]]
    type_layout = [
        [sg.Frame("", [
            [sg.Text("Select Pay Fixation Type :", size=(25, 1), font=('arial', 11)),
             sg.Combo(values=payFixationTypeList, default_value=payFixationTypeList[0], readonly=True,
                      enable_events=True, auto_size_text=True, key='pay-fix-type')]])]]
    main = [
        [sg.Column(layout=head, expand_x=True, justification='center', element_justification='center')],
        [sg.Column(layout=type_layout, expand_x=True, justification='center', element_justification='center')],
        [sg.Column([[sg.HSeparator()]], expand_x=True)]]
    return main


def fixationDatesLayout():
    """ Pay Fixation Dates Layout

        :return: Layout as a List
        :rtype: list
        """
    layout = [
        [sg.Frame("", [
            [sg.Text("Notional Date of Promotion : ", size=(27, 1)), sg.Input("", size=(12, 1), key='dt-notional'),
             sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", no_titlebar=False,
                               format='%d/%m/%Y', border_width=0, image_size=(20, 20), tooltip='Click Here to select date.')],
            [sg.Text("Actual Date of Promotion : ", size=(27, 1)),
             sg.Input("", size=(12, 1), key='dt-actual', enable_events=True),
             sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", no_titlebar=False,
                               format='%d/%m/%Y', border_width=0, image_size=(20, 20), key='dt-actual1',
                               enable_events=True, tooltip='Click Here to select date.')],
            [sg.Text("What are these dates?", size=(24, 1), key='--k--', text_color='red', enable_events=True,
                     font=('arial', 10, 'underline'), tooltip='Click here to learn more.')]]),
         sg.Frame("", [
             [sg.Text("Notional Date of Pay Fixation : ", size=(25, 1)),
              sg.Input("", size=(14, 1), key='dt-notional-pay-fix', disabled=False, enable_events=True, ),
              sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", no_titlebar=False,
                                format='%d/%m/%Y', border_width=0, image_size=(20, 20), tooltip='Click Here to select date.')],
             [sg.Text("Actual Date of Pay Fixation : ", size=(25, 1)),
              sg.Combo(values=['Select One', 'On the Date of Promotion', 'On the Date of Next Increment'],
                       default_value='Select One', key='option-choose', enable_events=True, readonly=True)],
             [sg.Text(size=(25, 1)), sg.Input("", size=(14, 1), key='dt-actual-pay-fix'),
              sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", no_titlebar=False,
                                format='%d/%m/%Y', border_width=0, image_size=(20, 20), tooltip='Click Here to select date.')]])]]
    return layout


def mCasLayout():
    """ Application Body Layout

        :return: Layout as a List
        :rtype: list
        """
    layout_head = [
        [sg.Frame("Pay Fixation Details", [
            [sg.Text("Select Existing Pay Level * :", size=(23, 1), font=('arial', 10)),
             sg.Combo(values=payLevels, default_value=payLevels[0], readonly=True, enable_events=True, key='old-level'),
             sg.Text(size=(23, 1)),
             sg.Text("Select New Pay Level * :", size=(18, 1), font=('arial', 10)),
             sg.Combo(values=payLevels, default_value=payLevels[0], readonly=True, enable_events=True,
                      key='new-level')],
            [sg.Text("Enter Pay Before Fixation * :", size=(23, 1), font=('arial', 10)),
             sg.Input(0, size=(20, 1), key='old-pay', enable_events=True),
             sg.Text(size=(15, 1)),
             sg.Text("Click Here If You are Fixing Pay From ROPA 2009", size=(45, 1), key='getRopa2009Win',
                     enable_events=True, font=('arial', 10, 'underline'), text_color='red')],
            [sg.Frame("Increment Details", [
                [sg.Column([
                    [sg.Text("1. Increment Date * :", size=(15, 1)), sg.Input("", size=(10, 1), key='ropa19incdate'),
                     sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", no_titlebar=False,
                                       format='%d/%m/%Y', border_width=0, image_size=(20, 20), tooltip='Click Here to select date.'),
                     sg.Text("Type :"), sg.Combo(values=incrementTypes, default_value=incrementTypes[0], readonly=True,
                                                 key='inc1', enable_events=True, size=(20, 1)),
                     sg.Text("Basic After Increment :", size=(17, 1)),
                     sg.Input(0, size=(12, 1), disabled=False, key='payAfterInc', enable_events=True),
                     sg.Text("Add Row", size=(10, 1), text_color='red', font=('arial', 8, 'underline'), key='addRow',
                             enable_events=True, tooltip='Click Here to add more rows.')]])],
                [sg.Column(key='rowAdded', layout=[], expand_y=False, vertical_scroll_only=True, )],
                [sg.Column([
                    [sg.Column([
                        [sg.Checkbox("Do you have any more no. of Increments in the Old Pay Level ?", default=False,
                                     key='extraInc', enable_events=True),
                         sg.Text("No. of Increments : "), sg.Input(0, disabled=True, key='extraIncNo', size=(12, 1)),
                         sg.Button("Save", size=(8, 1), border_width=0, font=('arial', 10), key='saveExtraInc',
                                   disabled=True)]],
                        visible=False, key='extraIncCol', )],
                    [sg.Text("Final Basic Pay in the Old Pay Level : ", size=(30, 1)),
                     sg.Input(0, disabled=True, enable_events=True, key='finalPayOld', size=(15, 1)),
                     sg.Checkbox("Click Here to Put Manually", default=False, size=(20, 1), key='putManual',
                                 enable_events=True)]])]], key='incrementFrame')],
            [sg.Frame("Final Pay Details", [
                [sg.Text("Final Basic Pay Fixed in the New Level : ", size=(30, 1)),
                 sg.Input(0, size=(15, 1), disabled=True, key='final-pay'), sg.Text("", size=(50, 1))]])],
            [sg.Column([
                [sg.Button("Fix Pay", key='cal', size=(20, 1), border_width=0),
                 sg.Button("Get Report", key='report1', size=(20, 1), border_width=0, disabled=True)]
            ], expand_x=True, justification='center', element_justification='center')]])]]
    main = [[sg.Column(layout=layout_head, expand_x=True)]]
    return main


def mCasRopa2009Layout():
    """ Application Body Layout of ROPA 2009

        :return: Layout as a List
        :rtype: list
        """
    layout1 = [
        [sg.Frame("Pay Fixation Details", [
            [sg.Text("Enter Your Old Band Pay * :", size=(22, 1)),
             sg.Input(0, size=(10, 1), key='oldBasicPay', enable_events=True),
             sg.Text("Enter Your Old Grade Pay * :", size=(22, 1)),
             sg.Input(0, size=(10, 1), key='oldGradePay', enable_events=True),
             sg.Text("Your Old Total Basic Pay :", size=(20, 1)),
             sg.Input(0, size=(10, 1), key='oldBandPay', enable_events=True, disabled=True)],
            [sg.Frame("Increment Details", [
                [sg.Column([
                    [sg.Text("1. Increment Date * :", size=(15, 1)), sg.Input("", size=(10, 1), key='incDate1'),
                     sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", no_titlebar=False,
                                       format='%d/%m/%Y', border_width=0, image_size=(20, 20), tooltip='Click Here to select date.'),
                     sg.Text("Type :"), sg.Combo(values=incrementTypes, default_value=incrementTypes[0], readonly=True,
                                                 key='inc09-1', enable_events=True, size=(20, 1)),
                     sg.Text("Basic After Increment :", size=(17, 1)),
                     sg.Input(0, size=(12, 1), disabled=False, key='payAfterInc09', enable_events=True),
                     sg.Text("Add Row", size=(10, 1), text_color='red', font=('arial', 8, 'underline'), key='addRow09',
                             enable_events=True, tooltip='Click Here to add more rows.')]])],
                [sg.Column(key='rowAdded09', layout=[], expand_y=False, vertical_scroll_only=True, )],
                [sg.Column([
                    [sg.Column([
                        [sg.Checkbox("Do you have any more no. of Increments in the Old Pay Level ?", default=False,
                                     key='extraInc09', enable_events=True),
                         sg.Text("No. of Increments : "), sg.Input(0, disabled=True, key='extraIncNo09', size=(12, 1)),
                         sg.Button("Save", size=(8, 1), border_width=0, font=('arial', 10), key='saveExtraInc09',
                                   disabled=True)]],
                        visible=False, key='extraIncCol09', )],
                    [sg.Text("Final Pay in ROPA 2009 : ", size=(30, 1)),
                     sg.Input(0, disabled=True, enable_events=True, key='finalPayOld09', size=(15, 1)),
                     sg.Checkbox("Click Here to Put Manually", default=False, size=(20, 1), key='putManual09',
                                 enable_events=True)],
                    [sg.Text("Select Pay Level Under ROPA 2019 * :", size=(30, 1), font=('arial', 10)),
                     sg.Combo(values=payLevels, default_value=payLevels[0], readonly=True, enable_events=True,
                              key='old-level09'),
                     sg.Text("New Basic Pay Under ROPA 2019 :", size=(30, 1)),
                     sg.Input(0, size=(15, 1), key='-pay19', disabled=True)]])]])],
            [sg.Column([
                [sg.Button("Fix Pay", size=(20, 1), border_width=0, key='ropa09Fix'),
                 sg.Button("Get Report", size=(20, 1), border_width=0, key='09Report', disabled=True),
                 sg.Button("ROPA 2019 Schedule", size=(20, 1), border_width=0, key='19menu')]
            ], expand_x=True, element_justification='center', justification='center')]])]]
    main = [[sg.Column(layout=layout1, expand_x=True)]]
    return main


def aboutWindowLayout():
    """ Application About Us Layout

        :return: Layout as a List
        :rtype: list
    """
    layout_1 = [
        [sg.Frame("", [
            [sg.Text("Automatic Pay Fixation Calculator v1.0.1", font=('arial', 15, 'underline'),
                     justification='center', size=(45, 1))],
            [sg.Text("Thank you for using this application.")],
            [sg.Text(
                "Automatic Pay Fixation Calculator is a Free and Open Source \nApplication which helps the employees "
                "of Govt. of WB to Calculate\ntheir New Basic Pay after any change in the pay level due to any\nPromotion, MCAS"
                " Confirmation, etc (ROPA 2009 and 2019).")],
            [sg.Text("Developed and Written by -"),
             sg.Text("SOURAV", text_color='red', font=('arial', 10, 'underline'), key='author', enable_events=True,
                     tooltip='Click here to learn more.')],
            [sg.Text("Copyright - Open Source Copy Left Licence\nLicence - GNU Public Licence ver. 3.\n"
                     "Programming Language - Python 3.9 ")],
            [sg.Text(
                "For any Feedback/Bug Reporting please write at \nloku-sama@outlook.com or happysourav96@gmail.com")],
            [sg.Text("Click Here to visit our GitHub page for more information", text_color='red',
                     font=('arial', 9, 'underline'), key='git', enable_events=True)],
            [sg.Text("Click Here to view the Licence Information", text_color='red', font=('arial', 9, 'underline'),
                     key='licence', enable_events=True)],
            [sg.Text("Free Software Movement", text_color='red', font=('arial', 9, 'underline'), key='fsm',
                     enable_events=True)]])]]
    main = [
        [sg.Column(layout=layout_1, expand_x=True, element_justification='center', justification='center')],
        [sg.Column([
            [sg.Button("Close", border_width=0, size=(15, 1), key='a_window_close')]], element_justification='center',
            justification='center', expand_x=True)]]
    return main


def helpWindowLayout():
    """ Help Window Layout

    :return: Help Layout as list
    :rtype: list
    """
    layout_1 = [
        [sg.Text("How to Calculate Date of Pay Fixation", font=('arial', 15, 'underline'), text_color='red',
                 justification='center')],
        [sg.Frame("", [
            [sg.Text("Dates of Promotion/CAS :", font=('arial', 11, 'underline'), text_color='red')],
            [sg.Text("Notional Date of Promotion:- Notional Date of Promotion implies the date with effect from which"
                     " the employee is\nentitled to be promoted or to get CAS.")],
            [sg.Text("Actual date of Promotion :- Actual date of Promotion implies the date on which he/she is actually"
                     " awarded\npromotion which may differ from the Notional date due to non fulfilment of conditions "
                     "like passing the\nDepartmental Examination or Extraordinary Leave etc. The employee is generally "
                     "entitled to get arrear\non the basis of revised pay due to such promotion from this date.")],
            [sg.Text("Dates of Pay Fixation :", font=('arial', 11, 'underline'), text_color='red')],
            [sg.Text("Notional date of Pay Fixation :- Notional date of Pay fixation is the date on which the pay of "
                     "the employee is to\nbe revised due to Promotion/CAS/Confirmation as if the employee does not opt "
                     "for any other date of pay fixation.")],
            [sg.Text("Actual Date of Pay Fixation :- Actual Date of Pay Fixation is the date from which the employee "
                     "actually opts\nto get his/her pay revised.")],
            [sg.Text("Example :", font=('arial', 11, 'underline'), text_color='red')],
            [sg.Text("Suppose an employee is appointed on 02/03/2010. His CAS is due on 02/03/2018 but he could not"
                     " clear\ndepartmental examination. He cleared his departmental examination on 15/04/2019. And he "
                     "is given CAS and\nhe chooses to take its effect from July next.\n\nHere are the dates of his/her "
                     "CAS - \n\n\tNotional Date of CAS: 02/03/2018\n\tActual date of CAS: 15/04/2019\n\t"
                     "Notional date of Pay fixation: 15/04/2019\n\tActual Date of Pay fixation: 01/07/2019")],
            [sg.Text("Source :- WBIFMS ", font=('arial', 11, 'underline'), text_color='red')]])]]
    main = [
        [sg.Column(layout=layout_1, expand_x=True, element_justification='center', justification='center')],
        [sg.Column([
            [sg.Button("Close", border_width=0, size=(15, 1), key='h_window_close')]], element_justification='center',
            justification='center', expand_x=True)]]
    return main


def mainAppLayout():
    """ Application Main Layout

        :return: Layout as a List
        :rtype: list
        """
    main_layout = [
        [sg.Menu(headMenu(), tearoff=False, key='head-menu')],
        [sg.Column(layout=header_layout(), justification='center', element_justification='center', pad=(0, 0))],
        [sg.Column(layout=fixationDatesLayout(), expand_x=True, pad=(0, 0))],
        [sg.Column(layout=mCasLayout(), expand_x=True, key='mCas-layout', visible=True, pad=(0, 0)),
         sg.Column(layout=mCasRopa2009Layout(), expand_x=True, key='ropa2009Layout', visible=False, pad=(0, 0))]]
    return main_layout


mainWindow = sg.Window("Automatic Pay Fixation calculator", layout=mainAppLayout(), finalize=True, size=(850, 700),
                       enable_close_attempted_event=True, modal=True, icon=icon)  # Application Main window
