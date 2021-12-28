"""
 #############################################   READ ME   ############################################
 # A simple Pay Fixation calculator app written in Python and for GUI I used PySimpleGUI module       #
 # This application is open source under GNU License v.3.                                             #
 # For any feedback or bug reporting, contact me at  - loku-sama@outlook.com                          #
 # Author : SOURAV, A newbie coder.                                                                   #
 # Dependencies to run the code : 1. Python3.9                                                        #
 #                                2. PySimpleGUI module to generate the User Interface                #
 #                                3. Pandas module to generate excel dataframes                       #
 #                                4. Default datetime, requests, math, os module                      #
 #                                5. jinja2 module to generate HTML reports                           #
 ######################################################################################################
"""
# Necessary Imports
import app_layout as al
import excelFunctionsMain
import math

i = 0  # For adding Increment Rows in first window (ROPA 2019)
k = 0  # For adding Increment Rows in second window (ROPA 2009)
increments = []  # For appending increments of ROPA 2019
increments_ropa_2009 = []  # For appending increments of ROPA 2009
icon = r'assets/new-icon.ico'  # Application icon file

# Start of Main loop
while True:
    event, values = al.mainWindow.read()  # Reads all events and values

    if event in [al.sg.WIN_CLOSE_ATTEMPTED_EVENT, 'Exit']:
        """ Closes main window when user tries to exit the application """
        confirm = al.sg.popup_yes_no("Do You Want to Close the Application?", title="Confirm Exit", icon=icon)
        if confirm == "Yes":
            break

    if event == 'Reload Excel Pay Matrix File':
        """ Reloads the excel pay matrix file if found """
        if excelFunctionsMain.reLoadPayMatrix():
            excelFunctionsMain.payMatrix = excelFunctionsMain.loadPayMatrix()

    if event == "Contact Us":
        """ Shows Contact Us Window """
        al.sg.popup("For any problem/bug reporting/help, please email me at - loku-sama@outlook.com \n"
                    "You can also visit my website www.lokusden.neocities.org for more information.",
                    title="Contact Us", custom_text="Thank You", icon=icon)

    if event == 'End-User Rights':
        """ Shows End-User Rights Window """
        al.sg.popup("This application is made available to you under the terms of the GNU General Public License ver. "
                    "3.\nThis means you may use, copy and distribute this application to others.\nYou are also welcome to"
                    " modify the source code of this application as you want to meet your needs.\nThe GNU General Public"
                    " License also gives you the right to distribute your modified versions.",
                    title="About Your Rights", custom_text="Close", icon=icon)

    if event == 'Check for Updates':
        excelFunctionsMain.check_update()

    if event == 'Fork Me on Github':
        al.sg.webbrowser.open(url='https://github.com/loku-sama/pay-fixation-wb', new=2, )

    if event in ['dt-actual1', 'dt-actual']:
        al.mainWindow['dt-notional-pay-fix'].update(values['dt-actual'])

    if event == 'option-choose':
        if values['option-choose'] == 'On the Date of Promotion':
            al.mainWindow['dt-actual-pay-fix'].update(values['dt-notional-pay-fix'])
        elif values['option-choose'] == 'On the Date of Next Increment':
            try:
                optionDate = al.getNextIncrementDate(values['dt-notional-pay-fix'])
                al.mainWindow['dt-actual-pay-fix'].update(optionDate)
            except:
                al.mainWindow['dt-actual-pay-fix'].update("")
        else:
            al.mainWindow['dt-actual-pay-fix'].update("")

    if event == 'addRow' and i == 0:
        """ Adds more increment rows """
        if i < 5:
            if values['ropa19incdate'] == "" or int(values['payAfterInc']) <= 0 or values['inc1'] == "Select One":
                al.sg.popup("Please Fill Up Details in the Previous Row to Add more Rows.", title="Error", icon=icon)
            else:
                al.mainWindow.extend_layout(al.mainWindow['rowAdded'], [[al.sg.Column([
                    [al.sg.Text(f"{i + 2}. Increment Date :", size=(15, 1)),
                     al.sg.Input("", size=(10, 1), key=f'incrementDate{i}'),
                     al.sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", image_size=(20, 20),
                                          no_titlebar=False, format='%d/%m/%Y', key=f'calBtn{i}', border_width=0,
                                          tooltip='Click Here to select date.'),
                     al.sg.Text("Type :"),
                     al.sg.Combo(values=al.incrementTypes, default_value=al.incrementTypes[0],
                                 readonly=True, key=f'incrementType{i}', enable_events=True, size=(20, 1)),
                     al.sg.Text("Basic After Increment :", size=(17, 1)),
                     al.sg.Input(0, size=(12, 1), disabled=False, key=f'payInc{i}', enable_events=True),
                     ]], key=f'addedCol{i}', expand_x=True, pad=(0, 0)), ]])
                i += 1
        else:
            al.sg.popup("Can not insert more rows. Please Click the checkbox to enter more Increments.", title="Opps!",
                        icon=icon)
            # al.mainWindow['extraIncCol'].update(visible=True)

    elif event == 'addRow' and i >= 1:
        """ Adds more increment rows """
        if i < 4:
            if values[f'incrementDate{i - 1}'] == "" or int(values[f'payInc{i - 1}']) <= 0 or \
                    values[f'incrementType{i - 1}'] == "Select One":
                al.sg.popup("Please Fill Up Details in the Previous Row to Add more Rows.", title="Error", icon=icon)
            else:
                try:
                    al.mainWindow.extend_layout(al.mainWindow['rowAdded'], [[al.sg.Column([
                        [al.sg.Text(f"{i + 2}. Increment Date :", size=(15, 1)),
                         al.sg.Input("", size=(10, 1), key=f'incrementDate{i}'),
                         al.sg.CalendarButton(image_filename=r'assets/calender.png', button_text="",
                                              image_size=(20, 20), tooltip='Click Here to select date.',
                                              no_titlebar=False, format='%d/%m/%Y', key=f'calBtn{i}', border_width=0),
                         al.sg.Text("Type :"),
                         al.sg.Combo(values=al.incrementTypes, default_value=al.incrementTypes[0],
                                     readonly=True, key=f'incrementType{i}', enable_events=True, size=(20, 1)),
                         al.sg.Text("Basic After Increment :", size=(17, 1)),
                         al.sg.Input(0, size=(12, 1), disabled=False, key=f'payInc{i}', enable_events=True),
                         ]], key=f'addedCol{i}', expand_x=True, pad=(0, 0)), ]])
                    i += 1
                except:
                    pass
        else:
            al.sg.popup("Can not insert more rows. Please Click the checkbox to enter more Increments.", title="Opps!",
                        icon=icon)
            al.mainWindow['extraIncCol'].update(visible=True)

    if event == 'old-pay' and int(values['payAfterInc']) <= 0:
        al.mainWindow[f'finalPayOld'].update(values['old-pay'])

    if event == 'inc1':
        if values['inc1'] != 'Select One':
            oldPayLvl = excelFunctionsMain.getPayLvl(values['old-level'])
            currRowNo = excelFunctionsMain.getPayRowNo(oldPayLvl, int(values['old-pay']))
            payAfterInc = excelFunctionsMain.getNormalInc(currRowNo, oldPayLvl, 1)
            al.mainWindow['payAfterInc'].update(payAfterInc[0])
            al.mainWindow['finalPayOld'].update(payAfterInc[0])
        else:
            al.mainWindow['payAfterInc'].update(0)
            try:
                for j in range(i):
                    al.mainWindow[f'payInc{j}'].update(0)
                    al.mainWindow[f'incrementType{j}'].update('Select One')
                al.mainWindow['finalPayOld'].update(values['payAfterInc'])
                al.mainWindow['finalPayOld'].update(values['old-pay'])
            except Exception as e:
                al.sg.popup(f"The following error is detected -\n{e}\nPlease report the bug.", title='Error', icon=icon)

    if event == f'incrementType0':
        if values['incrementType0'] != 'Select One':
            oldPayLvl = excelFunctionsMain.getPayLvl(values['old-level'])
            currRowNo = excelFunctionsMain.getPayRowNo(oldPayLvl, int(values['payAfterInc']))
            payAfterInc = excelFunctionsMain.getNormalInc(currRowNo, oldPayLvl, 1)
            al.mainWindow['payInc0'].update(payAfterInc[0])
            al.mainWindow['finalPayOld'].update(payAfterInc[0])
        else:
            try:
                al.mainWindow['payInc0'].update(0)
                for j in range(i - 1):
                    al.mainWindow[f'payInc{j + 1}'].update(0)
                    al.mainWindow[f'incrementType{j + 1}'].update('Select One')
                al.mainWindow['finalPayOld'].update(values['payAfterInc'])
            except:
                al.mainWindow['payInc0'].update(0)
                al.mainWindow['finalPayOld'].update(values['payAfterInc'])

    for m in range(i):
        if event == f'incrementType{m + 1}':
            if values[f'incrementType{m + 1}'] != 'Select One':
                oldPayLvl = excelFunctionsMain.getPayLvl(values['old-level'])
                currRowNo = excelFunctionsMain.getPayRowNo(oldPayLvl, int(values[f'payInc{m}']))
                payAfterInc = excelFunctionsMain.getNormalInc(currRowNo, oldPayLvl, 1)
                al.mainWindow[f'payInc{m + 1}'].update(payAfterInc[0])
                al.mainWindow[f'finalPayOld'].update(payAfterInc[0])
            elif values[f'incrementType{m + 1}'] == 'Select One':
                # al.mainWindow[f'payInc{m}'].update(0)
                for j in range(i - 2):
                    al.mainWindow[f'payInc{j + 2}'].update(0)
                    al.mainWindow[f'incrementType{j + 2}'].update('Select One')
                al.mainWindow[f'payInc{m + 1}'].update(0)
                al.mainWindow['finalPayOld'].update(values[f'payInc{m}'])

    if event == 'incrementType3':
        if values['incrementType3'] == "Select One":
            al.mainWindow['extraInc'].update(disabled=True)
            al.mainWindow['extraInc'].update(False)
            al.mainWindow['extraIncNo'].update(disabled=True)
            al.mainWindow['extraIncNo'].update(0)
            al.mainWindow['saveExtraInc'].update(disabled=True)
            increments.pop()
        else:
            al.mainWindow['extraInc'].update(disabled=False)

    if event == 'extraInc':
        if not values['extraInc']:
            try:
                al.mainWindow['finalPayOld'].update(values['payInc3'])
            except:
                pass

    if event == 'putManual':
        if values['putManual']:
            al.mainWindow[f'finalPayOld'].update(disabled=False)
        else:
            al.mainWindow[f'finalPayOld'].update(disabled=True)

    if event == 'cal':
        if values['old-level'] == 'Select One':
            al.sg.popup("Please select Existing Pay Level from the Dropdown.", title='Error', icon=icon)
            al.mainWindow['report1'].update(disabled=True)
        elif values['new-level'] == 'Select One':
            al.sg.popup("Please select New Pay Level from the Dropdown.", title='Error', icon=icon)
            al.mainWindow['report1'].update(disabled=True)
        elif not values['old-pay'].isdigit() or int(values['old-pay']) <= 0:
            al.sg.popup("Please Enter Your Basic Pay Before Fixation Correctly.", title='Error', icon=icon)
            al.mainWindow['report1'].update(disabled=True)
        elif excelFunctionsMain.getPayRowNo(excelFunctionsMain.getPayLvl(values['old-level']),
                                            int(values['old-pay'])) is not None:
            # b = \
            #     excelFunctionsMain.payMatrix[
            #         excelFunctionsMain.payMatrix[excelFunctionsMain.getPayLvl(values['new-level'])]
            #             .gt(int(values['finalPayOld']))].index[0]
            # Alternate method
            b = \
                excelFunctionsMain.payMatrix[
                    excelFunctionsMain.payMatrix[excelFunctionsMain.getPayLvl(values['new-level'])]
                    >= int(values['finalPayOld'])].index[0]
            extraInc = al.sg.popup_yes_no("Do you allowed any Increments in the New Pay Level?", title="Confirm",
                                          icon=icon)
            fixed_basic = round(excelFunctionsMain.payMatrix[excelFunctionsMain.getPayLvl(values['new-level'])][b])
            if extraInc == 'Yes':
                noOfInc = al.sg.popup_get_text("Enter no. of Increments : ", title='Increments Number', icon=icon)
                if noOfInc is None:
                    al.mainWindow['final-pay'].update(fixed_basic)
                elif not noOfInc.isdigit():
                    al.sg.popup("Please Enter Valid no. of Increments.", title="Error", icon=icon)
                    al.mainWindow['final-pay'].update(0)
                elif int(noOfInc) >= 1:
                    newPayInNextLvl = excelFunctionsMain.getNormalInc(b,
                                                                      excelFunctionsMain.getPayLvl(values['new-level']),
                                                                      int(noOfInc))
                    al.mainWindow['final-pay'].update(newPayInNextLvl[0])
                else:
                    al.mainWindow['final-pay'].update(fixed_basic)
            else:
                al.mainWindow['final-pay'].update(fixed_basic)
                noOfInc = 0
            al.mainWindow['report1'].update(disabled=False)

    if event == 'extraInc':
        if values['extraInc']:
            al.mainWindow['extraIncNo'].update(disabled=False)
            al.mainWindow['saveExtraInc'].update(disabled=False)
        else:
            al.mainWindow['extraIncNo'].update(0)
            al.mainWindow['extraIncNo'].update(disabled=True)
            al.mainWindow['saveExtraInc'].update(disabled=True)

    if event == 'saveExtraInc':
        if not values['extraIncNo'].isdigit():
            al.sg.popup("Please Enter no. of Increments Correctly.", title='Error', icon=icon)
        elif int(values['extraIncNo']) <= 0:
            try:
                al.mainWindow['finalPayOld'].update(values['payInc3'])
            except:
                pass
        elif int(values['extraIncNo']) >= 1:
            oldPayLvl = excelFunctionsMain.getPayLvl(values['old-level'])
            # currRowNo = excelFunctionsMain.getPayRowNo(oldPayLvl, int(values['finalPayOld']))
            currRowNo = excelFunctionsMain.getPayRowNo(oldPayLvl, int(values['payInc3']))
            payAfterInc = excelFunctionsMain.getNormalInc(currRowNo, oldPayLvl, int(values['extraIncNo']))
            al.mainWindow[f'finalPayOld'].update(payAfterInc[0])
        else:
            pass


    def generateReport19():
        """ Function for generating html report for ROPA 2019

        :return: True or False
        :rtype: bool
        """
        try:
            try:
                inc_new_level = noOfInc
            except:
                inc_new_level = 0
            html_out = al.report_template1.render(items=increments, fixation_type=values['pay-fix-type'],
                                                  notional_p=values['dt-notional'], acctual_p=values['dt-actual'],
                                                  notional_pf=values['dt-notional-pay-fix'],
                                                  acctual_pf=values['dt-actual-pay-fix'],
                                                  old_pay_level=values['old-level'], new_pay_level=values['new-level'],
                                                  old_pay=values['old-pay'], add_in_old_lvl=values['extraIncNo'],
                                                  old_final_pay=values['finalPayOld'], inc_new_level=inc_new_level,
                                                  new_final_pay=values['final-pay'], )
            file_name = f'Pay Fixation report ROPA 2019 {al.time_stamp}.html'
            with open(f"./reports/{file_name}", "w") as f:
                f.write(html_out)
            f.close()
            al.sg.webbrowser.open(url=f"{al.current_directory}/reports/{file_name}", new=2)
            return True
        except Exception as d:
            al.sg.popup_error(f"The following error is detected :\n{d}\nPlease report this bug to our GitHub.",
                              title="Error", icon=icon)
            return False


    if event == 'report1':
        try:
            increments.clear()  # Clear the Increments list
            increments.append([values['ropa19incdate'], values['inc1'], values['payAfterInc']])
            for f in range(i):
                increments.append([values[f'incrementDate{f}'], values[f'incrementType{f}'], values[f'payInc{f}']])
            generateReport19()
        except Exception as e:
            al.sg.popup_error(f"The following error is detected :\n{e}\nPlease report this bug to our GitHub.",
                              title="Error", icon=icon)

    ######################### ROPA 2009 Portion Start #############################
    if event == 'getRopa2009Win':
        al.mainWindow['mCas-layout'].update(visible=False)
        al.mainWindow['ropa2009Layout'].update(visible=True)

    if event == '19menu':
        al.mainWindow['mCas-layout'].update(visible=True)
        al.mainWindow['ropa2009Layout'].update(visible=False)

    if event in ['oldBasicPay', 'oldGradePay']:
        try:
            totalBandPayOld = int(values['oldBasicPay']) + int(values['oldGradePay'])
            al.mainWindow['oldBandPay'].update(totalBandPayOld)
            al.mainWindow['finalPayOld09'].update(totalBandPayOld)
        except:
            al.sg.popup("Please Enter Valid Details.", title='Error')
            al.mainWindow['oldBandPay'].update(0)

    if event == 'addRow09' and k == 0:
        if k < 5:
            if values['incDate1'] == "" or int(values['payAfterInc09']) <= 0 or values['inc09-1'] == 'Select One' or \
                    int(values['oldBandPay']) <= 0:
                al.sg.popup("Please Fill Up Details in Pay field or add details to the Previous Row to Add more Rows.",
                            title="Error", icon=icon)
            else:
                al.mainWindow.extend_layout(al.mainWindow['rowAdded09'], [[al.sg.Column([
                    [al.sg.Text(f"{k + 2}. Increment Date :", size=(15, 1)),
                     al.sg.Input("", size=(10, 1), key=f'incrementDate{k}09'),
                     al.sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", image_size=(20, 20),
                                          no_titlebar=False, format='%d/%m/%Y', key=f'calBtn{k}09', border_width=0,
                                          tooltip='Click Here to select date.'),
                     al.sg.Text("Type :"),
                     al.sg.Combo(values=al.incrementTypes, default_value=al.incrementTypes[0],
                                 readonly=True, key=f'incrementType{k}09', enable_events=True, size=(20, 1)),
                     al.sg.Text("Basic After Increment :", size=(17, 1)),
                     al.sg.Input(0, size=(12, 1), disabled=False, key=f'payInc{k}09', enable_events=True),
                     ]], key=f'addedCol{k}09', expand_x=True, pad=(0, 0)), ]])
                k += 1
        else:
            al.sg.popup("Can not insert more rows. Please Click the checkbox to enter more Increments.", title="Error",
                        icon=icon)
            al.mainWindow['extraIncCol09'].update(visible=True)

    if event == 'addRow09' and k >= 1:
        if k < 4:
            try:
                if values[f'incrementDate{k - 1}09'] == "" or int(values[f'payInc{k - 1}09']) <= 0 or values[
                    f'incrementType{k - 1}09'] \
                        == 'Select One':
                    al.sg.popup("Please Fill Up Details in the Previous Row to Add more Rows.", title="Error",
                                icon=icon)
                else:
                    al.mainWindow.extend_layout(al.mainWindow['rowAdded09'], [[al.sg.Column([
                        [al.sg.Text(f"{k + 2}. Increment Date :", size=(15, 1)),
                         al.sg.Input("", size=(10, 1), key=f'incrementDate{k}09'),
                         al.sg.CalendarButton(image_filename=r'assets/calender.png', button_text="", image_size=(20, 20)
                                              , no_titlebar=False, format='%d/%m/%Y', key=f'calBtn{k}09',
                                              border_width=0, tooltip='Click Here to select date.'),
                         al.sg.Text("Type :"),
                         al.sg.Combo(values=al.incrementTypes, default_value=al.incrementTypes[0],
                                     readonly=True, key=f'incrementType{k}09', enable_events=True, size=(20, 1)),
                         al.sg.Text("Basic After Increment :", size=(17, 1)),
                         al.sg.Input(0, size=(12, 1), disabled=False, key=f'payInc{k}09', enable_events=True),
                         ]], key=f'addedCol{k}09', expand_x=True, pad=(0, 0)), ]])
                    k += 1
            except:
                pass
        else:
            al.sg.popup("Can not insert more rows. Please Click the checkbox to enter more Increments.", title="Error",
                        icon=icon)
            al.mainWindow['extraIncCol09'].update(visible=True)

    if event == 'inc09-1':
        if values['inc09-1'] != 'Select One' and int(values['oldBandPay']) > 0 and int(values['oldGradePay']) > 0:
            payAfterInc09 = math.ceil(((int(values['oldBandPay']) * 103 / 100) / 10.0)) * 10
            al.mainWindow['payAfterInc09'].update(int(payAfterInc09))
            al.mainWindow['finalPayOld09'].update(int(payAfterInc09))
        elif int(values['oldBandPay']) <= 0 and int(values['oldGradePay']) <= 0:
            al.sg.popup("Please Enter Pay Details Properly.", title="Error", icon=icon)
            al.mainWindow['payAfterInc09'].update(0)
            al.mainWindow['finalPayOld09'].update(values['oldBandPay'])
        else:
            al.mainWindow['payAfterInc09'].update(0)
            al.mainWindow['finalPayOld09'].update(values['oldBandPay'])
            for h in range(k):
                al.mainWindow[f'incrementType{h}09'].update("Select One")
                al.mainWindow[f'payInc{h}09'].update(0)

    if event == 'incrementType009':
        if values['incrementType009'] != 'Select One':
            payAfterInc09 = math.ceil(((int(values['payAfterInc09']) * 103 / 100) / 10.0)) * 10
            al.mainWindow['payInc009'].update(int(payAfterInc09))
            al.mainWindow['finalPayOld09'].update(int(payAfterInc09))
        else:
            al.mainWindow['payInc009'].update(0)
            al.mainWindow['finalPayOld09'].update(
                math.ceil(((int(values[f'oldBandPay']) * 103 / 100) / 10.0)) * 10)
            for h in range(k - 1):
                al.mainWindow[f'incrementType{h + 1}09'].update("Select One")
                al.mainWindow[f'payInc{h + 1}09'].update(0)

    for r in range(k):
        if event == f'incrementType{r + 1}09':
            payAfterInc09 = math.ceil(((int(values[f'payInc{r}09']) * 103 / 100) / 10.0)) * 10
            if values[f'incrementType{r + 1}09'] != 'Select One':
                al.mainWindow[f'payInc{r + 1}09'].update(int(payAfterInc09))
                al.mainWindow['finalPayOld09'].update(int(payAfterInc09))
            else:
                if r == 0:
                    al.mainWindow[f'payInc109'].update(0)
                    al.mainWindow['finalPayOld09'].update(
                        math.ceil(((int(values['payAfterInc09']) * 103 / 100) / 10.0)) * 10)  # Round off amount to next 10
                else:
                    al.mainWindow[f'payInc{r + 1}09'].update(0)
                    al.mainWindow['finalPayOld09'].update(
                        math.ceil(((int(values[f'payInc{r - 1}09']) * 103 / 100) / 10.0)) * 10)
                for h in range(k - 2):
                    al.mainWindow[f'incrementType{h + 2}09'].update("Select One")
                    al.mainWindow[f'payInc{h + 2}09'].update(0)

    if event == 'incrementType309':
        if values['incrementType309'] == 'Select One':
            al.mainWindow['extraIncNo09'].update(disabled=True)
            al.mainWindow['extraInc09'].update(disabled=True)
            al.mainWindow['extraInc09'].update(False)
            al.mainWindow['extraIncNo09'].update(0)
            al.mainWindow['saveExtraInc09'].update(disabled=True)
            al.mainWindow['finalPayOld09'].update(values['payInc209'])

    if event == 'extraInc09':
        if values['extraInc09']:
            al.mainWindow['extraIncNo09'].update(disabled=False)
            al.mainWindow['saveExtraInc09'].update(disabled=False)
        else:
            al.mainWindow['extraIncNo09'].update(disabled=True)
            al.mainWindow['extraIncNo09'].update(0)
            al.mainWindow['saveExtraInc09'].update(disabled=True)
            al.mainWindow['finalPayOld09'].update(values['payInc309'])


    def getROPA2009Inc(pay, inc):
        """ Function to increase the basic with 3% and the roundup to nearest 10

        :param pay: Basic pay before increment
        :type pay: int
        :param inc: no. of increments
        :type inc: int
        :return: Basic pay after increment
        :rtype: int
        """
        for z in range(inc):
            x = math.ceil(((pay * 103 / 100) / 10.0)) * 10
            pay = x
        return pay


    if event == 'saveExtraInc09':
        if not values['extraIncNo09'].isdigit():
            al.sg.popup("Please Enter no. of Increments Correctly.", title='Error', icon=icon)
        else:
            if k == 0:
                finalBasic09 = getROPA2009Inc(int(values['payAfterInc09']), int(values['extraIncNo09']))
                al.mainWindow['finalPayOld09'].update(finalBasic09)
            else:
                for g in range(k):
                    if int(values[f'payInc{g}09']) > 0:
                        finalBasic09 = getROPA2009Inc(int(values[f'payInc{g}09']), int(values['extraIncNo09']))
                        al.mainWindow['finalPayOld09'].update(finalBasic09)

    if event == 'putManual09':
        if values['putManual09']:
            al.mainWindow['finalPayOld09'].update(disabled=False)
        else:
            al.mainWindow['finalPayOld09'].update(disabled=True)

    if event == 'ropa09Fix':
        try:
            if values['old-level09'] == 'Select One':
                al.sg.popup("Please select a Pay Level to Fix the Pay.", title='Error', icon=icon)
                al.mainWindow['09Report'].update(disabled=True)
            elif int(values['oldGradePay']) <= 0:
                al.sg.popup("Please enter Pay Details Properly.", title='Error', icon=icon)
                al.mainWindow['09Report'].update(disabled=True)
            else:
                # newPayInRopa2019 = \
                # excelFunctionsMain.payMatrix[excelFunctionsMain.payMatrix[excelFunctionsMain.getPayLvl(
                #     values['old-level09'])].gt(round(int(values['finalPayOld09']) * 2.57))].index[0]
                newPayInRopa2019 = \
                    excelFunctionsMain.payMatrix[excelFunctionsMain.payMatrix[excelFunctionsMain.getPayLvl(
                        values['old-level09'])] >= (round(int(values['finalPayOld09']) * 2.57))].index[0]
                fixed_basic = round(
                    excelFunctionsMain.payMatrix[excelFunctionsMain.getPayLvl(values['old-level09'])][newPayInRopa2019])
                al.mainWindow['old-level'].update(value=values['old-level09'])
                al.mainWindow['old-pay'].update(fixed_basic)
                al.mainWindow['finalPayOld'].update(fixed_basic)
                al.mainWindow['-pay19'].update(fixed_basic)
                al.mainWindow['09Report'].update(disabled=False)
        except:
            al.sg.popup("Something went wrong. Please report this bug.", title='Error', icon=icon)
            al.mainWindow['09Report'].update(disabled=True)


    def generateReport09():
        """ Function for generating html report for ROPA 2009

        :return: True or False
        :rtype: bool
        """
        try:
            html_out = al.report_template2.render(items=increments_ropa_2009, fixation_type=values['pay-fix-type'],
                                                  notional_p=values['dt-notional'], acctual_p=values['dt-actual'],
                                                  notional_pf=values['dt-notional-pay-fix'],
                                                  acctual_pf=values['dt-actual-pay-fix'],
                                                  old_basic=values['oldBasicPay'], grade_pay=values['oldGradePay'],
                                                  old_total_pay=values['oldBandPay'],
                                                  add_in_old_lvl_09=values['extraIncNo09'],
                                                  old_final_pay=values['finalPayOld09'],
                                                  ropa09_lvl=values['old-level09'],
                                                  new_final_pay09=values['-pay19'], )
            file_name = f'Pay Fixation report ROPA 2009 {al.time_stamp}.html'
            with open(f"./reports/{file_name}", "w") as a:
                a.write(html_out)
            a.close()
            al.sg.webbrowser.open(url=f"{al.current_directory}/reports/{file_name}", new=2)
            return True
        except Exception as d:
            al.sg.popup_error(f"The following error is detected :\n{d}\nPlease report this bug to our GitHub.",
                              title="Error", icon=icon)
            return False


    if event == '09Report':
        try:
            increments_ropa_2009.clear()
            increments_ropa_2009.append([values['incDate1'], values['inc09-1'], values['payAfterInc09']])
            for f in range(k):
                increments_ropa_2009.append([values[f'incrementDate{f}09'], values[f'incrementType{f}09'],
                                             values[f'payInc{f}09']])
            generateReport09()
        except Exception as e:
            al.sg.popup_error(f"The following error is detected :\n{e}\nPlease report this bug to our GitHub.",
                              title="Error", icon=icon)


    def showAboutWindow():
        """ Generates About the Application Window

        :return: window events and values
        :rtype: list
        """
        about_window = al.sg.Window("About the Application", layout=al.aboutWindowLayout(), size=(450, 400),
                                    finalize=True, modal=True, icon=icon)
        while True:
            a_event, a_values = about_window.read()
            if a_event in [al.sg.WIN_CLOSED, 'a_window_close']:
                break

            if a_event == 'fsm':
                o = al.sg.popup(
                    "The free software movement is a social movement with the goal of obtaining and guaranteeing "
                    "certain "
                    "freedoms for software users, namely the freedom to run the software, to study the software, "
                    "to modify the software, to share possibly modified copies of the software. Software which "
                    "meets these requirements is termed free software. The word 'free' is ambiguous in English, "
                    "although in this context, it means 'free as in freedom', not 'free as in zero price'."
                    "\nSource: Wikipedia", custom_text="Know More", title="Free Software Movement", icon=icon)
                if o:
                    al.sg.webbrowser.open('https://en.wikipedia.org/wiki/Free_software_movement', new=2)

            if a_event == 'licence':
                al.sg.webbrowser.open('https://www.gnu.org/licenses/gpl-3.0.en.html', new=2)

            if a_event == 'git':
                al.sg.webbrowser.open('https://github.com/loku-sama/pay-fixation-wb', new=2)

            if a_event == 'author':
                al.sg.webbrowser.open('https://lokusden.neocities.org/', new=2)

        about_window.close()
        return a_event, a_values


    if event == 'About the App':
        showAboutWindow()


    def showHelpWin():
        """ Generates Help Window

        :return: window events and values
        :rtype: list
        """
        help_window = al.sg.Window(layout=al.helpWindowLayout(), finalize=True, modal=True, size=(700, 580),
                                   title="Help : Automatic Pay Fixation Calculator", icon=icon)

        while True:
            h_events, h_values = help_window.read()
            if h_events in [al.sg.WIN_CLOSED, 'h_window_close']:
                break

        help_window.close()
        return h_events, h_values

    if event in ['--k--']:
        showHelpWin()  # Shows the Help window

al.mainWindow.close()
