import openpyxl
#import BarChart, Reference, Series


# ------------------------------------------------------------------------
#                          METHODS
# -----------------------------------------------------------------------
def sheet_setup(sheet, teamnum):
    # Method sets up sheet for each team

    # set top of sheet to say team number
    sheet.cell(1,5).value = teamnum

    # --------------------------------------------------
    #         AUTO
    # --------------------------------------------------
    # col 1 - match num
    # col 2 - switch
    # col 3 - scale
    # col 4 - total cubes
    # col 5 - baseline (Passed/ Not passed)

    # col 7 - avg switch
    # col 8 - avg scale


    sheet.cell(2,2).value = "AUTO"
    # AUTO

    sheet.cell(3, 1).value = "Match Num"
    sheet.cell(3, 2).value = "Switch"
    sheet.cell(3, 3).value = "Scale"
    sheet.cell(3, 4).value = "Total"
    sheet.cell(3, 5).value = "Baseline"

    sheet.cell(3, 7).value = "AVG Switch"
    sheet.cell(3, 8).value = "AVG Scale"

    # ----------------------------------------------------------
    #                     TELEOP
    # ----------------------------------------------------------
    # col 1 - match num
    # col 2 - switch
    # col 3 - vault
    # col 4 - scale
    # col 5 - total cubes
    # col 6 - climb
    # col 7 - result
    # col 9 - avg switch
    # col 10 - avg vault
    # col 11 - avg scale
    # input is the Excel sheet
    # TELEOP
    sheet.cell(19, 2).value = "TELEOP"

    sheet.cell(20,1).value = "Match Num"
    sheet.cell(20,2).value = "Switch"
    sheet.cell(20,3).value = "Vault"
    sheet.cell(20,4).value = "Scale"
    sheet.cell(20,5).value = "Total"
    sheet.cell(20,6).value = "Climb"
    sheet.cell(20,7).value = "Result"
    sheet.cell(20,9).value = "AVG Switch"
    sheet.cell(20,10).value = "AVG Vault"
    sheet.cell(20,11).value = "AVG Scale"

# --------------------------------------------------------------------------------------

def transfer_tele_data(data_sheet, output_sheet, d_row, o_row):
    # transfer tele data from data sheet to the output sheet for each team
    # in data excel sheet
    # col 2 - team num
    # col 3 - match num
    # col 4 - cubes in switch
    # col 5 - cubes in vault
    # col 6 - cubes in scale
    # col 7 - Climb (Yes/No)
    # col 8 - Result (Win/Lost/Tie)

    # OUTPUT SHEET
    # col 1 - match num
    # col 2 - switch
    # col 3 - vault
    # col 4 - scale
    # col 5 - total cubes
    # col 6 - climb
    # col 7 - result
    # col 9 - avg switch
    # col 10 - avg vault
    # col 11 - avg scale

    #   set MATCH NUM
    output_sheet.cell(o_row, 1).value = data_sheet.cell(d_row, 3).value
#   set SWITCH
    output_sheet.cell(o_row, 2).value = data_sheet.cell(d_row, 4).value
#   set VAULT
    output_sheet.cell(o_row, 3).value = data_sheet.cell(d_row, 5).value
#   set SCALE
    output_sheet.cell(o_row, 4).value = data_sheet.cell(d_row, 6).value
#   set TOTAL
    output_sheet.cell(o_row, 5).value = int(data_sheet.cell(d_row, 6).value) + int(data_sheet.cell(d_row, 5).value) + int(data_sheet.cell(d_row, 4).value)

#   set CLIMB
    output_sheet.cell(o_row, 6).value = data_sheet.cell(d_row, 7).value
#   set RESULT
    output_sheet.cell(o_row, 7).value = data_sheet.cell(d_row, 8).value

# -----------------------------------------------------------------------

def transfer_auto_data(data_sheet, output_sheet, d_row, o_row):
    # transfer auto data from data sheet to the output sheet for each team
    # in data excel sheet
    # col 2 - team num
    # col 3 - match num
    # col 4 - cubes in switch
    # col 5 - baseline
    # col 6 - cubes in scale

    # OUTPUT SHEET
    # col 1 - match num
    # col 2 - switch
    # col 3 - scale
    # col 4 - total cubes
    # col 5 - baseline (Passed/ Not passed)

    # col 7 - avg switch
    # col 8 - avg scale

#   set MATCH NUM
    output_sheet.cell(o_row, 1).value = data_sheet.cell(d_row, 3).value
#   set SWITCH
    output_sheet.cell(o_row, 2).value = data_sheet.cell(d_row, 4).value
#   set SCALE
    output_sheet.cell(o_row, 3).value = data_sheet.cell(d_row, 6).value
#   set TOTAL
    output_sheet.cell(o_row, 4).value = int(data_sheet.cell(d_row, 6).value) + int(data_sheet.cell(d_row, 4).value)

def calc_avg(output_sheet, matches_played):
    '''
    Calculates both averages from auto and teleop data
    :param output_sheet: team sheet
    :param matches_played: how long the row goes for
    :param col:
    :return: prints the averages to the sheet
    '''

    # AUTO
    # ----------------------------
    # average switch
    a_switch_sum = 0
    a_scale_sum = 0
  # finds sum of all the switch and scale cubes
    for i in range(4, 4 + matches_played - 1, 1):
        a_switch_sum += output_sheet.cell(i,2).value
        a_scale_sum += output_sheet.cell(i,3).value
# prints out average to the Excel sheet
    if matches_played == 0:
        output_sheet.cell(4, 7).value = 0
        output_sheet.cell(4, 8).value = 0
    else:
        output_sheet.cell(4, 7).value = float(a_switch_sum)/matches_played
        output_sheet.cell(4, 8).value = float(a_scale_sum)/matches_played


    # TELE
    # ------------------------------
    t_switch_sum = 0
    t_vault_sum = 0
    t_scale_sum = 0

    for i in range(21, 21 + matches_played, 1):
        t_switch_sum += output_sheet.cell(i, 2).value
        t_vault_sum += output_sheet.cell(i, 3).value
        t_scale_sum += output_sheet.cell(i, 4).value

    if matches_played == 0:
        output_sheet.cell(21, 9).value = 0
        output_sheet.cell(21, 10).value = 0
        output_sheet.cell(21, 11).value = 0
    else:
        output_sheet.cell(21, 9).value = float(t_switch_sum)/matches_played
        output_sheet.cell(21, 10).value = float(t_vault_sum)/matches_played
        output_sheet.cell(21, 11).value = float(t_scale_sum)/matches_played



#
#
#
# ----------------------------------------------------------------------------------------------------------------------
#                                               MAIN CODE
# ----------------------------------------------------------------------------------------------------------------------
# instantiate new variables
# data is the excel sheet with all the scouting data
data = openpyxl.load_workbook('2018_LA_Regional_Data.xlsx')
# auto sheet of the scouting data Excel file
auto = data['Auto']
# tele sheet of the scouting data Excel file
tele = data['Tele']


# Creating new workbook to store organized data
teamOutput = openpyxl.Workbook()
teamOutput_name = 'LATeamDataTest7.xlsx'
teamSheet = teamOutput.active

# ------------------------
#          TELEOP
# ------------------------
# in data excel sheet
# col 2 - team num
# col 3 - match num
# col 4 - cubes in switch
# col 5 - cubes in vault
# col 6 - cubes in scale
# col 7 - Climb (Yes/No)
# col 8 - Result (Win/Lost/Tie)

# --------------------
# Set up First team
# --------------------
# set lastTeam to the first team to start
lastTeam = tele.cell(4,2).value
teamSheet.title = str(lastTeam)  # set the first team sheet title to the first team
sheet_setup(teamSheet, lastTeam)


sum = [0]
counter = 0
played = 0
# data to keep overall track of
teams_switch_avg = []
teams_scale_avg = []
teams_vault_avg = []
# format of output sheet
# col 1 - match num
# col 2 - switch
# col 3 - vault
# col 4 - scale
# col 5 - climb
# col 6 - result
# col 8 - avg switch
# col 9 - avg vault
# col 10 - avg scale

for i in range(4, 250, 1): # row 4 is start of data
    # ------------------------------------------------------------------
    # SAME TEAM NUMBER
    if lastTeam == tele.cell(i,2).value: #  if same team num add to sum counter

        sum[counter] += tele.cell(i,4).value
        # Set match num
        #ROW starts at 3
        # auto starts at 4
        # teleop starts at 21
        transfer_auto_data(auto, teamSheet, i, 4 + played)
        transfer_tele_data(tele, teamSheet, i, 21 + played)

        played += 1
    # -------------------------------------------------------------------
    # NEW TEAM NUMBER
    else:
        # calculate the averages of the previous team
        calc_avg(teamSheet, played)
        teamSheet.cell(1,1).value = played
        lastTeam = tele.cell(i,2).value #  start of different team num
        teamOutput.create_sheet(str(lastTeam))
        teamSheet = teamOutput[str(lastTeam)]
        sheet_setup(teamSheet, lastTeam)

        counter += 1
        sum.append(0)
        played = 0
print sum
print counter

teamOutput.save(teamOutput_name)

#sheet1['S4'] = counter
#print sheet1['S4'].value

