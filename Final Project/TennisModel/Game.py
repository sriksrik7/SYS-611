from openpyxl import Workbook, load_workbook
import random
from Player import Tennis_Player


# Global Variables
players = []
first_round_winners = []
scnd_round_winners = []
quarter_final_winners = []
semi_final_winners = []

################################ Start of function ##############################
# Function Name: is_player1_win_coin_flip
# Description: Function to determine if player 1 has won the toss or not
# Return: returns true or false
#################################################################################
def is_player1_win_coin_flip():
    heads = 1
    coin = random.randint(0,1)
    if coin == 1:
        return True
    else:
        return False
################################ End of function ################################

################################ Start of function ##############################
# Function Name: simulate_set
# Description: Function to simulate a set in Tennis
# Return: returns a string contains the winner's rank/index followed by the
#        set's result.
#        Eg: Wining player rank/index, player 1 score - player 2 score
#        Eg: 2,4-6
#################################################################################
def simulate_set(player1_idx, player2_idx):

    # Fetch player 1 stats weightage
    player1_weightage = players[player1_idx].weightage_calculation()
    # Fetch player 2 stats weightage
    player2_weightage = players[player2_idx].weightage_calculation()

    # holds player1 and player2's game points for a set
    player1_gamepoints = 0
    player2_gamepoints = 0

    for game_index in range(12):
        rand_fp = random.uniform(0.0, 100.0)
        # Even serves are serviced by player 1
        if (game_index % 2) == 0:
            if rand_fp <= player1_weightage:
                player1_gamepoints += 1
            else:
                player2_gamepoints += 1
        # ODD serves are serviced by player 2
        else:
            if rand_fp <= player2_weightage:
                player2_gamepoints += 1
            else:
                player1_gamepoints += 1
        # If a player has scored 6 points and if he is 1 set ahead of the opponent,
        # then the set is complete. Hence break from the for loop to finalise the set.
        # Eg: 6-0,6-1,6-2,6-3 and 6-4
        if player1_gamepoints == 6 and game_index >= 5 and game_index != 10:
            break
        # Eg: 0-6,1-6,2-6,3-6 and 4-6
        elif player2_gamepoints == 6 and game_index >= 5 and game_index != 10:
            break

    # Check if both players scored 6-6
    # Based on higher weightage of a player, determine the tie breaker
    if player1_gamepoints == player2_gamepoints:
        if player1_weightage > player2_weightage:
            player_won_idx = player1_idx
            # Eg: 7-6
            player1_gamepoints += 1
        else:
            player_won_idx = player2_idx
            # Eg: 6-7
            player2_gamepoints += 1
    elif player1_gamepoints > player2_gamepoints:
        player_won_idx = player1_idx
    else:
        # if player2_gamepoints > player1_gamepoints
        player_won_idx = player2_idx

    # String contains the winner rank/index followed by the set result
    # Eg: Wining player rank/index, player 1 score - player 2 score
    # Eg: 2,4-6
    ret_string = f"{player_won_idx},{player1_gamepoints}-{player2_gamepoints}"
    # print(ret_string)
    return ret_string
################################ End of function ################################

################################ Start of function ##############################
# Function Name: simulate_games
# Description: Function to simulate tennis games
# Return: returns a list of winners
#################################################################################
def simulate_games(title_name, num_matches, match_detail_list, mtc_result_wb_obj):
    list_of_winners = []
    # write to new excel with match schedule for initial/second round
    mtc_result_ws = mtc_result_wb_obj
    mtc_result_ws = mtc_result_ws.create_sheet("Mysheet")
    mtc_result_ws.title = title_name

    mtc_result_ws[f'A1'] = "Match"
    mtc_result_ws[f'B1'] = "Player1 Name"
    mtc_result_ws[f'C1'] = "Player2 Name"
    mtc_result_ws[f'D1'] = "Set 1"
    mtc_result_ws[f'E1'] = "Set 2"
    mtc_result_ws[f'F1'] = "Set 3"
    mtc_result_ws[f'G1'] = "Set 4"
    mtc_result_ws[f'H1'] = "Set 5"
    mtc_result_ws[f'I1'] = "Player Won"

    Cell_num = 0
    for match_list_idx in range(num_matches):
        # Increment the cell number
        Cell_num += 1

        # Fetch player 1 and player 2's indices/ranks from the match schedule
        if is_player1_win_coin_flip():
            p1_index = match_detail_list[match_list_idx]['player1_index']
            p2_index = match_detail_list[match_list_idx]['player2_index']
        else:
            p2_index = match_detail_list[match_list_idx]['player1_index']
            p1_index = match_detail_list[match_list_idx]['player2_index']

        mtc_result_ws[f'A{Cell_num+1}'] = match_detail_list[match_list_idx]['match_index']
        mtc_result_ws[f'B{Cell_num+1}'] = players[p1_index].get_name()
        mtc_result_ws[f'C{Cell_num+1}'] = players[p2_index].get_name()

        # Holds player1 and player 2's win count
        p1_win_count = 0
        p2_win_count = 0
        # Simulate first 3 sets to check if any player has already won all three sets.
        for match_played in range(3):
            set_result = simulate_set(p1_index,p2_index)
            part_set_result = set_result.partition(',')
            # print(f"p1_index:{p1_index} = {part_set_result[0]}")
            if p1_index == int(part_set_result[0]):
                p1_win_count += 1
            else:
                p2_win_count += 1

            if match_played == 0:
                mtc_result_ws[f'D{Cell_num + 1}'] = f"{part_set_result[2]}"
            elif match_played == 1:
                mtc_result_ws[f'E{Cell_num + 1}'] = f"{part_set_result[2]}"
            elif match_played == 2:
                mtc_result_ws[f'F{Cell_num + 1}'] = f"{part_set_result[2]}"


        # print(f"p1_win_count= {p1_win_count}; p2_win_count= {p2_win_count}")
        # check the any player has won the 3 sets already after 3 sets. If so match is finished
        if p1_win_count == 3:
            mtc_result_ws[f'I{Cell_num + 1}'] = f"{players[p1_index].get_name()}"
            list_of_winners.append(players[p1_index].get_name())
            continue
        elif p2_win_count == 3:
            mtc_result_ws[f'I{Cell_num + 1}'] = f"{players[p2_index].get_name()}"
            list_of_winners.append(players[p2_index].get_name())
            continue

        # If none of the player has won 3 sets continue to simulate 4th set
        set4_result = simulate_set(p1_index,p2_index)
        part_set4_result = set4_result.partition(',')
        mtc_result_ws[f'G{Cell_num + 1}'] = f"{part_set4_result[2]}"

        if p1_index == int(part_set4_result[0]):
            p1_win_count += 1
        else:
            p2_win_count += 1

        # print(f"p1_win_count= {p1_win_count}; p2_win_count= {p2_win_count}")
        # check the any player has won the 3 sets already. If so match is finished
        if p1_win_count == 3:
            mtc_result_ws[f'I{Cell_num + 1}'] = f"{players[p1_index].get_name()}"
            list_of_winners.append(players[p1_index].get_name())
            continue
        elif p2_win_count == 3:
            mtc_result_ws[f'I{Cell_num + 1}'] = f"{players[p2_index].get_name()}"
            list_of_winners.append(players[p2_index].get_name())
            continue

        # If none of the player has won 3 sets continue to simulate final set
        set5_result = simulate_set(p1_index,p2_index)
        part_set5_result = set5_result.partition(',')
        mtc_result_ws[f'G{Cell_num + 1}'] = f"{part_set5_result[2]}"

        if p1_index == int(part_set5_result[0]):
            p1_win_count += 1
        else:
            p2_win_count += 1

        # print(f"p1_win_count= {p1_win_count}; p2_win_count= {p2_win_count}")
        if p1_win_count == 3:
            mtc_result_ws[f'I{Cell_num + 1}'] = f"{players[p1_index].get_name()}"
            list_of_winners.append(players[p1_index].get_name())
        elif p2_win_count == 3:
            mtc_result_ws[f'I{Cell_num + 1}'] = f"{players[p2_index].get_name()}"
            list_of_winners.append(players[p2_index].get_name())

    return list_of_winners
################################ End of function ################################

################################ Start of function ##############################
# Function Name: get_player_index
# Description: function to get player index from a list of 32 players
# Return: returns player index from a list of 32 players
#################################################################################
def get_player_index(p_name):
    ret_player_index = int()
    for p_idx in range(32):
        # print(f"p_idx->{p_idx}: players->{players[p_idx].get_name()}: p_name->{p_name}")
        if (players[p_idx].get_name() == p_name):
            ret_player_index = p_idx
            break
    return ret_player_index
################################ End of function ################################



#################################################################################
########################     Main  Starts Here       ############################
#################################################################################


# Take input from the user
print("Enter the year to simulate Wimbeldon tennis tournament")
sheet_input = input("Choose from following years: 2019, 2018, 2017:\n")
print("Please be patient it takes about 3 minutes to simulate")

if sheet_input == "2019":
    sheet_title = "2019"
elif sheet_input == "2018":
    sheet_title = "2018"
elif sheet_input == "2017":
    sheet_title = "2017"
else:
    raise ValueError(f"Invalid input {sheet_input}, Please try again...")

# Load the work book:
wb_obj = load_workbook('Statistics_Data.xlsx')
# Read the active sheet:
stat_sheet = wb_obj[sheet_title]

players = [Tennis_Player() for i in range(32)]

cellnum = 1
for player in players:
    cellnum +=1
    player.set_name(stat_sheet[f'A{cellnum}'].value)
    player.set_stat(player.ACE, stat_sheet[f'D{cellnum}'].value*100)
    player.set_stat(player.DOUBLE_FAULT, stat_sheet[f'E{cellnum}'].value*100)
    player.set_stat(player.FIRST_SERVE, stat_sheet[f'F{cellnum}'].value*100)
    player.set_stat(player.FIRST_SERVE_WON, stat_sheet[f'G{cellnum}'].value*100)
    player.set_stat(player.SECOND_SERVE_WON, stat_sheet[f'H{cellnum}'].value*100)
    player.set_stat(player.BREAK_POINT_SAVED, stat_sheet[f'I{cellnum}'].value*100)
    player.set_stat(player.SERVICE_POINTS_WON, stat_sheet[f'J{cellnum}'].value*100)
    player.set_stat(player.SERVICE_GAMES_WON, stat_sheet[f'K{cellnum}'].value*100)
    player.set_stat(player.ACE_AGAINST, stat_sheet[f'L{cellnum}'].value*100)
    player.set_stat(player.FIRST_SERVE_RET_WON, stat_sheet[f'M{cellnum}'].value*100)
    player.set_stat(player.SECOND_SERVE_RET_WON, stat_sheet[f'N{cellnum}'].value*100)
    player.set_stat(player.BREAK_POINTS_WON, stat_sheet[f'O{cellnum}'].value*100)
    player.set_stat(player.RET_POINTS_WON, stat_sheet[f'P{cellnum}'].value*100)
    player.set_stat(player.RET_GAMES_WON, stat_sheet[f'Q{cellnum}'].value*100)

# Close the workbook after reading
wb_obj.close()

#########################################################
####### Populate and process first round matches #######

####### Bracket seeding for 32 players #######
# Top 16 out of 32 players are played againts the lower 16 players.
# Lower 16 players are randomly allocated to play against the top 16 players
# This is to avoid strong players playing against each other in the intial stages of tournament.
# This type seeding is followed in Wembledom tournament to make the final games challening.

# write to new excel with match schedule for initial/second round
sch_res_wb = Workbook()
rd1_sch_ws = sch_res_wb.active
rd1_sch_ws.title = "RD1_Schedule"

rd1_sch_ws[f'A1'] = "Match"
rd1_sch_ws[f'B1'] = "Rank"
rd1_sch_ws[f'C1'] = "Player Name"

# Create a match details list to incude match number and player 1 ranking.
# This match details list can be used to subscript each player statistics data based on their ranking.
# Holds match number along with player1 and player 2's index/rank details
first_rd_match_list = []
rd1_cell =1
for index in range(16):
    rd1_cell+=1
    rd1_sch_ws[f'A{rd1_cell}'] = index+1
    rd1_sch_ws[f'B{rd1_cell}'] = index+1
    rd1_sch_ws[f'C{rd1_cell}'] = players[index].get_name()
    first_rd_match_list.append({'match_index':(index+1),'player1_index':(index)})
    rd1_cell+=1

# Randomly assign the lower 16 players to play against the top 16 players.
processed_numbers = [16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
rd1_cell_sch = 1
for sch_index in range(16):
    rank = random.choice(processed_numbers)
    processed_numbers.remove(rank)
    rd1_cell_sch+=2
    rd1_sch_ws[f'B{rd1_cell_sch}'] = rank+1
    rd1_sch_ws[f'C{rd1_cell_sch}'] = players[rank].get_name()
    first_rd_match_list[sch_index]['player2_index'] = rank

# print(first_rd_match_list)
# Save the bracket seeding schedule into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

# Simulate matches for first round
first_round_winners = simulate_games("FirstRoundResults", 16, first_rd_match_list, sch_res_wb)

# Save the results of first round matches
sch_res_wb.save("Schedule_Results.xlsx")

#########################################################
####### Populate and process second round matches #######
rd2_sch_ws = sch_res_wb.create_sheet("Mysheet")
rd2_sch_ws.title = "RD2_Schedule"

rd2_sch_ws[f'A1'] = "Match"
rd2_sch_ws[f'B1'] = "Rank"
rd2_sch_ws[f'C1'] = "Player Name"

# wb_sch_res_obj = load_workbook('Schedule_Results.xlsx')
# Read FirstRoundResults sheet to get the wining players from round 1 matches
# rd1_sheet = wb_sch_res_obj["FirstRoundResults"]

# Schedule second round matches
scnd_rd_match_list = []

rd2_top_cell = 1
for win_rd1_top_index in range(8):
    rd2_top_cell+=1
    # Fetch player indices from the list of 32 players to simulate second round matches
    pl1_index = get_player_index(first_round_winners[win_rd1_top_index])

    # Populate second round schedule from first round winners.
    rd2_sch_ws[f'A{rd2_top_cell}'] = win_rd1_top_index+1
    rd2_sch_ws[f'B{rd2_top_cell}'] = pl1_index+1
    rd2_sch_ws[f'C{rd2_top_cell}'] = first_round_winners[win_rd1_top_index]
    # Generate a list of dictionaries for second round matches
    scnd_rd_match_list.append({'match_index':(win_rd1_top_index+1),'player1_index':(pl1_index)})
    rd2_top_cell+=1

rd2_low_cell = 18
for win_rd1_low_index in reversed(range(8)):
    rd2_low_cell-=1
    pl2_index = get_player_index(first_round_winners[win_rd1_low_index+4])
    rd2_sch_ws[f'B{rd2_low_cell}'] = pl2_index+1
    rd2_sch_ws[f'C{rd2_low_cell}'] = first_round_winners[win_rd1_low_index+4]
    # Generate a list of dictionaries for second round matches
    scnd_rd_match_list[win_rd1_low_index]['player2_index'] = pl2_index
    rd2_low_cell-=1

# Save the second round schedule into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

# Simulate matches for second round
scnd_round_winners = simulate_games("SecondRoundResults", 8, scnd_rd_match_list, sch_res_wb)

# Save the second round results into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

##########################################################
####### Populate and process quarter final matches #######
qf_sch_ws = sch_res_wb.create_sheet("Mysheet")
qf_sch_ws.title = "QF_Schedule"

qf_sch_ws[f'A1'] = "Match"
qf_sch_ws[f'B1'] = "Rank"
qf_sch_ws[f'C1'] = "Player Name"

# Schedule quarter final matches
qf_match_list = []
qf_cell =1
win_rd2_index = 0
for index in range(4):
    qf_cell+=1
    # Fetch player indices from the list of 32 players to simulate quarter final matches
    pl1_index = get_player_index(scnd_round_winners[win_rd2_index])
    pl2_index = get_player_index(scnd_round_winners[win_rd2_index+1])

    # Populate quarter final schedule from second round winners.
    qf_sch_ws[f'A{qf_cell}'] = index+1
    qf_sch_ws[f'B{qf_cell}'] = pl1_index+1
    qf_sch_ws[f'C{qf_cell}'] = scnd_round_winners[win_rd2_index]
    qf_sch_ws[f'B{qf_cell+1}'] = pl2_index+1
    qf_sch_ws[f'C{qf_cell+1}'] = scnd_round_winners[win_rd2_index+1]
    # Generate a list of dictionaries for quarter final matches
    qf_match_list.append({'match_index':(index+1),'player1_index':(pl1_index),'player2_index':(pl2_index)})
    qf_cell+=1
    win_rd2_index+=2

# Save the quarter final schedule into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

# Simulate matches for quarter final
quarter_final_winners = simulate_games("QuarterFinalResults", 4, qf_match_list, sch_res_wb)

# Save the quarter final results into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

##########################################################
####### Populate and process semi final matches #######
sf_sch_ws = sch_res_wb.create_sheet("Mysheet")
sf_sch_ws.title = "SemiFinal_Schedule"

sf_sch_ws[f'A1'] = "Match"
sf_sch_ws[f'B1'] = "Rank"
sf_sch_ws[f'C1'] = "Player Name"

# Schedule semi final matches
sf_match_list = []
sf_cell =1
win_qf_index = 0
for index in range(2):
    sf_cell+=1
    # Fetch player indices from the list of 32 players to simulate semi final matches
    pl1_index = get_player_index(quarter_final_winners[win_qf_index])
    pl2_index = get_player_index(quarter_final_winners[win_qf_index+1])

    # Populated semi final schedule from quarter final winners.
    sf_sch_ws[f'A{sf_cell}'] = index+1
    sf_sch_ws[f'B{sf_cell}'] = pl1_index+1
    sf_sch_ws[f'C{sf_cell}'] = quarter_final_winners[win_qf_index]
    sf_sch_ws[f'B{sf_cell+1}'] = pl2_index+1
    sf_sch_ws[f'C{sf_cell+1}'] = quarter_final_winners[win_qf_index+1]
    # Generate a list of dictionaries for semi final matches
    sf_match_list.append({'match_index':(index+1),'player1_index':(pl1_index),'player2_index':(pl2_index)})
    sf_cell+=1
    win_qf_index+=2

# Save the semi final schedule into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

# Simulate matches for semi final
semi_final_winners = simulate_games("SemiFinalResults", 2, sf_match_list, sch_res_wb)

# Save the semi final results into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

##########################################################
####### Populate and process final match #################
final_sch_ws = sch_res_wb.create_sheet("Mysheet")
final_sch_ws.title = "Final_Schedule"

final_sch_ws[f'A1'] = "Match"
final_sch_ws[f'B1'] = "Rank"
final_sch_ws[f'C1'] = "Player Name"

# Schedule final matche
final_match_list = []

# Fetch player indices from the list of 32 players to simulate final match
pl1_index = get_player_index(semi_final_winners[0])
pl2_index = get_player_index(semi_final_winners[1])

# Populate final schedule from semi final winners.
final_sch_ws[f'A2'] = 1
final_sch_ws[f'B2'] = pl1_index+1
final_sch_ws[f'C2'] = semi_final_winners[0]
final_sch_ws[f'B3'] = pl2_index+1
final_sch_ws[f'C3'] = semi_final_winners[1]
# Generate a list of dictionaries for final matche
final_match_list.append({'match_index':(1),'player1_index':(pl1_index),'player2_index':(pl2_index)})

# Save the final schedule into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

# Simulate matches for final
final_winner = simulate_games("FinalResults", 1, final_match_list, sch_res_wb)

print(final_winner)
# Save the final results into a spread sheet.
sch_res_wb.save("Schedule_Results.xlsx")

# Close the workbook after writing simulated data
sch_res_wb.close()
#################################################################################
########################         End of Main         ############################
#################################################################################