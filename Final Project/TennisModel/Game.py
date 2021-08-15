from openpyxl import Workbook, load_workbook
import random
from Player import Tennis_Player

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
#        Eg: Wining player rank/index, player 1 score , player 2 score
#        Eg: 2,4,6
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
    # Eg: Wining player rank/index, player 1 score , player 2 score
    # Eg: 2,4,6
    ret_string = f"{player_won_idx},{player1_gamepoints},{player2_gamepoints}"
    # print(ret_string)
    return ret_string
################################ End of function ################################

################################ Start of function ##############################
# Function Name: simulate_games
# Description: Function to simulate tennis games
# Return: returns a list of winners
#################################################################################
def simulate_games(cell_num, num_matches, match_detail_list, mtc_result_ws):
    list_of_winners = []
    cell_aphabet = ["D", "E", "F", "G", "H"]

    mtc_result_ws[f'D{cell_num}'] = "Set 1"
    mtc_result_ws[f'E{cell_num}'] = "Set 2"
    mtc_result_ws[f'F{cell_num}'] = "Set 3"
    mtc_result_ws[f'G{cell_num}'] = "Set 4"
    mtc_result_ws[f'H{cell_num}'] = "Set 5"
    mtc_result_ws[f'I{cell_num}'] = "Player Won"

    for match_list_idx in range(num_matches):
        # Increment the cell number
        cell_num += 1

        #Is player1 and player2 flipped at coin toss
        #if flipped player 2 is servicing first else player 1 is servicing first
        is_player_index_flipped = False

        # Fetch player 1 and player 2's indices/ranks from the match schedule
        if is_player1_win_coin_flip():
            is_player_index_flipped = False
            p1_index = match_detail_list[match_list_idx]['player1_index']
            p2_index = match_detail_list[match_list_idx]['player2_index']
        else:
            is_player_index_flipped = True
            p2_index = match_detail_list[match_list_idx]['player1_index']
            p1_index = match_detail_list[match_list_idx]['player2_index']

        # Holds player1 and player 2's win count
        p1_win_count = 0
        p2_win_count = 0
        # Simulate first 3 sets to check if any player has already won all three sets.
        for match_played in range(3):
            set_result = simulate_set(p1_index,p2_index)
            part_set_result = set_result.split(',')
            # print(f"p1_index:{p1_index} = {part_set_result[0]}")
            if p1_index == int(part_set_result[0]):
                p1_win_count += 1
            else:
                p2_win_count += 1

            # If the player 1 has lost the toss, flip and write the results in assigned cell numbers
            if is_player_index_flipped == True:
                mtc_result_ws[f'{cell_aphabet[match_played]}{cell_num}'] = int(part_set_result[2])
                mtc_result_ws[f'{cell_aphabet[match_played]}{cell_num + 1}'] = int(part_set_result[1])
            else:
                mtc_result_ws[f'{cell_aphabet[match_played]}{cell_num}'] = int(part_set_result[1])
                mtc_result_ws[f'{cell_aphabet[match_played]}{cell_num + 1}'] = int(part_set_result[2])

        # print(f"p1_win_count= {p1_win_count}; p2_win_count= {p2_win_count}")
        # check the any player has won the 3 sets already after 3 sets. If so match is finished
        if p1_win_count == 3:
            mtc_result_ws[f'I{cell_num}'] = f"{players[p1_index].get_name()}"
            list_of_winners.append(players[p1_index].get_name())
            # Increment the cell number for next match
            cell_num += 1
            continue
        elif p2_win_count == 3:
            mtc_result_ws[f'I{cell_num}'] = f"{players[p2_index].get_name()}"
            list_of_winners.append(players[p2_index].get_name())
            # Increment the cell number for next match
            cell_num += 1
            continue

        # If none of the player has won 3 sets continue to simulate 4th set
        set4_result = simulate_set(p1_index,p2_index)
        part_set4_result = set4_result.split(',')

        # If the player 1 has lost the toss, flip and write the results in assigned cell numbers
        if is_player_index_flipped == True:
            mtc_result_ws[f'G{cell_num}'] = int(part_set4_result[2])
            mtc_result_ws[f'G{cell_num + 1}'] = int(part_set4_result[1])
        else:
            mtc_result_ws[f'G{cell_num}'] = int(part_set4_result[1])
            mtc_result_ws[f'G{cell_num + 1}'] = int(part_set4_result[2])

        if p1_index == int(part_set4_result[0]):
            p1_win_count += 1
        else:
            p2_win_count += 1

        # print(f"p1_win_count= {p1_win_count}; p2_win_count= {p2_win_count}")
        # check the any player has won the 3 sets already. If so match is finished
        if p1_win_count == 3:
            mtc_result_ws[f'I{cell_num}'] = f"{players[p1_index].get_name()}"
            list_of_winners.append(players[p1_index].get_name())
            # Increment the cell number for next match
            cell_num += 1
            continue
        elif p2_win_count == 3:
            mtc_result_ws[f'I{cell_num }'] = f"{players[p2_index].get_name()}"
            list_of_winners.append(players[p2_index].get_name())
            # Increment the cell number for next match
            cell_num += 1
            continue

        # If none of the player has won 3 sets continue to simulate final set
        set5_result = simulate_set(p1_index,p2_index)
        part_set5_result = set5_result.split(',')

        # If the player 1 has lost the toss, flip and write the results in assigned cell numbers
        if is_player_index_flipped == True:
            mtc_result_ws[f'H{cell_num}'] = int(part_set5_result[2])
            mtc_result_ws[f'H{cell_num + 1}'] = int(part_set5_result[1])
        else:
            mtc_result_ws[f'H{cell_num}'] = int(part_set5_result[1])
            mtc_result_ws[f'H{cell_num + 1}'] = int(part_set5_result[2])

        if p1_index == int(part_set5_result[0]):
            p1_win_count += 1
        else:
            p2_win_count += 1

        if p1_win_count == 3:
            mtc_result_ws[f'I{cell_num}'] = f"{players[p1_index].get_name()}"
            list_of_winners.append(players[p1_index].get_name())
        elif p2_win_count == 3:
            mtc_result_ws[f'I{cell_num}'] = f"{players[p2_index].get_name()}"
            list_of_winners.append(players[p2_index].get_name())

        # Increment the cell number for next match
        cell_num += 1

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


# Global Variables
players = []
first_round_winners = []
scnd_round_winners = []
quarter_final_winners = []
semi_final_winners = []

rd1_cell_index = 2
rd2_cell_index = 37
qf_cell_index = 56
sf_cell_index = 67
final_cell_index = 74
champ_cell_index = 78

# Take input from the user
print("Enter the year to simulate Wimbledon tennis tournament")
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
sch_res_ws = sch_res_wb.active
sim_num = 1
sch_res_ws.title = f"Sim{sim_num}_results"

sch_res_ws[f'A{rd1_cell_index-1}'] = "Wimbledon Round 1"
sch_res_ws[f'A{rd1_cell_index}'] = "Match"
sch_res_ws[f'B{rd1_cell_index}'] = "Rank"
sch_res_ws[f'C{rd1_cell_index}'] = "Player Name"

# Create a match details list to incude match number and player 1 ranking.
# This match details list can be used to subscript each player statistics data based on their ranking.
# Holds match number along with player1 and player 2's index/rank details
first_rd_match_list = []
rd1_cell = rd1_cell_index
for index in range(16):
    rd1_cell+=1
    sch_res_ws[f'A{rd1_cell}'] = index+1
    sch_res_ws[f'B{rd1_cell}'] = index+1
    sch_res_ws[f'C{rd1_cell}'] = players[index].get_name()
    first_rd_match_list.append({'match_index':(index+1),'player1_index':(index)})
    rd1_cell+=1

# Randomly assign the lower 16 players to play against the top 16 players.
processed_numbers = [16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31]
cell_sch = rd1_cell_index
for sch_index in range(16):
    rank = random.choice(processed_numbers)
    processed_numbers.remove(rank)
    cell_sch+=2
    sch_res_ws[f'B{cell_sch}'] = rank+1
    sch_res_ws[f'C{cell_sch}'] = players[rank].get_name()
    first_rd_match_list[sch_index]['player2_index'] = rank

# print(first_rd_match_list)
# Save the bracket seeding schedule into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

# Simulate matches for first round
first_round_winners = simulate_games(rd1_cell_index, 16, first_rd_match_list, sch_res_ws)

# Save the results of first round matches
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

#########################################################
####### Populate and process second round matches #######

sch_res_ws[f'A{rd2_cell_index-1}'] = "Wimbledon Round 2"
sch_res_ws[f'A{rd2_cell_index}'] = "Match"
sch_res_ws[f'B{rd2_cell_index}'] = "Rank"
sch_res_ws[f'C{rd2_cell_index}'] = "Player Name"

# Schedule second round matches
scnd_rd_match_list = []
rd2_top_cell = rd2_cell_index
for win_rd1_top_index in range(8):
    rd2_top_cell+=1
    # Fetch player indices from the list of 32 players to simulate second round matches
    pl1_index = get_player_index(first_round_winners[win_rd1_top_index])

    # Populate second round schedule from first round winners.
    sch_res_ws[f'A{rd2_top_cell}'] = win_rd1_top_index+1
    sch_res_ws[f'B{rd2_top_cell}'] = pl1_index+1
    sch_res_ws[f'C{rd2_top_cell}'] = first_round_winners[win_rd1_top_index]
    # Generate a list of dictionaries for second round matches
    scnd_rd_match_list.append({'match_index':(win_rd1_top_index+1),'player1_index':(pl1_index)})
    rd2_top_cell+=1

rd2_low_cell = (rd2_cell_index+17)
for win_rd1_low_index in reversed(range(8)):
    rd2_low_cell-=1
    pl2_index = get_player_index(first_round_winners[win_rd1_low_index+8])
    sch_res_ws[f'B{rd2_low_cell}'] = pl2_index+1
    sch_res_ws[f'C{rd2_low_cell}'] = first_round_winners[win_rd1_low_index+8]
    # Generate a list of dictionaries for second round matches
    scnd_rd_match_list[win_rd1_low_index]['player2_index'] = pl2_index
    rd2_low_cell-=1

# Save the second round schedule into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

# Simulate matches for second round
scnd_round_winners = simulate_games(rd2_cell_index, 8, scnd_rd_match_list, sch_res_ws)

# Save the second round results into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

##########################################################
####### Populate and process quarter final matches #######
sch_res_ws[f'A{qf_cell_index-1}'] = "Wimbledon Quarter Finals"
sch_res_ws[f'A{qf_cell_index}'] = "Match"
sch_res_ws[f'B{qf_cell_index}'] = "Rank"
sch_res_ws[f'C{qf_cell_index}'] = "Player Name"

# Schedule quarter final matches
qf_match_list = []
qf_cell = qf_cell_index
win_rd2_index = 0
for index in range(4):
    qf_cell+=1
    # Fetch player indices from the list of 32 players to simulate quarter final matches
    pl1_index = get_player_index(scnd_round_winners[win_rd2_index])
    pl2_index = get_player_index(scnd_round_winners[win_rd2_index+1])

    # Populate quarter final schedule from second round winners.
    sch_res_ws[f'A{qf_cell}'] = index+1
    sch_res_ws[f'B{qf_cell}'] = pl1_index+1
    sch_res_ws[f'C{qf_cell}'] = scnd_round_winners[win_rd2_index]
    sch_res_ws[f'B{qf_cell+1}'] = pl2_index+1
    sch_res_ws[f'C{qf_cell+1}'] = scnd_round_winners[win_rd2_index+1]
    # Generate a list of dictionaries for quarter final matches
    qf_match_list.append({'match_index':(index+1),'player1_index':(pl1_index),'player2_index':(pl2_index)})
    qf_cell+=1
    win_rd2_index+=2

# Save the quarter final schedule into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

# Simulate matches for quarter final
quarter_final_winners = simulate_games(qf_cell_index, 4, qf_match_list, sch_res_ws)

# Save the quarter final results into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

##########################################################
####### Populate and process semi final matches #######
sch_res_ws[f'A{sf_cell_index-1}'] = "Wimbledon Semi Final"
sch_res_ws[f'A{sf_cell_index}'] = "Match"
sch_res_ws[f'B{sf_cell_index}'] = "Rank"
sch_res_ws[f'C{sf_cell_index}'] = "Player Name"

# Schedule semi final matches
sf_match_list = []
sf_cell = sf_cell_index
win_qf_index = 0
for index in range(2):
    sf_cell+=1
    # Fetch player indices from the list of 32 players to simulate semi final matches
    pl1_index = get_player_index(quarter_final_winners[win_qf_index])
    pl2_index = get_player_index(quarter_final_winners[win_qf_index+1])

    # Populated semi final schedule from quarter final winners.
    sch_res_ws[f'A{sf_cell}'] = index+1
    sch_res_ws[f'B{sf_cell}'] = pl1_index+1
    sch_res_ws[f'C{sf_cell}'] = quarter_final_winners[win_qf_index]
    sch_res_ws[f'B{sf_cell+1}'] = pl2_index+1
    sch_res_ws[f'C{sf_cell+1}'] = quarter_final_winners[win_qf_index+1]
    # Generate a list of dictionaries for semi final matches
    sf_match_list.append({'match_index':(index+1),'player1_index':(pl1_index),'player2_index':(pl2_index)})
    sf_cell+=1
    win_qf_index+=2

# Save the semi final schedule into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

# Simulate matches for semi final
semi_final_winners = simulate_games(sf_cell_index, 2, sf_match_list, sch_res_ws)

# Save the semi final results into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

##########################################################
####### Populate and process final match #################
sch_res_ws[f'A{final_cell_index-1}'] = "Wimbledon Final"
sch_res_ws[f'A{final_cell_index}'] = "Match"
sch_res_ws[f'B{final_cell_index}'] = "Rank"
sch_res_ws[f'C{final_cell_index}'] = "Player Name"

# Schedule final matche
final_match_list = []

# Fetch player indices from the list of 32 players to simulate final match
pl1_index = get_player_index(semi_final_winners[0])
pl2_index = get_player_index(semi_final_winners[1])

# Populate final schedule from semi final winners.
sch_res_ws[f'A{final_cell_index+1}'] = 1
sch_res_ws[f'B{final_cell_index+1}'] = pl1_index+1
sch_res_ws[f'C{final_cell_index+1}'] = semi_final_winners[0]
sch_res_ws[f'B{final_cell_index+2}'] = pl2_index+1
sch_res_ws[f'C{final_cell_index+2}'] = semi_final_winners[1]
# Generate a list of dictionaries for final matche
final_match_list.append({'match_index':(1),'player1_index':(pl1_index),'player2_index':(pl2_index)})

# Save the final schedule into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

# Simulate matches for final
final_winner = simulate_games(final_cell_index, 1, final_match_list, sch_res_ws)

# Record the simulate wimbledon champion
# Fetch player indices from the list of 32 players to simulate final match
Champ_pl_index = get_player_index(final_winner[0])
sch_res_ws[f'A{champ_cell_index}'] = "Wimbledon Champion"
sch_res_ws[f'B{champ_cell_index}'] = "Rank"
sch_res_ws[f'C{champ_cell_index}'] = "Player Name"

sch_res_ws[f'B{champ_cell_index+1}'] = Champ_pl_index+1
sch_res_ws[f'C{champ_cell_index+1}'] = final_winner[0]

print(f"The wimbledon Champion is {final_winner[0]}")
# Save the final results into a spread sheet.
sch_res_wb.save("Wimbledon_Model_Results.xlsx")

# Close the workbook after writing simulated data
sch_res_wb.close()
#################################################################################
########################         End of Main         ############################
#################################################################################