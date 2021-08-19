################################## Start of file ##################################
# File Name: Player.py
# Description: This is file contains all methods for class Tennis_Player
###################################################################################

class Tennis_Player:

  # Statistic type constants
  ACE = "ACE%"
  DOUBLE_FAULT = "Double Fault%"
  FIRST_SERVE = "1st Serve %"
  FIRST_SERVE_WON = "1st Serve WON%"
  SECOND_SERVE_WON = "2nd Serve WON%"
  BREAK_POINT_SAVED = "Break Point Saved%"
  SERVICE_POINTS_WON = "Service points Won%"
  SERVICE_GAMES_WON = "Service Games Won%"
  ACE_AGAINST = "ACE Against%"
  FIRST_SERVE_RET_WON = "1st Serve Return WON%"
  SECOND_SERVE_RET_WON = "2nd Serve Return WON%"
  BREAK_POINTS_WON = "Break Points Won %"
  RET_POINTS_WON = "Return Points won%"
  RET_GAMES_WON = "Return Games Won%"

  # Weightage constants
  ACE_WEIGHT = 2/100
  DOUBLE_FAULT_WEIGHT  = 2/100
  FIRST_SERVE_WEIGHT  = 2/100
  FIRST_SERVE_WON_WEIGHT  = 7/100
  SECOND_SERVE_WON_WEIGHT  = 12/100
  BREAK_POINT_SAVED_WEIGHT  = 9/100
  SERVICE_POINTS_WON_WEIGHT  = 9/100
  SERVICE_GAMES_WON_WEIGHT = 15/100
  ACE_AGAINST_WEIGHT  = 2/100
  FIRST_SERVE_RET_WON_WEIGHT  = 9/100
  SECOND_SERVE_RET_WON_WEIGHT  = 5/100
  BREAK_POINTS_WON_WEIGHT  = 7/100
  RET_POINTS_WON_WEIGHT = 7/100
  RET_GAMES_WON_WEIGHT  = 12/100

  ################################ Start of function ##############################
  # Function Name: __init__
  # Description: This is a constructor for class Tennis_Player to initializes all
  #              class variables
  # Return: None
  #################################################################################
  def __init__(self):
    self.player_name = "null"
    self.ace_percent = int(0)
    self.double_fault = int(0)
    self.first_serve = int(0)
    self.first_serve_won = int(0)
    self.scnd_serve_won = int(0)
    self.break_points_saved = int(0)
    self.service_points_won = int(0)
    self.service_games_won = int(0)
    self.ace_against = int(0)
    self.first_serve_ret_won = int(0)
    self.scnd_serve_ret_won = int(0)
    self.break_points_won = int(0)
    self.ret_points_won = int(0)
    self.ret_games_won = int(0)

  ################################ End of function ################################

  ################################ Start of function ##############################
  # Function Name: get_name
  # Description: This function returns a player's name for a given Tennis_Player
  #              class object
  # Return: None
  #################################################################################
  def get_name(self):
    return(self.player_name)
  ################################ End of function ################################

  ################################ Start of function ##############################
  # Function Name: set_name
  # Description: This function sets a player's name for a given Tennis_Player
  #              class object
  # Return: None
  #################################################################################
  def set_name(self, name):
    self.player_name = name

  ################################ End of function ################################

  ################################ Start of function ##############################
  # Function Name: set_stat
  # Description: This function updates the player's statistics information based on
  #              the given statistics type and data
  # Return: None
  #################################################################################
  def set_stat(self, stat_type,val):

      if stat_type == "":
          raise ValueError(f"Invalid stat_type: {stat_type}")

      if val == None:
          raise ValueError(f"Invalid stat value: {val}")

      if stat_type == "ACE%":
          self.ace_percent = val

      if stat_type == "Double Fault%":
          self.double_fault = val

      if stat_type == "1st Serve %":
          self.first_serve = val

      if stat_type == "1st Serve WON%":
          self.first_serve_won = val

      if stat_type == "2nd Serve WON%":
          self.scnd_serve_won = val

      if stat_type == "Break Point Saved%":
          self.break_points_saved = val

      if stat_type == "Service points Won%":
          self.service_points_won = val

      if stat_type == "Service Games Won%":
          self.service_games_won = val

      if stat_type == "ACE Against%":
          self.ace_against = val

      if stat_type == "1st Serve Return WON%":
          self.first_serve_ret_won = val

      if stat_type == "2nd Serve Return WON%":
          self.scnd_serve_ret_won = val

      if stat_type == "Break Points Won %":
          self.break_points_won = val

      if stat_type == "Return Points won%":
          self.ret_points_won = val

      if stat_type == "Return Games Won%":
          self.ret_games_won = val

  ################################ End of function ################################

  ################################ Start of function ##############################
  # Function Name: weightage_calculation
  # Description: This function calculates the win weigthage for a tennis player
  #              based on the provided statistics information
  # Return: win weigthage number
  #################################################################################
  def weightage_calculation(self):

      ret_weightage = self.ace_percent * self.ACE_WEIGHT
      ret_weightage += self.double_fault * self.DOUBLE_FAULT_WEIGHT
      ret_weightage += self.first_serve * self.FIRST_SERVE_WEIGHT
      ret_weightage += self.first_serve_won * self.FIRST_SERVE_WON_WEIGHT
      ret_weightage += self.scnd_serve_won * self.SECOND_SERVE_WON_WEIGHT
      ret_weightage += self.break_points_saved * self.BREAK_POINT_SAVED_WEIGHT
      ret_weightage += self.service_points_won * self.SERVICE_POINTS_WON_WEIGHT
      ret_weightage += self.service_games_won * self.SERVICE_GAMES_WON_WEIGHT
      ret_weightage += self.ace_against * self.ACE_AGAINST_WEIGHT
      ret_weightage += self.first_serve_ret_won * self.FIRST_SERVE_RET_WON_WEIGHT
      ret_weightage += self.scnd_serve_ret_won * self.SECOND_SERVE_RET_WON_WEIGHT
      ret_weightage += self.break_points_won * self.BREAK_POINTS_WON_WEIGHT
      ret_weightage += self.ret_points_won * self.RET_POINTS_WON_WEIGHT
      ret_weightage += self.ret_games_won * self.RET_GAMES_WON_WEIGHT

      return ret_weightage
  ################################ End of function ################################

################################## End of file ####################################


