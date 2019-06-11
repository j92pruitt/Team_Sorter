"""
Team Sorter

Author: Joseph Pruitt (j92pruitt@gmail.com)
"""

from random import randint
import heapq
from openpyxl import load_workbook


class Player:
    """
    Class holding information about a player for sorting.

    Attributes:
        first_name (string)
        last_name (string)
        rating (float)
    """

    def __init__(self, row_idx, first_name, last_name, rating=0, age=0, request=None):
        self.first_name = first_name
        self.last_name = last_name
        self.rating = rating
        self.row_idx = row_idx
        self.age = age
        self.request = request

    def __repr__(self):
        return "{first} {last}".format(
                first=self.first_name, last=self.last_name)


class Team:
    """
    A list of players, with useful methods and data for
    sorting stored.

    Attributes:
        number (int)
        player_list (list)
        player_count (int)
    """

    def __init__(self, number):
        self.number = number
        self.player_list = []
        self.player_count = 0

    def __repr__(self):
        return "Team {}".format(self.number)

    def __lt__(self, team):
        return self.player_count < team.player_count

    def add_player(self, player):
        self.player_list.append(player)
        self.player_count += 1

    def avg_rating(self):
        total_rating = 0
        for player in self.player_list:
            total_rating += player.rating
        return total_rating/self.player_count

    def avg_age(self):
        total_age = 0
        for player in self.player_list:
            total_age += player.age
        return total_age/self.player_count

    def copy(self):
        team_copy = Team(self.number)
        for player in self.player_list:
            team_copy.add_player(player)
        return team_copy


def sort(playerpool, team_heap):
    """
    This function randomly divides a list of players evenly 
    into a given number of teams.

    Parameters:
        playerpool (list): List of Players to be sorted
        num (int): Number of teams to sort into.

    Returns:
        team_heap (list): List of teams.
    """

    while playerpool:
        idx = randint(0, len(playerpool)-1)
        team = heapq.heappop(team_heap)
        player = playerpool.pop(idx)

        team.add_player(player)
        if player.request:
            team.add_player(player.request)
            playerpool.remove(player.request)
        heapq.heappush(team_heap, team)

    return team_heap


def team_sort(playerpool, team_heap, number_of_sorts = 1000):
    """ 
    Executes the sort function number_of_sorts times and returns
    the best result. 
    """

    best_score = 0
    for i in range(number_of_sorts):
        team_heap_copy = [team.copy() for team in team_heap]
        playerpool_copy = playerpool.copy()
        team_list = sort(playerpool_copy, team_heap_copy)

        if i == 0 or sort_score(team_list) < best_score:
            best_teams = team_list
            best_score = sort_score(best_teams)
    return best_teams


def sort_score(team_list):
    """
    Scores a list of team objects, a score of 0 is perfectly
    balanced list of teams. 
    """
    score = 0
    avg_ratings = [team.avg_rating() for team in team_list]
    avg_ages = [team.avg_age() for team in team_list]

    for i in avg_ratings:
        for j in avg_ratings:
            score += (i - j) ** 2
    
    for i in avg_ages:
        for j in avg_ages:
            score += ((i-j) ** 2)/4
    
    player_counts = [team.player_count for team in team_list]
    score *= max(player_counts) - min(player_counts) + 1
    return score


def load_players(worksheet, first_name_col, last_name_col, rating_col, age_col, request_col, num):
    playerpool = []

    team_heap = []
    for i in range(num):
        team_heap.append(Team(i+1))

    for i, row in enumerate(worksheet.values):
        if i > 0:
            playerpool.append(
                Player(i, row[first_name_col], row[last_name_col], row[rating_col], row[age_col])
            )
        request_string = row[request_col]
        if request_string:
            if request_string[:9].lower() == "req coach":
                request_team_num = int(request_string[10:])
                team_heap[request_team_num-1].add_player(playerpool[-1])
                print("{} want to play on team {}".format(playerpool[-1].first_name, request_team_num))
                playerpool.pop()
    
    # Checks row by row for player requests and accomadates them.
    for i, row in enumerate(worksheet.values):
        request_string = row[request_col]
        if request_string:
            if request_string[:10].lower() == "req player":
                request_name = request_string[11:]
                request_name_split = request_name.split()
                first = request_name_split[0].lower()
                last = request_name_split[1].lower()
                for player in playerpool:
                    if player.first_name.lower() == first:
                        if player.last_name.lower() == last:
                            playerpool[i-1].request = player
                            player.request = playerpool[i-1]

    heapq.heapify(team_heap)
    
    return playerpool, team_heap


col_num = {}
alphabet_string = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
for i, letter in enumerate(alphabet_string):
    col_num[letter] = i


print()
while True:
    filename = input("What file would you like to sort?: ")
    print()

    try:
        wb = load_workbook(filename, data_only=True)
        break

    except:
        print("Error: Could not load file.\n")

print("File Loaded. \n")
print("Available Worksheets are:")

for i in wb.sheetnames:
    print(i)
print()

while True:
    sheetname = input("Which worksheet would you like to sort?:")
    print()

    try:
        ws = wb[sheetname]
        break
    except KeyError:
        print("Error: Incorrect worksheet name.")

number_of_teams = int(
    input("How many teams would you like for this sort?:")
)

playerpool, team_heap = load_players(ws, col_num["D"], col_num["E"], col_num["L"], col_num["O"], col_num["P"], number_of_teams)

print("There are {} players detected in worksheet".format(len(playerpool)))

team_list = team_sort(playerpool, team_heap)

for team in team_list:
    print("Team {}".format(team.number))
    for player in team.player_list:
        print(player)
    print("------------------------------")
    