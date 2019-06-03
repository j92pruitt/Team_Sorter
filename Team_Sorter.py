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

    def __init__(self, first_name, last_name, rating):
        self.first_name = first_name
        self.last_name = last_name
        self.rating = rating

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


def sort(playerpool, num):
    """
    This function randomly divides a list of players evenly 
    into a given number of teams.

    Parameters:
        playerpool (list): List of Players to be sorted
        num (int): Number of teams to sort into.

    Returns:
        team_heap (list): List of teams.
    """
    
    team_heap = []
    for i in range(num):
        heapq.heappush(team_heap, Team(i+1))

    while playerpool:
        idx = randint(0, len(playerpool)-1)
        team = heapq.heappop(team_heap)

        team.add_player(playerpool.pop(idx))
        heapq.heappush(team_heap, team)

    return team_heap
    