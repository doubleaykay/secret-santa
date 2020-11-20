# Secret Santa Generator and Mailer
# Anoush Khan, November 2020

import random
import pandas as pd
import pickle

# establish arguments and grab them from commandline
# arg: path to excel
# arg: optional no go text file (add logic to not run no go check if need be)
# arg: path to pickle file

# load excel into pandas data frame
path_to_excel = 'Secret_Santa_2020_Responses.xlsx'
data_pd = pd.read_excel(path_to_excel)

# get Names column as list
names = data_pd['Name'].tolist()
# print(names)

# get and construct no-go pairings
path_no_go = 'no-go.txt'
no_go_list = []
with open(path_no_go) as no_go_txt:
    no_go_lines = no_go_txt.readlines()
for line in no_go_lines:
    tokens = line.split(', ')
    no_go_list.append( (tokens[0], tokens[1]) )
    no_go_list.append( (tokens[1], tokens[0]) )
# print(no_go_list)

# shuffle names and generate pairings
good = False

while not good:
    names_shuffled = names
    random.shuffle(names_shuffled)
    random.shuffle(names_shuffled)
    random.shuffle(names_shuffled) # 3 times just to be safe
    # print(names_shuffled)

    pairings = []
    for x, y in zip(names_shuffled, names_shuffled[1:]):
        pairings.append( (x, y))
    pairings.append((names_shuffled[-1], names_shuffled[0]))

    no_go_in_pairings = any(item in pairings for item in no_go_list)
    if no_go_in_pairings:
        # print('oopsie')
        good = False
    else:
        good = True

print(pairings)

# pickle the list of pairings

# send emails to the pairs