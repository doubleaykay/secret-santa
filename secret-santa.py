# Secret Santa Generator and Mailer
# Anoush Khan, November 2020

import random
import pandas as pd
import pickle
from os import path
import smtplib

# establish arguments and grab them from commandline
# arg: path to excel
path_to_excel = 'Secret_Santa_2020_Responses.xlsx'
# arg: optional no go text file (add logic to not run no go check if need be)
path_no_go = 'no-go.txt'
# arg: path to pickle file
path_pickle_file = 'matches'

# get gmail username and password
gmail_user = input('Gmail Username: ')
gmail_password = input('Gmail password: ')

# load excel into pandas data frame
data_pd = pd.read_excel(path_to_excel)

# get Names column as list
names = data_pd['Name'].tolist()
# print(names)

# get and construct no-go pairings
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
# if not path.exists(path_pickle_file):
pickle.dump( pairings, open( path_pickle_file, "wb+" ) )

# initiate SMTP connection
try:
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
except:
    print('Something went wrong...')

# send emails to the pairs
for pair in pairings:
    # get dataframe rows for both people
    p1 = data_pd.loc[data_pd['Name'] == pair[0]]
    p2 = data_pd.loc[data_pd['Name'] == pair[1]]
    
    # get email address of person 1
    p1_email = p1['Email Address'].values[0]
    print(p1_email)

    # get info for person 2
    p2_email = p2['Email Address'].values[0]
    p2_address = p2['Address'].values[0]
    # print(p2_address)

    # tell person 1 they need to gift to person 2
    message_body = f"""Subject: SECRET SANTA MATCH

    Hey {pair[0]}! The person you are sending a gift to is {pair[1]}.
    Their email address is {p2_email}.
    Their address is {p2_address}.

    Make sure the gift is delivered to them NO LATER than December 20, 2020!

    Message Anoush if you have any questions.

    Happy gifting!
    """

    # send the email
    server.sendmail(gmail_user, p1_email, message_body)

# close the mail server
server.close()
