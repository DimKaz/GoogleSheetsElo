from __future__ import print_function
import pickle
import os.path
import trueskill
from trueskill import (
    quality, quality_1vs1, rate, rate_1vs1, Rating, setup, TrueSkill)
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
ELOSheetID = '140bFCoPDWE6gl4iaq0xUiE7j6TWsacAJ1w654bLnFQI'

creds = None

def sheetsSetup():
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    global service
    service = build('sheets', 'v4', credentials=creds)



def processNewTourney():
    # Get matches
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ELOSheetID,
                                range='TrueSkill!A3:F').execute()
    matchList = result.get('values', [])
    #End of Google Sheets API Code

    #Get players
    result = sheet.values().get(spreadsheetId=ELOSheetID,
                                range='TrueSkill!L3:N').execute()
    tempPlayerList = result.get('values',[])

    #Put players in string list
    playerList = list()
    ratingList = list()
    for player in tempPlayerList:
        playerList.append(player[0])                     #
        ratingList.append(Rating(float(player[1]), float(player[2])))  # Create new player rating

    #print(playerList)
    #print(ratingList)
    #print('\n')
    #print(matchList)
    #print('\n')

    index = 0
    for match in matchList:
        WinnerName = match[0]
        LoserName = match[1]
        WinnerIndex = playerList.index(WinnerName)
        LoserIndex = playerList.index(LoserName)
        ratingList[WinnerIndex], ratingList[LoserIndex] = rate_1vs1(ratingList[WinnerIndex], ratingList[LoserIndex])
        print('%s was winner. Their new rating is %f' % (WinnerName, ratingList[WinnerIndex]))
        print('%s was loser. Their new rating is %f' % (LoserName, ratingList[LoserIndex]))
        matchList[index][4] = ratingList[WinnerIndex].mu
        matchList[index][5] = ratingList[LoserIndex].mu
        index = index + 1
    print()
    
    #Format data into json format
    data = [
        {
            'range': 'TrueSkill!A3:F',
            'values': matchList
        },
    ]
    body = {
        'valueInputOption': 'RAW',
        'data': data
    }
    
    result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=ELOSheetID, body=body).execute()

    env = TrueSkill()
    env.create_rating()
    leaderboard = sorted(ratingList, key=env.expose, reverse=True)
    position=1
    mu = list()
    sigma = list()
    
    print('---Leaderboard---')
    for player in playerList:
        mu.append(ratingList[playerList.index(player)].mu)
        sigma.append(ratingList[playerList.index(player)].sigma)
        print('Rank %d: %s - %f' % (position, player, leaderboard[playerList.index(player)]))
        position=position+1
        
    
    #Update final ratings
    #Format data into json 
    data = [
        {
            'range': 'TrueSkill!H3:J',
            'values': list(zip(playerList,mu,sigma))
        },
    ]
    body = {
        'valueInputOption': 'RAW',
        'data': data
    }
    
    result = service.spreadsheets().values().batchUpdate(
        spreadsheetId=ELOSheetID, body=body).execute()
    
    
    
    
    

	
if __name__ == '__main__':
    sheetsSetup()
    processNewTourney()
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	