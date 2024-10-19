import requests
import logging
from datetime import datetime
import xlwt 
from xlwt import Workbook 
from openpyxl import Workbook
from openpyxl import load_workbook
import os

sheetName = "showdown.xlsx"
accounts = []
gameDownloadsPath = "./replay_downloads/"
with open('./accounts.txt') as f:
    allLines = f.readlines()
    if len(allLines) == 0:
        raise ValueError("No User Accounts Specified in Accounts")
    else:
        accounts = allLines

def intToColumnLetter(int):
    #only does up to two digits which isnt ideal but i doubt i would ever have more than 26^2 columns

    alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    begining = ''
    end = alphabet[int % 26]

    if int >= 26:
        begining = alphabet[(int // 26) - 1]

    return begining + end

def logDictionaryToInputList(logDictionary):
    outList = []
    dictionaryKeys = list(logDictionary.keys())
    for x in dictionaryKeys:
        if not (x == "p1PokemonList" or x == "p2PokemonList"):
            outList.append(logDictionary[x])

    return outList

def gameLogTodictionary(fileName, accountList):

    outList = None

    with open(fileName, encoding="utf-8", errors='ignore') as f:
        outList = f.read().split('\n')

    gameLogDictionary = {
        "format" : (fileName.split("/")[-1].split("-")[0]),
        "p1": None,
        "p2": None,
        "p1PokemonList": [],
        "p1Poke0" : None, "p1Poke1" : None, "p1Poke2" : None, "p1Poke3" : None, "p1Poke4" : None, "p1Poke5" : None,
        "p2Poke0" : None, "p2Poke1" : None, "p2Poke2" : None, "p2Poke3" : None, "p2Poke4" : None, "p2Poke5" : None,
        "p2PokemonList": [],
        "p1RatingStart": None,
        "p1RatingEnd": None,
        "p2RatingStart": None,
        "p2RatingEnd": None,
    }

    allTimes = []
    turnTimes = []
    
    for x in outList:
        if x[0:6] == "|poke|":
            pokemonName = x[9:].split(',')[0].split('|')[0]

            if x[6:8] == "p1":
                gameLogDictionary["p1PokemonList"].append(pokemonName)
            if x[6:8] == "p2":
                gameLogDictionary["p2PokemonList"].append(pokemonName)

        if x[0:4] == "|t:|":
            allTimes.append(x[4:])
        
        if x[0:6] == "|turn|":
            turnTimes.append(allTimes[-1])

        if x[0:5] == "|raw|" and not x[0:9] == "|raw|<div" and not x[0:10] == "|raw|<font": 
            ratingSplit = x.split("rating:")
            if len(ratingSplit) >1:
                
                endRating = ratingSplit[1].split("<")[1].split(">")[1]
                if gameLogDictionary["p1RatingEnd"] == None:
                    gameLogDictionary["p1RatingEnd"] = endRating
                else:
                    gameLogDictionary["p2RatingEnd"] = endRating

        if x[0:8] == "|player|":
            if len(x)>16:
                rating = x[-4:]

                playerName = x[11:].split("|")[0]

                if x[0:11] == "|player|p1|" and gameLogDictionary["p1"] == None:
                    gameLogDictionary["p1"] = playerName
                    gameLogDictionary["p1RatingStart"] = rating
                elif x[0:11] == "|player|p2|" and gameLogDictionary["p2"] == None:
                    gameLogDictionary["p2"] = playerName
                    gameLogDictionary["p2RatingStart"] = rating

    if len(turnTimes) == 0:
        gameLogDictionary["date"] = datetime.fromtimestamp(int(allTimes[0])).strftime("%Y-%m-%d") 
        gameLogDictionary["timeStart"] = datetime.fromtimestamp(int(allTimes[0])).strftime("%H:%M:%S") 
        gameLogDictionary["timeFinish"] = datetime.fromtimestamp(int(allTimes[-1])).strftime("%H:%M:%S") 
        gameLogDictionary["turnCount"] = 0 
    else:
        gameLogDictionary["date"] = datetime.fromtimestamp(int(allTimes[0])).strftime("%Y-%m-%d") 
        gameLogDictionary["timeStart"] = datetime.fromtimestamp(int(allTimes[0])).strftime("%H:%M:%S") 
        gameLogDictionary["timeFinish"] = datetime.fromtimestamp(int(allTimes[-1])).strftime("%H:%M:%S") 
        gameLogDictionary["turnCount"] = len(turnTimes)

    if any(x.lower().rstrip() == gameLogDictionary["p1"].lower().rstrip() for x in accountList):
        pass
    elif any(x.lower().rstrip() == gameLogDictionary["p2"].lower().rstrip() for x in accountList):

        p1Save = gameLogDictionary["p1"]
        p1RatingStartSave = gameLogDictionary["p1RatingStart"]
        p1RatingEndSave = gameLogDictionary["p1RatingEnd"]
        p1PokemonListSave = gameLogDictionary["p1PokemonList"]


        gameLogDictionary["p1"] = gameLogDictionary["p2"]
        gameLogDictionary["p1RatingStart"] = gameLogDictionary["p2RatingStart"]
        gameLogDictionary["p1RatingEnd"] = gameLogDictionary["p2RatingEnd"]
        gameLogDictionary["p1PokemonList"] = gameLogDictionary["p2PokemonList"]


        gameLogDictionary["p2"] = p1Save
        gameLogDictionary["p2RatingStart"] = p1RatingStartSave
        gameLogDictionary["p2RatingEnd"] = p1RatingEndSave
        gameLogDictionary["p2PokemonList"] = p1PokemonListSave
        
    else:
        raise ValueError("no user recognized")

    teamKeys = ["Poke" + str(x) for x in range(6)]

    for x in range (2):
        for y in range (len(teamKeys)):
            currentPlayerPokeList = gameLogDictionary["p"+str(x+1)+"PokemonList"]

            if len(currentPlayerPokeList) <= y:
                gameLogDictionary["p"+str(x+1)+str(teamKeys[y])] = "N/A"
            else:
                gameLogDictionary["p"+str(x+1)+str(teamKeys[y])] = currentPlayerPokeList[y]

    return(gameLogDictionary)

def addGameToSheet(fileName, gameDirectory, listOfAccounts):
    rowToCheck = 2
    unfilledRowFound = False

    gameDictionary = gameLogTodictionary(gameDirectory, listOfAccounts)
    outputList = logDictionaryToInputList(gameDictionary)
    gameListKeys = list(gameDictionary.keys())

    workbook = load_workbook(filename=fileName)
    workbook.sheetnames
    sheet = workbook.active

    while not unfilledRowFound:
        if (sheet["A"+str(rowToCheck)].value == None ):
            unfilledRowFound = True
            keyIndex = 0
            sheetIndex = 0
            sheet["A"+str(rowToCheck)] = sheet["A"+str(rowToCheck-1)].value + 1

            while keyIndex != len(gameListKeys):
                indexOfCurrentKey = gameListKeys[keyIndex]

                if not (indexOfCurrentKey == "p1PokemonList" or indexOfCurrentKey == "p2PokemonList"):
                    sheet[intToColumnLetter(sheetIndex+1) + str(rowToCheck)] =  indexOfCurrentKey
                    sheetIndex += 1
                keyIndex = keyIndex + 1

            for x in range (len(outputList)):
                sheet[intToColumnLetter(x+1) + str(rowToCheck)] =  outputList[x]
                
        else:
            rowToCheck += 1

    workbook.save(filename=fileName)

    return True

def addAllGamesToSheet():
    dir_list = os.listdir(gameDownloadsPath)

    for game in dir_list:
        if(addGameToSheet(sheetName, (gameDownloadsPath + game), accounts)):
            os.remove(gameDownloadsPath + game)



addAllGamesToSheet()