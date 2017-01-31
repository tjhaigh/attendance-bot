#!/usr/bin/python3

import discord
import asyncio
import httplib2
import os
import logging
import time
import configparser
from datetime import datetime, timedelta
from apiclient import discovery
import oauth2client
from oauth2client import client,tools

# Get keys from seperate file
config = configparser.ConfigParser()
config.readfp(open('config.ini'))
SHEET_ID = config.get('Keys', 'sheet_id')
TOKEN = config.get('Keys', 'token')

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'secret.json'
APPLICATION_NAME = 'Attendance Bot'
LOG_FILE = 'attendance.txt'
DISCORD_CHANNEL = 'Ready for Invite'
#SHEET_ID = '1lHNhWJu2W0oT49RGieyH7PUYgPl-d9saFzRyNcQMu_w'
# Test ^^^
MISSING_FILE = 'missing.txt'

HELP_MSG = 'attendance-bot help:\n' \
         + 'Commands (must be done in "missingraid"):\n' \
         + '\t.missing - lets officers know when you will be missing raid\n' \
         + '\t\tsyntax: .missing <mm/dd/yy or mm/dd/yy-mm/dd/yy>  <reason>\n' \
         + '\t.late - lets officers know when you will be late to raid\n' \
         + '\t\tCOMING SOON\n' \
         + 'Officer Commands (must be done in "officers"):\n' \
         + '\t.checkmissing - gives a list of who will be missing on a given date\n' \
         + '\t\tsyntax: .checkmissing mm/dd/yy\n' \
         + '\t.attendance - starts the attendance process' \

client = discord.Client()
here = []
missing = []

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-python-quickstart.json')
    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = oauth2client.client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        credentials = tools.run_flow(flow, store)
        print('Storing credentials to ' + credential_path)

    return credentials

def take_attendance():
    for server in client.servers:
        for member in server.members:
            mv = member.voice
            if str(mv.voice_channel) == DISCORD_CHANNEL:
                name = str(member.nick)
                if name == 'None':
                    name = str(member)
                    name = name[:-5]
                here.append(name.lower())
    print(here)
    # Log who it saw in discord
    with open(LOG_FILE, 'a') as of:
        of.write('[' + get_time() + '] Discord: ' + ','.join(here))
    absent = update_sheet()
    return absent

def update_sheet():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?version=v4')
    service = discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)
    
    rangeName = 'A:D'
    result = service.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=rangeName).execute()
    values = result.get('values', [])
    print(values)

    # lists to check what was changed on the sheet
    heresheet = []
    absentsheet = []

    for idx, val in  enumerate(values):
        if len(val) < 2 or val[0] == 'Toon':
            if val[0] == 'Inactive Raiders': break
            continue
        if val[0].lower() in here:
            heresheet.append(val[0])
            rangeName = 'B' + str(idx+1)
            newval = [[str(int(val[1]) + 1)]]
        else:
            absentsheet.append(val[0]) 
            rangeName = 'D' + str(idx+1)
            newval = [[str(int(val[3]) + 1)]]
        body = { 'values': newval }
        result = service.spreadsheets().values().update(spreadsheetId=SHEET_ID, range=rangeName, valueInputOption='USER_ENTERED', body=body).execute()
        
    # log the changes
    with open(LOG_FILE, 'a') as of:
        t = get_time()
        of.write('\n[' + t + '] Here: ' + ','.join(heresheet)) 
        of.write('\n[' + t + '] Absent: ' + ','.join(absentsheet))
        of.write('\n----------------------------------------------------------\n')
        return absentsheet

def write_missing(date, author, reason):
    missing.append([date, author, reason])
    with open(MISSING_FILE, 'a') as missing_file:
        ds = date.strftime('%m/%d/%y')
        missing_file.write(ds + ' ' + author + ' ' + reason + '\n')


def read_missing_file():
    with open(MISSING_FILE, 'r') as missing_file:
        for line in missing_file:
            parts = line.split()
            ds = parts[0]
            date = datetime.strptime(ds, '%m/%d/%y')
            player = parts[1]
            reason = ""
            if len(parts) > 2:
                reason = ' '.join(parts[2:])
            missing.append([date, player, reason])

@client.event
async def on_ready():
    print(client.user.name)
    print(client.user.id)
    print('--------')
    read_missing_file()
    print(missing)
    #take_attendance()

@client.event
async def on_message(message):
    # Takes attendance
    # Replies with each absent player
    if message.content.startswith('.attendance') and message.channel.name == 'officers':
        await client.send_message(message.channel, 'Taking attendance now')
        absent = take_attendance()
        await client.send_message(message.channel, 'Attendance Done')
        await client.send_message(message.channel, 'Absent players:\n' + ', '.join(absent))
    
    # Checks the list to see who is missing for a given date
    # Replies with each player and the given reason
    if message.content.startswith('.checkmissing') and message.channel.name == 'officers':
        parts = message.content.split()
        try:
            datestring = parts[1]
            date = datetime.strptime(datestring, '%m/%d/%y')
            msg = 'Missing Players for ' + datestring + ': \n'
            for p in missing:
                if p[0] == date:
                    msg += p[1] + ': ' + p[2] + '\n'
            await client.send_message(message.channel, msg)
        except:
            await client.send_message(message.channel, 'Invalid command format. please use ".checkmissing mm/dd/yy" (eg .checkmissing 1/24/17)')

    # Records when a player will be missing raid
    # syntax: .missing mm/dd/yy <reason>(optional)
    if message.content.startswith('.missing') and message.channel.name == 'missingraid':
        parts = message.content.split()
        author = message.author.nick or str(message.author)[:-5]
        try:
            datestring = parts[1]
            reason = ""
            if len(parts) > 2:
                reason = ' '.join(parts[2:])
            if '-' in datestring:
                dates = datestring.split('-')

                d1 = datetime.strptime(dates[0], '%m/%d/%y')
                d2 = datetime.strptime(dates[1], '%m/%d/%y')
                delta = d2 - d1
                for i in range(delta.days + 1):
                    write_missing(d1 + timedelta(days=i), author, reason)
            else:
                date = datetime.strptime(datestring, '%m/%d/%y')
                write_missing(date, author, reason)
            await client.send_message(message.channel, 'Recorded that ' + author + ' will be missing raid on ' + datestring)
        except Exception as e:
            print(e)
            await client.send_message(message.channel, 'Invalid command format. Please use ".missing mm/dd/yy" (eg .missing 1/24/17)')

    # Print Help message
    if message.content.startswith('.help'):
        await client.send_message(message.channel, HELP_MSG)    


def get_time():
    t = time.strftime('%d-%m-%Y %H:%M:%S')
    return t

def main():
    client.run(TOKEN)

if __name__ == '__main__':
    main()
