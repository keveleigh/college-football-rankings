#!/usr/bin/env python

"""
This scrapes college football team information from
ESPN. It uses BeautifulSoup 4 and xlwt.

In order to use this, you will need to download bs4, lxml, and xlwt.

Team dictionary format: {Team Name: [ESPN ID, FBS/FCS, Wins, Losses, {Opponent1, Opponent2, ... , OpponentN}, [Opponent1, Outcome1], [Opponent2, Outcome2], ... , [OpponentN, OutcomeN]]}
ID dictionary format: {ESPN ID: Team Name}
Ranks dictionary format: {Team Name: Rank Stat}

Command format: python schedules.py [reuse] [year]
"""

import urllib2
import re
import datetime
import xlwt
import operator
import sys
import ast
import os
from bs4 import BeautifulSoup as bs

allSchools = {}
allIDs = {}
allRanks = {}

def _format_schedule_url(year, idNum):
    """Format ESPN link to scrape individual records from."""
    link = ['http://espn.go.com/college-football/team/schedule/_/id/' + idNum + '/year/' + str(year) + '/'];
    return link[0];

def scrape_links(school, espn_schedule):
    """Scrape ESPN's pages for data."""
    global allSchools;
    url = urllib2.urlopen(espn_schedule);
##    print url.geturl();
    soup = bs(url.read(), ['fast', 'lxml']);

    record = soup.find('div', attrs={'id':'showschedule'})
    record = record.find_all(text=re.compile('\d{1,2}-\d{1,2}'));
    record = record[len(record)-1].string.encode('ascii');
    record = record.split(' ');
    record = record[0].split('-');
    allSchools[school].append(record[0]);
    allSchools[school].append(record[1]);

    opponents = soup.find_all("li", "team-name");
    outcomes = soup.find_all("ul", re.compile('game-schedule'));
    allSchools[school].append({});
    i = 5;
    j = 1;
    for opp in opponents:
        tempName = re.split('[><]', opp.encode('ascii'));
        oppID = re.split('[/]', tempName[3]);
        if len(oppID) >= 8:
            oppID = oppID[7]
            oppName = allIDs[oppID]
        else:
            oppName = tempName[2]
            oppName = re.sub('&amp;', '&', oppName, flags=re.IGNORECASE)
        allSchools[school].append([oppName]);
        if(j < len(outcomes)):
            temp = re.split('[><]', outcomes[j].encode('ascii'));
            if temp[6] == 'W' or temp[6] == 'L':
                allSchools[school][i].append(temp[6]);
                allSchools[school][4][oppName] = temp[6]; # Will cause issues when team is played twice
                i+=1;
                j+=2;
            elif temp[4] == 'Postponed':
                del allSchools[school][i];
                j+=2;

def calculate_score(school1):
    global allSchools;
    
    school1Info = allSchools[school1];
    school1Ops = school1Info[4];
    teamScore = 0;
    for school2, result in school1Ops.iteritems():
        if school2 not in allSchools or allSchools[school2][1] != 'FBS':
            continue;
        
        school2OP = calculate_op(school1, school2);
        
        if result == 'W':
            outcome = 1;
        else:
            outcome = 0;
            
        teamScore = teamScore + (outcome * school2OP);
        if school1 == 'Georgia Tech':
            print school1 + ' teamScore = ' + str(teamScore);
    
    return teamScore;

def calculate_op(school1, school2):
    global allSchools;
    
    school2Info = allSchools[school2];
    school2Ops = school2Info[4];
    teamOP = 0;
    numGames = len(school2Ops);
    for school3, result in school2Ops.iteritems():
        if school3 not in allSchools or allSchools[school3][1] != 'FBS':
            continue;
        
        if school3 == school1:
            numGames = numGames - 1;
            continue;
        
        school3OOP = calculate_oop(school1, school2, school3);
            
        if result == 'W':
            outcome = 1;
        else:
            outcome = 0;
            
        teamOP = teamOP + (outcome * school3OOP);
        if school1 == 'Georgia Tech':
            print school2 + ' OP = ' + str(teamOP) + ' ' + str(outcome) + ' ' + str(school3OOP);
        
    return teamOP/numGames;
    
def calculate_oop(school1, school2, school3):
    global allSchools;
    
    school3Info = allSchools[school3];
    school3Ops = school3Info[4];
    teamOOP = 0;
    numGames = len(school3Ops);
    for school4, result in school3Ops.iteritems():
        if school4 not in allSchools or allSchools[school4][1] != 'FBS':
            continue;
        
        if school4 == school1:
            numGames = numGames - 1;
            continue;
        if school4 == school2:
            numGames = numGames - 1;
            continue;
                
        school4Info = allSchools[school4];
        school4W = int(school4Info[2]);
        school4L = int(school4Info[3]);
        
        if school2 in school4Info[4]:
            if school4Info[4][school2] == 'W':
                school4W = school4W - 1;
            else:
                school4L = school4L - 1;
            
        if school1 in school4Info[4]:
            if school4Info[4][school1] == 'W':
                school4W = school4W - 1;
            else:
                school4L = school4L - 1;
            
        if result == 'W':
            outcome = 1;
            school4L = school4L - 1;
        else:
            outcome = 0;
            school4W = school4W - 1;
            
        school4WinPerc = float(school4W) / (school4W+school4L);
        teamOOP = teamOOP + (outcome * school4WinPerc);
        if school1 == 'Georgia Tech':
            print school3 + ' teamOOP ' + school4 + ' = ' + str(teamOOP);
        
    return teamOOP/numGames;

def get_schools():
    global allSchools
    global allIDs
    global allRanks
    url = urllib2.urlopen('http://espn.go.com/college-football/teams');
    print url.geturl();
    soup = bs(url.read(), ['fast', 'lxml']);
    #divisions = soup.
    school_links = soup.find_all(href=re.compile("football/team/_/"));
    for school in school_links[0:126]:
        lenId = len(unicode(school)) - 50
        lenId = lenId + 27
        schID = (school['href'].split('/')[7])
        school = school.string.encode('ascii')
        allSchools[school] = [schID,'FBS']
        allIDs[schID] = school
    for school in school_links[126:]:
        lenId = len(unicode(school)) - 50
        lenId = lenId + 27
        schID = (school['href'].split('/')[7])
        school = school.string.encode('ascii')
        allSchools[school] = [schID,'FCS']
        allIDs[schID] = school

def main(argv):
    global allSchools
    year = datetime.date.today().year;
    if len(argv) < 2:
        year = str(2012)
    else:
        year = argv[1]
    if len(argv) > 0 and argv[0] == 'reuse' and os.path.isfile('teams'+str(year)+'.txt'):
        f = open('teams'+str(year)+'.txt', 'r')
        allSchools = ast.literal_eval(f.read())
        for key in allSchools:
            print key
    else:
        get_schools()
        for school in allSchools:
            print school
            scrape_links(school, _format_schedule_url(year, allSchools[school][0]))
        f = open('teams'+str(year)+'.txt', 'w')
        f.write(str(allSchools))
        f.close()

    j = 1;
    wb = xlwt.Workbook()
    items = allSchools.items()
    items.sort()
    ws1 = wb.add_sheet('FBS Main Page');
    ws1Col = 0;
    ws1Row = 0;

    for key, value in items:
        if value[1] == 'FBS':
            wb.add_sheet(key);

    ws2 = wb.add_sheet('FCS Main Page');

    for key, value in items:
        if value[1] == 'FCS':
            wb.add_sheet(key);

    for key, value in items:
        if value[1] == 'FBS':
            longestOpp = 0;
            ws = wb.get_sheet(j);
            ws.write(0, 0, xlwt.Formula('HYPERLINK("#' + "'FBS Main Page'" + '!A1";"' + key + '")'));
            ws.write(1, 0, value[1]);
            ws.write(2, 0, value[2] + '-' + value[3]);
            teamRec = float(value[2]) / (int(value[2]) + int(value[3])) # Fix?
            ws.write(3, 0, teamRec)

            i = 1;
            oppWins = 0
            oppLoss = 0
            for team in value[5:]:
                teamName = team[0];
                teamWin = 0
                teamLoss = 0
                if len(team) > 1:
                    teamOutc = team[1];
                else:
                    teamOutc = '';
                if(teamName == 'TX A&M-Commerce') and year == '2012':
                    teamName = 'Texas A&M-Commerce'
                    ws.write(i, 1, teamName);
                    teamWin = 1
                    teamLoss = 9
                elif teamName == 'NW Oklahoma St' and year == '2012':
                    teamName = 'Northwestern Oklahoma State'
                    ws.write(i, 1, teamName);
                    teamWin = 4
                    teamLoss = 7
                elif(teamName == 'West Alabama') and year == '2011':
                    ws.write(i, 1, teamName);
                    teamWin = 8
                    teamLoss = 4
                elif teamName == 'Henderson St' and year == '2011':
                    teamName = 'Henderson State'
                    ws.write(i, 1, teamName);
                    teamWin = 6
                    teamLoss = 4
                elif teamName == 'Tarleton St' and year == '2011':
                    teamName = 'Tarleton State'
                    ws.write(i, 1, teamName);
                    teamWin = 6
                    teamLoss = 5
                elif teamName == 'NE St' and year == '2011':
                    teamName = 'Northeastern State'
                    ws.write(i, 1, teamName);
                    teamWin = 7
                    teamLoss = 5
                else:
                    ws.write(i, 1, xlwt.Formula('HYPERLINK("#' + "'" + teamName + "'" + '!A1";"' + teamName + '")'));
                    teamWin = int(allSchools[teamName][2])
                    teamLoss = int(allSchools[teamName][3])
                ws.write(i, 2, teamOutc);
                if teamOutc != '':
                    if teamOutc == 'W':
                        offset = 0;
                    else:
                        offset = 1;
                    ws.write(i, 3, teamWin-offset);
                    ws.write(i, 4, teamLoss-(1-offset));
                    if teamName in allSchools and allSchools[teamName][1] == 'FBS':
                        oppWins = oppWins + teamWin-offset
                        oppLoss = oppLoss + teamLoss-(1-offset)
                    elif teamOutc == 'L':
                        oppLoss = oppLoss + 12
                try: #For non-DII/III Schools
                    ws.write(i, 5, xlwt.Formula("'" + teamName + "'!A6"));
                except:
                    ws.write(i, 5, 0)
                    print teamName;
                if len(teamName) > longestOpp:
                    longestOpp = len(teamName);
                i+=1;
            oppRec = float(oppWins) / (oppWins + oppLoss)
            ws.write(i+1, 3, oppWins);
            ws.write(i+1, 4, oppLoss);
            ws.write(4, 0, oppRec);
            ws.write(5, 0, teamRec * oppRec);
            ws.col(0).width = len(key)*350;
            ws.col(1).width = longestOpp*300;
            
#             score = teamRec*oppRec;
            score = calculate_score(key);
#             print key;
#             print score;
            allRanks[key] = score;
            j+=1;

    j+=1;
    for key, value in items:
        if value[1] == 'FCS':
            longestOpp = 0;
            ws = wb.get_sheet(j);
            ws.write(0, 0, xlwt.Formula('HYPERLINK("#' + "'FCS Main Page'" + '!A1";"' + key + '")'));
            ws.write(1, 0, value[1]);
            ws.write(2, 0, value[2] + '-' + value[3]);
            if(value[3] == '0'):
                ws.write(3, 0, 1.000);
            else:
                ws.write(3, 0, float(value[2])/(float(value[2])+float(value[3])));
            i = 1;
            for team in value[5:]:
                teamName = team[0];
                if len(team) > 1:
                    teamOutc = team[1];
                else:
                    teamOutc = '';
                if len(teamName) > longestOpp:
                    longestOpp = len(teamName);
                ws.write(i, 1, xlwt.Formula('HYPERLINK("#' + "'" + teamName + "'" + '!A1";"' + teamName + '")'));
                ws.write(i, 2, teamOutc);
                i+=1;
            ws.col(0).width = len(key)*350;
            ws.col(1).width = longestOpp*300;
            j+=1;

    sortedRanks = sorted(allRanks.items(), key=operator.itemgetter(1), reverse=True)
#     highestRank = sortedRanks[0][1];
    for rank in sortedRanks:
        ws1.write(ws1Row, ws1Col, rank[1]);
        ws1.write(ws1Row, ws1Col+1, xlwt.Formula('HYPERLINK("#' + "'" + rank[0] + "'" + '!A1";"' + rank[0] + '")'));
        #if ws1Row == 19:
        #    ws1Row = 0;
        #    ws1Col+=3;
        #else:
        ws1Row+=1;
    wb.save('AllTeams'+str(year)+'.xls');

if __name__ == '__main__':
    import time
    start = time.time()
    main(sys.argv[1:])
    print time.time() - start, 'seconds'
