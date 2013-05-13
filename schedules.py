#!/usr/bin/env python

"""
This scrapes college football team information from
ESPN. It uses BeautifulSoup Alpha 4 and xlwt.

In order to use this, you will need to download bs4, lxml, and xlwt.

Team dictionary format: {Team Name: [ESPN ID, FBS/FCS, Wins, Losses, [Opponent1, Outcome1, Opponent1W, Opponent1L], [Opponent2, Outcome2, Opponent2W, Opponent2L], ... , [OpponentN, OutcomeN, OpponentNW, OpponentNL]]}
ID dictionary format: {ESPN ID: Team Name}
Ranks dictionary format: {Rank Stat: Team Name}
"""

import urllib2
import re
import datetime
import xlwt
import operator
from urlparse import urlparse
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
    i = 4;
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
    outcomes = soup.find_all("ul", re.compile('game-schedule'));
    for oc in outcomes:
        temp = re.split('[><]', oc.encode('ascii'));
        if temp[6] == 'W' or temp[6] == 'L':
            allSchools[school][i].append(temp[6]);
            i+=1;
        elif temp[4] == 'Postponed':
            del allSchools[school][i];

def get_schools():
    global allSchools
    global allIDs
    global allRanks
    url = urllib2.urlopen('http://espn.go.com/college-football/teams');
    print url.geturl();
    soup = bs(url.read(), ['fast', 'lxml']);
    #divisions = soup.
    school_links = soup.find_all(href=re.compile("football/team/_/"));
    for school in school_links[0:124]:
        lenId = len(unicode(school)) - 50
        lenId = lenId + 27
        schID = (school['href'].split('/')[7])
        school = school.string.encode('ascii')
        allSchools[school] = [schID,'FBS']
        allIDs[schID] = school
    for school in school_links[124:]:
        lenId = len(unicode(school)) - 50
        lenId = lenId + 27
        schID = (school['href'].split('/')[7])
        school = school.string.encode('ascii')
        allSchools[school] = [schID,'FCS']
        allIDs[schID] = school

def main():
    get_schools();
    #year = datetime.date.today().year;
    year = 2012
    for school in allSchools:
        print school;
        scrape_links(school, _format_schedule_url(year, allSchools[school][0]));
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
            for team in value[4:]:
                teamName = team[0];
                if len(team) > 1:
                    teamOutc = team[1];
                else:
                    teamOutc = '';
                if(teamName == 'TX A&M-Commerce'):
                    teamName = 'Texas A&M-Commerce'
                    ws.write(i, 1, teamName);
                    ws.write(i, 2, teamOutc);
                    ws.write(i, 3, int(1));
                    ws.write(i, 4, int(9));
                    oppWins = oppWins + 1
                    oppLoss = oppLoss + 9
                elif teamName == 'NW Oklahoma St':
                    teamName = 'Northwestern Oklahoma State'
                    ws.write(i, 1, teamName);
                    ws.write(i, 2, teamOutc);
                    ws.write(i, 3, int(4));
                    ws.write(i, 4, int(7));
                    oppWins = oppWins + 4
                    oppLoss = oppLoss + 7
                else:
                    ws.write(i, 1, xlwt.Formula('HYPERLINK("#' + "'" + teamName + "'" + '!A1";"' + teamName + '")'));
                    ws.write(i, 2, teamOutc);
                    if teamOutc != '':
                        if teamOutc == 'W':
                            offset = 0;
                        else:
                            offset = 1;
                        ws.write(i, 3, int(allSchools[teamName][2])-offset);
                        ws.write(i, 4, int(allSchools[teamName][3])-(1-offset));
                        oppWins = oppWins + int(allSchools[teamName][2])-offset
                        oppLoss = oppLoss + int(allSchools[teamName][3])-(1-offset)
                try: #For non-DII/III Schools
                    ws.write(i, 5, xlwt.Formula("'" + teamName + "'!A6"));
                except:
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
            allRanks[key] = teamRec*oppRec
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
            for team in value[4:]:
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
    for rank in sortedRanks:
        print rank
        ws1.write(ws1Row, ws1Col, rank[1]);
        ws1.write(ws1Row, ws1Col+1, xlwt.Formula('HYPERLINK("#' + "'" + rank[0] + "'" + '!A1";"' + rank[0] + '")'));
        #if ws1Row == 19:
        #    ws1Row = 0;
        #    ws1Col+=3;
        #else:
        ws1Row+=1;
    wb.save('AllTeams.xls');

if __name__ == '__main__':
    import time
    start = time.time()
    main()
    print time.time() - start, 'seconds'
