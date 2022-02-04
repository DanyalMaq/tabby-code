import json
from urllib.response import addinfo
import requests
import openpyxl as op
from pprint import pprint

def parseExcel(path):
    sheet = op.load_workbook(path)['Sheet1']
    institutions = []
    teams = []
    for row in sheet.iter_rows(values_only=True):
        temp = [item for item in row if item != None]
        if len(temp) == 0:
            break
        institution = temp[0]
        if temp[0] not in institutions: institutions.append(temp.pop(0)) 
        else: temp.pop(0)
        teamName = temp.pop(0)
        names = [item for item in temp if ('@' not in item)]
        emails = [item for item in temp if ('@' in item)]
        teams.append({
            'team': teamName,
            'names': names,
            'emails': emails,
            'institution': institution
        })
    institutions.pop(0)
    teams.pop(0)
    return {'institutions': institutions, 'teams': teams}

def addInstitution(site, token, name, code, region):
    url = f'https://{site}.calicotab.com/api/v1/institutions'
    myheader = {'Authorization': f'{token}'}
    obj = {
        'name':f'{name}',
        'code':f'{code}',
        'region':f'{region}'
    }
    print(obj)
    requests.post(url, json=obj, headers=myheader)

def makeSpeakers(names, emails):
    speakerList = []
    for name, email in zip(names, emails):
        obj = {
            "name": f'{name}',
            "gender": "",
            "email": f'{email}',
            "phone": "",
            "anonymous": False,
            "pronoun": "",
            "categories": [],
            "url_key": ""
        }
        speakerList.append(obj)

    return speakerList

def addTeam(site, tournament, token, teamName, shortName, allInstitutions, institution, names, emails):
    url = f'https://{site}.calicotab.com/api/v1/tournaments/{tournament}/teams'
    myheader = {'Authorization': f'{token}'}
    institutionURI = getInstitutionURI(allInstitutions, institution)
    speakers = makeSpeakers(names, emails)
    obj = {
        "reference": teamName,
        "short_reference": shortName,
        "code_name": "",
        "emoji": "",
        "institution": institutionURI,
        "speakers": speakers,
        "use_institution_prefix": False,
        "break_categories": [],
        "institution_conflicts": []
    }
    requests.post(url, json=obj, headers=myheader)
        

def getInstitutionURI(institutions, name):
    for item in institutions:
        if str(item['name']) == str(name):
            return item['url']

def getInstitutionList(site, tournament, token):
    url = f'https://{site}.calicotab.com/api/v1/institutions'
    myheader = {'Authorization': f'{token}'}
    institutions = requests.get(url, headers=myheader).json()
    return institutions

def main():
    token = "Token b537c72bd8e11e324bdd6d10fb8b9bba1c7a4ed5"
    site = 'tns'
    tournament = 'tns22'
    institutions = getInstitutionList(site, tournament, token)
    data = parseExcel('test.xlsx')
    
    # for item in data['institutions']:
    #     addInstitution(site, token, item, item[0:2], 'Pakistan')

    # for item in data['teams']:
    #     addTeam(site, tournament, token, item['team'], item['team'], institutions, item['institution'], item['names'], item['emails'])
    
main()