# -*- coding: utf-8 -*-

import os, json,argparse, getpass
import xlrd
import linkedin_api
from linkedin_api import Linkedin
import pandas as pd


parser = argparse.ArgumentParser()
parser.add_argument("infile", type=str,
                    help="input xlsx file")
#parser.add_argument("--nqueries", type=int, 
#                    help="number of lines to be read from input file", default=10)
parser.add_argument("--firstRow", type=int, 
                    help="first row to be considered in the input file", default=1)
parser.add_argument("--lastRow", type=int, 
                    help="last row to be considered in the input file", default=11)
usern=input("Username (email):")
passw = getpass.getpass()

args = parser.parse_args()

infile = xlrd.open_workbook(args.infile)
insheet = infile.sheet_by_index(0)
personcodes = insheet.col_values(0)
fullnames = insheet.col_values(4)
names = insheet.col_values(5)
surnames = insheet.col_values(6)
companies = insheet.col_values(10)

personcodes = personcodes[args.firstRow:args.lastRow+1]
fullnames   = fullnames[args.firstRow:args.lastRow+1]      
names       = names[args.firstRow:args.lastRow+1]          
surnames    = surnames[args.firstRow:args.lastRow+1]       
companies   = companies[args.firstRow:args.lastRow+1]      


fullnames = [f.lower().split() for f in fullnames]
names = [n.lower().split() for n in names]
surnames = [s.lower().split() for s in surnames]
companies = [c.lower().split() for c in companies]

for f in fullnames:
    for ff in f:
        if ff in ["mrs","mr","di", "de"]:
            f.remove(ff)

for c in companies:
    for cc in c:
        if cc in ["srl","a","limitata", "s.r.l.", "societa", "responsabilita", "responsabilita'"]:
            c.remove(cc)            

names = [n  if n != [] else [f[0]] for n,f in zip(names,fullnames)]
search_keys=[]
for f,c in zip(fullnames,companies):
    thisk=[f[0], f[-1]]
    for cc in c:
        thisk.append(cc)
    search_keys.append(thisk)

#search_keys = search_keys[1:]
#fullnames = fullnames[1:]
#personcodes = personcodes[1:]
#

#lenSearch=min(len(search_keys),args.nqueries)
lenSearch=len(search_keys)
#profiles=[[]]*len(search_keys)
profiles=[[]]*lenSearch
for k in range(lenSearch):
    print(search_keys[k])
    prevRes=[0,0]
    for ik in range(len(search_keys[k]) - 1):
        res=[0,0]
        nkeys=2+ik
        s = ",".join(search_keys[k][0:nkeys])
        print(s)
        os.system('python3 singleSearch.py %s %s %s'%(s, usern, passw))
        with open('searchResult.json', 'r') as fh:
            res = json.load(fh)
            if ik == 0:
                prevRes = res.copy()
        print(res)
        if len(res)==1: #if a unique match, we take it
            profiles[k]=[res[0]]
            profiles[k][0]["nkeys"]=nkeys
            break
        if len(res)==0: #if no match
            if nkeys == 2: #if only two keys, no chance,fill empty
                profiles[k]=[]
                break
            else: #if more than two keys, take a step back
                res = prevRes.copy()
                profiles[k]=[res[0]]
                profiles[k][0]["nkeys"]=nkeys-1
##                s = ",".join(search_keys[k][0:nkeys-1])
##                print(s)
##                os.system('python3 singleSearch.py %s %s %s'%(s, usern, passw))
##                with open('searchResult.json', 'r') as fh:
##                    res = json.load(fh)
##                    profiles[k]=[res[0]]
                break
            prevRes=res.copy()
#        if nkeys == len(search_keys[k]):
#            if len(res)>0:
#                profiles[k]=[res[0]]
#            else:
#                profiles[k]=[]
            
out_dict = {"personCode": [], "firstName":[], "lastName": [], "location": [], "linkedinProfile": [], "nkeys": [],
            "experience1_location": [], "experience1_companyName": [], "experience1_title": [], "experience1_startDate": [], "experience1_endDate": [], "experience1_description": [],
            "experience2_location": [], "experience2_companyName": [], "experience2_title": [], "experience2_startDate": [], "experience2_endDate": [], "experience2_description": [],
            "experience3_location": [], "experience3_companyName": [], "experience3_title": [], "experience3_startDate": [], "experience3_endDate": [], "experience3_description": [],
            "experience4_location": [], "experience4_companyName": [], "experience4_title": [], "experience4_startDate": [], "experience4_endDate": [], "experience4_description": [],
            "experience5_location": [], "experience5_companyName": [], "experience5_title": [], "experience5_startDate": [], "experience5_endDate": [], "experience5_description": [],
            "experience6_location": [], "experience6_companyName": [], "experience6_title": [], "experience6_startDate": [], "experience6_endDate": [], "experience6_description": [],
            "experience7_location": [], "experience7_companyName": [], "experience7_title": [], "experience7_startDate": [], "experience7_endDate": [], "experience7_description": [],
            "experience8_location": [], "experience8_companyName": [], "experience8_title": [], "experience8_startDate": [], "experience8_endDate": [], "experience8_description": [],
            "experience9_location": [], "experience9_companyName": [], "experience9_title": [], "experience9_startDate": [], "experience9_endDate": [], "experience9_description": [],
            "experience10_location": [], "experience10_companyName": [], "experience10_title": [], "experience10_startDate": [], "experience10_endDate": [], "experience10_description": [],


            "education1_schoolName": [], "education1_degreeName": [], "education1_fieldOfStudy": [], "education1_grade": [], "education1_startDate": [], "education1_endDate": [], "education1_description": [],
            "education2_schoolName": [], "education2_degreeName": [], "education2_fieldOfStudy": [], "education2_grade": [], "education2_startDate": [], "education2_endDate": [], "education2_description": [],
            "education3_schoolName": [], "education3_degreeName": [], "education3_fieldOfStudy": [], "education3_grade": [], "education3_startDate": [], "education3_endDate": [], "education3_description": [],
            "education4_schoolName": [], "education4_degreeName": [], "education4_fieldOfStudy": [], "education4_grade": [], "education4_startDate": [], "education4_endDate": [], "education4_description": [],
            "education5_schoolName": [], "education5_degreeName": [], "education5_fieldOfStudy": [], "education5_grade": [], "education5_startDate": [], "education5_endDate": [], "education5_description": [],
}

###linkedinProfile="/www.linkedin.com/in/"+profiles[k][0]["public_id"]+"/"

##with open('searchResult.json', 'r') as fh:
##    res = json.load(fh)
##    profiles[4]=[res[0]]

for k in range(lenSearch):
    print("k,key, profile",k,search_keys[k],profiles[k])
api = Linkedin(usern, passw, refresh_cookies=True)
for k in range(lenSearch):
#for k in range(4,5):
    for dictk in out_dict.keys():
        out_dict[dictk].append("")
    if len(profiles[k])==0:
        out_dict["firstName"][-1] = (" ".join(fullnames[k]))
        out_dict["personCode"][-1] = personcodes[k]
        out_dict["firstName"][-1]       =     out_dict["firstName"][-1]      .replace('`','')
        out_dict["nkeys"][-1] = 0
        continue
    print(fullnames[k],profiles[k])
    profile = api.get_profile(profiles[k][0]["public_id"])
#    print(profile)

    
    out_dict["firstName"][-1] = profile["firstName"] if "firstName" in profile.keys() else ""
    out_dict["lastName"][-1] = profile["lastName"] if "lastName" in profile.keys() else ""
    out_dict["personCode"][-1] = personcodes[k]
    out_dict["linkedinProfile"][-1] = "/www.linkedin.com/in/"+profiles[k][0]["public_id"]+"/"
    out_dict["location"][-1] = profile["locationName"] if "locationName" in profile.keys() else ""
    out_dict["nkeys"][-1] = profiles[k][0]["nkeys"] 


    #removing all backticks since they mess up exporting to stata format
    out_dict["firstName"][-1]       =     out_dict["firstName"][-1]      .replace('`','')
    out_dict["lastName"][-1]        =     out_dict["lastName"][-1]       .replace('`','')
    out_dict["personCode"][-1]      =     out_dict["personCode"][-1]     .replace('`','')
    out_dict["linkedinProfile"][-1] =     out_dict["linkedinProfile"][-1].replace('`','')
    out_dict["location"][-1]        =     out_dict["location"][-1]       .replace('`','')
    
    theExperience = profile["experience"] if "experience" in profile.keys() else {}
    theExperience = theExperience[:10]

    for iexp in range(len(theExperience)):
        exp                                               = theExperience[iexp]
        out_dict["experience%s_location"%(iexp+1)][-1]    = exp["geoLocationName"] if "geoLocationName" in exp.keys() else ""
        out_dict["experience%s_companyName"%(iexp+1)][-1] = exp["companyName"] if "companyName" in exp.keys() else ""
        out_dict["experience%s_title"%(iexp+1)][-1]       = exp["title"] if "title" in exp.keys() else ""
 #       out_dict["experience%s_description"%(iexp+1)][-1] = exp["description"] if "description" in exp.keys() else ""
        
        timePeriod                                        = exp["timePeriod"] if "timePeriod" in exp.keys() else {}
        startDate =  timePeriod["startDate"] if "startDate" in timePeriod.keys() else {}
        monthS = startDate["month"] if "month" in startDate.keys() else "x"
        yearS = startDate["year"] if "year" in startDate.keys() else "x"
        out_dict["experience%s_startDate"%(iexp+1)][-1]   = str(monthS)+"."+str(yearS)

        endDate =  timePeriod["endDate"] if "endDate" in timePeriod.keys() else {}
        monthE = endDate["month"] if "month" in endDate.keys() else "x"
        yearE = endDate["year"] if "year" in endDate.keys() else "x"
        out_dict["experience%s_endDate"%(iexp+1)][-1]     =  str(monthE)+"."+str(yearE)

        #removing all backticks since they mess up exporting to stata format
        out_dict["experience%s_location"%(iexp+1)][-1]    = out_dict["experience%s_location"%(iexp+1)][-1]    .replace('`','') 
        out_dict["experience%s_companyName"%(iexp+1)][-1] = out_dict["experience%s_companyName"%(iexp+1)][-1] .replace('`','')
        out_dict["experience%s_title"%(iexp+1)][-1]       = out_dict["experience%s_title"%(iexp+1)][-1]       .replace('`','')
        out_dict["experience%s_startDate"%(iexp+1)][-1]   = out_dict["experience%s_startDate"%(iexp+1)][-1]   .replace('`','')
        out_dict["experience%s_endDate"%(iexp+1)][-1]     = out_dict["experience%s_endDate"%(iexp+1)][-1]     .replace('`','')
        
    theEducation = profile["education"] if "education" in profile.keys() else {}
    theEducation = theEducation[:5]
    for ied in range(len(theEducation)):
        ed = theEducation[ied]
        out_dict["education%s_schoolName"%(ied+1)][-1]   = ed["schoolName"] if "schoolName" in ed.keys() else ""
        out_dict["education%s_degreeName"%(ied+1)][-1]   = ed["degreeName"] if "degreeName" in ed.keys() else ""
        out_dict["education%s_fieldOfStudy"%(ied+1)][-1] = ed["fieldOfStudy"] if "fieldOfStudy" in ed.keys() else ""
        out_dict["education%s_grade"%(ied+1)][-1]        = ed["grade"] if "grade" in ed.keys() else ""
#        out_dict["education%s_description"%(ied+1)][-1]  = ed["description"] if "description" in ed.keys() else ""

        timePeriod                                        = ed["timePeriod"] if "timePeriod" in ed.keys() else {}
        startDate =  timePeriod["startDate"] if "startDate" in timePeriod.keys() else {}
        monthS = startDate["month"] if "month" in startDate.keys() else "x"
        yearS = startDate["year"] if "year" in startDate.keys() else "x"
        out_dict["education%s_startDate"%(ied+1)][-1]   = str(monthS)+"."+str(yearS)

        endDate =  timePeriod["endDate"] if "endDate" in timePeriod.keys() else {}
        monthE = endDate["month"] if "month" in endDate.keys() else "x"
        yearE = endDate["year"] if "year" in endDate.keys() else "x"
        out_dict["education%s_endDate"%(ied+1)][-1]     =  str(monthE)+"."+str(yearE)

        out_dict["education%s_schoolName"%(ied+1)][-1]   = out_dict["education%s_schoolName"%(ied+1)][-1]  .replace('`','')
        out_dict["education%s_degreeName"%(ied+1)][-1]   = out_dict["education%s_degreeName"%(ied+1)][-1]  .replace('`','')
        out_dict["education%s_fieldOfStudy"%(ied+1)][-1] = out_dict["education%s_fieldOfStudy"%(ied+1)][-1].replace('`','')
        out_dict["education%s_grade"%(ied+1)][-1]        = out_dict["education%s_grade"%(ied+1)][-1]       .replace('`','')
        out_dict["education%s_startDate"%(ied+1)][-1]    = out_dict["education%s_startDate"%(ied+1)][-1]   .replace('`','')
        out_dict["education%s_endDate"%(ied+1)][-1]      = out_dict["education%s_endDate"%(ied+1)][-1]     .replace('`','')
        

        
df = pd.DataFrame.from_dict(out_dict)
print(df)
df.to_stata('linkedinData_rows%sto%s.dta'%(args.firstRow,args.lastRow), version=118) 
