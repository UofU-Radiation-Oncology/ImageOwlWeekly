
import sys
import os
import time
import openpyxl
import shutil
import requests
import datetime
sys.path.append(os.path.abspath("C:/Users/u6011999/source/repos/mattlwhitaker-IO/pyTQA"))
import tqa 
tqa.client_id = '184:ImageSender'
tqa.client_key = 'e2d960b802dcf51763774c1f4271536ab63ded88c6f88d3e433c47c6346bc26b'


class Machine:
    def __init__(Machine, name, id, mlc, wl, mpcfl, mpc):
        # name is the name for the machine
        # id is the id for the ImageOwl schedule to which we are uploading
        # mlc is the path to the folder containing the picket fence dicom files
        # wl is the path to the folder containing the wl files
        Machine.name = name
        Machine.id = id
        Machine.mlc = mlc
        Machine.wl = wl
        Machine.mpcfl = mpcfl
        Machine.mpc = mpc


class tqau:
    def __init__(self):
        if access_token == '':
                tqa.set_tqa_token()
        else:
                close_time_delta = datetime.timedelta(seconds =(1-token_exp_margin)*token_duration)
                if datetime.datetime.now() > token_exp_time - close_time_delta:
                        tqa.set_tqa_token()
    
    def log(comment):
        log = open(r"\\hci.utah.edu\dfs\RadOnc\Physics\ImageOwl\log.txt","a")
        log.write(datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S.%f") + "  " + comment + "\n")
        print(comment)
        log.close()
        return
                        
    def eqhub(schedule_id,eqid,date):
        headers = tqa.get_standard_headers()
        url_ext = ''.join(['/schedules/',str(schedule_id),'/equipment/',eqid,'/equipment-hub-results'])
        url = ''.join([tqa.base_url,url_ext])
        info = {
            #"date": datetime.datetime.now().strftime("%Y/%m/%d %H:%M"),
            "deviceDateFilter": date,
            #"comment": "Added via API script",
            "mode": "save_append",
            "finalize": 0
            }
        #headers.update(info)
        return requests.post( url, headers=headers, json = info)

    def mpc(machine):
        if machine.mpcfl is None:
            return " "
        os.chdir('\\\\hci-eclipse-fs\\va_transfer\\TDS\\' + machine.mpcfl + '\\MPCChecks\\')
        list_of_fldrs = [d for d in os.listdir(os.getcwd()) if os.path.isdir(d)]
        mpc_last = list_of_fldrs[-1]
        mpcdate = mpc_last[15:25]
        #mpc_dates = [x[20:25] for x in list_of_fldrs[-5:]]
        #comment = "The last 5 entries for MPC on this machine were on: " + mpc_dates[4] + ", " + mpc_dates[3] + ", " + mpc_dates[2] + ", " + mpc_dates[1] + ", & " + mpc_dates[0] + ".       "
        if (time.time()-os.path.getmtime(mpc_last))/60/60/24 < 4 :
            s = tqau.eqhub(machine.id,machine.mpc,mpcdate)
            if not s.ok :
                comment = "Error Code Uploading MPC: " + str(s.status_code) + "  Reason: " + s.reason + "       "
                tqau.log(comment)
            else :
                comment = "MPC uploaded. "
                tqau.log(comment)                            
        else:
            comment = "Last MPC was performed on " + time.strftime("%m/%d",time.localtime(os.path.getmtime(mpc_last))) + ". Please ask therapists to do a new MPC.  "
            tqau.log(comment)
        
        return comment
        
    def mlc(machine):
        if machine.name == "V1" or machine.name == "V5":
            pf = "picketfenceHD_"
        else:
            pf = "picketfence_"
        headers = tqa.get_standard_headers
        os.chdir(machine.mlc)
        list_of_files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        for file in list_of_files:
            if "picketfence" not in file:
                os.rename(file,pf+file)
        list_of_files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        latest_file = list_of_files[-1]
        
        if (time.time()-os.path.getmtime(latest_file))/60/60/24 < 4 :
            s = tqa.upload_analysis_file(machine.id,machine.mlc + latest_file)
            if not s.ok :
                comment = "Error Code Uploading Picket Fence: " + str(s.status_code) + "  Reason: " + s.reason + "       "
                tqau.log(comment)
            else :
                comment = "MLC Picket Fence uploaded. "
                tqau.log(comment)
                s = tqa.get_upload_status(machine.id)
                while not s['raw'].ok:
                    s = tqa.get_upload_status(machine.id)                
        else:
            comment = "Last MLC Picket Fence was performed on " + time.strftime("%m/%d",time.localtime(os.path.getmtime(latest_file))) + ". Please ask therapists to do a new MLC Picket Fence. After they have finished go to " + machine.mlc + " to find the image, then drag and drop it into the box above."
            tqau.log(comment)
        tqa.start_processing(machine.id)
        return comment

    def wl(machine):
        headers = tqa.get_standard_headers
        os.chdir(machine.wl)
        if "transfer" in machine.wl:
            listfldrs = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
            subfldr = ''
            for fldr in listfldrs:
                if (time.time()-os.path.getmtime(fldr))/60/60/24 < 7 :
                    subfldr = fldr
            if subfldr == '':
                tqau.log("No WL acquired for " + machine.name + " in the past week.")
            else:
                nwfldr = '\\\\hci.utah.edu\\dfs\\RadOnc\\Physics\\WinstonLutzTest\\' + machine.name + '\\Monthly\\20' + subfldr[:2] + '\\' + subfldr
                shutil.copytree(machine.wl + subfldr, nwfldr)
                shutil.rmtree(machine.wl + subfldr)
                os.chdir(nwfldr)
        list_of_files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        for file in list_of_files:
            if "winstonlutz" not in file:
                os.rename(file,r'winstonlutz_'+file)
        list_of_files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        newfiles = []
        for file in list_of_files:
            if (time.time()-os.path.getmtime(file))/60/60/24 < 7 :
                newfiles.append(file)
            
        if len(newfiles)>0 :
            for file in newfiles:
                s = tqa.upload_analysis_file(machine.id,file)
                if not s.ok :
                    tqau.log("Error Code: " + str(s.status_code) + "  Reason: " + s.reason)
                else :
                    tqau.log(machine.name + ' WL upload ' + str(newfiles.index(file) +1) + ' of ' + str(len(newfiles)) + ': ' + s.json()['results'])
            t = tqa.get_upload_status(machine.id)
            while t['json']['uploads'][0]['files']<len(newfiles):
                t = tqa.get_upload_status(machine.id)                
        else:
            tqau.log("No WL acquired for " + machine.name + " in the past week.")

    def add_comment(schedule_id,test,comment):
        headers = tqa.get_standard_headers()
        url = ''.join([tqa.base_url,'/schedules/',str(schedule_id),'/add-results'])
        date = datetime.datetime.now().isoformat(' ',timespec='minutes')
        data = {
            "date": date,
            "comment": "",
            "finalize": 0,
            "mode": "save_append",
     	    "variables": [{
                 "id": test,
                 "comment": comment,
                 "value": "Script ran"}]}
        return requests.post( url, headers=headers, json=data)

    def sch(machine):
        #url = '/equipment' 
        #url = '/schedules/' + machine.id + '/variables'
        #variable 288 for MLC 
        #url = '/report-data/113408'
        #url = '/reports?schedule=' + machine.id
        s = add_comment(machine.id,'11396','Testing to see if this works')
        #s = add_comment(machine.id,'11396',' to append text')
        #s = get_request(url)
        if not s['raw'].ok :
            print("Error Code: " + str(s.status_code) + "  Reason: " + s.reason)
        else :
            for report in s['json']['variables']:
                print(report['id'], report['name'], report['description'])
                



 

excel = openpyxl.load_workbook(filename="\\\\hci.utah.edu\\dfs\\RadOnc\\Physics\\ImageOwl\\Machines.xlsx")
allmachines = []
for row in excel.active.iter_rows(min_row=2, values_only=True):
    allmachines.append(Machine(row[0],str(row[1]),row[2],row[3],row[4],str(row[5])))

for machine in allmachines:
    #if machine.name == "V3":
    tqau.log(machine.name)
    b = tqau.mlc(machine)
    c = tqau.mpc(machine)
    comment = b+c
    #tqau.wl(machine)
    tqau.add_comment(machine.id,'11396',comment)
    if comment.find("Last") < 0 :
        s = tqa.finalize_report(machine.id)
        if not s.ok :
            tqau.log("Error Code: " + str(s.status_code) + "  Reason: " + s.reason)
        else :
            tqau.log("Weekly QA finalized for " + machine.name)
    print()
 
    
   

print("All machines looped through")


