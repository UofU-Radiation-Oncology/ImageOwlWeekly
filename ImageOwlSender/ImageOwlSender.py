import json
import requests
import os
import time
import openpyxl
import base64
import datetime

#TQA : Module for Connecting to Total QA API


client_id = '184:ImageSender'
client_key = 'e2d960b802dcf51763774c1f4271536ab63ded88c6f88d3e433c47c6346bc26b'
base_url = 'https://tqa.imageowl.com/api/rest'
_oauth_ext = '/oauth'
_grant_type = 'client_credentials'
access_token = ''
token_duration = 0
token_type = ''
token_exp_time = ''
token_exp_margin = 0.9

def set_tqa_token():
    
        payload = {"client_id": client_id,
                    "client_secret": client_key,
                    "grant_type": _grant_type}

        request_url = base_url + _oauth_ext
        r = requests.post(request_url ,data = payload)

        j = r.json()
        global access_token
        access_token = j['access_token']
        
        global token_duration
        token_duration = j['expires_in']
        
        global token_type
        token_type = j['token_type']

        global token_exp_time
        token_exp_time = datetime.datetime.now()+datetime.timedelta(seconds=token_duration)

    

def load_json_credentials(credential_file):
        with open(credential_file) as cred_file:
                cred = json.load(cred_file)
                tqaCred = cred['TQACredentials']
                
                global client_id
                client_id = tqaCred['ClientID']

                global base_url
                base_url = tqaCred['BaseURL']

                global _oauth_ext
                _oauth_ext = tqaCred['OauthURL']

                key_bytes = base64.b64decode(tqaCred['APIKey'])
                global client_key
                client_key = key_bytes.decode('UTF-8')

                set_tqa_token()
     
def save_json_credentials(credential_file):
    #encode the key in base 64
    key_bytes = client_key.encode('UTF-8')
    key_bytes_b64 = base64.b64encode(key_bytes)
    base64_key = key_bytes_b64.decode('UTF-8')

    cred_info = {
        "ClientID":client_id,
        "APIKey": base64_key,
        "BaseURL": base_url,
        "OauthURL": _oauth_ext}
    
    tqa_cred_dict = {"TQACredentials":cred_info}

    json_out_file = open(credential_file, "w")
    json_out_file.write(json.dumps(tqa_cred_dict, indent=4, sort_keys=True))
    json_out_file.close()

def get_standard_headers():
        if access_token == '':
                set_tqa_token()
        else:
                close_time_delta = datetime.timedelta(seconds =(1-token_exp_margin)*token_duration)
                if datetime.datetime.now() > token_exp_time - close_time_delta:
                        set_tqa_token()

        bearer_token = 'Bearer ' + access_token
        headers = {
            'authorization': bearer_token,
            'content-type': "application/json",
            'accept': "application/json",
        }
        return headers
def get_mpc_headers():
        if access_token == '':
                set_tqa_token()
        else:
                close_time_delta = datetime.timedelta(seconds =(1-token_exp_margin)*token_duration)
                if datetime.datetime.now() > token_exp_time - close_time_delta:
                        set_tqa_token()

        mdate = datetime.datetime.today()-datetime.timedelta(days=1)
        mpcdate = mdate.isoformat();
        bearer_token = 'Bearer ' + access_token
        headers = {
            'authorization': bearer_token,
            'content-type': "application/json",
            'accept': "application/json",
            'deviceDateFilter': mpcdate
        }
        return headers

def get_request(url_ext):
        url = base_url + url_ext
        response = requests.request("GET",url,headers = get_standard_headers())
        
        return {'json':response.json(),
                'status':response.status_code,
                'raw':response}

def get_sites():
        return get_request('/sites')

def get_users(user_id = -1):
        if user_id == -1:
                return get_request('/users')
        else:
                return get_request('/users/'+str(user_id))

def get_machines(active = -1,site = -1, device_type = -1):
        #build the filter
        filter = ''
        if not active == -1:
                filter = filter + 'active=' +str(active)

        if not site == -1:
                if len(filter) > 0: filter += '&'
                filter = filter + 'site=' + str(site)

        if not device_type == -1:
                if len(filter) > 0: filter += '&'
                filter = filter + 'device_type=' + str(device_type)

        if len(filter) > 0: filter = '?' + filter

        url_ext = '/machines'+filter
        
        return get_request(url_ext)

def get_report_data(report_id):
        return get_request('/report-data/'+str(report_id))


def eqhub(schedule_id,schedule_eqh):
        headers = get_mpc_headers()
        url_ext = ''.join(['/schedules/',str(schedule_id),'/variables/6661/equipment/',str(schedule_eqh),'/equipment-hub-results'])
        #url_ext = ''.join(['/schedules/',str(schedule_id),'/equipment-hub-results'])
        url = ''.join([base_url,url_ext])
        
        return requests.post( url, headers=headers, data = {})


def add_comment(schedule_id,test,comment):
        headers = get_standard_headers()
        url_ext = ''.join(['/schedules/',str(schedule_id),'/add-results'])
        url = ''.join([base_url,url_ext])
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
        return requests.post( url, headers=headers, data = json.dumps(data))


def upload_analysis_file(schedule_id,file_path):
        headers = get_standard_headers()
        #remove the content-type for this call
        del headers['content-type']
        url_ext = ''.join(['/schedules/',str(schedule_id),'/upload-images'])
        url = ''.join([base_url,url_ext])
        files = [
            ('file',(file_path,open(file_path,'rb'),'application/octet-stream'))
            ]
        return requests.post( url, headers=headers, data = {}, files = files)

def start_processing(schedule_id):
        headers = get_standard_headers()
        url_ext = ''.join(['/schedules/',str(schedule_id),'/start-processing'])
        url_process = ''.join([base_url,url_ext])
        return requests.post( url_process, headers=headers, data = {})

def finalize_report(schedule_id):
        headers = get_standard_headers()
        url_ext = ''.join(['/schedules/',str(schedule_id),'/finalize-results'])
        url_process = ''.join([base_url,url_ext])
        return requests.post( url_process, headers=headers, data = {})        

def get_upload_status(schedule_id):
        return get_request(''.join(['/schedules/',str(schedule_id),'/upload-images']))

class Machine:
    def __init__(Machine, name, id, mlc, wl, mpc):
        # name is the name for the machine
        # id is the id for the ImageOwl schedule to which we are uploading
        # mlc is the path to the folder containing the picket fence dicom files
        # wl is the path to the folder containing the wl files
        Machine.name = name
        Machine.id = id
        Machine.mlc = mlc
        Machine.wl = wl
        Machine.mpc = mpc


class tqa:
    def __init__(self):
        if access_token == '':
                set_tqa_token()
        else:
                close_time_delta = datetime.timedelta(seconds =(1-token_exp_margin)*token_duration)
                if datetime.datetime.now() > token_exp_time - close_time_delta:
                        set_tqa_token()

    def mpc(machine):
        if machine.mpc is None:
            return " "
        os.chdir('\\\\hci-eclipse-fs\\va_transfer\\TDS\\' + machine.mpc + '\\MPCChecks\\')
        list_of_fldrs = [d for d in os.listdir(os.getcwd()) if os.path.isdir(d)]
        mpc_dates = [x[20:25] for x in list_of_fldrs[-5:]]
        comment = "The last 5 entries for MPC on this machine were on: " + mpc_dates[4] + ", " + mpc_dates[3] + ", " + mpc_dates[2] + ", " + mpc_dates[1] + ", & " + mpc_dates[0] + ".       "
        return comment
        
    def mlc(machine):
        if machine.name == "V1" or machine.name == "V5":
            pf = "picketfenceHD_"
        else:
            pf = "picketfence_"
        headers = get_standard_headers
        os.chdir(machine.mlc)
        list_of_files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        for file in list_of_files:
            if "picketfence" not in file:
                os.rename(file,pf+file)
        list_of_files = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
        latest_file = list_of_files[-1]
        
        if (time.time()-os.path.getmtime(latest_file))/60/60/24 < 7 :
            s = upload_analysis_file(machine.id,machine.mlc + latest_file)
            if not s.ok :
                comment = "Error Code Uploading Picket Fence: " + str(s.status_code) + "  Reason: " + s.reason + "       "
                print(comment)
            else :
                comment = machine.name + ' MLC Picket Fence upload: ' + s.json()['results'] + "       "
                print (comment)
                s = get_upload_status(machine.id)
                while not s['raw'].ok:
                    s = get_upload_status(machine.id)                
        else:
            comment = "No MLC Picket Fence acquired for " + machine.name + " since " + time.localtime(os.path.getmtime(latest_file)).strftime('%m-%d') + ".       "
            print(comment)
        print()
        return comment

    def wl(machine):
        if not machine.wl:
            comment = "No Directory saved for location of WL files for " + machine.name + "       "
            print (comment)
        else:
            headers = get_standard_headers
            os.chdir(machine.wl)
            if "transfer" in machine.wl:
                listfldrs = sorted(os.listdir(os.getcwd()), key=os.path.getmtime)
                subfldr = ''
                for fldr in listfldrs:
                    if (time.time()-os.path.getmtime(fldr))/60/60/24 < 7 :
                        subfldr = fldr
                if subfldr == '':
                    print("No WL acquired for " + machine.name + " in the past week.")
                else:
                    os.chdir(machine.wl + subfldr)
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
                    s = upload_analysis_file(machine.id,file)
                    if not s.ok :
                        print("Error Code: " + str(s.status_code) + "  Reason: " + s.reason)
                    else :
                        print (machine.name + ' WL upload ' + str(newfiles.index(file) +1) + ' of ' + str(len(newfiles)) + ': ' + s.json()['results'])
                t = get_upload_status(machine.id)
                while t['json']['uploads'][0]['files']<len(newfiles):
                    t = get_upload_status(machine.id)                
            else:
                print("No WL acquired for " + machine.name + " in the past week.")

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
    allmachines.append(Machine(row[0],str(row[1]),row[2],row[3],row[4]))

for machine in allmachines:
    #tqa.sch(machine)
    b = tqa.mlc(machine)
    c = tqa.mpc(machine)
    #tqa.wl(machine)
    
    s = add_comment(machine.id,'11396',b + c)
        
    s = start_processing(machine.id)
    if s.ok :
        print(s.json()['message']+' on any '+machine.name+' uploaded weekly files.')
    else:
        print(s.json()['detail'])
        print(json.dumps(s.json(), indent=4, sort_keys=True))
    print()
    
print("All machines looped through")


