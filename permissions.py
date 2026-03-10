import json
from datetime import datetime
from pathlib import Path
from typing import Optional

from app.utils.config import ACCESSREQUESTSFILE, USERPERMISSIONSFILE, AVAILABLEAPPS
from app.auth.login_encryption import generaterequestid

class PermissionsManager:
    def __init__(self):
        self.requestsfile = ACCESSREQUESTSFILE
        self.permissionsfile = USERPERMISSIONSFILE
        self.ensurefiles()
        
    def ensurefiles(self):
        if not self.requestsfile.exists():
            with open(self.requestsfile, 'w') as f:
                json.dump([], f)
        if not self.permissionsfile.exists():
            with open(self.permissionsfile,'w') as f:
                json.dump({}, f)
        else:
            try:
                with open(self.permissionsfile, 'r') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        print("Fixing permissions file format...")
                        with open(self.permissionsfile, 'w') as f:
                            json.dump({}, f)
            except:
                with open(self.permissionsfile, 'w') as f:
                    json.dump({}, f)
                
    def submitaccessrequest(self, userid: str, username: str, appname: str) -> str:
        with open(self.requestsfile, 'r') as f:
            requests = json.load(f)
        for req in requests:
            if req['userid'] == userid and req['requestedapp'] == appname and req['status'] == 'pending':
                return req['requestid']
        requestid = generaterequestid()
        newrequest = {
            "requestid": requestid,
            "userid": userid,
            "username": username,
            "requestedapp": appname,
            "requestedat": datetime.now().isoformat(),
            "status": "pending",
            "approvedby": None,
            "approvedat": None
            }
        requests.append(newrequest)
        with open(self.requestsfile,'w') as f:
            json.dump(requests, f, indent=2)
        return requestid
    
    def checkaccess(self, userid: str, appname: str) -> bool:
        try:
            with open(self.permissionsfile, 'r') as f:
                permissions = json.load(f)
            if userid not in permissions:
                return False
            userperms = permissions[userid]
            if 'apps' not in userperms:
                return False
            if appname not in userperms:
                return False
            return userperms['apps'][appname].get('access',False)
        except Exception as e:
            print(f"Error checking access: {e}")
            return False
    def getuserapps(self, userid: str) -> list:
        try:
            with open(self.permissionsfile, 'r') as f:
                permissions = json.load(f)
            if userid not in permissions:
                return []
            userperms = permissions[userid]
            if 'apps' not in userperms:
                return []
            return [app for app, data in userperms['apps'].items() if data.get('access', False)]
        except Exception as e:
            print(f"Error getting user apps: {e}")
            return []
    def haspendingrequest(self, userid: str, appname: str) -> bool:
        try:
            with open(self.requestsfile, 'r') as f:
                requests = json.load(f)
            for req in requests:
                if (req['userid'] == userid and
                    req['requestedapp'] == appname and
                    req['status'] == 'pending'):
                    return True
            return False
        except Exception as e:
            print(f"Error checking pending requests: {e}")
            return False
    
    def getpendingrequests(self) -> list:
        try:
            with open(self.requestsfile, 'r') as f:
                requests = json.load(f)
            return [req for req in requests if req['status'] == 'pending']
        except Exception as e:
            print(f"Error getting pending requests: {e}")
            return []
    
    def approverequest(self, requestid: str, adminusername: str) -> bool:
        try:
            with open(self.requestsfile, 'r') as f:
                requests = json.load(f)
            request = None
            for req in requests:
                if req['requestid'] == requestid:
                    request = req
                    break
            if not request:
                return False
            request['status'] = 'approved'
            request['approvedby'] = adminusername
            request['approvedat'] = datetime.now().isoformat()
            with open(self.requestsfile, 'w') as f:
                json.dump(requests, f, indent=2)
            self.grantpermission(
                request['userid'],
                request['username'],
                request['requestedapp'],
                adminusername
            )
            return True
        except Exception as e:
            print(f"Error approving request: {e}")
            return False
        
    def denyrequest(self, requestid: str, adminusername: str) -> bool:
        try:
            with open(self.requestsfile, 'r') as f:
                requests = json.load(f)
            for req in requests:
                if req['requestid'] == requestid:
                    req['status'] = 'denied'
                    req['approvedby'] = adminusername
                    req['approvedat'] = datetime.now().isoformat()
                    break
            with open(self.requestsfile, 'w') as f:
                json.dump(requests, f, indent=2)
            return True
        except Exception as e:
            print(f"Error denying request: {e}")
            return False
        
    def grantpermission(self, userid: str, username: str, appname: str, grantedby: str):
        with open(self.permissionsfile, 'r') as f:
            permissions = json.load(f)
        if userid not in permissions:
            permissions[userid] = {
                "username": username,
                "apps": {}
                }
        permissions[userid]['apps'][appname] = {
            "access": True,
            "grantedat": datetime.now().isoformat(),
            "grantedby": grantedby
            }
        with open(self.permissionsfile, 'w') as f:
            json.dump(permissions, f, indent=2)