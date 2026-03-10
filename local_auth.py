import json
from datetime import datetime
from pathlib import Path

from app.utils.config import LOCALUSERCONFIG
from app.auth.login_encryption import hashpassword, verifypassword, generateuserid

class LocalAuth:
    def __init__(self):
        self.configfile = LOCALUSERCONFIG
    
    def userexists(self) -> bool:
        return self.configfile.exists()
    
    def createuser(self, username: str, password: str) -> dict:
        if self.userexists():
            raise ValueError("User already exists on this machine")
        userid = generateuserid()
        passwordhash = hashpassword(password)
        userdata = {
            "username": username,
            "passwordhash": passwordhash,
            "userid": userid,
            "createdat": datetime.now().isoformat()
            }
        with open(self.configfile, 'w') as f:
            json.dump(userdata, f, indent=2)
        return userdata
    
    def authenticate(self, username: str, password: str) -> tuple[bool, dict]:
        if not self.userexists():
            return False, {}
        with open(self.configfile, 'r') as f:
            userdata = json.load(f)
            
        print("=== DEBUG: User Data ===")
        print(f"Keys in userdata: {list(userdata.keys())}")
        print(f"Username: {userdata.get('username', 'MISSING')}")
        print(f"Password hash exists: {'passwordhash' in userdata}")
        if 'passwordhash' in userdata:
            print(f"Password hash length: {len(userdata['passwordhash'])}")
            print(f"Password hash starts with: {userdata['passwordhash'][:10]}...")
        else:
            print("ERROR: No passwordhash key found!")
            print("========================")
            
        if userdata['username'] != username:
            return False, {}
        if not verifypassword(password, userdata['passwordhash']):
            return False, {}
        return True, userdata
    
    def getuserdata(self) -> dict:
        if not self.userexists():
            return {}
        with open(self.configfile, 'r') as f:
            return json.load(f)
        
    def deleteuser(self):
        if self.configfile.exists():
            self.configfile.unlink()