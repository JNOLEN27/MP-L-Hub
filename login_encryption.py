import bcrypt
import uuid

def hashpassword(password: str) -> str:
    passwordbytes = password.encode('utf-8')
    salt = bcrypt.gensalt(rounds=12)
    hashed = bcrypt.hashpw(passwordbytes, salt)
    return hashed.decode('utf-8')

def verifypassword(password: str, hashed: str) -> bool:
    try:
        passwordbytes = password.encode('utf-8')
        hashedbytes = hashed.encode('utf-8')
        return bcrypt.checkpw(passwordbytes, hashedbytes)
    except Exception as e:
        print(f"Error verifying password: {e}")
        return False

def generateuserid() -> str:
    return str(uuid.uuid4())

def generaterequestid() -> str:
    return f"req_{uuid.uuid4().hex[:12]}"