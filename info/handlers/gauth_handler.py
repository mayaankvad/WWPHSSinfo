from google.oauth2 import id_token
from google.auth.transport import requests
from WWPHSSinfo import settings


def authenticate(token):
    idinfo = id_token.verify_oauth2_token(token, requests.Request(), settings.GAUTH_KEY)

    try:
        if idinfo['hd'] == settings.GSUITE_DOMAIN:
            return True
        return False
    except:
        return False


