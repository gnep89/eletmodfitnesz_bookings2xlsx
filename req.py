import requests
import pwd

BASE_URL = "https://eletmodfitnesz.booked4.us/rest-v2/api/"


# get token to authorize
def get_token():
    get_token_url = BASE_URL + "token"
    get_token_body = {
        'grant_type': 'password',
        'username': pwd.EMAIL,
        'password': pwd.PASSWORD,
        'client_id': pwd.API_KEY,
        'client_secret': pwd.API_SECRET
    }

    request = requests.post(get_token_url, data=get_token_body)
    token = request.json()['access_token']
    return token


# get reservations by date
def get_reservations(from_date, to_date, calendar_id, token):
    get_reservations_url = f"{BASE_URL}calendars/{calendar_id}/Reservations"
    get_reservations_headers = {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
    }
    get_reservations_body = {
        'StartDateTime': from_date,
        'EndDateTime': to_date
    }
    request = requests.get(get_reservations_url, headers=get_reservations_headers, data=get_reservations_body, params=get_reservations_body)
    return request.json()