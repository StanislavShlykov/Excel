import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

CREDENTIALS_FILE = 'C:/Users/Stas/PycharmProjects/script_exel/cred.json'

spread_id = '1ouqhTyrBL2cbdHh9vJM2A4-YIaUB2fyTKRQ0l7ULzvU'

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE, [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive'])

httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)


def data():
    values = service.spreadsheets().values().batchGet(
        spreadsheetId=spread_id,
        ranges=['F2:F', 'L2:L', 'O2:O', 'AI2:AI'],
        majorDimension='COLUMNS',
    ).execute()
    main_data = []
    for i in values['valueRanges']:
        main_data.append(i['values'][0])

    return main_data
