from pprint import pprint
import httplib2
import googleapiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


class Spreadsheet:
    def __init__(self, json_file_name, debug_mode=False):
        self.debug_mode = debug_mode
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(
            json_file_name, ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
        self.httpAuth = self.credentials.authorize(httplib2.Http())
        self.service = googleapiclient.discovery.build('sheets', 'v4', http=self.httpAuth)
        self.drive_service = None   # Needed for sharing
        self.spreadsheet_id = None
        self.sheet_id = None
        self.sheet_title = None
        self.requests = []
        self.values_ranges = []

    def create(self, title, sheet_title, rows=1000, cols=26, locale='en_US'):
        spreadsheet = self.service.spreadsheets().create(body={
            'properties': {'title': title, 'locale': locale},
            'sheets': [{'properties': {'sheetType': 'GRID', 'sheetId': 0, 'title': sheet_title,
                                       'gridProperties': {'rowCount': rows, 'columnCount': cols}}}]
        }).execute()
        if self.debug_mode:
            pprint(spreadsheet)
        self.spreadsheet_id = spreadsheet['spreadsheetId']
        self.sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']
        self.sheet_title = spreadsheet['sheets'][0]['properties']['title']

    def share(self, share_request_body):
        if self.drive_service is None:
            self.drive_service = googleapiclient.discovery.build('drive', 'v3', http=self.httpAuth)
        share_res = self.drive_service.permissions().create(
            fileId=self.spreadsheet_id,
            body=share_request_body,
            fields='id'
        ).execute()
        if self.debug_mode:
            pprint(share_res)

    def share_with_email_for_reading(self, email):
        self.share({'type': 'user', 'role': 'reader', 'emailAddress': email})

    def share_with_email_for_writing(self, email):
        self.share({'type': 'user', 'role': 'writer', 'emailAddress': email})

    def share_with_anybody_for_reading(self):
        self.share({'type': 'anyone', 'role': 'reader'})

    def share_with_anybody_for_writing(self):
        self.share({'type': 'anyone', 'role': 'writer'})

    def get_sheet_url(self):
        return 'https://docs.google.com/spreadsheets/d/' + self.spreadsheet_id + '/edit#gid=' + str(self.sheet_id)

    # Sets current spreadsheet by id; set current sheet as first sheet of this spreadsheet
    def set_spreadsheet_by_id(self, spreadsheet_id):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        if self.debug_mode:
            pprint(spreadsheet)
        self.spreadsheet_id = spreadsheet['spreadsheetId']
        self.sheet_id = spreadsheet['sheets'][0]['properties']['sheetId']
        self.sheet_title = spreadsheet['sheets'][0]['properties']['title']

    def set_sheet_by_title(self, sheet_title):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        sheet_id = None
        for sheet in spreadsheet['sheets']:
            if sheet['properties']['title'] == sheet_title:
                sheet_id = sheet['properties']['sheetId']
        self.sheet_id = sheet_id
        self.sheet_title = sheet_title

    # spreadsheets.batchUpdate and spreadsheets.values.batchUpdate
    def run_prepared(self, value_input_option='USER_ENTERED'):
        upd1res = {'replies': []}
        upd2res = {'responses': []}
        try:
            if len(self.requests) > 0:
                upd1res = self.service.spreadsheets().batchUpdate(
                    spreadsheetId=self.spreadsheet_id, body={'requests': self.requests}).execute()
                if self.debug_mode:
                    pprint(upd1res)
            if len(self.values_ranges) > 0:
                upd2res = self.service.spreadsheets().values().batchUpdate(
                    spreadsheetId=self.spreadsheet_id, body={'valueInputOption': value_input_option,
                                                             'data': self.values_ranges}).execute()
                if self.debug_mode:
                    pprint(upd2res)
        finally:
            self.requests = []
            self.values_ranges = []
        return upd1res['replies'], upd2res['responses']

    def prepare_add_sheet(self, sheet_title, rows=1000, cols=26):
        self.requests.append({"addSheet": {"properties": {"title": sheet_title,
                                                          'gridProperties': {'rowCount': rows, 'columnCount': cols}}}})

    # Adds new sheet to current spreadsheet, sets as current sheet and returns it's id
    def add_sheet(self, sheet_title, rows=1000, cols=26):
        self.prepare_add_sheet(sheet_title, rows, cols)
        added_sheet = self.run_prepared()[0][0]['addSheet']['properties']
        self.sheet_id = added_sheet['sheetId']
        self.sheet_title = added_sheet['title']
        return self.sheet_id

    def prepare_set_dimension_pixel_size(self, dimension, start_index, end_index, pixel_size):
        self.requests.append({'updateDimensionProperties': {
            'range': {'sheetId': self.sheet_id,
                      'dimension': dimension,
                      'startIndex': start_index,
                      'endIndex': end_index},
            'properties': {'pixelSize': pixel_size},
            'fields': 'pixelSize'}})

    def prepare_set_columns_width(self, start_col, end_col, width):
        self.prepare_set_dimension_pixel_size('COLUMNS', start_col, end_col + 1, width)

    def prepare_set_column_width(self, col, width):
        self.prepare_set_columns_width(col, col, width)

    def prepare_set_rows_height(self, start_row, end_row, height):
        self.prepare_set_dimension_pixel_size('ROWS', start_row, end_row + 1, height)

    def prepare_set_row_height(self, row, height):
        self.prepare_set_rows_height(row, row, height)

    def prepare_set_values(self, cells_range, values, major_dimension='ROWS'):
        self.values_ranges.append({'range': self.sheet_title + '!' + cells_range,
                                   'majorDimension': major_dimension, 'values': values})

    def clear_sheet(self):
        range_all = '{0}!A1:Z'.format(self.sheet_title)
        body = {}
        self.service.spreadsheets( ).values( ).clear(spreadsheetId=self.spreadsheet_id,
                                                     range=range_all, body=body).execute()
