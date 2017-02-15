# -*- coding: UTF-8 -*-
from __future__ import unicode_literals

import argparse
import pprint
from functools import partial

import httplib2
from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage

from com.properties import SheetProperties, GridProperties
from com.request_properties import DuplicateSheetRequest, DeleteSheetRequest


class GoogleSheetException(Exception):
    pass


class SpreadSheetException(Exception):
    pass


class SheetException(Exception):
    pass


class GoogleSheets(object):
    # SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
    SCOPES = 'https://spreadsheets.google.com'
    CLIENT_SECRET_FILE = 'client_secret.json'
    APPLICATION_NAME = 'GoogleSheets'

    def __init__(self, credential_path=None, scopes=None, application_name=None, flags=None):
        self.credential_path = self.CLIENT_SECRET_FILE if credential_path is None else credential_path
        self.scopes = self.SCOPES if scopes is None else scopes
        self.application_name = self.APPLICATION_NAME if application_name is None else application_name
        self.spreadsheets = None
        self.flags = flags if flags else argparse.ArgumentParser(parents=[tools.argparser]).parse_args()

        self._init_service_and_spreadsheets()

    def get_credentials(self):
        credential_path = self.credential_path

        store = Storage(credential_path)
        try:
            credentials = store.get()
        except KeyError:
            flow = client.flow_from_clientsecrets(self.credential_path, self.scopes)
            flow.user_agent = self.application_name
            credentials = tools.run_flow(flow, store, self.flags)
        else:
            if credentials.invalid:
                flow = client.flow_from_clientsecrets(self.credential_path, self.scopes)
                flow.user_agent = self.application_name
                credentials = tools.run_flow(flow, store, self.flags)

        print('Storing credentials to ' + credential_path)

        return credentials

    def _init_service_and_spreadsheets(self):
        credentials = self.get_credentials()
        http = credentials.authorize(httplib2.Http())
        service = discovery.build('sheets', 'v4', http=http)

        self.spreadsheets = service.spreadsheets()

    def get_spreadsheet_by_id(self, spreadsheet_id, **kwargs):
        """
        Returns the spreadsheet at the given ID. The caller must specify the spreadsheet ID.
        """

        # self.get_sheet_by_spreadsheetId.__doc__ = inspect.getdoc(self.spreadsheets.get)

        if 'spreadsheetId' in kwargs:
            raise GoogleSheetException("spreadsheetId should not contain in operational arguments")
        spreadsheet = self.spreadsheets.get(spreadsheetId=spreadsheet_id, **kwargs)
        return SpreadSheet(self, spreadsheet, self.spreadsheets)

    def get_spreadsheet_by_name(self, name, **kwargs):
        """
        Returns the spreadsheet at the given ID. The caller must specify the spreadsheet ID.
        """

        sheets = self.spreadsheets.sheets()
        return

    def create_new_spreadsheet(self, *args, **kwargs):
        raise NotImplementedError

    def get_all_spreadsheet(self):
        # sheets = self.spreadsheets.sheets()
        # pprint.pprint(sheets.__dict__)
        # for value in sheets.values():
        #     print(value)
        raise NotImplementedError


def parse_response_replies(name):
    def wrapper(func):
        def inner(self, *args, **kwargs):
            response = func(self, *args, **kwargs)
            properties = dict()
            for item in response['replies']:
                if name in item:
                    properties.update(item[name]['properties'])
            return Sheet(self, self.google_spreadsheets, spreadsheetId=response['spreadsheetId'], **properties)

        return inner

    return wrapper


class SpreadSheet(object):
    """
    Create one Spread Sheet.
    """

    def __init__(self, parent, spreadsheet, google_spreadsheets):
        """
        Initialize SpreadSheet, and support some common functions to control the google
        SpreadSheet like operation in local.
        :param spreadsheet: The single SpreadSheet Request.
        :param parent: The valid Google SpreadSheets object.
        :param google_spreadsheets: The source object for valid google docs
        """
        self.spreadsheet = spreadsheet
        self.parent = parent
        self.google_spreadsheets = google_spreadsheets

        self._parse_spreadsheet()
        self._partial_parent_with_spreadsheet_id()

        self.change_flag = False

    def _partial_parent_with_spreadsheet_id(self):
        self.batchUpdate = partial(self.google_spreadsheets.batchUpdate, spreadsheetId=self.spreadsheetId)

    def _parse_spreadsheet(self):
        self.data = self.spreadsheet.execute()
        self.title = self.data['properties']['title']
        self.spreadsheetId = self.data['spreadsheetId']

        self.change_flag = False

    def get_sheet_id_by_sheet_name(self, sheet_name):
        for sheet in self.data['sheets']:
            if sheet_name == sheet['properties']['title']:
                return sheet['properties']['sheetId']

    def get_one_sheet_by_name(self, sheet_name):
        """
        Get one sheet by name, the name must be string type and if the name not in this spreadsheet, it will case a
        Exception.
        """
        for sheet in self.data['sheets']:
            if sheet_name == sheet['properties']['title']:
                return Sheet(self, self.google_spreadsheets, spreadsheetId=self.spreadsheetId, **sheet['properties'])

    def get_one_sheet_by_id(self, sheet_id):
        for sheet in self.data['sheets']:
            if sheet_id == sheet['properties']['sheetId ']:
                return Sheet(self, self.google_spreadsheets, spreadsheetId=self.spreadsheetId, **sheet['properties'])

    @parse_response_replies('duplicateSheet')
    def create_new_sheet(self, **kwargs):
        """
        Create a new sheet in spreadsheet and return the newer sheet object.
        :param kwargs:
        :return: Sheet object
        """
        requests = list()
        requests.append({
            'addSheet': {
                'properties': SheetProperties(**kwargs).to_dict()
            }
        })

        return self.batch_update(requests=requests)

    def delete_one_sheet(self, sheetId):
        requests = list()
        requests.append({
            'deleteSheet': DeleteSheetRequest(sheetId=sheetId).to_dict()
        })
        return self.batch_update(requests)

    @parse_response_replies('duplicateSheet')
    def duplicate_one_sheet(self, sourceSheetId, **kwargs):
        requests = list()
        requests.append({
            'duplicateSheet': DuplicateSheetRequest(sourceSheetId, **kwargs).to_dict()
        })
        return self.batch_update(requests)

    def get_all_sheets(self):
        """
        Get current all sheet info from current spreadsheet, each item type SingleSheet
        :return:
        """
        sheets = list()
        for sheet in self.data['sheets']:
            sheets.append(Sheet(self, self.google_spreadsheets, **sheet['properties']))

        return sheets

    def batch_update(self, requests):
        body = {
            'requests': requests
        }
        # response = self.batchUpdate(body=body).execute()
        # return response
        try:
            response = self.batchUpdate(body=body).execute()
        except Exception as e:
            raise e
        else:
            self.change_flag = True
            self._parse_spreadsheet()
            return response


def execute(func):
    def wrapper(*args, **kwargs):
        return func(*args, **kwargs).execute()

    return wrapper


class Sheet(object):
    def __init__(self, parent, google_spreadsheets, **kwargs):

        self.parent = parent
        self.google_spreadsheets = google_spreadsheets

        self._parse_params(**kwargs)

    def _parse_params(self, **kwargs):
        if 'gridProperties' in kwargs:
            self.__dict__['gridProperties'] = GridProperties(**kwargs.pop('gridProperties'))
        for key, value in kwargs.iteritems():
            self.__dict__[key] = value

    def __getattr__(self, item):
        if hasattr(self.google_spreadsheets.values(), item):
            return partial(getattr(self.google_spreadsheets.values(), item),
                           spreadsheetId=self.spreadsheetId)

    def _get_range_name(self, source_range):

        if not source_range:
            return self.title
        elif self.title in source_range:
            return source_range
        else:
            return '{}!{}'.format(self.title, source_range)

    get_range_name = _get_range_name

    @execute
    def update_cells(self, range_name, valueInputOption='RAW', **kwargs):
        return self.update(range=self._get_range_name(range_name), valueInputOption=valueInputOption, **kwargs)

    @execute
    def get_cells(self, range_name, **kwargs):
        return self.get(range=self._get_range_name(range_name), **kwargs)

    def get_cell_position_by_value(self, check, range_name=''):
        res = self.get_cells(range_name=range_name)

        for i, raw in enumerate(res['values']):
            for j, col in enumerate(raw):
                if col == check:
                    return i, j


def main():
    """Shows basic usage of the Sheets API"""
    google_sheet = GoogleSheets()

    spreadsheetId = '1MtftCdoMseNbSZ7XvH4iBRBsQx3jDmRi5Vf7AobRDcI'
    spreadsheet = google_sheet.get_spreadsheet_by_id(spreadsheetId)

    # sheets = test_sheet.get_all_sheets()
    # pprint.pprint(sheets[0].title)

    def update_fields():
        sheet = spreadsheet.get_one_sheet_by_name('sheet2')
        position = sheet.get_position_by_string('BPV Light-DUW2')
        pprint.pprint(position)

        body = {
            'values': [
                ['', '284d97a'],
                ['Total/Reruns/Failed/Skipped/Dismissed', '184/7/5/0/0'],
                ['print first rerun still failed test cases', ],
                ['SPECMAN_CRS_HSDPCCH_CHAR_CQI_DETECTION_IS_LDR_HDR_2', 'UBP Scoreboard validation failed'],
                ['SPECMAN_BB_CRS_EDCH_TRANS_21', 'UBP Scoreboard validation failed'],
                ['SPECMAN_CRS_EDCH_CHAR_ULCOMP_9', 'UBP Scoreboard validation failed'],
                ['SPECMAN_CRS_EDCH_CHAR_ULCOMP_10', 'UBP Scoreboard validation failed'],
                ['SPECMAN_BB_CRS_UPSWITCH_PLOT_TPC_OFF_9', 'UBP Scoreboard validation failed'],
                ['Brief analysis of first original run:'],
                ['SPECMAN_BB_CRS_RACH_PREAMBLE_DETECTION_1', 'Command dpadcli returned false'],
                ['SPECMAN_BB_CRS_UPSWITCH_PLOT_TPC_ON_7', 'Command dpadcli returned false'],
            ]
        }

        sheet.update('sheet2!D7', body=body)

    def create_new_sheet():
        sheet = spreadsheet.create_new_sheet(title='sheet14', )
        body = {
            'values': [
                ['', '284d97a'],
                ['Total/Reruns/Failed/Skipped/Dismissed', '184/7/5/0/0'],
                ['print first rerun still failed test cases', ],
                ['SPECMAN_CRS_HSDPCCH_CHAR_CQI_DETECTION_IS_LDR_HDR_2', 'UBP Scoreboard validation failed'],
                ['SPECMAN_BB_CRS_EDCH_TRANS_21', 'UBP Scoreboard validation failed'],
                ['SPECMAN_CRS_EDCH_CHAR_ULCOMP_9', 'UBP Scoreboard validation failed'],
                ['SPECMAN_CRS_EDCH_CHAR_ULCOMP_10', 'UBP Scoreboard validation failed'],
                ['SPECMAN_BB_CRS_UPSWITCH_PLOT_TPC_OFF_9', 'UBP Scoreboard validation failed'],
                ['Brief analysis of first original run:'],
                ['SPECMAN_BB_CRS_RACH_PREAMBLE_DETECTION_1', 'Command dpadcli returned false'],
                ['SPECMAN_BB_CRS_UPSWITCH_PLOT_TPC_ON_7', 'Command dpadcli returned false'],
            ]
        }

        sheet.update_cells('A1', body=body)

    def duplicate_one_sheet():
        source_sheet_id = spreadsheet.get_sheet_id_by_sheet_name('sheet2')
        sheet = spreadsheet.duplicate_one_sheet(source_sheet_id, )
        pprint.pprint(sheet)

    def delete_one_sheet():
        source_sheet_id = spreadsheet.get_sheet_id_by_sheet_name('sheet2（副本）')

        sheet = spreadsheet.delete_one_sheet(source_sheet_id)
        pprint.pprint(sheet)

    def get_cells():
        sheet = spreadsheet.get_one_sheet_by_name('sheet2')
        res = sheet.get_cells('D1:D')
        pprint.pprint(res)

    def get_spreadsheet_by_name():
        spreadsheet = google_sheet.get_spreadsheet_by_name('FOR PYTHON SPREAD TEST')

        sheet = spreadsheet.get_one_sheet_by_name('sheet2')
        res = sheet.get_cells('D1:D')
        pprint.pprint(res)

    get_spreadsheet_by_name()


if __name__ == '__main__':
    main()
    # sheet = Sheet(SpreadSheet, a=10)
