from properties import Properties


class DuplicateSheetRequest(Properties):
    _dynamic_attr = ['insertSheetIndex', 'newSheetId', 'newSheetName', 'sourceSheetId']

    def __init__(self, sourceSheetId, **kwargs):
        super(DuplicateSheetRequest, self).__init__(sourceSheetId=sourceSheetId, **kwargs)


class DeleteSheetRequest(Properties):
    _dynamic_attr = ['sheetId']

    def __init__(self, sheetId):
        super(DeleteSheetRequest, self).__init__(sheetId=sheetId)
