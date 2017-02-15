import abc


class PropertiesException(Exception):
    pass


class Properties(object):
    __metaclass__ = abc.ABCMeta

    _dynamic_attr = []

    def __init__(self, **kwargs):
        self.update(**kwargs)

    def __setattr__(self, key, value):
        if key not in self._dynamic_attr:
            raise PropertiesException(
                "Class {} not support current attribute {}".format(str(self.__class__.__name__), key))

        self.__dict__[key] = value

    def __getattr__(self, item):
        if item in self:
            return self[item]

        raise PropertiesException('Properties has no attribute {}'.format(str(item)))

    def __repr__(self):
        return str(self.__dict__)

    def __str__(self):
        return str(self.__dict__)

    def update(self, **kwargs):
        for key, value in kwargs.iteritems():
            self.__setattr__(key, value)

    def to_dict(self):
        return self.__dict__


class SheetProperties(Properties):
    """
    u'properties': {
        u'gridProperties': {
            u'$ref': u'GridProperties',
            u'description': u'Additional properties of the sheet if this sheet is a grid.
        u'hidden': {
            u'description': u"True if the sheet is hidden in the UI, false if it's visible.",
            u'type': u'boolean'},
        u'index': {
            u'description': u'The index of the sheet within the spreadsheet.',
            u'format': u'int32',
            u'type': u'integer'},
        u'rightToLeft': {
            u'description': u'True if the sheet is an RTL sheet instead of an LTR sheet.',
            u'type': u'boolean'},
        u'sheetId': {
            u'description': u'The ID of the sheet. Must be non-negative.\nThis field cannot be changed once set.',
            u'format': u'int32',
            u'type': u'integer'},
        u'sheetType': {
            u'description': u'The type of sheet. Defaults to GRID.\nThis field cannot be changed once set.',
            u'enum': [u'SHEET_TYPE_UNSPECIFIED',
                      u'GRID',
                      u'OBJECT'],
            u'enumDescriptions': [
                u'Default value, do not use.',
                u'The sheet is a grid.',
                u'The sheet has no grid and instead has an object like a chart or image.'],
            u'type': u'string'},
        u'tabColor': {u'$ref': u'Color',
                      u'description': u'The color of the tab in the UI.'},
        u'title': {u'description': u'The name of the sheet.',
                   u'type': u'string'}},
    """

    _dynamic_attr = ['gridProperties', 'hidden', 'index', 'rightToLeft', 'sheetId', 'sheetType', 'tabColor', 'title']

    def __init__(self, **kwargs):
        super(SheetProperties, self).__init__(**kwargs)


class GridProperties(Properties):
    """
        u'properties': {
            u'columnCount': {u'description': u'The number of columns in the grid.',
                                         u'format': u'int32',
                                         u'type': u'integer'},
            u'frozenColumnCount': {u'description': u'The number of columns that are frozen in the grid.',
                                   u'format': u'int32',
                                   u'type': u'integer'},
            u'frozenRowCount': {u'description': u'The number of rows that are frozen in the grid.',
                                u'format': u'int32',
                                u'type': u'integer'},
            u'hideGridlines': {u'description': u"True if the grid isn't showing gridlines in the UI.",
                               u'type': u'boolean'},
            u'rowCount': {u'description': u'The number of rows in the grid.',
                          u'format': u'int32',
                          u'type': u'integer'}},
    """
    _dynamic_attr = ['columnCount', 'frozenColumnCount', 'frozenRowCount', 'hideGridlines', 'rowCount']

    def __init__(self, **kwargs):
        """
        Basic contain row_count and column_count
        :param row_count:
        :param column_count:
        :param args:
        :param kwargs:
        """
        super(GridProperties, self).__init__(**kwargs)


class ValueRange(Properties):
    _dynamic_attr = ['majorDimension', 'range', 'values']

    def __init__(self, **kwargs):
        """
        Basic contain row_count and column_count
        :param row_count:
        :param column_count:
        :param args:
        :param kwargs:
        """
        super(ValueRange, self).__init__(**kwargs)




if __name__ == '__main__':
    p = GridProperties()
    p.columnCount = 10
    p.rowCount = 10

    print(str(p))
