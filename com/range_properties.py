from properties import Properties


class GridRange(Properties):
    _dynamic_attr = ['sheetID', 'startRowIndex', 'endRowIndex', 'startColumnIndex', 'endColumnIndex']

    def __init__(self, **kwargs):
        super(GridRange, self).__init__(**kwargs)


class NamedRange(Properties):
    _dynamic_attr = ['namedRangeId', 'name', 'range']

    def __init__(self, **kwargs):
        super(NamedRange, self).__init__(**kwargs)
