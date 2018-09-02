from pygsheets.utils import format_addr
from pygsheets.cell import Cell


class Chart(object):
    """
    Represents a chart in a sheet.

    :Param Worksheet:       Represents the current working worksheet

    :Param domain:          Cell range of the desired chart domain in the form of list of tuples

    :Param ranges:          Cell ranges of the desired ranges in the form of list of list of tuples

    :Param chart_type:      The supported chart types are given in the link below-
                            https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#BasicChartType

    :Param title:           Title of the chart

    :Param anchor_cell:     position of the left corner of the chart in the form of cell address or cell object

    :Param chart_data:      Represents a json (dictionary) structure of the chart
    """
    def __init__(self, Worksheet, domain, ranges, chart_type, title=None, anchor_cell=None, chart_data=None):
        self._title = title
        self._chart_type = chart_type
        self._domain = domain
        self._ranges = ranges
        self._worksheet = Worksheet
        self._title_font_family = 'Roboto'
        self._font_name = 'Roboto'
        self._legend_position = 'RIGHT_LEGEND'
        self._chart_id = None
        self._anchor_cell = anchor_cell
        if chart_data is None:
            self._create_chart()
        else:
            self.set_json(chart_data)

    @property
    def title(self):
        """Title of the chart"""
        return self._title

    @title.setter
    def title(self, new_title):
        temp = self._title
        self._title = new_title
        try:
            self.update_chart()
        
        except:
            self._title = temp

    @property
    def domain(self):
        """Domain of the chart"""
        return self._domain

    @domain.setter
    def domain(self, new_domain):
        temp = self._domain
        self._domain = new_domain
        try:
            self.update_chart()
        except:
            self._domain = temp

    @property
    def chart_type(self):
        """Type of the chart"""
        return self._chart_type

    @chart_type.setter
    def chart_type(self, new_chart_type):
        temp = self._chart_type
        self._chart_type = new_chart_type
        try:
            self.update_chart()
        except:
            self._chart_type = temp

    @property
    def ranges(self):
        """Ranges of the chart"""
        return self._ranges

    @ranges.setter
    def ranges(self, new_ranges):
        temp = self._ranges
        self._ranges = new_ranges
        try:
            self.update_chart()
        except:
            self._ranges = temp

    @property
    def title_font_family(self):
        """
        Font family of the title.
        Default value is set to 'Roboto'
        """
        return self._title_font_family

    @title_font_family.setter
    def title_font_family(self, new_title_font_family):
        temp = self._title_font_family
        self._title_font_family = new_title_font_family
        try:
            self.update_chart()
        except:
            self._title_font_family = temp

    @property
    def font_name(self):
        """
        Font name for the chart.
        Default value is set to 'Roboto'
        """
        return self._font_name

    @font_name.setter
    def font_name(self, new_font_name):
        temp = self._font_name
        self._font_name = new_font_name
        try:
            self.update_chart()
        except:
            self._font_name = temp

    @property
    def legend_position(self):
        """
        Legend postion of the chart.
        The available options are given in the below link-
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#BasicChartLegendPosition
        """
        return self._legend_position

    @legend_position.setter
    def legend_position(self, new_legend_position):
        temp = self._legend_position
        self._legend_position = new_legend_position
        try:
            self.update_chart()
        except:
            self._legend_position = temp

    @property
    def chart_id(self):
        """Id of the current chart."""
        return self._chart_id

    @property
    def anchor_cell(self):
        """position of the left corner of the chart in the form of cell address or cell object"""
        return self._anchor_cell

    @anchor_cell.setter
    def anchor_cell(self, new_anchor_cell):
        temp = self._anchor_cell
        try:
            if type(new_anchor_cell) is Cell:
                self._anchor_cell = (new_anchor_cell.row, new_anchor_cell.col)
                self._update_position()
            else:
                self._anchor_cell = format_addr(new_anchor_cell, 'tuple')
                self._update_position()
        except:
            self._anchor_cell = temp

    def delete_chart(self):
        request = {
            "deleteEmbeddedObject": {
                "objectId": self._chart_id
            }
        }
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def refresh(self):
        chart_data = self._worksheet.client.sheet.get(self._worksheet.spreadsheet.id,fields='sheets(charts,properties)')
        sheet_list = chart_data.get('sheets')
        for sheet in sheet_list:
            if sheet.get('properties',{}).get('sheetId') is self._worksheet.id:
                chart_list = sheet.get('charts')
                if chart_list:
                    for chart in chart_list:
                        if chart.get('chartId') == self._chart_id:
                            self.set_json(chart)

    def _get_anchor_cell(self):
        """Returns the cell in the form of a dictionary structure.

        The cell holds the left corner of the chart."""

        if self._anchor_cell is None:
            return {
                "columnIndex": self._domain[1][1]-1,
                "rowIndex": self._domain[1][0],"sheetId": self._worksheet.id}
        else:
            if type(self._anchor_cell) is Cell:
                return {
                "columnIndex": self._anchor_cell.col-1,
                "rowIndex": self._anchor_cell.row-1,"sheetId": self._worksheet.id}
            else:
                cell = format_addr(self._anchor_cell, 'tuple')
                return {
                    "columnIndex": cell[1]-1,
                    "rowIndex": cell[0]-1,"sheetId": self._worksheet.id}

    def get_ranges_request(self):
        """Returns a list of dictionary structured ranges for the desired chart."""

        ranges_request_list = []
        for i in range(len(self._ranges)):
            req = {
                'series':{
                    'sourceRange':{
                        'sources':[self._worksheet.get_gridrange(self._ranges[i][0], self._ranges[i][1])]
                    }
                },
            }
            ranges_request_list.append(req)
        return ranges_request_list

    def _create_chart(self):
        """Creates a chart in the working Google sheet."""

        request = {
          "addChart": {
            "chart": {
              "spec": {
                "title": self._title,
                "basicChart": {
                  "chartType": self._chart_type,
                  "domains": [
                    {
                      "domain": {
                        "sourceRange": {
                          "sources": [
                              self._worksheet.get_gridrange(self._domain[0], self._domain[1])
                          ]
                        }
                      }
                    }
                  ],
                 "series": self.get_ranges_request()
                }
              },
              "position": {
                "overlayPosition": {
                  "anchorCell": self._get_anchor_cell()
                }
              }
            }
          }
        }
        response = self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)
        chart_data_list = response.get('replies')
        chart_json = chart_data_list[0].get('addChart',{}).get('chart')
        self.set_json(chart_json)

    def _update_position(self):
        request = {
            "updateEmbeddedObjectPosition": {
                "objectId": self._chart_id,
                "newPosition": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId": self._worksheet.id,
                            "rowIndex": self._anchor_cell[0]-1,
                            "columnIndex": self._anchor_cell[1]-1
                        }
                    }
                },
                "fields": "*"
        }} 
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_chart(self):
        """updates the applied changes to the sheet."""
        request = {
            'updateChartSpec':{
                'chartId': self._chart_id, "spec": self.get_json(),}
        }
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def get_json(self):
        """Returns the chart as a dictionary structured like the Google Sheets API v4."""

        domains = [{'domain': {'sourceRange': {'sources': [
            self._worksheet.get_gridrange(self._domain[0], self._domain[1])]}}}]
        ranges = self.get_ranges_request()
        spec = dict()
        spec['title'] = self._title
        spec['basicChart'] = dict()
        spec['titleTextFormat'] = dict()
        spec['basicChart']['chartType'] = self._chart_type
        spec['basicChart']['legendPosition'] = self._legend_position
        spec['titleTextFormat']['fontFamily'] = self._title_font_family
        spec['fontName'] = self._font_name
        spec['basicChart']['domains'] = domains
        spec['basicChart']['series'] = ranges
        return spec

    def set_json(self, chart_data):
        """
        Reads a json-dictionary returned by the Google Sheets API v4 and initialize all the properties from it.

        :param chart_data:   The chart data.
        """
        anchor_cell_data = chart_data.get('position',{}).get('overlayPosition',{}).get('anchorCell')
        self._anchor_cell = (anchor_cell_data.get('rowIndex',0)+1, anchor_cell_data.get('columnIndex',0)+1)
        self._title = chart_data.get('spec',{}).get('title',None)
        self._chart_id = chart_data.get('chartId',None)
        self._title_font_family = chart_data.get('spec',{}).get('titleTextFormat',{}).get('fontFamily',None)
        self._font_name = chart_data.get('spec',{}).get('titleTextFormat',{}).get('fontFamily',None)    
        basic_chart = chart_data.get('spec',{}).get('basicChart',None)
        self._chart_type = basic_chart.get('chartType',None)
        self._legend_position = basic_chart.get('legendPosition',None)
        domain_list = basic_chart.get('domains')
        for d in domain_list:
            source_list = d.get('domain',{}).get('sourceRange',{}).get('sources',None)
            for source in source_list:
                start_row = source.get('startRowIndex',0)
                end_row = source.get('endRowIndex',0)
                start_column = source.get('startColumnIndex',0)
                end_column = source.get('endColumnIndex',0)
                self._domain = [(start_row+1, start_column+1),(end_row, end_column)]
        range_list = basic_chart.get('series')
        self._ranges = []
        for r in range_list:
            source_list = r.get('series',{}).get('sourceRange',{}).get('sources',None)
            for source in source_list:
                start_row = source.get('startRowIndex',0)
                end_row = source.get('endRowIndex',0)
                start_column = source.get('startColumnIndex',0)
                end_column = source.get('endColumnIndex',0)
                self._ranges.append([(start_row+1, start_column+1),(end_row, end_column)])
