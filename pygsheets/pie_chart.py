from pygsheets.chart import Chart


class PieChart(Chart):
    """
    Represents a Pie Chart in a worksheet.

    :param worksheet:         Worksheet object in which the chart resides
    :param domain:            Cell range of the desired chart domain in the form of tuple of tuples
    :param chart_range:       Cell ranges of the desired (singular) range in the form of a tuple of tuples
    :param title:             Title of the chart
    :param anchor_cell:       Position of the left corner of the chart in the form of cell address or cell object
    :param three_dimensional  True if the pie is three dimensional
    :param pie_hole           (float) The size of the hole in the pie chart (defaults to 0). Must be between 0 and 1.
    :param json_obj:          Represents a json structure of the chart as given in `api <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#BasicChartSpec>`__.
    """
    def __init__(self, worksheet, domain=None, chart_range=None, title='', anchor_cell=None, three_dimensional=False,
                 pie_hole=0, json_obj=None):
        self._three_dimensional = three_dimensional
        self._pie_hole = pie_hole

        if self._pie_hole < 0 or self._pie_hole > 1:
            raise ValueError("Pie Chart's pie_hole must be between 0 and 1.")

        super().__init__(worksheet, domain, ranges=[chart_range], chart_type=None, title=title,
                         anchor_cell=anchor_cell, json_obj=json_obj)

    def get_json(self):
        """Returns the pie chart as a dictionary structured like the Google Sheets API v4."""

        domains = [{'domain': {'sourceRange': {'sources': [
            self._worksheet.get_gridrange(self._domain[0], self._domain[1])]}}}]
        ranges = self._get_ranges_request()
        spec = dict()
        spec['title'] = self._title
        spec['pieChart'] = dict()
        spec['pieChart']['legendPosition'] = self._legend_position
        spec['fontName'] = self._font_name
        spec['pieChart']['domain'] = domains[0]["domain"]
        spec['pieChart']['series'] = ranges[0]["series"]
        spec['pieChart']['threeDimensional'] = self._three_dimensional
        spec['pieChart']['pieHole'] = self._pie_hole
        return spec

    def _create_chart(self):
        domains = []
        if self._domain:
            domains.append({
                "domain": {
                    "sourceRange": {
                        "sources": [self._worksheet.get_gridrange(self._domain[0], self._domain[1])]
                    }
                }
            })

        request = {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": self._title,
                        "pieChart": {
                            "domain": domains[0]["domain"] if domains else None,
                            "series": self._get_ranges_request()[0]["series"],
                            "threeDimensional": self._three_dimensional,
                            "pieHole": self._pie_hole
                        },
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

    def set_json(self, chart_data):
        """
        Reads a json-dictionary returned by the Google Sheets API v4 and initialize all the properties from it.

        :param chart_data:   The chart data as json specified in sheets api.
        """
        anchor_cell_data = chart_data.get('position',{}).get('overlayPosition',{}).get('anchorCell')
        self._anchor_cell = (anchor_cell_data.get('rowIndex',0)+1, anchor_cell_data.get('columnIndex',0)+1)
        self._title = chart_data.get('spec',{}).get('title',None)
        self._chart_id = chart_data.get('chartId',None)
        self._title_font_family = chart_data.get('spec',{}).get('titleTextFormat',{}).get('fontFamily',None)
        self._font_name = chart_data.get('spec',{}).get('titleTextFormat',{}).get('fontFamily',None)
        pie_chart = chart_data.get('spec',{}).get('pieChart', None)
        self._legend_position = pie_chart.get('legendPosition', None)
        domain = pie_chart.get('domain', {})
        source_list = domain.get('sourceRange', {}).get('sources', None)
        for source in source_list:
            start_row = source.get('startRowIndex',0)
            end_row = source.get('endRowIndex',0)
            start_column = source.get('startColumnIndex',0)
            end_column = source.get('endColumnIndex',0)
            self._domain = [(start_row+1, start_column+1),(end_row, end_column)]
        range = pie_chart.get('series', {})
        self._ranges = []
        source_list = range.get('sourceRange',{}).get('sources',None)
        for source in source_list:
            start_row = source.get('startRowIndex',0)
            end_row = source.get('endRowIndex',0)
            start_column = source.get('startColumnIndex',0)
            end_column = source.get('endColumnIndex',0)
            self._ranges.append([(start_row+1, start_column+1), (end_row, end_column)])
