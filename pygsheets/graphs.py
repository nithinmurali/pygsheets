#from pygsheets.worksheet import Worksheet

class graphs(object):
	def __init__(self, Worksheet, chart_type, domain, range1, title, chart_data=None):
		self._title = title
		self._chart_type = chart_type
		self._domain = domain
		self._range1 = range1
		self._worksheet = Worksheet
		if chart_data is None:
			self._create_chart()
		else:
			self.set_json(chart_data)

	@property
	def title(self):
		return self._title

	@property
	def domain(self):
		return self._domain

	@property
	def chart_type(self):
		return self._chart_type

	#@TODO return a list of ranges
	@property
	def range(self):
		return self._range1


	def _create_chart(self):
		request = {
		  "addChart": {
			"chart": {
			  "spec": {
				"title": self._title,
				"basicChart": {
				  "chartType": self._chart_type,
				  "axis": [
					{
					  "position": "BOTTOM_AXIS",
					  #"title": "x"
					},
					{
					  "position": "LEFT_AXIS",
					  #"title": "y"
					}
				  ],
				  "domains": [
					{
					  "domain": {
						"sourceRange": {
						  "sources": [
							{
							  "sheetId": self._worksheet.id,
							  "startRowIndex": self._domain[0][0]-1,
							  "endRowIndex": self._domain[1][0],
							  "startColumnIndex": self._domain[0][1]-1,
							  "endColumnIndex": self._domain[1][1],
							}
						  ]
						}
					  }
					}
				  ],
				  "series": [
					{
					  "series": {
						"sourceRange": {
						  "sources": [
							{
							  "sheetId": self._worksheet.id,
							  "startRowIndex": self._range1[0][0]-1,
							  "endRowIndex": self._range1[1][0],
							  "startColumnIndex": self._range1[0][1]-1,
							  "endColumnIndex": self._range1[1][1],
							}
						  ]
						}
					  },
					  "targetAxis": "LEFT_AXIS"
					}
				  ]
				}
			  },
			  "position": {
				"overlayPosition": {
				  "anchorCell": {
					"columnIndex": 3,
					"rowIndex": 0,
					"sheetId": 0
				  }
				}
			  }
			}
		  }
		}
		self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)


	def set_json(self,chart_data):
		sheet_list = chart_data.get('sheets',None)
		if sheet_list:
			for sheet in sheet_list:
				chart_list = []
				chart_list = (sheet.get('charts',None))
				if chart_list:
					for chart in chart_list:
						if (chart.get('spec',{}).get('title',None) == self._title):
							self._chart_type = chart.get('spec',{}).get('basicChart',{}).get('chartType')
						else:
							pass
				else:
					print("no more charts")
		else:
			print('no more sheets')