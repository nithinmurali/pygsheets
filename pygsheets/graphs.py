#from pygsheets.worksheet import Worksheet

class graphs(object):
	def __init__(self, Worksheet, chart_type, domain, range1, title):
		self.title = title
		self.chart_type = chart_type
		self.domain = domain
		self.range1 = range1
		self.worksheet = Worksheet
		print(self.title)
		self.create_chart()

	def create_chart(self):
		request = {
		  "addChart": {
			"chart": {
			  "spec": {
				"title": self.title,
				"basicChart": {
				  "chartType": self.chart_type,
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
							  "sheetId": self.worksheet.id,
							  "startRowIndex": self.domain[0][0]-1,
							  "endRowIndex": self.domain[1][0],
							  "startColumnIndex": self.domain[0][1]-1,
							  "endColumnIndex": self.domain[1][1],
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
							  "sheetId": self.worksheet.id,
							  "startRowIndex": self.range1[0][0]-1,
							  "endRowIndex": self.range1[1][0],
							  "startColumnIndex": self.range1[0][1]-1,
							  "endColumnIndex": self.range1[1][1],
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
		# return request
		self.worksheet.client.sheet.batch_update(self.worksheet.spreadsheet.id, request)


	
