#from pygsheets.worksheet import Worksheet

class charts(object):
	def __init__(self, Worksheet, chart_type, domain, range1, title, chart_data=None):
		self._title = title
		self._chart_type = chart_type
		self._domain = domain
		self._range1 = range1
		self._worksheet = Worksheet
		self._title_font_family = 'Roboto'
		self._font_name = 'Roboto'
		self._legend_position = 'RIGHT_LEGEND'
		self._chart_id = None
		if chart_data is None:
			self._create_chart()
		else:
			self.set_json(chart_data)

	@property
	def title(self):
		return self._title

	@title.setter
	def title(self, new_title):
		self._title = new_title
		self.update_chart()

	@property
	def domain(self):
		return self._domain

	@domain.setter
	def domain(self, new_domain):
		self._domain = new_domain
		self.update_chart()

	@property
	def chart_type(self):
		return self._chart_type

	@chart_type.setter
	def chart_type(self, new_chart_type):
		self._chart_type = new_chart_type
		self.update_chart()

	#@TODO return a list of ranges
	@property
	def range(self):
		return self._range1

	@range.setter
	def range(self, new_range1):
		self._range1 = new_range1
		self.update_chart()

	@property
	def title_font_family(self):
		return self._title_text_format

	@title_font_family.setter
	def title_font_family(self, new_title_font_family):
		self._title_text_format = new_title_font_family
		self.update_chart()

	@property
	def font_name(self):
		return self._font_name

	@font_name.setter
	def font_name(self, new_font_name):
		self._font_name = new_font_name
		self.update_chart()

	@property
	def legend_position(self):
		return self._legend_position

	@legend_position.setter
	def legend_position(self, new_legend_position):
		self._legend_position = new_legend_position
		self.update_chart()

	@property
	def chart_id(self):
		return self._chart_id

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


	def update_chart(self):
		request = {
			'updateChartSpec':{
				'chartId': self._chart_id, "spec": self.get_json(),}}
		self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)


	def get_json(self):
		domains = [{'domain':{'sourceRange':{'sources':[{
			'startRowIndex': self._domain[0][0]-1,
			'endRowIndex': self._domain[1][0],
			'startColumnIndex': self._domain[0][1]-1,
			'endColumnIndex': self._domain[1][1],
			'sheetId': self._worksheet.id
		}]}}}]
		ranges = [{'series':{'sourceRange':{'sources':[{
			'startRowIndex': self._range1[0][0]-1,
			'endRowIndex': self._range1[1][0],
			'startColumnIndex': self._range1[0][1]-1,
			'endColumnIndex': self._range1[1][1],
			'sheetId': self._worksheet.id
		}]}}}]
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


	def set_json(self,chart_data):
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
				start_row = source.get('startRowIndex',None)
				end_row = source.get('endRowIndex',None)
				start_column = source.get('startColumnIndex',None)
				end_column = source.get('endColumnIndex',None)
				self._domain = [(start_row+1, start_column+1),(end_row, end_column)]
		range_list = basic_chart.get('series')
		for r in range_list:
			source_list = r.get('series',{}).get('sourceRange',{}).get('sources',None)
			for source in source_list:
				start_row = source.get('startRowIndex',None)
				end_row = source.get('endRowIndex',None)
				start_column = source.get('startColumnIndex',None)
				end_column = source.get('endColumnIndex',None)
				self._range1 = [(start_row+1, start_column+1),(end_row, end_column)]