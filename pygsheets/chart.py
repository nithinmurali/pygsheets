from pygsheets.utils import format_addr

class Chart(object):
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

	@property
	def ranges(self):
		return self._ranges

	@ranges.setter
	def ranges(self, new_ranges):
		self._ranges = new_ranges
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

	def get_anchor_cell(self):
		if self._anchor_cell is None:
			return {
				"columnIndex": self._domain[1][1]-1,
				"rowIndex": self._domain[1][0],"sheetId": self._worksheet.id}
		else:
			cell = format_addr(self._anchor_cell)
			return {
				"columnIndex": cell[1],
				"rowIndex": cell[0]+1,"sheetId": self._worksheet.id}

	def get_ranges_request(self):
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
				  "anchorCell": self.get_anchor_cell()
				}
			  }
			}
		  }
		}
		response = self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)
		chart_data_list = response.get('replies')
		chart_json = chart_data_list[0].get('addChart',{}).get('chart')
		self.set_json(chart_json)

	def update_chart(self):
		request = {
			'updateChartSpec':{
				'chartId': self._chart_id, "spec": self.get_json(),}}
		self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

	def get_json(self):
		domains = [{'domain':{'sourceRange':{'sources':[
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
		self._ranges = []
		for r in range_list:
			source_list = r.get('series',{}).get('sourceRange',{}).get('sources',None)
			for source in source_list:
				start_row = source.get('startRowIndex',None)
				end_row = source.get('endRowIndex',None)
				start_column = source.get('startColumnIndex',None)
				end_column = source.get('endColumnIndex',None)
				self._ranges.append([(start_row+1, start_column+1),(end_row, end_column)])