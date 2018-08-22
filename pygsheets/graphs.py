class graphs(object):
	# List of properties-
	# 	chart id
	# 	chart spec-
	# 		title
	# 		alt Text
	# 		title text format
	# 			foreground color : r,b,g,alpha
	# 			font family: string
	# 			font size: num
	# 			bold : bool
	# 			italic : bool
	# 			strike thru : bool
	# 			underline : bool
	# 		title text position
	# 			horizontal alignment : enum(4 options)

	# 		subtitle : string
	# 		subtitle text format
	# 			foreground color : r,b,g,alpha
	# 			font family: string
	# 			font size: num
	# 			bold : bool
	# 			italic : bool
	# 			strike thru : bool
	# 			underline : bool
	# 		subtitle text position
	# 			horizontal alignment : enum(4 options)
	# 		font name : string
	# 		maximized : bool
	# 		background color : r,b,g,alpha
	# 		hidden dimension strategy : enum()
	# 		chart type 

	# 	position-
	# 		sheet id : num
	# 		overlay position 
	# 			anchor cell 
	# 				sheet id: num
	# 				row index : num
	# 				col index : num
	# 			offset x pixels : num
	# 			offset y pixels : num
	# 			width pixel: num
	# 			height pixel : num
	# 		new sheet : bool

	def __init__(self):
		self._value = 5


	def basic_chart(self, chart_catagory, chart_type, start_row_index, end_row_index, target_col_index):
		request = {
		  "addChart": {
		    "chart": {
		      "spec": {
		        "title": "practice",
		        "basicChart": {
		          "chartType": "LINE",
		          "axis": [
		            {
		              "position": "BOTTOM_AXIS",
		              "title": "x"
		            },
		            {
		              "position": "LEFT_AXIS",
		              "title": "y"
		            }
		          ],
		          "domains": [
		            {
		              "domain": {
		                "sourceRange": {
		                  "sources": [
		                    {
		                      "sheetId": 0,
		                      "startRowIndex": 0,
		                      "endRowIndex": 3,
		                      "startColumnIndex": 0,
		                      "endColumnIndex": 1
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
		                      "sheetId": 0,
		                      "startRowIndex": 0,
		                      "endRowIndex": 3,
		                      "startColumnIndex": 1,
		                      "endColumnIndex": 2
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
		return request



