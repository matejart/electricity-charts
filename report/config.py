"""
Configurable settings
"""

DEBUG = False

DEFAULT_SHEET_NAME = 'Vnos meritev'
DATA_START_ROW = 5
DATE_COLUMN = 'B'

DATASETS = [
	{
		'title': 'MERILNO MESTO 42A',
		'VT column': 'C',
		'MT column': 'F',
	},
	{
		'title': 'MERILNO MESTO 42B',
		'VT column': 'I',
		'MT column': 'L',
	},
]

PDF_HEADER_TITLE = "Skupni prostori stavbe 42a in 42b"

try:
	from local_config import *
except ImportError:
	pass
