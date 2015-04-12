# This was originally written around the xlwt module, updated to use xlsxwriter.

import xlsxwriter  # Import xlsxWriter, used for creating xlsx Excel files


class ExcelWriter(object):
	def __init__(self, filename):
		"""
		Creates New Excel Workbook
		Usage:
		>> xls = ExcelWriter(filename)
		>> xls.writesheet('sheet1', ARRAYofDictionaries=[{'a':1,'b':2},{'a':4,'b':10], KnownHeader)
		>> xls.defaultformatting('sheet1')
		>> xls.save()
		"""
		self.filename = filename
		self.workbook = xlsxwriter.Workbook(self.filename)
		self.sheets = {}
		self.sheetoptions = {}
		self.columnstyles = {}
		self.headerindex = {}
		self.columnwidth = {}
		self._headerwritten = 0

	###############
	# Definitions #
	###############
	def add_sheet(self, sheetname):
		if sheetname in self.sheets:
			raise Exception("Sheet already exists: [%s]" % sheetname)
		self.sheets[sheetname] = self.workbook.add_worksheet(sheetname)

	def add_sheet_option(self, sheetname, option):
		if sheetname not in self.sheetoptions:
			self.sheetoptions[sheetname] = {}
		self.sheetoptions[sheetname].update(option)

	def sheet_header(self, sheetname, header):
		self.add_sheet_option(sheetname, {'Header': header})

	def add_column_style(self, sheetname, columnname, style):
		self.add_sheet_option(sheetname, {columnname: style})

	###########
	# Writers #
	###########
	def writecell(self, sheetname, column, row, value, style=None):
		self.sheets[sheetname].write(row, column, value, style)

	def writeheader(self, sheetname, header=None, rowindex=0, columnindex=0):
		if self._headerwritten == 1:
			return 0
		if header is None:
			try:
				header = self.sheetoptions[sheetname]['Header']
			except KeyError:
				return 1
		self.headerindex[sheetname] = {}
		for idx, columnname in enumerate(header):
			self.headerindex[sheetname][columnname] = columnindex + idx
			self.writecell(sheetname, self.headerindex[sheetname][columnname], rowindex, columnname)
		self._headerwritten = 1
		return 0

	def writerow(self, sheetname, rowindex, dictrow):
		if dictrow is None:
			return
		for key, value in dictrow.items():
			style = None
			if sheetname in self.sheetoptions:
				if key in self.sheetoptions[sheetname]:
					style = self.sheetoptions[sheetname][key]
			try:
				self.writecell(sheetname, self.headerindex[sheetname][key], rowindex, value, style)
				try:
					self.update_column_width(sheetname, key, len(value))
				except TypeError:
					self.update_column_width(sheetname, key, 0)
			except KeyError:
				pass

	def writerows(self, sheetname, dictlist, rowindex=1):
		for dictrow in dictlist:
			self.writerow(sheetname, rowindex, dictrow)
			rowindex += 1

	def writesheet(self, sheetname, data, header=None):
		self.add_sheet(sheetname=sheetname)
		hw = self.writeheader(sheetname=sheetname, header=header)
		rowindex = 1
		if hw != 0:
			if isinstance(data, types.GeneratorType) or isinstance(data, types.InstanceType):
				dictrow = data.next()
			elif isinstance(data, types.ListType):
				dictrow = data.pop()
			header = dictrow.keys()
			self.sheet_header(sheetname, header)
			self.writeheader(sheetname=sheetname, header=header)
			self.writerow(sheetname=sheetname, rowindex=rowindex, dictrow=dictrow)
			rowindex = 2
		self.writerows(sheetname=sheetname, dictlist=data, rowindex=rowindex)

	##############
	# Formatting #
	##############
	def style_dateformat(self, dateformat):
		style = self.workbook.add_format()
		style.set_num_format(dateformat)
		return style

	def update_column_width(self, sheetname, colname, width):
		if sheetname not in self.columnwidth:
			self.columnwidth[sheetname] = {}
		if colname not in self.columnwidth[sheetname]:
			self.columnwidth[sheetname][colname] = len(colname)
		if self.columnwidth[sheetname][colname] < width < 200:
			self.columnwidth[sheetname][colname] = width

	def adjust_column_width(self, sheetname):
		for colname, width in self.columnwidth[sheetname].items():
			self.sheets[sheetname].set_column(self.headerindex[sheetname][colname],
			                                  self.headerindex[sheetname][colname], width)

	def freeze_top_row(self, sheetname):
		self.sheets[sheetname].freeze_panes(1, 0)

	def default_formatting(self, sheetname):
		self.adjust_column_width(sheetname)
		self.freeze_top_row(sheetname)

	#########
	# Close #
	#########
	def save(self):
		#self.workbook.save(self.filename)  # Out-Dated
		self.workbook.close()

	def close(self):
		self.save()