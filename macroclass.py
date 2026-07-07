class Workbook:
	# ----- Contructors for workbook class -----
	def __init__(self):
		self.ctx = XSCRIPTCONTEXT   # libreoffice API
		self.smgr = self.ctx.getComponentContext().getServiceManager()    # Libreoffice service manager API
		self.doc = self.ctx.getDocument()       # gets the current working document
		self.active = self.doc.CurrentController    # gets the current controller which is responsible for getting active cells or sheets
		self.Sheets = self.doc.Sheets   # gets all the sheets in the current working doucments as objects
		self.ActiveSheet = self.active.getActiveSheet()     # gets active sheet in the current working document
		self.SheetsCount = self.Sheets.getCount()   # gets the total number of sheets of the current working document
		self.ActiveSheetName = self.ActiveSheet.getName()   # gets the current working sheet's name 
		self.ActiveSheetIndex = self.Sheets.getElementNames().index(self.ActiveSheetName)+1     # gets the current working sheet's postions from the sheets 
		self.MaxRows = self.ActiveSheet.Rows.Count      # gets the total number of rows in the current working sheet
		self.MaxColumns = self.ActiveSheet.Columns.Count	# gets the total number of columns in the current working sheet
		self.NameRanges = self.doc.getPropertyValue("NamedRanges")		# Gets all the defined name for the cells
		self.NameRangesCount = self.NameRanges.getCount()		# Gets count of all the Names in the NameRanges


	# ----- Macro Functions -----

	# FIXED: Selects Cell, Range or Even Sheets.
	def Select(self, obj):
		try:
			self._select_core(obj)
		except Exception:
			try:
				self._select_core(obj)
			except:
				pass
		return obj	

	# Return a Cell with given row and column as an Object Eg Calc.Cell(1 , 1) will return cell A1
	def Cell(self, row=None, col=None):
		return self.ActiveSheet.getCellByPosition(col-1, row-1)

	def Name(self, oCell):
		for i in range(self.NameRangesCount):
			nr = self.NameRanges.getByIndex(i)
			if nr.Content == oCell.AbsoluteName:
				return nr.Name
		return None

	# IMPROVED Now it can set value to a Range Eg: Calc.Range("A1").String = "Hello word" or Calc.Range("A1:"B10") = "Hello world" or Calc.Range(Cell(1 , 1) , Cell(20 , 1))
	def Range(self, fromCell=None , toCell=None): 	
		if toCell==None:
			rng = self.ActiveSheet.getCellRangeByName(fromCell)
		else:
			rng = self.ActiveSheet.getCellRangeByName(fromCell.AbsoluteName+":"+toCell.AbsoluteName)
		class _R:
			def __init__(self, r): self._r = r
			def __getattr__(self, n): return getattr(self._r, n)
			def __setattr__(self, n, v):
				if n == "_r": super().__setattr__(n, v)
				elif n in ("String", "Value", "Formula"):
					rows = self._r.Rows.Count
					cols = self._r.Columns.Count
					row_tuple = (v,) * cols
					data = (row_tuple,) * rows
					if n == "String": self._r.setDataArray(data)
					elif n == "Value": self._r.setDataArray(data)
					elif n == "Formula": self._r.setFormulaArray(data)
				else:
					setattr(self._r, n, v)
		return _R(rng)

	def MsgBox(self, text, title="MsgBox"):
		toolkit = self.smgr.createInstanceWithContext(
			"com.sun.star.awt.Toolkit",
			self.ctx
		)
		parent = self.active.getFrame().getContainerWindow()
		msgbox = toolkit.createMessageBox(
			parent,
			0,
			1,
			title,
			text
		)
		return msgbox.execute()

	# This acts like ActiveCell.Offset in VBA, eg: Calc.Offset(0 , 1) will return the cell from the next column as an object
	def Offset(self, rowIndex, colIndex): 
		maxColumns = self.ActiveSheet.Columns.Count - 1
		maxRows = self.ActiveSheet.Rows.Count - 1
		newRow = self.ActiveCell.RangeAddress.StartRow + rowIndex
		newCol = self.ActiveCell.RangeAddress.StartColumn + colIndex
		newRow = max(0, min(newRow, maxRows))
		newCol = max(0, min(newCol, maxColumns))
		return self.ActiveSheet.getCellByPosition(newCol, newRow)
			
	# Returns the Row Index for the given Cell/Range
	def Row(self, oCell): 
		return 	oCell.RangeAddress.StartRow + 1

	# Returns the Column Index for a given Cell/Range
	def Column(self, oCell): 
		return oCell.RangeAddress.StartColumn + 1

	# Returns the End Row Index for a given Cell/Range
	def EndRow(self , oCell): 
		return oCell.RangeAddress.EndRow + 1

	# Returns the End Column Index for a given Cell/Range
	def EndColumn(self , oCell):
		return oCell.RangeAddress.EndColumn + 1

	# Returns the Address given Cell/Range
	def RangeAddress(self , absolute=False):
		if absolute:
			return self.getActiveSelection().AbsoluteName 
		return (self.active.getSelection().AbsoluteName).replace("$" , "")

	# Returns the Sheet as an Object searched by Sheet Name
	def SheetName(self , sheetName):
		return self.Sheets.getByName(sheetName)

	# Returns the Sheet as an Object searched by Sheet Index
	def SheetIndex(self , index):
		return self.Sheets.getByIndex(index-1)

	# NEW Creates a new sheet by giving sheet name, and using set before attribute Eg Calc.CreateSheet("SHEETNAME" , 1) then the sheet will be created on index 2
	def CreateSheet(self, sheetName, setBefore=None):
		if setBefore is None:
			setBefore = self.ActiveSheetIndex - 1
		elif setBefore > self.SheetsCount:
			setBefore = self.SheetsCount 
		self.Sheets.insertNewByName(sheetName, setBefore)
		newSheet = self.Sheets.getByIndex(setBefore)
		self.active.setActiveSheet(newSheet)
		self.ActiveSheet = newSheet
		self.ActiveSheetName = newSheet.getName()
		self.ActiveSheetIndex = setBefore
		self.SheetsCount = self.Sheets.getCount()
		return newSheet

	# Renames a Sheet by replacing its old name with a new one
	def RenameSheet(self , oldName , newName):
		Sheet = self.SheetName(oldName)
		Sheet.setName(newName)

	# Deletes a sheet by addr, an addr could be sheet's name or index
	def DeleteSheet(self , addr):
		if isinstance(addr , int):
			addr -= 1
			name = self.Sheets.getByIndex(addr).getName()
			self.Sheets.removeByName(name)
		elif isinstance(addr , str):
			self.SearchSheetByName(addr).removeByName(addr)

	# Returns a cell which is at the end of data. Just like VBA's ActiveCell.End(xlToRight/xlToLeft/xlUp/xlDown) Eg: Calc.Select(Calc.End("Right"))
	def End(self, direction):
		direction = direction.lower()
		start_row = self.ActiveCell.RangeAddress.StartRow
		start_col = self.ActiveCell.RangeAddress.StartColumn
		CONTENT_FLAGS = 23 
		if direction == "right":
			if start_col >= self.MaxColumns - 1: return self.ActiveCell
			scan_range = self.ActiveSheet.getCellRangeByPosition(start_col + 1, start_row, self.MaxColumns - 1, start_row)
		elif direction == "left":
			if start_col <= 0: return self.ActiveCell
			scan_range = self.ActiveSheet.getCellRangeByPosition(0, start_row, start_col - 1, start_row)
		elif direction == "down":
			if start_row >= self.MaxRows - 1: return self.ActiveCell
			scan_range = self.ActiveSheet.getCellRangeByPosition(start_col, start_row + 1, start_col, self.MaxRows - 1)
		elif direction == "up":
			if start_row <= 0: return self.ActiveCell
			scan_range = self.ActiveSheet.getCellRangeByPosition(start_col, 0, start_col, start_row - 1)
		else:
			return self.ActiveCell
		filled_ranges = scan_range.queryContentCells(CONTENT_FLAGS).getRangeAddresses()
		empty_ranges = scan_range.queryEmptyCells().getRangeAddresses()
		is_start_empty = len(self.ActiveCell.getString()) == 0
		if direction == "right":
			if is_start_empty:
				# Jump straight to the beginning of the next data cluster
				if filled_ranges:
					return self.ActiveSheet.getCellByPosition(filled_ranges[0].StartColumn, start_row)
				return self.ActiveSheet.getCellByPosition(self.MaxColumns - 1, start_row)
			else:
				# We are sitting on data; find where this current continuous run breaks
				if empty_ranges and empty_ranges[0].StartColumn == start_col + 1:
					# Next cell is already empty! Find where data starts back up again
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(filled_ranges[0].StartColumn, start_row)
					return self.ActiveSheet.getCellByPosition(self.MaxColumns - 1, start_row)
				elif empty_ranges:
					# There is continuous data; jump right to the cell before the first empty gap
					return self.ActiveSheet.getCellByPosition(empty_ranges[0].StartColumn - 1, start_row)
				else:
					# Data runs all the way cleanly to the end edge of the sheet
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(filled_ranges[-1].EndColumn, start_row)
					return self.ActiveSheet.getCellByPosition(self.MaxColumns - 1, start_row)
		elif direction == "left":
			if is_start_empty:
				if filled_ranges:
					return self.ActiveSheet.getCellByPosition(filled_ranges[-1].EndColumn, start_row)
				return self.ActiveSheet.getCellByPosition(0, start_row)
			else:
				if empty_ranges and empty_ranges[-1].EndColumn == start_col - 1:
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(filled_ranges[-1].EndColumn, start_row)
					return self.ActiveSheet.getCellByPosition(0, start_row)
				elif empty_ranges:
					return self.ActiveSheet.getCellByPosition(empty_ranges[-1].EndColumn + 1, start_row)
				else:
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(filled_ranges[0].StartColumn, start_row)
					return self.ActiveSheet.getCellByPosition(0, start_row)
		elif direction == "down":
			if is_start_empty:
				if filled_ranges:
					return self.ActiveSheet.getCellByPosition(start_col, filled_ranges[0].StartRow)
				return self.ActiveSheet.getCellByPosition(start_col, self.MaxRows - 1)
			else:
				if empty_ranges and empty_ranges[0].StartRow == start_row + 1:
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(start_col, filled_ranges[0].StartRow)
					return self.ActiveSheet.getCellByPosition(start_col, self.MaxRows - 1)
				elif empty_ranges:
					return self.ActiveSheet.getCellByPosition(start_col, empty_ranges[0].StartRow - 1)
				else:
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(start_col, filled_ranges[-1].EndRow)
					return self.ActiveSheet.getCellByPosition(start_col, self.MaxRows - 1)
		elif direction == "up":
			if is_start_empty:
				if filled_ranges:
					return self.ActiveSheet.getCellByPosition(start_col, filled_ranges[-1].EndRow)
				return self.ActiveSheet.getCellByPosition(start_col, 0)
			else:
				if empty_ranges and empty_ranges[-1].EndRow == start_row - 1:
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(start_col, filled_ranges[-1].EndRow)
					return self.ActiveSheet.getCellByPosition(start_col, 0)
				elif empty_ranges:
					return self.ActiveSheet.getCellByPosition(start_col, empty_ranges[-1].EndRow + 1)
				else:
					if filled_ranges:
						return self.ActiveSheet.getCellByPosition(start_col, filled_ranges[0].StartRow)
					return self.ActiveSheet.getCellByPosition(start_col, 0)
		return self.ActiveCell

	# Calc.WorksheetFunctions("functionName in capital" , [args as tuple]) for eg: Calc.Offset(0 , -1).Value = Calc.WorksheetFunctions("ISBLANK" , [Calc.ActiveCell.getString()].
	# IF you are giving cell ref. always add .getString() ahead to get the cell contents as String
	def WorksheetFunction(self, functionName , args):
		smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
		func_access = smgr.createInstanceWithContext(
			"com.sun.star.sheet.FunctionAccess",
			XSCRIPTCONTEXT.getComponentContext()
		)
		return func_access.callFunction(functionName, tuple(args))

	def Copy(self, oRange=None):
		if oRange is None:
			oRange = self.ActiveRange
		if hasattr(oRange, "_r"):
			oRange = oRange._r
		if self.IsSheet(oRange):
			oRange = oRange.getDataArea()
		self._clipboard = {
			"data": oRange.getDataArray(),
			"values": oRange.getDataArray(),
			"formulas": oRange.getFormulaArray()
		}
		return self._clipboard

	def Paste(self, oRange=None, special=None):
		if oRange is None:
			oRange = self.ActiveRange
		if hasattr(oRange, "_r"):
			oRange = oRange._r
		if self._clipboard is None:
			return
		if special is None or special.lower() == "all":
			oRange.setDataArray(self._clipboard["data"])
		elif special.lower() == "values":
			oRange.setDataArray(self._clipboard["values"])
		elif special.lower() == "formulas":
			oRange.setFormulaArray(self._clipboard["formulas"])
		return oRange


	# ----- Funtions for workbook class extension -----

	# Rentrus Boolean value to check the 
	def IsNull(self, oCell):
		return True if len(oCell.getString())==0 else False
	
	def IsSheet(self, obj):
    	return hasattr(obj, "getCellRangeByPosition") and not hasattr(obj, "RangeAddress")

	def IsRange(self, obj):
		return hasattr(obj, "RangeAddress")

	# Fixed: Returns ActiveCell as and Object
	@property
	def ActiveCell(self):
		return self.active.getSelection()
	
	@property
	def ActiveRange(self):
		return self.active.getSelection()

	# Fixed: Raw Select Function that selects Cell/Range. NOTE: using this function might get error while using  
	def _select_core(self, obj):
		if hasattr(obj, "_r"):
			obj = obj._r
		if hasattr(obj, "getCellByPosition") and not hasattr(obj, "RangeAddress"):
			self.ActiveSheet = obj
			self.active.setActiveSheet(obj)
			self.ActiveCell = self.Cell(self.Row(obj), self.Column(obj))
		elif hasattr(obj, "RangeAddress"):
			self.active.select(obj)
			self.ActiveCell = self.ActiveSheet.getCellByPosition(
				obj.RangeAddress.StartColumn,
				obj.RangeAddress.StartRow
			)

# Add the below code at the bottom
Calc = Workbook()

ctx = Calc.ctx
smgr = Calc.smgr
doc = Calc.doc
active = Calc.active
Sheets = Calc.Sheets
ActiveSheet = Calc.ActiveSheet
SheetsCount = Calc.SheetsCount
ActiveSheetName = Calc.ActiveSheetName
ActiveSheetInd = Calc.ActiveSheetIndex
MaxRows = Calc.MaxRows
MaxColumns = Calc.MaxColumns
NameRanges = Calc.NameRanges
NameRangesCount = Calc.NameRangesCount
Select = Calc.Select
Cell = Calc.Cell
Name = Calc.Name
Range = Calc.Range
Offset = Calc.Offset
Row = Calc.Row
Column = Calc.Column
EndRow = Calc.EndRow
EndColumn = Calc.EndColumn
RangeAddress = Calc.RangeAddress
SheetName = Calc.SheetName
SheetIndex = Calc.SheetIndex
CreateSheet = Calc.CreateSheet
RenameSheet = Calc.RenameSheet
DeleteSheet = Calc.DeleteSheet
End = Calc.End
GetDataArray = Calc.GetDataArray
SetDataArray = Calc.SetDataArray
WorksheetFunction = Calc.WorksheetFunction
IsNull = Calc.IsNull
IsSheet = Calc.IsSheet
IsRange = Calc.IsRange
ActiveCell = Calc.ActiveCell
ActiveRange = Calc.ActiveRange



# Your Macro will start from any function you create outside the Class
def Automate():
	# Your Macro start from here
	Select(Range("A1"))

	
