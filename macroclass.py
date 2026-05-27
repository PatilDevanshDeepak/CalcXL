class Workbook:
	# ----- Contructors for workbook class -----
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()		# gets the current working document
		self.active = self.doc.CurrentController 	# gets the current controller which is responsible for getting active cells or sheets
		self.Sheets = self.doc.Sheets 	# gets all the sheets in the current working doucments as objects
		self.ActiveSheet = self.active.getActiveSheet() 	# gets active sheet in the current working document
		self.SheetsCount = self.Sheets.getCount() 	# gets the total number of sheets of the current working document
		self.ActiveSheetName = self.ActiveSheet.getName() 	# gets the current working sheet's name 
		self.ActiveSheetIndex = self.Sheets.getElementNames().index(self.ActiveSheetName)+1 	# gets the current working sheet's postions from the sheets 
		self.MaxRows = self.ActiveSheet.Rows.Count 		# gets the total number of rows in the current working sheets
		self.MaxColumns = self.ActiveSheet.Columns.Count	 # gets the totol number of columns in the current working sheets


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
	
	# Return a Cell with given row and column as an Object Eg Calc.Cell(1 , 1) will return cell A1
	def Cell(self, row, col):
		return self.ActiveSheet.getCellByPosition(col-1, row-1)
		
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

	# Clears the content from given Cell/Range. NOTE: THIS FUNCTION HAS BUGS
	def ClearContents(self, obj, clearAll=False):
		# Clear All Values can be:
		FLAG_STRING = 1
		FLAG_VALUE = 2
		FLAG_FORMULA = 4
		FLAG_HARDATTR = 16  # formatting
		FLAG_ALL = 7	# clear all

		## NOTE: if the clearAll is not given it will consider 7 as value
		if clearAll:
			flags = FLAG_STRING | FLAG_VALUE | FLAG_FORMULA | FLAG_HARDATTR
		else:
			flags = FLAG_STRING | FLAG_VALUE | FLAG_FORMULA   # = 7
	
	# This function is like xlToRight in VBA. NOTE: THIS FUNCTION HAS BUGS
	def clcToRight(self):
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndColumn()-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(0 , count-1 if not self.isNull(self.ActiveCell) else count)
	
	# This function is like xlToLeft in VBA. NOTE: THIS FUNCTION HAS BUGS
	def clcToLeft(self):
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while abs(count) <= self.getStartColumn()-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count -= 1
		return self.Offset(0 , count+1 if not self.isNull(self.ActiveCell) else count)
	
	# This function is like xlDown in VBA. 
	def clcDown(self):
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndRow()-1:
			nullCell = self.isNull(self.Offset(count , 0))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(count-1 if not self.isNull(self.ActiveCell) else count , 0)
	
	# This function is like xlUp in VBA. NOTE: THIS FUNCTION HAS BUGS
	def clcUp(self):
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getStartRow()-1:
			nullCell = self.isNull(self.Offset(count , 0))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count -= 1
		return self.Offset(count+1 if not self.isNull(self.ActiveCell) else count , 0)

	# Calc.WorksheetFunctions("functionName in capital" , [args as tuple]) for eg: Calc.Offset(0 , -1).Value = Calc.WorksheetFunctions("ISBLANK" , [Calc.ActiveCell.getString()].
	# IF you are giving cell ref. always add .getString() ahead to get the cell contents as String
	def WorksheetFunctions(self, functionName , args):
		smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
		func_access = smgr.createInstanceWithContext(
			"com.sun.star.sheet.FunctionAccess",
			XSCRIPTCONTEXT.getComponentContext()
		)
		return func_access.callFunction(functionName, tuple(args))


	# ----- Funtions for workbook class extension -----

	# Rentrus Boolean value to check the 
	def isNull(self, oCell):
		return True if len(oCell.getString())==0 else False

	# Fixed: Returns ActiveCell as and Object
	@property
	def ActiveCell(self):
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


# Your Macro will start from any function you create outside the Class
def Automate():
	# Your Macro start from here
	Calc.ActiveCell.String = Calc.maxColumns


# If you want to add another macro, you can create a new function
def test():
	Calc.Select(Calc.Range("A1"))
