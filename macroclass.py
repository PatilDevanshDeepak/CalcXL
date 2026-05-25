class WorkBook:
	def __init__(self):				# Constructors of Workbook class
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.active = self.doc.CurrentController
		self.Sheets = self.doc.Sheets
		self.ActiveSheet = self.active.getActiveSheet()
		self.SheetsCount = self.Sheets.getCount()
		self.ActiveSheetName = self.ActiveSheet.getName()
		self.ActiveSheetIndex = self.Sheets.getElementNames().index(self.ActiveSheetName)+1


	# Funtions for workbook class development
	def isNull(self, oCell):
		return True if len(oCell.getString())==0 else False

	@property
	def ActiveCell(self):
		return self.active.getSelection()

	
	# Macro Functionalities
	def Select(self, obj):					# IMROVED Now you can Select Cell/Range or a Sheet, Eg: Calc.Select(Calc.Cell(1 , 1)) or Calc.Select(Calc.Range("A1:C10")) or Calc.Select(Calc.SheetName())       
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
	
	def Cell(self, row, col):					# Return a Cell with given row and column as an Object Eg Calc.Cell(1 , 1) will return cell A1
		return self.ActiveSheet.getCellByPosition(col-1, row-1)
		
	def Range(self, fromCell=None , toCell=None):					# IMPROVED Now it can set value to a Range Eg: Calc.Range("A1:A10").String = "Hello word"  
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
	
	def Offset(self, rowIndex, colIndex):					# This acts as ActiveCell.Offset in VBA 
		maxColumns = self.ActiveSheet.Columns.Count - 1
		maxRows = self.ActiveSheet.Rows.Count - 1
		newRow = self.ActiveCell.RangeAddress.StartRow + rowIndex
		newCol = self.ActiveCell.RangeAddress.StartColumn + colIndex
		newRow = max(0, min(newRow, maxRows))
		newCol = max(0, min(newCol, maxColumns))
		return self.ActiveSheet.getCellByPosition(newCol, newRow)
			
	def Row(self, oCell):					# Returns the Start Row of a Cell/Range 
		return 	oCell.RangeAddress.StartRow + 1

	def Column(self, oCell):		# Returns the Start Column of a Cell/Range 
		return oCell.RangeAddress.StartColumn + 1
	
	def EndRow(self , oCell):					# NEW Returns End Row of a Range as an int 
		return self.RangeAddress.EndRow + 1
	
	def EndColumn(self , oCell):					# NEW Returns End Column of a Range as an int 
		return self.RangeAddress.EndColumn + 1
	
	def RangeAddress(self , absolute=False):			# NEW Returns the Address of a Cell/Range 
		if absolute:
			return self.getActiveSelection().AbsoluteName 
		return (self.active.getSelection().AbsoluteName).replace("$" , "")

	def SheetName(self , sheetName):					# NEW Returns the Sheet as an Object searched by Sheet Name
		return self.Sheets.getByName(sheetName)
	
	def SheetIndex(self , index):					# NEW Returns the Sheet as an Object searched by Sheet Index
		return self.Sheets.getByIndex(index-1)

	def CreateSheet(self, sheetName, setBefore=None):					# NEW Creates a new sheet by giving sheet name, and using set before attribute Eg Calc.CreateSheet("SHEETNAME" , 1) then the sheet will be created on index 2
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
	
	def RenameSheet(self , oldName , newName):
		Sheet = self.SheetName(oldName)
		Sheet.setName(newName)
	
	def DeleteSheet(self , addr):
		if isinstance(addr , int):
			addr -= 1
			name = self.Sheets.getByIndex(addr).getName()
			self.Sheets.removeByName(name)
		elif isinstance(addr , str):
			self.SearchSheetByName(addr).removeByName(addr)

	def ClearContents(self, obj, clearAll=False):		# Note: This Function Have Bugs !!!
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
	
	def clcToRight(self , oCell=None):	# This function is like xlToRight in VBA   Note: This Function Have Bugs !!!
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndColumn()-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(0 , count-1 if not self.isNull(self.ActiveCell) else count)
	
	def clcToLeft(self):	# This function is like xlToLeft in VBA    Note: This Function Have Bugs !!!
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while abs(count) <= self.getStartColumn()-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count -= 1
		return self.Offset(0 , count+1 if not self.isNull(self.ActiveCell) else count)
	
	def clcDown(self):	# This function is like xlDown in VBA		Note: This Function Have Bugs !!!
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndRow()-1:
			nullCell = self.isNull(self.Offset(count , 0))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(count-1 if not self.isNull(self.ActiveCell) else count , 0)
	
	def clcUp(self):	# This function is like xlUp in VBA		Note: This Function Have Bugs In It !!!
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getStartRow()-1:
			nullCell = self.isNull(self.Offset(count , 0))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count -= 1
		return self.Offset(count+1 if not self.isNull(self.ActiveCell) else count , 0)

	def WorksheetFunctions(self, functionName , args):       # Calc.WorksheetFunctions("functionName in capital" , [args as tuple]) for eg: Calc.Offset(0 , -1).Value = Calc.WorksheetFunctions("ISBLANK" , [Calc.ActiveCell.getString()]. IF you are giving cell ref. always add .getString() ahead
		smgr = XSCRIPTCONTEXT.getComponentContext().ServiceManager
		func_access = smgr.createInstanceWithContext(
			"com.sun.star.sheet.FunctionAccess",
			XSCRIPTCONTEXT.getComponentContext()
		)
		return func_access.callFunction(functionName, tuple(args))


Calc = WorkBook()

def Automate():
	Calc.Range(fromCell=Calc.Cell(1 , 1) , toCell=Calc.Cell(10 , 2)).String = "Hello world"
