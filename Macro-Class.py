import uno

class WorkBook:
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.active = self.doc.CurrentController
		self.ActiveSheet = self.getActiveSheet()
		self.ActiveCell = self.getActiveSelection()


	# Macro Functions a
	def Select(self, oCell):	# Select a given cell
		self.active.select(oCell)
		self.ActiveCell = oCell.getCellByPosition(0 , 0)
	
	def Cell(self, row, col):	# Returns a cell as an object
		return self.ActiveSheet.getCellByPosition(col-1, row-1)
		
	def Range(self, rangeName=None , fromCell=None , toCell=None):	# Returns ranges like A1:F5 as an object
		if self.isRangeNameGiven(isRowIndex=fromCell , isColumnIndex=toCell , isRangeName=rangeName):
			return self.ActiveSheet.getCellRangeByName(rangeName)
		else:
			return self.ActiveSheet.getCellRangeByName(fromCell.AbsoluteName+":"+toCell.AbsoluteName)
		
	def Offset(self, rowIndex, colIndex):	# jumps as given rowIndex/column index
		maxColumns = self.ActiveSheet.Columns.Count - 1
		maxRows = self.ActiveSheet.Rows.Count - 1
		newRow = self.ActiveCell.RangeAddress.StartRow + rowIndex
		newCol = self.ActiveCell.RangeAddress.StartColumn + colIndex
		newRow = max(0, min(newRow, maxRows))
		newCol = max(0, min(newCol, maxColumns))
		return self.ActiveSheet.getCellByPosition(newCol, newRow)
		
	def Row(self, oCell):	# Returns the Row number of a given cell
		return 	oCell.RangeAddress.StartRow + 1
	
	def Column(self, oCell):	# Returns the Column number of a given cell
		return oCell.RangeAddress.StartColumn + 1
	
	def MsgBox(self, message="none", title="Information", msgBoxType=0):	# Issues with this function
		self.active.getFrame().getComponentWindow().showMessageBox(message, title, msgBoxType)

	def getStartColumn(self):	# Returns starting column of ActiveCell as a number, suppose ActiveCell = G1, starting column (in number) = 1 as column "A"
		return oCell.RangeAddress.StartColumn

	def getStartRow(self):	# Returns starting Row of ActiveCell as a number, suppose ActiveCell = A15, starting row (in number) = 1 as row "1"
		return oCell.RangeAddress.StartRow

	def getEndColumn(self):
		return self.ActiveSheet.Columns.Count

	def getEndRow(self):
		return self.ActiveSheet.Rows.Count

	def clcToRight(self):	# This function is like xlToRight in VBA
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndColumn()-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(0 , count-1 if not self.isNull(self.ActiveCell) else count)

	def clcToLeft(self):	# This function is like xlToLeft in VBA
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while abs(count) <= self.getStartColumn()-1:
			nullCell = self.isNull(self.Offset(0 , count))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count -= 1
		return self.Offset(0 , count+1 if not self.isNull(self.ActiveCell) else count)

	def clcDown(self):	# This function is like xlDown in VBA
		nullCell = self.isNull(self.ActiveCell)
		count = 0
		while count <= self.getEndRow()-1:
			nullCell = self.isNull(self.Offset(count , 0))
			if nullCell != self.isNull(self.ActiveCell):
				break
			count += 1
		return self.Offset(count-1 if not self.isNull(self.ActiveCell) else count , 0)

	def clcUp(self):	# This function is like xlUp in VBA
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


	# Funtions for workbook class development
	def getActiveSelection(self):	
		return self.active.getSelection()

	def getActiveSheet(self):
		return self.active.getActiveSheet()

	def isNull(self, oCell):
		return True if len(oCell.getString())==0 else False

	def isRangeNameGiven(self, isRowIndex=None , isColumnIndex=None , isRangeName=None):
		if isRangeName!=None and (isRowIndex==None and isColumnIndex==None):
			return True
		elif isRangeName==None and (isRowIndex!=None and isColumnIndex!=None):
			return False

Calc = WorkBook()

def Automate():
	Calc.Select(Calc.Range(fromCell=Calc.Cell(1 , 1) , toCell=Calc.Cell(10 , 5)))	# selects range form cell A1 to E10
