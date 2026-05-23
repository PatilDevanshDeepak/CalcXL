# Libreoffice Calc Python Macros version 0.3 - Dev

## Features:
### 1. Now select function can select Cells/Range and worksheet  Eg: Calc.Select(Calc.Cell(1 , 1)) or Calc.Select(Calc.Range("A1:C10")) or Calc.Select(Calc.SheetName())
### 2. Now you can set value to a range Eg: Calc.Range("A1:A10").String = "Hello word"
### 3. New function RangeAddress returns the Address of a cell
### 4. New function SheetName returns worksheet as an object by name and you can select a worksheet  Eg. Calc.Select(Calc.SheetName("nameofsheet"))
### 5. New function SheetIndex returns worksheet as an object by index and you can select a worksheet  Eg. Calc.Select(Calc.SheetName(indexofsheet))
### 6. New function CreateSheet creates a new sheet by giving sheet name, and using set before attribute Eg Calc.CreateSheet("SHEETNAME" , 1) then the sheet will be created on index 2
### 7. New fuction RenameSheet renames a sheet  Eg. Calc.RenameSheet("oldname" , "newname")
### 8. New fucntion DeleteSheet deletes a sheet  Eg. Calc.DeleteSheet(addr) where addr can be your sheet index or sheet name
### 9. New function ClearContents clears the contents of a Range/Cell  Eg. Calc.ClearContents(obj, clearAll) where obj is your cell/range and clearAll is your flag which dicides what to clear String, int, formula or everything. Note: This function may have bugs

<img width="188" height="225" alt="image" src="https://github.com/user-attachments/assets/2ff26132-ac05-4801-83b5-ed1b17fb98b1" />




## Now in your WorkBook class, new constructors are:
### Sheets which have all the sheets stored as a objects in a variable
### ActiveSheet which have current sheet stored as a object in a variable
### SheetCount returns the total number of sheets in your workbook
### ActiveSheetName returns the name of the current sheet
### ActiveSheetIndex returns the index of the current sheet
