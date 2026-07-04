# Libreoffice Calc Python Macros version 0.4

📖 **[How to Use Guide](./HOWTOUSE.md)** | Quick Start

## Anouncements
- An alternative function to VBA's End function has been added to v0.4 which simulates the action of Ctrl + Right/Left/Up/Down arrowkeys
- Dev branch is been deleted and all the new function are added to this branch
- Clear Contents function contains bugs and will be added soon after fixing
- Developemn in progress for new features like: Fill Right/Fill Left/Fill Up/Fill Down


## Summary: The development branch mostly focuses on stability & bug fixes.

New:
* End function added, which is just like Excel's ActiveCell.End(xlToRight/xlToLeft/xlUp/xlDown) where you can simulate the action of ctrl + right/left/up/down
  <img width="1919" height="1042" alt="image" src="https://github.com/user-attachments/assets/dc2906cb-dcd3-49e6-ba24-dc36a9a12dfd" />

  Name Function added, which is just like Excel's Range("celladdr").Name.Name which returns the name of a cell

* Contructors Added:
  - smgr: Libreoffice service manager API
  - MaxRows: Gets total rows from the current worksheet,
  - MaxColumns: Gets total column from the current worksheet
  - NameRanges: Gets all the defined name for the cells
  - NameRangesCount: Gets count of all the Names in the NameRanges
  <img width="1412" height="308" alt="image" src="https://github.com/user-attachments/assets/2fa59eb4-b067-4588-b876-ed542c49219f" />

## Bug Fixed:
- Fixed the Select Function
- Fixed ActiveCell Constructor
- The clcToRight, clcToLeft, clcUp, clcDown functions are deprecated from the project and have a new alternative 'End' function for that 
