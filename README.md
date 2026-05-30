# Libreoffice Calc Python Macros version 0.4 - Dev

📖 **[How to Use Guide](./HOWTOUSE.md)** | Quick Start

## Summary: The development branch mostly focuses on testing and adding new features. This branch migh contain bugs as new features are added and tested in this branch.

New:
* End function added, which is just like Excel's ActiveCell.End(xlToRight/xlToLeft/xlUp/xlDown) where you can simulate the action of ctrl + right/left/up/down
* Contructors Added:
  - MaxRows: Gets total rows from the current worksheet,
  - MaxColumns: Gets total column from the current worksheet
 

## Bug Fixed:
- Fixed the Select Function
- Fixed ActiveCell Constructor
- Removed clcToRight, clcToLeft, clcUp, clcDown functions from stable branch and added to this branch
- Remove ClearContents function from the stable branch and added to this branch
