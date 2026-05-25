## How to use the class file

Note this is only for Libreoffice Calc (its an Microsoftoffice excel alternative for linux/macos/windows users)
And works on Libreoffice & Openoffice

INSTALLATION:
In Libreoffice or Openoffice you just need this extension:
https://extensions.libreoffice.org/en/extensions/show/apso-alternative-script-organizer-for-python
<img width="1599" height="817" alt="image" src="https://github.com/user-attachments/assets/c35be396-d9fc-428b-baff-7a6347407efb" />

Download it and go to Libreoffice/Openoffice, open tools menu and then open Extensions option.
<img width="1600" height="859" alt="image" src="https://github.com/user-attachments/assets/a0895ad1-c0b7-4fd1-8a10-f4410f2ba863" />

It will have 'Add' button, click it and find the path where you installed the extension.
And then open the extension (APSO.oxt file), now the Libreoffice/Openoffice installs the extension.
This will restart the LibreCalc App.
<img width="1600" height="857" alt="image" src="https://github.com/user-attachments/assets/98d53866-7ab8-46da-a441-2a600d96e5c3" />

USE:
In the menu bar go to, Tools > Macros > Organise Python Macros (This option will appear only after you install the extension)
Now right click 'My Macros' and select 'Create a module' option, and then name the module
Now inside 'My Macros' you will have your module, now right click the module and click the edit option
Paste the MacroClass file's code inside that module, and at the bottom add:
  Calc = Workbook
  def automate():
    Calc.ActiveCell.String = "Hello world"
    Calc.Range("C1:D10").String = "Hello World"
Save it and again go to Tools > Macros > Organise Python Macros.
Go to 'My Macros' > Automate and run it.
<img width="1599" height="860" alt="image" src="https://github.com/user-attachments/assets/5a10be91-ef89-4bcd-aa5b-9d8cda2c65ef" />
<img width="1600" height="860" alt="image" src="https://github.com/user-attachments/assets/88f5f8f7-f5ad-47ba-a0c5-9ee50cd95e36" />
<img width="1600" height="859" alt="image" src="https://github.com/user-attachments/assets/c290c524-6058-4a5d-bfad-cacbe69158ff" />


EXPLAINATION:
What's happening is you created a Module in 'My Macros', and inside that you paste the Macros Class code
Now you have VBA function, but you have to write a Macro
What you will do is create user defined function: def automate() or with any name you like
Inside that funtion you will wirte your macro:
  def automate():
    Workbook.ActiveCell.String = "Hello world"    # Will insert hello world in Active Cell
    Workbook.Range("A1").String = "Hello world"    # will insert hello world in Range("A1") or Cell A1
    Workbook.Range("A1:B10").String = "Hello world"  # this will add hello world to Range A1:B10
So you can use multiple functions from Workbook class

Also you can create multiple user defined function like def automate(), def fetchData(), ...
Which will create options in 'My Macros' > your created module > automate, fetchData, ...
Now if you click the def automate() code gets execute!
