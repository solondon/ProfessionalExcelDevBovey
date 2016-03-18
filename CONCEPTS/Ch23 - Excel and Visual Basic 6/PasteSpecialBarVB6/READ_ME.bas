Attribute VB_Name = "READ_ME"
'This COM Addin was created by copying the code from the PasteSpecialBar example
'addin workbook, described in Chapter 8 - Advanced Command Bar Handling, then
'modifying it to work as a COM Addin.  The following changes were required:
'
' - Add a project reference to the Microsoft Excel object library
'
' - Instead of the Auto_Open and Auto_Close routines in the MOpenClose module,
'   we use the OnConnection and OnDisconnection events in the Addin Designer
'   Class, dsrConnect.  These two events are the COM Addin equivalents of
'   Auto_Open and Auto_Close, or Workbook_Open and Workbook_Close.
'
' - COM Addins don't automatically have access to Excel's global objects, such as
'   the Application and Commandbars objects.  Instead, the Application object is
'   given to use within the OnConnection event, which we store in a new global
'   variable gxlApp.  We then use 'gxlApp' instead of 'Application' throughout the
'   COM Addin code and ensure all our references to global objects are prefixed by gxlApp:
'       Application.         ->      gxlApp.
'       CommandBars.         ->      gxlApp.CommandBars.
'       Selection.           ->      gxlApp.Selection.
'
' - COM Addins don't have a spreadsheet handy to store the menu definition table used
'   by the command bar builder, so we have two alternatives:
'    1. Modify the command bar builder, so we can pass all the columns of information
'       as parameters to a new procedure, or
'    2. Modify the command bar builder, to retrieve the definition table from a database, or
'
'    3. Don't use the command bar builder, creating the menu items individually
'   In this example, we're creating a toolbar of seven essentially identical buttons, so
'   it will easier for us to create the menu items individually (option 3).  Hence, the
'   MCommandBars module has been simplified to add our menu items directly and the
'   MPastePicture module is no longer required.
'   The toolbar images that were also stored in the menu definition worksheet have been
'   copied to a resource file in the VB project
'
' - The MErrorHandler module contained a few references to ThisWorkbook, for the file name
'   and path, which were replaced by App (which is the VB6 equivalent of ThisWorkbook).
'
' - Added the MCopyTransparent module, copied from KB article 288771, to allow copying of
'   custom faces for pasting to Excel 2000's toolbars.
