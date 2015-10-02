Attribute VB_Name = "FolderFunctionForExcel"
Option Explicit

'Originally from
'Demonstration of the Windows Shell Browse for Folder function for Excel 97 and Excel 2000.
'(Revision 1)
'By Jim Rech (jarech@kpmg.com)

'NOTE: The brilliant AddrOf function herein contained is the work of Ken Getz and
'Michael Kaplan.  Published in the May 1998 issue of
'Microsoft Office & Visual Basic for Applications Developer (page 46).

'Office 97 does not support the "AddressOf" operator which is needed to tell Windows
'where our "call back" function is.  Getz and Kaplan figured out a workaround.

'The rest of this module is entirely their work.






'ulFlag CONTANTS
'***************

'Only return file system directories.
'If the user selects folders that are not part of the file system, the OK button is grayed.
Public Const BIF_RETURNONLYFSDIRS = &H1

'Do not include network folders below the domain level in the tree view control.
Public Const BIF_DONTGOBELOWDOMAIN = &H2

'Include a status area in the dialog box.
'The callback function can set the status text by sending messages to the dialog box.
Public Const BIF_STATUSTEXT = &H4

'Only return file system ancestors.
'An ancestor is a subfolder that is beneath the root folder in the namespace hierarchy.
'If the user selects an ancestor of the root folder that is not part of the file system
  
  'the OK button is grayed.
Public Const BIF_RETURNFSANCESTORS = &H8

'Version 4.71. The browse dialog includes an edit control
'in which the user can type the name of an item.
Public Const BIF_EDITBOX = &H10

'Version 4.71. If the user types an invalid name into the edit box,
'the browse dialog will call the application's BrowseCallbackProc
'with the BFFM_VALIDATEFAILED message.
'This flag is ignored if BIF_EDITBOX is not specified
Public Const BIF_VALIDATE = &H20

'Version 5.0. New dialog style with context menu that can be resized with dag & drop.
'Does not show files. Has a new folder button.
Public Const BIF_NEWDIALOGSTYLE = &H40

'Version 5.0. Uses The New User Interface, Including Edit Box.
'This Flag Is Equivalent To BIF_EDITBOX | BIF_NEWDIALOGSTYLE
Public Const BIF_USENEWUI = &H50

'Version 5.0. Allow URLs to be displayed or entered. Requires BIF_USENEWUI &
  'BIF_BROWSEINCLUDEFILES.
Public Const BIF_BROWSEINCLUDEURLS = &H80

'Version 6.0. When combined with BIF_NEWDIALOGSTYLE adds a usage hint to the dialog box,
  'in place of the edit box.
'BIF_EDITBOX overrides this flag.
Public Const BIF_UAHINT = &H100

'Version 6.0. Do not include the new folder button in the browser dialog box.
'BIF_EDITBOX overrides this flag.
Public Const BIF_NONEWFOLDERBUTTON = &H200

'Version 6.0. When the selected item is a shortcut,
  'return the PIDL of the shortcut rather than its target.
'BIF_EDITBOX overrides this flag.
Public Const BIF_NOTRANSLATETARGETS = &H400

'Only return computers.
'If the user selects anything other than a computer, the OK button is grayed.
Public Const BIF_BROWSEFORCOMPUTER = &H1000

'Only return printers.
'If the user selects anything other than a printer, the OK button is grayed.
Public Const BIF_BROWSEFORPRINTER = &H2000

'The browse dialog will display files as well as folders. Dialog box not expandable.
Public Const BIF_BROWSEINCLUDEFILES = &H4000

'Version 5.0. Allow display of remote shareable resources.  Requires BIF_USENEWUI.
Public Const BIF_SHAREABLE = &H8000

'Win 7 & later. Allow folder junctions such as library or a compressed file
  'with a .zip extension to be browsed.
Public Const BIF_BROWSEFILEJUNCTIONS = &H10000


'-------------------------------------------------------------------------------------------------------------------
'   Declarations
'
'   These function names were puzzled out by using DUMPBIN /exports
'   with VBA332.DLL and then puzzling out parameter names and types
'   through a lot of trial and error and over 100 IPFs in MSACCESS.EXE
'   and VBA332.DLL.
'
'   These parameters may not be named properly but seem to be correct in
'   light of the function names and what each parameter does.
'
'   EbGetExecutingProj: Gives you a handle to the current VBA project
'   TipGetFunctionId: Gives you a function ID given a function name
'   TipGetLpfnOfFunctionId: Gives you a pointer a function given its function ID
'
'-------------------------------------------------------------------------------------------------------------------
Private Declare PtrSafe Function GetCurrentVbaProject _
 Lib "vba332.dll" Alias "EbGetExecutingProj" _
 (hProject As LongPtr) As LongPtr
Private Declare PtrSafe Function GetFuncID _
 Lib "vba332.dll" Alias "TipGetFunctionId" _
 (ByVal hProject As LongPtr, ByVal strFunctionName As String, _
 ByRef strFunctionId As String) As LongPtr
Private Declare PtrSafe Function GetAddr _
 Lib "vba332.dll" Alias "TipGetLpfnOfFunctionId" _
 (ByVal hProject As LongPtr, ByVal strFunctionId As String, _
 ByRef lpfn As LongPtr) As LongPtr



Public Type BROWSEINFO
    'Handle to the owner window for the dialog box.
    hOwner As LongPtr
    
    'Address of an ITEMIDLIST structure specifying the location
    'of the root folder from which to browse.
    'Only the specified folder and its subfolders appear in the dialog box.
    'This member can be NULL; in that case, the namespace root (the desktop folder) is used.
    pidlRoot As LongPtr
    
    'Address of a buffer to receive the display name of the folder selected by the user.
    'The size of this buffer is assumed to be MAX_PATH bytes.
    pszDisplayName As String
    
    'Address of a null-terminated string that is displayed above
    'the tree view control in the dialog box.
    'This string can be used to specify instructions to the user.
    lpszTitle As String
    
    'Flags specifying the options for the dialog box. See constants below.
    ulFlags As Long
    
    'Address of an application-defined function that the dialog box calls when an event occurs.
    'For more information, see the BrowseCallbackProc function.
    'This member can be NULL.
    lpfn As LongPtr
    
    'Application-defined value that the dialog box passes to the
    'callback function (in pData), if one is specified.
    lParam As LongPtr
    
    'Variable to receive the image associated with the selected folder.
    'The image is specified as an index to the system image list.
    iImage As Long
End Type

Public Const WM_USER = &H400
Public Const MAX_PATH = 260





'Message from browser to callback function constants

'Indicates the browse dialog box has finished initializing. The lParam parameter is NULL.
Public Const BFFM_INITIALIZED = 1

'Indicates the selection has changed.
'The lParam parameter contains the address of the item identifier list
'for the newly selected folder.
Public Const BFFM_SELCHANGED = 2

'Version 4.71. Indicates the user typed an invalid name into the edit box of the browse dialog.
'The lParam parameter is the address of a character buffer that contains the invalid name.
'An application can use this message to inform the user that the name entered was not valid.
'Return zero to allow the dialog to be dismissed or nonzero to keep the dialog displayed.
Public Const BFFM_VALIDATEFAILED = 3

' messages to browser from callback function
Public Const BFFM_SETSTATUSTEXTA = WM_USER + 100
Public Const BFFM_ENABLEOK = WM_USER + 101
Public Const BFFM_SETSELECTIONA = WM_USER + 102
Public Const BFFM_SETSELECTIONW = WM_USER + 103
Public Const BFFM_SETSTATUSTEXTW = WM_USER + 104

Public Const LMEM_FIXED = &H0
Public Const LMEM_ZEROINIT = &H40
Public Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)

'Main Browse for directory function
Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" _
 Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As LongPtr                                    'Checked PtrSafe
'Gets path from pidl
Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" _
  Alias "SHGetPathFromIDListA" (ByVal pidl As LongPtr, ByVal pszPath As String) As Long                'Checked PtrSafe althogh as long was boolean.
'Used by callback function to communicate with the browser
Declare PtrSafe Function SendMessage Lib "user32" _
 Alias "SendMessageA" (ByVal hWnd As LongPtr, _
  ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As Any) As LongPtr                         'Checked PtrSafe

Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, _
      hpvSource As String, ByVal cbCopy As LongPtr)

Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

Public Declare PtrSafe Function LocalAlloc Lib "kernel32" _
   (ByVal uFlags As LongPtr, _
    ByVal uBytes As Long) As LongPtr
    
Public Declare PtrSafe Function LocalFree Lib "kernel32" _
   (ByVal hMem As LongPtr) As LongPtr


''The following declarations for the option to center the dialog in the user's screen
Public Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hWnd As LongPtr, lpRect As rect) As LongPtr

Public Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const SM_CXFULLSCREEN = 16
Public Const SM_CYFULLSCREEN = 17

Public Type rect
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, _
 ByVal x As Long, _
   ByVal Y As Long, _
    ByVal nWidth As Long, _
     ByVal nHeight As Long, ByVal bRepaint As Long) As LongPtr
'End of dialog centering declarations

Dim CntrDialog As Boolean




'-------------------------------------------------------------------------------------------------------------------
'   AddrOf
'
'   Returns a function pointer of a VBA public function given its name. This function
'   gives similar functionality to VBA as VB5 has with the AddressOf param type.
'
'   NOTE: This function only seems to work if the proc you are trying to get a pointer
'       to is in the current project. This makes sense, since we are using a function
'       named EbGetExecutingProj.
'-------------------------------------------------------------------------------------------------------------------
Public Function AddrOf(strFuncName As String) As LongPtr
    Dim hProject As LongPtr
    Dim lngResult As LongPtr
    Dim strID As String
    Dim lpfn As LongPtr
    Dim strFuncNameUnicode As String
    
    Const NO_ERROR = 0
    
    ' The function name must be in Unicode, so convert it.
    strFuncNameUnicode = StrConv(strFuncName, vbUnicode)
    
End Function





Function GetDirectory(initDir As String, Flags As Long, CntrDlg As Boolean, Msg) As String
    Dim bInfo As BROWSEINFO
    Dim pidl As LongPtr, lpInitDir As LongPtr
    
    CntrDialog = CntrDlg  'Copy dialog centering setting to module level variable so callback function can see it
    With bInfo
        .pidlRoot = 0 'Root folder = Desktop
        .lpszTitle = Msg
        .ulFlags = Flags

        lpInitDir = LocalAlloc(LPTR, Len(initDir) + 1)
        CopyMemory ByVal lpInitDir, ByVal initDir, Len(initDir) + 1
        .lParam = lpInitDir
        
        'Modified this If Statement for use with none Microsoft Office products.
        'If this is Office 97 or earlier.
        If Left(Application.name, 9) = "Microsoft" And Val(Application.Version) < 8 Then
            .lpfn = AddrOf("BrowseCallBackFunc")
        Else
            .lpfn = CLng(BrowseCallBackFuncAddress)
        End If
    End With
    'Display the dialog
    pidl = SHBrowseForFolder(bInfo)
    'Get path string from pidl
    GetDirectory = GetPathFromID(pidl)
    CoTaskMemFree pidl
    LocalFree lpInitDir
End Function

'Windows calls this function when the dialog events occur
Function BrowseCallBackFunc(ByVal hWnd As LongPtr, ByVal Msg As LongPtr, ByVal lParam As LongPtr, ByVal pData As LongPtr) As LongPtr
    Select Case Msg
        Case BFFM_INITIALIZED
            'Dialog is being initialized. I use this to set the initial directory and to center the dialog if the requested
            SendMessage hWnd, BFFM_SETSELECTIONA, 1, pData 'Send message to dialog
            CenterDialog hWnd
        Case BFFM_SELCHANGED
            'User selected a folder - change status text ("show status text" option must be set to see this)
            SendMessage hWnd, BFFM_SETSTATUSTEXTA, 0, GetPathFromID(lParam)
            CenterDialog hWnd
        Case BFFM_VALIDATEFAILED
            'This message is sent  to the callback function only if "Allow direct entry" and
            '"Validate direct entry" have been be set on the Demo worksheet
            'and the user's direct entry is not valid.
            '"Show status text" must be set on to see error message we send back to the dialog
            Beep
            SendMessage hWnd, BFFM_SETSTATUSTEXTA, 0, "Bad Directory"
            BrowseCallBackFunc = 1 'Block dialog closing
            Exit Function
    End Select
    BrowseCallBackFunc = 0 'Allow dialog to close
End Function

'Converts a PIDL to a string
Function GetPathFromID(id As LongPtr) As String
    Dim Result As Boolean, Path As String * MAX_PATH
    Result = SHGetPathFromIDList(id, Path)
    If Result Then
        GetPathFromID = Left(Path, InStr(Path, Chr$(0)) - 1)
    Else
        GetPathFromID = ""
    End If
End Function

'XL8 is very unhappy about using Excel 9's AddressOf operator, but as longptr as it is in a
' function that is not called when run on XL8, it seems to allow it to exist.
Function BrowseCallBackFuncAddress() As LongPtr
    BrowseCallBackFuncAddress = LongPtr2LongPtr(AddressOf BrowseCallBackFunc)
End Function

'It is not possible to assign the result of AddressOf (which is a longptr) directly to a member
'of a user defined data type.  This explicitly "converts" it to a longptr and
'allows the assignment
Function LongPtr2LongPtr(x As LongPtr) As LongPtr
    LongPtr2LongPtr = x
End Function

'Centers dialog on desktop
Sub CenterDialog(hWnd As LongPtr)
    Dim WinRect As rect, ScrWidth As Integer, ScrHeight As Integer
    Dim DlgWidth As Integer, DlgHeight As Integer
    GetWindowRect hWnd, WinRect
    
    DlgWidth = WinRect.Right - WinRect.Left
    DlgHeight = WinRect.Bottom - WinRect.Top
    ScrWidth = GetSystemMetrics(SM_CXFULLSCREEN)
    ScrHeight = GetSystemMetrics(SM_CYFULLSCREEN)
    MoveWindow hWnd, (ScrWidth - DlgWidth) / 2, (ScrHeight - DlgHeight) / 2, DlgWidth, DlgHeight, 1
End Sub


Function FOLDER_SELECTION(strINITIAL_DIR As String, FOLDER_TYPE As Long, MESSAGE As String) As String
    'strINITIAL_DIR can be CurDir or any specific address such as "C:\"
    Dim Flags As Long, DoCenter As Boolean
    
    Select Case strINITIAL_DIR
        Case "CurDir"
            'Use the next line for Word, CorelDraw etc but not Excel
            strINITIAL_DIR = ThisDocument.FilePath & ThisDocument.name
            'Use the next line instead of the one above for Excel.
            'strINITIAL_DIR = ThisWorkbook.Path & ThisWorkbook.Name
        Case "MyDocuments"
            strINITIAL_DIR = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
        Case "Desktop"
            strINITIAL_DIR = CreateObject("WScript.Shell").SpecialFolders("Desktop")
        Case "AllUsersDesktop"
            strINITIAL_DIR = CreateObject("WScript.Shell").SpecialFolders("AllUsersDesktop")
        Case Else
            strINITIAL_DIR = strINITIAL_DIR
    End Select
    
    Select Case FOLDER_TYPE
        Case 1
            'Show only folders and have the ability to add new folders.
            Flags = BIF_NEWDIALOGSTYLE
        Case 2
            'Show only folders and not have the ability to add new folders.
            Flags = BIF_NEWDIALOGSTYLE + BIF_NONEWFOLDERBUTTON
        Case 3
            'Show folders and files and have the ability to add new folders.
            Flags = BIF_NEWDIALOGSTYLE + BIF_BROWSEINCLUDEFILES
        Case 4
            'Show folders and files and not have the ability to add new folders.
            Flags = BIF_NEWDIALOGSTYLE + BIF_NONEWFOLDERBUTTON + BIF_BROWSEINCLUDEFILES
        Case Else
            Flags = FOLDER_TYPE
    End Select
    
    'The meaning of BIF_EDITBOX etc is shown above in the ulFlag CONTANTS section.
    
    'The initial dialog folder is set to CurDir the current folder.
    'That's a good default but you can use a hard wired string like "C:\FILES" if you want.
    
    'FOLDER_SELECTION is the full name & path of the selected folder. Use it in your procedures.
    FOLDER_SELECTION = GetDirectory(strINITIAL_DIR, Flags, DoCenter, MESSAGE)
End Function

Sub TEST_DIALOG()
    'The following opens a dialog box showing a directory tree.
    'You can select either a single folder or a single file.
    'The initial folder can be preselected, the type of dialog box and a message added.
    'When the OK is pressed the full path to the folder or file is passed to strDESTINATION
    'Of course you do not have to use strDESTINATION. Substitute your own Variable.
    
    'FOLDER_SELECTION(strINITIAL_DIR As String, FOLDER_TYPE As Long, MESSAGE As String)
        'strINITIAL_DIR
            'CurDir or "MyDocuments", "Desktop", "AllUsersDesktop" or an address such as "C:\"
            'if the address does not exist then it will initialise on My Computer.
            'CurDir is a standard term for the current directory.
            'MyDocuments & Desktop are found using script this program.
              
        'FOLDER_TYPE
            '1 Show only folders and have the ability to add new folders.
            '2 Show only folders and not have the ability to add new folders.
            '3 Show folders and files and have the ability to add new folders.
            '4 Show folders and files and not have the ability to add new folders.
            'BIF_USENEWUI + BIF_NONEWFOLDERBUTTON + BIF_BROWSEINCLUDEFILES for example will give custom views.
            'BIF_STATUSTEXT + BIF_BROWSEINCLUDEFILES will open zip files and
              'the initial folder will be intially expanded but
              'the dialog box is of fixed size.
        'MESSAGE
            'The instruction text in the dialog box.
            
    'In every call to FOLDER_SELECTION you should if not must have the following line of code.
    Dim strDESTINATION As String
    
    'Only one additional line of code is required and is dependant on the application and the relative postion of the code calling FOLDER_SELECTION.
    'The samples below open at the user's desktop and show a dialog box showing folders & files. A new folder can be created.
    
    
    'CorelDraw
    '*********
    'Use this next line if your macro is saved within the same VBAProject as the Module Folders.
    strDESTINATION = FOLDER_SELECTION("CurDir", 4, "Make your selection.")
    
    'Use this next line if your macro is not saved within GlobalMacros such as in your CorelDraw file.
    'strDESTINATION = GMSManager.RunMacro("GlobalMacros", "Folders.FOLDER_SELECTION", "Desktop", 4, "Make your selection.")
    
    'Alternatively you can use the following but you but your VBAProject must make reference to GlobalMacros.
    'strDESTINATION = GlobalMacros.FOLDER_SELECTION("Desktop", 4, "Make your selection.")
    
    
    'Corel PhotoPaint
    '****************
    'Use this next line if your macro is saved within the same VBAProject as the Module Folders.
    'strDESTINATION = FOLDER_SELECTION("Desktop", 4, "Make your selection.")
    
    'Unlike CorelDraw PhotoPaint does not have a GMSManager.
    'If your code is not in GlobalMacros then the VBAProject where you have your code must reference GlobalMacros."
    'strDESTINATION = GlobalMacros.FOLDER_SELECTION("Desktop", 4, "Make your selection.")
    
    
    'EXCEL
    '*****
    'Excel does not have a GlobalMacros folder as does CorelDraw.
    'Instead you can create an Excel file that can contain your global macros and place it in the Excel Start-Up folder.
    'In Office 2000 & 2003 save as .xls and in Office 2007 & 2010 save as .xlsm with any name you like.
    'The Start-Up folder is located for
    
    'Win XP, Office 2003 & 2010 at C:\Documents and Settings\<User Name>\Application Data\Microsoft\Excel\XLSTART
    'Win 7, Office 2010 at C:\Users\<User Name>\AppData\Roaming\Microsoft\Excel\XLSTART

    'When you next open Excel your Excel file will open.
    'In view hide this window and close Excel so that it can contain macros.
    'Now whenever Excel opens your Excel file will open but will be hidden.
    'In Excel's VBA IDE window the hidden Excel file will display its VBAProject.
    'Use this Excel file just as you use the GlobalMacros VBAProject in Corel.
    'If you want a macro to be available to any Excel file you open plae the macrto in this VBAProject.
    
    'If you import this Folder.bas into the VBAProject of your hidden Excel file and you call it from another VBAProject use the code below.
    'Use either of the next two lines depending on your version of Excel, if your macro is not saved within the same file.
    'Here PERSONAL is the name of my hidden Excel file.
    'strDESTINATION = Run("PERSONAL.xls!Folders.FOLDER_SELECTION", "Desktop", 4, "Make your selection.")
    'strDESTINATION = Run("PERSONAL.xlsm!Folders.FOLDER_SELECTION", "Desktop", 4, "Make your selection.")
    
    'Use this next line if your macro is saved within the same VBAProject as the Module Folders.
    'strDESTINATION = FOLDER_SELECTION("Desktop", 4, "Make your selection.")
    
    
    'WORD
    '****
    'For Word import Folders into the Normal template.
    'No special reference to the location of the macro Folders is required.
    'Use the following.
    'strDESTINATION = FOLDER_SELECTION("Desktop", 4, "Make your selection.")
    
    'PUBLISHER
    '*********
    'Publisher does not have a GlobalMacros folder as does CorelDraw.
    'Either import Folder.bas into your Publisher file or save a Publisher template with a Folder module.
    'If you need Folders in a Publisher document open the Publisher template to automatically have the Folder module in your document.
    'The following will run in this Publisher document.
    'strDESTINATION = FOLDER_SELECTION("Desktop", 4, "Make your selection.")
    
End Sub


