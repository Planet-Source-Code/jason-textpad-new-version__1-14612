Attribute VB_Name = "Modmain"
' **************************************************************************
' You have A roal Free right too use ANY Or ALL of this source code in your Programs
' It would be nice if i get credit ; but anyway , Read ALL the code before you go around
' Copying stuff Seperate , You might miss some IMPORTANT
' functions and what not ;  Enjoy !!!
'****************************************************************************


'***************************************************************************
' TextPad Version 4.120 Beta 23 revision 2
' Last Update January ,21 ,2001 9:36 PM
'***************************************************************************
' Fixes :
'
' Mainly Code ( Well Most Of it ) Has Been Restructured ,.
' Also Numerous Bugz Were fixed For example When Launching The External
' editor TextPad would Use the long File Name Format Instead of the short Dos Format,
' Which Would Cause the External Editor, Not Being able too open the file
'***************************************************************************

Public strfind As String
Option Explicit
Type fileopened
dirty  As Integer
End Type

Option Compare Text

Public Const Normal_Cdlogflags = cdlOFNHideReadOnly + cdlOFNFileMustExist + cdlOFNLongNames

'**Public Currentfilename As String
Public ExternalEditorPath As String
Type filestring ' The heart of textpads BOOLEAN Memory
dirty As Integer ' Without This Then TextPad wouldnt
End Type ' Know if a file was changed or not
Public fstate As filestring

Public Const URL = "http://www.vb-world.net"
Public Const email = "Cyberarea@hotmail.com"

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
' constant(s) for Shell execute
Public Const SW_SHOWNORMAL = 1

' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1 'Unicode nul terminated string
Public Const REG_BINARY = 3 'Free form binary
Public Const REG_DWORD = 4 '32-bit number
Public Const ERROR_SUCCESS = 0&

' Constants for shell execute info
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400

Private Type SHELLEXECUTEINFO ' Type Of SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    Hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" (sei As SHELLEXECUTEINFO) As Long

Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal Hkey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" _
Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey _
As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal _
lpValueName As String) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" _
Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey _
As String, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName _
As String, ByVal lpReserved As Long, lpType As Long, lpData _
As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal _
lpValueName As String, ByVal Reserved As Long, ByVal _
dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Sub gotoweb()
Dim success As Long

success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)

End Sub

Public Sub sendemail()
Dim success As Long
success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub


Sub findnexttext()
    If strfind <> "" Then
    Dim Search, Where   ' Declare variables.
    ' Get search string from user.
    Search = strfind
    Where = InStr(Form1.ActiveControl.Text, Search)   ' Find string in text.
    If Where Then   ' If found,
        Form1.ActiveControl.SelStart = Where - 1  ' set selection start and
       Form1.ActiveControl.SelLength = Len(Search)   ' set selection length.
    Else
        MsgBox "Cannot find  " & Chr(34) & Search & Chr(34) _
        , vbInformation, "TextPad" ' Notify user.
    End If
Else
If strfind = "" Then
Load frmfind
frmfind.Show (0), Form1

End If
End If

End Sub


  Sub findit()
       
    strfind = frmfind.txtfind.Text
    Dim Search, Where     ' Declare variables.
    ' Get search string from user.
    Search = frmfind.txtfind.Text
    Where = InStr(Form1.ActiveControl, Search) ' Find string in text.
    
    If Where Then   ' If found,
        Form1.ActiveControl.SelStart = Where - 1  ' set selection start and
      Form1.ActiveControl.SelLength = Len(Search)   ' set selection length.
    Form1.SetFocus
    
    strfind = frmfind.txtfind.Text
    Else
        MsgBox "Cannot find " & Chr(34) & Search & Chr(34) _
        , vbInformation, "Text Pad" ' Notify user.
    End If
  End Sub


'*******************************************************
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' Old Open File Command Line COde ( No longer In use )
'*******************************************************
'Sub OpenFilecommandline()
'Dim Text_control1, Text_Control2
'Set Text_control1 = Form1.Text1
'Set Text_Control2 = Form1.Text2

'Open Command$ For Binary Access Read As #1
'        If FileLen(Command$) > 65000 Then GoTo toobig:
'On Error GoTo toobig:
'Text_control1.Text = Input(LOF(1), 1)
    
'   Text_Control2.Text = Text_control1.Text
'    Text_control1.Text = Text_Control2.Text
'    Form1.caption = Command$ & " - TextPad"
'    fstate.dirty = False
'    Form1.lblfilename.caption = Command$
'Close #1
'toobig:
'If Err.Number <> 0 Then

''Form1.lblfilename.caption = ""
'Close #1
'Reset
'Close #1
'Exit Sub
'End If
'End Sub
'**********************************************************
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'***********************************************************
Sub Resizenotewithtoolbar()
On Error GoTo Errresize:
If Form1.WindowState = vbMinimized Then Exit Sub
    If Form1.Toolbar.visible And Form1.txt(2).visible = True Then
        Form1.txt(2).Height = Form1.ScaleHeight - Form1.Toolbar.Height
        Form1.txt(2).Width = Form1.ScaleWidth
        Form1.txt(2).Top = Form1.Toolbar.Height
    Else
        If Form1.Toolbar.visible And Form1.txt(1).visible = True Then
        Form1.txt(1).Height = Form1.ScaleHeight - Form1.Toolbar.Height
        Form1.txt(1).Width = Form1.ScaleWidth
        Form1.txt(1).Top = Form1.Toolbar.Height

    Else
        If Form1.txt(2).visible = True Then
        Form1.txt(2).Height = Form1.ScaleHeight
        Form1.txt(2).Width = Form1.ScaleWidth
        Form1.txt(2).Top = 0
        Else
        If Form1.txt(1).visible = True Then
        Form1.txt(1).Height = Form1.ScaleHeight
        Form1.txt(1).Width = Form1.ScaleWidth
        Form1.txt(1).Top = 0
       
       End If
      End If
    End If

Errresize:
If Err.Number <> 0 Then
Exit Sub
End If
End If
End Sub
Sub check_if_textpad_is_associated()

Dim msg, response
' if getsetting = NOT associated then we simply exit the Sub
If GetSetting("TextPad", "associations", "isassociated", "") = 1 _
And GetSetting("TextPad", "Chckassociations", "show", "") = 0 Then
Exit Sub
' else if the getsetting function returns false (0) and  true (1) then
ElseIf GetSetting("TextPad", "associations", "isassociated", "") = 0 And GetSetting("TextPad", "Chckassociations", "show", "") = 1 Then
msg = "Textpad is not currently associated with Text Files Would you like it too be ?" _
& Chr(13) & Chr(10) & Chr(13) & Chr(10) & "For this Message Not too show the next time you Run text pad," _
& Chr(13) & Chr(10) & "Check off in the options window :" _
& Chr(13) & Chr(10) & "(Textpad should check wether it is the default text viewer) "

response = MsgBox(msg, vbYesNo + vbInformation + vbDefaultButton1, "Text Pad")

Select Case response
Case vbYes
SaveSettingString HKEY_CLASSES_ROOT, _
"Txtfile\shell\open\command", _
"", App.Path & "\" & App.EXEName & ".EXE" & " %1"

Case vbNo
Exit Sub
Resume Next
End Select
End If

End Sub
Sub Readonlyerror()
Close #1
If Err.Number <> 0 Then
MsgBox "Error, The file you are trying too save too Exists with read only attributes please select a different filename.", vbExclamation, "Error,TextPad"
Form1.CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI Configuration Files (*.INI) |*.INI| All Files (*.*) |*.* "
Form1.CommonDialog1.Flags = cdlOFNHideReadOnly
On Error GoTo dialogerror:
Form1.CommonDialog1.ShowSave
If Form1.CommonDialog1.Filename <> "" Then
Open Form1.CommonDialog1.Filename For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
Form1.lblfilename.caption = Form1.CommonDialog1.Filename
fstate.dirty = False
dialogerror:
If Err.Number <> 0 Then
Exit Sub
End If
End If
End If

End Sub
Sub Filequicksave()
Close #1

Dim strFileName As String

On Error GoTo Readonlyerr:
If Form1.lblfilename.caption <> "" Then
strFileName = Form1.lblfilename.caption
Open strFileName For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
Else
Form1.CommonDialog1.Filter = "Text documents (*.TXT) |*.TXT| INI Configuration Files (*.INI) |*.INI| Log Files (*.LOG) |*.LOG| All Files (*.*) |*.* "
Form1.CommonDialog1.DialogTitle = "Save As"
Form1.CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt
On Error GoTo dialogerror
Form1.CommonDialog1.ShowSave
If Form1.CommonDialog1.Filename <> "" Then
Open Form1.CommonDialog1.Filename For Output As #1
Print #1, Form1.ActiveControl.Text
Close #1
Form1.lblfilename.caption = Form1.CommonDialog1.Filename
Form1.caption = Form1.lblfilename.caption & " - TextPad"

Readonlyerr:
If Err.Number <> 0 Then
Readonlyerror
End If
End If
End If

dialogerror:
If Err.Number <> 0 Then
Exit Sub
End If
End Sub

Sub newfile()
Dim msg, response
If Form1.lblfilename.caption <> "" Then '2
 msg = "The Text in " _
 & Form1.lblfilename.caption & " File has changed" _
& Chr(13) & Chr(13) & "Do you wish too save The changes ?"
Else ' * If there isnt then Display The One Below
msg = "The Text in the untitled file has changed" _
& Chr(13) & Chr(13) & "Do you wish too save the changes?"
End If

Beep
response = MsgBox(msg, vbYesNoCancel + vbQuestion + vbDefaultButton2, "New File ")

Select Case response
Case vbYes     ' User chose Yes.
Filequicksave 'call quicksave procedure in modquicksave
Form1.caption = "Untitled - TextPad"
Form1.ActiveControl.Text = ""
fstate.dirty = False
Form1.lblfilename.caption = ""



Case vbNo ' user chose No.
Form1.caption = "Untitled - TextPad"
fstate.dirty = False
Form1.ActiveControl.Text = ""
Form1.lblfilename.caption = ""
Form1.CommonDialog1.Filename = ("")
End Select
End Sub

' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

Public Function GetSettingString(Hkey As Long, _
strPath As String, strValue As String, Optional _
Default As String) As String
Dim hCurKey As Long
Dim lResult As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

'Set up default value
If Not IsEmpty(Default) Then
GetSettingString = Default
Else
GetSettingString = ""
End If

lRegResult = RegOpenKey(Hkey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, _
lValueType, ByVal 0&, lDataBufferSize)

If lRegResult = ERROR_SUCCESS Then

If lValueType = REG_SZ Then

strBuffer = String(lDataBufferSize, " ")
lResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, _
ByVal strBuffer, lDataBufferSize)

intZeroPos = InStr(strBuffer, Chr$(0))
If intZeroPos > 0 Then
GetSettingString = Left$(strBuffer, intZeroPos - 1)
Else
GetSettingString = strBuffer
End If

End If
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

Else
MsgBox "An error has occured while calling the api function , settings cannot be saved .", vbCritical, "Error,api"

End If

lRegResult = RegCloseKey(hCurKey)
End Function
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

Public Sub SaveSettingString(Hkey As Long, strPath _
As String, strValue As String, strdata As String)
Dim hCurKey As Long
Dim lRegResult As Long

lRegResult = RegCreateKey(Hkey, strPath, hCurKey)

lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, _
ByVal strdata, Len(strdata))

If lRegResult <> ERROR_SUCCESS Then
'there is a problem
End If
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!

lRegResult = RegCloseKey(hCurKey)
End Sub ' Thank you again VBWORLD.com !!!!!!!
' Credit is given where credit is due !!!!
' Registry source code from Vbworld.com !! Thank you !!!




' \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' commented out for now
'Sub Check_For_Registry_Entrys()

'If GetSetting("Textpad", "Toolbar", "Visible") = "" _
'And GetSetting("Textpad", "chckassociations", "show") = "" _
'And GetSetting("Textpad", "Font", "Font") = "" _
'And GetSetting("Textpad", "Font", "Fontsize") = "" _
'And GetSetting("TextPad", "Wordwrap", "Wordwrap") = "" _
'And GetSetting("TextPad", "UseExternalEditor", "Path") = "" Then
'Load Frmnoreg
'Frmnoreg.Show
 'Form1.Hide
'End If
'End Sub '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
' above commented out for now

Sub Main() ' main startup  Procedures and Misc. for The Project
   Dim onoffwordwrap As Boolean
    On Error GoTo Inputpastendoffile:
    
' check if theres any registry entrys if not then load the Setup
' Dialog (Frmnoreg)
   If GetSetting("Textpad", "Toolbar", "Visible") = "" _
    And GetSetting("Textpad", "chckassociations", "show") = "" _
    And GetSetting("Textpad", "Font", "Font") = "" _
    And GetSetting("Textpad", "Font", "Fontsize") = "" _
    And GetSetting("TextPad", "Wordwrap", "Wordwrap") = "" _
    And GetSetting("TextPad", "UseExternalEditor", "Path") = "" _
    And GetSetting("TextPad", "associations", "isassociated") = "" Then
  Load Frmnoreg
   Frmnoreg.Show
    Form1.Hide
     Exit Sub
   End If


      RetrieveALLSettings 'get all the programs settings that were
      'saved in the registry
      
      ' togglewordwrap from usewordwraps value form modoptions BOOLEAN
       Togglewordwrap (Usewordwrap)

    If Command$ <> "" Then '  Check If Command Line Is being Used
     CommandLine_OpenFile ' If it is then Goto The CommandLine_OpenFile Procedure In modmain (This Module)
  End If


     'Declare Public Variables Boolean
     fstate.dirty = False

    check_if_textpad_is_associated


   'Recent files menu
   If GetSetting("TextPad", "Recentfiles", "1") <= "" Then
     Form1.mnurecentfile(1).visible = False
   Else
      Form1.mnurecentfile(1).caption = GetSetting("TextPad", "RecentFiles", "1")
       Form1.mnurecentfile(1).visible = True
        Form1.mnurecentfile(0).visible = True
    End If
 
 Form1.Show
'////////////////////////////////////////////////////////
'| Error Control                                        |
'////////////////////////////////////////////////////////
Inputpastendoffile:
     If Err.Number <> 0 Then

Dim response

  response = MsgBox("TextPad has encountered the following error(s) while loading :" _
  & vbCrLf & vbCrLf & "Error :  " & Err.Description & _
  Chr(13) & Chr(10) & "source :  " & Err.Source & vbCrLf & vbCrLf & "Would you like too Continue loading Text pad anyway ?" _
  , vbYesNo + vbExclamation + vbDefaultButton2, "TextPad")
 
   Select Case response
     Case vbYes
      Close #1
      Form1.Show
     Case vbNo
      End
   End Select

    End If
End Sub
Sub Openfile(sFileName As String)
Dim Strfilerecent, regval
On Error GoTo Filenotfound:
    If sFileName <> "" Then
     Open sFileName For Binary Access Read As #1
      If FileLen(sFileName) > 65000 Then GoTo outofmemory:
        On Error GoTo outofmemory:
        Form1.Show
         Form1.ActiveControl.Text = Input(LOF(1), 1)
          fstate.dirty = False
           Form1.caption = UCase$(sFileName) & "  - TextPad"
         Strfilerecent = sFileName
       Form1.mnurecentfile(1).caption = Strfilerecent
    SaveRegistryString "Recentfiles", "1", Strfilerecent
  Form1.mnurecentfile(1).visible = True
    Form1.mnurecentfile(0).visible = True
      Form1.mnurecentfile(1).caption = sFileName
       Form1.lblfilename.caption = sFileName
        fstate.dirty = False
         Close #1
          Exit Sub
           End If

     
outofmemory:      ' error That occurs when textpad runs out of memory
  Form1.Hide
   Query_TooBig (GetShortPath(sFileName))
Filenotfound:
        If Err.Number = 75 Then
         MsgBox "The File You Attempted too open Could Not Be found." & vbNewLine & " Please Check The Name And Path And Try Again.", vbCritical, "Error,TextPad"
        Exit Sub
       End If
        
 End Sub
Sub CommandLine_OpenFile()
     Dim thecontrol As Boolean
       thecontrol = Usewordwrap
   On Error GoTo Filetoobig:
       If FileLen(Command$) >= 65000 Then: Form1.Hide: Query_TooBig (GetShortPath(Command$))

      Open Command$ For Binary Access Read As #1
        Select Case thecontrol
         Case True ' The Control Accepting the File Will be......
           Form1.txt(1).Text = Input(LOF(1), 1)
         Case False ' The Control Accepting the File Will be......
           Form1.txt(2).Text = Input(LOF(1), 1)
         End Select
        Form1.caption = Command$ + " - TextPad" 'set the forms caption
       Form1.lblfilename.caption = Command$ ' set the labels Caption
      Form1.Show ' Finally Show the form
   Close #1 ' close the file
   
Filetoobig:
    If Err.Number <> 0 Then
       Form1.Hide
       Query_TooBig (GetShortPath(Command$))
    End If
End Sub

Function openrecentfile()
Close #1
Dim regfilename As String
Dim Strfilerecent As String
'***********************************
'Declare RegFilename as A string
'Data Type : String
'We first Check If the file exists If
'The vba Command : Filedatetime([Expression]) doesnt
'find it Then an error Will
'occur and the error Control will
'Pick it up and Prompt the user about it.
'************************************
On Error GoTo Filenotfound: ' error control
 regfilename = GetSetting("TextPad", "Recentfiles", "1", Strfilerecent)
FileDateTime (regfilename) ' check if file exists first
 Open GetSetting("TextPad", "Recentfiles", "1", "") For Binary Access Read As #1
   On Error GoTo outofmemory:
   Form1.ActiveControl.Text = Input(LOF(1), 1)
Close #1
Form1.caption = regfilename & " - TextPad"
Form1.lblfilename.caption = regfilename
fstate.dirty = False

Filenotfound: ' error control
If Err.Number <> 0 Then

MsgBox "File not found" _
& vbNewLine & "The File may have been moved renamed or Deleted", vbCritical, "TextPad"

SaveRegistryString "Recentfiles", "1", ""
Form1.mnurecentfile(1).visible = False
Form1.mnurecentfile(1).caption = ""
Form1.mnurecentfile(0).visible = False
Form1.caption = "Untitled - TextPad"
'Form1.line8.visible = False

Exit Function
End If

outofmemory: ' error That occurs when File overloads textpads limit
If Err.Number = 7 Then
  Form1.caption = "Untitled - TextPad"
 Form1.ActiveControl.Text = ("")
MsgBox regfilename & _
 vbNewLine & "Is too large Too be opened Here ", vbCritical, "Error"
SaveRegistryString "Recentfiles", "1", ""
Form1.lblfilename.caption = ("")
Close #1
Form1.CommonDialog1.Filename = ("")
 fstate.dirty = False
Exit Function
End If
End Function
Function ShellNewTextPad(Thewindowstyle As VbAppWinStyle) As Long

On Error GoTo Filenotfound:

Dim strapp As String
  strapp = App.Path & "\" & App.EXEName
  Shell strapp, Thewindowstyle

Filenotfound:
 If Err.Number <> 0 Then
 MsgBox "TextPad Cannot find its Own executable File.", vbCritical, "Error, New instance"
 Exit Function
End If

End Function
Sub TextChangecontrol()
On Error GoTo outofmemoryerror:
If fstate.dirty = False Then
fstate.dirty = True
End If
outofmemoryerror:
If Err.Number <> 0 Then
MsgBox "TextPad Has Encountered The Following Error(s) While Performing The operation you requested : " _
& Chr(13) & Chr(13) & Err.Description & Chr(13) & Chr(13) & "If The Error Is " & Chr(34) & "Out of memory" & Chr(34) & " Then TextPad Cannot Place Anymore Text Into The Text Box Because It Has Run Out of memory", vbCritical, "TextPad "
Exit Sub
End If
End Sub
Sub SaveRegistryString(Thesection As String, thekey As Variant _
, thesetting As Variant)
SaveSetting "TextPad", Thesection, thekey, thesetting
End Sub
Function ExecuteExternalEditor(Currentfilename As String) As String
Dim thepathname As String
Dim success As Long
  thepathname = GetSetting("TextPad", "UseExternalEditor", "path", "")
success = Shell(thepathname & Space(1) & Currentfilename, vbNormalFocus)
End Function
Sub DetectExternalEditor()
' **************************************************
' Boy this Only Took Me About a Minute or so Too
' Write But It Actually WORKS !!!!
' What It Basically does is Use the Two Textboxes
' On the Dumby Form then Sets It Towards them in code
' then it Finds The String We Dont Want And Just
' Strips It too a Null Zero Length String ""
' And It actually works !!! Kudos For Me !!!!
' **************************************************
Dim retval As String
Dim ThisTextBox1
Dim ThisTextBox2
Set ThisTextBox1 = FrmDumbyform.TxtConvert1
Set ThisTextBox2 = FrmDumbyform.TxtConvert2
Dim SearchforWhat, Where
Dim i As Integer
retval = GetSettingString(HKEY_CLASSES_ROOT, _
"rtffile\shell\open\command", _
"", "")
ThisTextBox1.Text = Chr(34) & "%1" & Chr(34)
ThisTextBox2.Text = retval
For i = 1 To 3
SearchforWhat = ThisTextBox1.Text
Where = InStr(ThisTextBox2.Text, SearchforWhat)
If Where Then
     ThisTextBox2.SelStart = Where - 1  ' set selection start and
     ThisTextBox2.SelLength = Len(SearchforWhat)   ' set selection length.
     ThisTextBox2.SelText = ""
     retval = ThisTextBox2.Text
ExternalEditorPath = retval
SaveRegistryString "UseExternalEditor", "path", retval
End If
Next i
End Sub
Sub Query_TooBig(sFileName As String)
    Dim regval As String
     Dim msg, response

  'set regval variables
    regval = GetSetting("TextPad", "UseExternaleditor", "Path", "")

      If UseExternalEditor.use = True And regval = "" Then
      Form1.Show
       No_Externaleditor_Detected
        Form1.lblfilename.caption = ("")
         Close #1
          Reset
           fstate.dirty = False
          Exit Sub
      End If

      Select Case UseExternalEditor.use
        
        Case True ' Case is True *******
         GoTo Query:
             Exit Sub
   
         Case False ' Case is False ******
          Form1.Show
          MsgBox sFileName _
           & vbNewLine & "Is too large For TextPad too open." _
           & vbNewLine & vbNewLine & "TIP : Be Sure too have " & vbNewLine & Chr$(34) & "Use external editor When opening files too large for textpad too open." & Chr$(34) & _
           vbNewLine & " Enabled in the options Dialog .", vbExclamation, "Error,TextPad"
            Form1.lblfilename.caption = ("")
              Form1.txt(1).Text = ("")
               Form1.txt(2).Text = ("")
                Form1.caption = "Untitled - TextPad"
                 fstate.dirty = False
                  Exit Sub
        End Select
        
'**************************************************************
'Nothing Suspicous Was Detected ( If Code Gets Below This Line
'**************************************************************
Query:
   msg = "This File Is Too large For TextPad Too Open." _
   & vbNewLine & vbNewLine & "Would You Like The External Editor too Open it ?"
    response = MsgBox(msg, vbDefaultButton3 + vbQuestion + vbYesNoCancel, "TextPad")

      Select Case response
        Case vbYes
         ExecuteExternalEditor (sFileName)
         End
        Case vbNo
         Form1.Show
         Exit Sub
        Case vbCancel
         Form1.Show
         Exit Sub
      End Select
End Sub
Sub No_Externaleditor_Detected()
Dim msg, response
msg = "This File is too Large for textpad too open. and also ," _
& vbNewLine & vbNewLine & "No External Editor Has been Detected" _
& vbNewLine & vbNewLine & "This means That If a file is too Large too open You will continue too see this message Until You Have TextPad Detect An External Editor." _
& vbNewLine & vbNewLine & "Do you wish Too Have TextPad Detect One For you Now ? (Recommended) "
response = MsgBox(msg, vbDefaultButton3 + vbExclamation + vbYesNo, "TextPad")
Select Case response
Case vbYes
DetectExternalEditor
MsgBox "External Editor Has been Succesfully Detected", vbInformation, "TextPad"
Exit Sub
Case vbNo
Exit Sub
End Select
End Sub
Sub Togglewordwrap(Optional ONOrOFF As Boolean)
On Error GoTo Err:
Dim fontname, fontsize As Integer
'onoff = IIf(us, True, False)
 fontname = GetSetting("TextPad", "Font", "font")
 fontsize = GetSetting("TextPad", "Font", "Fontsize")

Select Case ONOrOFF
Case False
Resizenotewithtoolbar
Form1.txt(1).visible = False
Resizenotewithtoolbar

Form1.txt(2).fontname = Form1.txt(1).fontname
Form1.txt(2).fontsize = Form1.txt(1).fontsize

Form1.txt(2).fontname = fontname
Form1.txt(2).fontsize = fontsize

Form1.txt(2).visible = True

If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If

Resizenotewithtoolbar
Form1.txt(2).Text = Form1.txt(1).Text
If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If

'Check off Menu too match current state
Form1.mnuwordwrap.Checked = False


Case True

'resize form too match current state
Resizenotewithtoolbar
Form1.txt(2).visible = False
Resizenotewithtoolbar
'set font name and size from registry
Form1.txt(1).fontname = Form1.txt(2).fontname
Form1.txt(1).fontsize = Form1.txt(2).fontsize
Form1.txt(1).visible = True
Form1.txt(1).fontname = fontname
Form1.txt(1).fontsize = fontsize
If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If
Resizenotewithtoolbar
Form1.txt(1).Text = Form1.txt(2).Text
If fstate.dirty = True Then
fstate.dirty = True
Else
If fstate.dirty = False Then
fstate.dirty = False
End If
End If
Form1.mnuwordwrap.Checked = True



'error Handler
Err:
Resume Next
End Select

End Sub

Public Function DetectFextension(sFileName As String) As String
DetectFextension = UCase$(Right(Dir(sFileName), 3))
End Function
Public Function GetShortPath(strFileName As String) As String
    Dim lngRes As Long, strPath As String
    'Create a small buffer
    strPath = String$(165, 0)
    'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
    'remove all unnecessary chr$(0)'s
    GetShortPath = Left$(strPath, lngRes)
End Function

Public Function ShowFileproperties(ownerHwnd As Long, Filename As String)
Dim sei As SHELLEXECUTEINFO
Dim R As Long
With sei
       'Set the structure's size
        .cbSize = Len(sei)
        'Sett the mask
        .fMask = SEE_MASK_NOCLOSEPROCESS Or _
         SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        'Set the owner window
        .Hwnd = ownerHwnd
        'Show the properties
        .lpVerb = "properties"
        'Set the Filename
        .lpFile = Filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
End With
R = ShellExecuteEx(sei)
End Function


'********************************************************************
'////////////////////////////////////////////////////////
' **** Commented out because the Function has not been completed yet
'Public Function AddtooRecentfiles(sFileName As String, Optional nPOS As Integer)
'Static i As Integer
'i = nPOS
'Dim retval As String
'retval = sFileName
'SaveRegistryString "Recentfiles", i, sFileName
'Form1.mnurecentfile(0).visible = True
'Form1.mnurecentfile(i).visible = True
'Form1.mnurecentfile(i).caption = retval
'End Function
'/////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////
' **** Commented out because the Function has not been completed yet
'Public Function Getrecentfiles()
'Dim i As Integer
'Dim thefiles As Variant
' If GetSetting("TextPad", "Recentfiles", "1", "") = Empty Then Exit Function
'thefiles = GetAllSettings("TextPad", "Recentfiles")
'
'For i = 1 To UBound(thefiles, 1)
'Form1.mnurecentfile(0).visible = True
'Form1.mnurecentfile(i).caption = thefiles(i, 1)
'Form1.mnurecentfile(i).visible = True
'Next i
'End Function
'/////////////////////////////////////////////////////////




