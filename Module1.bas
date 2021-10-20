Attribute VB_Name = "Module1"
'#Refference#
'   Windows script host object model
'#Purpose#
'   Extract code / deobfuscate / analyze without opening the file

Option Explicit

Sub test()
Dim OUTPUT As String
OUTPUT = oleVBA("G:\My Drive\SOFTWARE\EXCEL\0 Alex\CodePrinter\CodeArchive.xlsm")
'Debug.Print str

Overwrite ThisWorkbook.path & "\" & "TestOutput.txt", OUTPUT
End Sub

Sub LoopFilesFromSheet()
Dim FILE_PATH   As String
Dim FILE_NAME   As String
Dim OUT_PATH    As String
Dim OUTPUT      As String
Dim WORK_SHEET  As Worksheet:   Set WORK_SHEET = ActiveSheet
Dim CURRENT_CELL As Range:      Set CURRENT_CELL = WORK_SHEET.Range("A2")
Dim OUT_STATUS As Range:        Set OUT_STATUS = CURRENT_CELL.Offset(0, 1)
Do While CURRENT_CELL <> ""
    If OUT_STATUS = "" Then
    
        FILE_PATH = CURRENT_CELL
        FILE_NAME = getFilePartName(FILE_PATH)
        OUTPUT = oleVBA(FILE_PATH)
        OUT_PATH = ThisWorkbook.path & "\" & FILE_NAME & ".txt"
        
        Overwrite OUT_PATH, OUTPUT
        
        With WORK_SHEET
            .Hyperlinks.Add Anchor:=CURRENT_CELL.Offset(0, 1), _
            Address:=OUT_PATH, _
            ScreenTip:="", _
            TextToDisplay:=FILE_NAME
        End With
               
    End If
    
    Set CURRENT_CELL = CURRENT_CELL.Offset(1, 0)
    Set OUT_STATUS = CURRENT_CELL.Offset(0, 1)
Loop
End Sub

Function oleVBA(path As String) As String
    ' /C will execute the command and Terminate the window
    Dim Q As String
    Q = """"
    oleVBA = ShellText("cmd.exe /c olevba " & Q & path & Q)
End Function

Public Function ShellText(FuncExec As String) As String
    Dim wsh As Object, wshOut As Object, sShellOut As String, sShellOutLine As String
    
    'Create object for Shell command execution
    Set wsh = CreateObject("WScript.Shell")

    'Run Excel VBA shell command and get the output string
    Set wshOut = wsh.Exec(FuncExec).StdOut

    'Read each line of output from the Shell command & Append to Final Output Message
    While Not wshOut.AtEndOfStream
        sShellOutLine = wshOut.ReadLine
        If sShellOutLine <> "" Then
            sShellOut = sShellOut & sShellOutLine & vbCrLf
        End If
    Wend

    'Return the Output of Command Prompt
    ShellText = sShellOut
End Function

'---------------------------------------------------------------------------------------
' Procedure : Overwrite
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Output Data to an external file (*.txt or other format)
'             ***Do not forget about access' DoCmd.OutputTo Method for
'             exporting objects (queries, report,...)***
'             Will overwirte any data if the file already exists
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites (links
'             back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - name of the file that the text is to be output to including the full path
' sText - text to be output to the file
'
' Usage:
' ~~~~~~
' Call Overwrite("C:\Users\Vance\Documents\EmailExp2.txt", "Text2Export")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2012-Jul-06                 Initial Release
'---------------------------------------------------------------------------------------
Function Overwrite(sFile As String, sText As String)
On Error GoTo Err_Handler
    Dim FileNumber As Integer
 
    FileNumber = FreeFile                   ' Get unused file number
    Open sFile For Output As #FileNumber    ' Connect to the file
    Print #FileNumber, sText                ' Append our string
    Close #FileNumber                       ' Close the file
 
Exit_Err_Handler:
    Exit Function
 
Err_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: Overwrite" & vbCrLf & _
           "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function getFilePartName(fileNameWithExtension As String, Optional IncludeExtension As Boolean) As String
If InStr(1, fileNameWithExtension, "\") > 0 Then
    getFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "\"))
Else
    getFilePartName = fileNameWithExtension
End If
    If IncludeExtension = False Then getFilePartName = Left(getFilePartName, InStr(1, getFilePartName, ".") - 1)
End Function

