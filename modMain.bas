Attribute VB_Name = "modMain"
Option Explicit
Public clApplication As Application

Sub Main()
'Purpose: Extract all mails from a unix mailbox (mbox file)
'Inputs: command line: mbox file name
'Outputs:
'Author: Krohn, Martin - 14/07/03
'Last modification: Krohn, Martin - 29/07/03
            
    ' toma el command line:
    If Command = "" Then
        MsgBox "Usage: " & vbCrLf & "* " & App.EXEName & " filename." & vbCrLf & "* Drag mbox file and drop on this app icon.", vbInformation
        Exit Sub
    End If
                                
    Set clApplication = New Application
    
    clApplication.sCommandLine = Command
    
    clApplication.Run
        
    Set clApplication = Nothing
End Sub




