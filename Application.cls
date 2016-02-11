VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents clMailbox As Mailbox
Attribute clMailbox.VB_VarHelpID = -1
Public sCommandLine As String

Private Sub Class_Initialize()
    Set clMailbox = New Mailbox
End Sub

Private Sub Class_Terminate()
    Set clMailbox = Nothing
End Sub

Private Sub clMailbox_OnCancel(Cancel As Boolean)
    If vbYes = MsgBox("¿ Desea cancelar el proceso ?", vbYesNo + vbQuestion) Then
        Cancel = True
    Else
        Cancel = False
    End If
End Sub

Private Sub clMailbox_OnError(Code As Long, Source As String, Description As String)
    MsgBox "Error " & CStr(Code) & ": " & Description & vbCrLf & "Source: " & Source, vbCritical
End Sub

Private Sub clMailbox_OnMessageAdded(Nro As Long, Subject As String)
    Debug.Print Nro & ": " & Subject
End Sub

Private Sub clMailbox_OnProcessEnd(Status As ProcStatus, MessagesProcessed As Long)
    
    Select Case Status
        Case Error:
            MsgBox "El proceso ha terminado con errores", vbInformation
        Case Succesful:
            If MessagesProcessed = 0 Then
                MsgBox "El mailbox está vacío.", vbInformation
            End If
        Case Canceled:
            MsgBox "El proceso ha sido cancelado", vbInformation
    End Select
    
    If MessagesProcessed > 0 Then
        clMailbox.OpenOutDir
    End If
    
End Sub

Public Sub Run()
    clMailbox.Path = sCommandLine
    clMailbox.Process
End Sub
