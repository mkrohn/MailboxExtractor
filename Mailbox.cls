VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mailbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Mensajes"
Option Explicit

Public Enum ProcStatus
    Succesful = 0
    Canceled = 1
    Error = 2
End Enum

Private mvarPath As String
Private mvarSize As Long
Private mvarContents As String
Private mvarOutDir As String

Private reBoundary As RegExp
Private mcBoundaries As MatchCollection

' mail boundary is From user@domain.com  Thu Jul 10 17:56:22 2003
Private Const MAIL_BOUNDARY_RE = "(^|\n\n)From [\t ]*[^\t\n ]+[\t ]+[^\n\t ]+.*"
' (In VB regexes, ^$ signs indicate begin-end of the whole string)

Private WithEvents clMessage As Message
Attribute clMessage.VB_VarHelpID = -1
Private mvarMessages As Messages

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Event OnProcessEnd(Status As ProcStatus, MessagesProcessed As Long)
Public Event OnMessageAdded(Nro As Long, Subject As String)
Public Event OnCancel(Cancel As Boolean)
Public Event OnError(Code As Long, Source As String, Description As String)

Public Property Let OutDir(ByVal vData As String)
    mvarOutDir = vData
End Property

Public Property Get OutDir() As String
    OutDir = mvarOutDir
End Property

Public Property Get Messages() As Messages
    If mvarMessages Is Nothing Then
        Set mvarMessages = New Messages
    End If

    Set Messages = mvarMessages
End Property

Public Property Set Messages(vData As Messages)
    Set mvarMessages = vData
End Property

Private Sub Class_Initialize()
    Set reBoundary = New RegExp
    Load frmStatus
End Sub

Private Sub Class_Terminate()
    Set mvarMessages = Nothing
    Set mcBoundaries = Nothing
    Set reBoundary = Nothing
    Unload frmStatus
End Sub

Private Function LoadFile(Optional sPath As String) As Boolean
    Dim iFileIn As Long
    
    iFileIn = FreeFile
    
    If Trim(sPath) <> "" Then
        mvarPath = sPath
    End If
    
    mvarPath = TrimChr34(mvarPath)
    
    If Not FileExists(mvarPath) Then
        RaiseEvent OnError(1000, "Mailbox:LoadFile", "Archivo no encontrado: " & mvarPath)
        LoadFile = False
        Exit Function
    End If
    
    mvarSize = FileLen(mvarPath)
        
    Open mvarPath For Binary As #iFileIn
    mvarContents = Space(mvarSize)
    Get #iFileIn, , mvarContents
    Close #iFileIn
    LoadFile = True
End Function

Public Property Let Contents(ByVal vData As String)
    mvarContents = vData
End Property

Public Property Get Contents() As String
Attribute Contents.VB_UserMemId = 0
    Contents = mvarContents
End Property

Public Property Get Size() As Long
    Size = mvarSize
End Property

Public Property Let Path(ByVal vData As String)
    mvarPath = vData
End Property

Public Property Get Path() As String
    Path = mvarPath
End Property

Public Sub Process(Optional sPath As String)
    Dim lEmlStart As Long
    Dim lEmlEnd As Long
    Dim lEmlSize As Long
    Dim I As Long
    Dim sProximoBoundary As String
    Dim bConfirmaFin As Boolean
            
    If Not LoadFile(sPath) Then
        RaiseEvent OnProcessEnd(Error, 0)
        Exit Sub
    End If
    
    ' if it is a DOS file
    If InStr(1, mvarContents, vbCrLf, vbBinaryCompare) > 0 Then
        ' convert it to unix file
        mvarContents = Replace(mvarContents, vbCrLf, vbLf)
        mvarSize = Len(mvarContents)
    End If
            
    'find all boundaries and store them in mcBoundaries
    Set reBoundary = New RegExp
    reBoundary.Global = True
    reBoundary.Pattern = MAIL_BOUNDARY_RE
    Set mcBoundaries = reBoundary.Execute(mvarContents)
            
    If mcBoundaries.Count > 20 Then
        ' show progress window
        frmStatus.prgExtracc.Max = mcBoundaries.Count
        frmStatus.Show
    Else
        If mcBoundaries.Count = 0 Then
            RaiseEvent OnProcessEnd(Succesful, 0)
            Exit Sub
        End If
    End If
    
    'if outdir is not specified
    If Trim(mvarOutDir) = "" Then
        ' build it using mailbox name
        BuildOutDir
    End If
    
    ' create dir if not exists
    If Not MakeOutDir Then
        RaiseEvent OnProcessEnd(Error, I)
        Exit Sub
    End If
    
    lEmlEnd = 1
    ' for each boundary found
    For I = 0 To mcBoundaries.Count - 1
        frmStatus.prgExtracc.Value = I + 1
        frmStatus.lblNums = CStr(I + 1) & " of " & mcBoundaries.Count
        DoEvents
        
        Set clMessage = New Message
        
        clMessage.Boundary = mcBoundaries.Item(I).Value
        
        '  search Message start
        lEmlStart = lEmlEnd ' from the end of last Message
        
        '  find Message size
        If I < (mcBoundaries.Count - 1) Then ' if not the last Message
            sProximoBoundary = mcBoundaries.Item(I + 1).Value
            lEmlEnd = InStr(lEmlStart + 1, mvarContents, sProximoBoundary, vbBinaryCompare)
            lEmlSize = lEmlEnd - lEmlStart
        Else
            lEmlSize = mvarSize - lEmlStart
        End If
            
        clMessage.Content = Mid(mvarContents, lEmlStart, lEmlSize)
        clMessage.SearchDate
        If Not clMessage.Save(mvarOutDir, I, mcBoundaries.Count) Then
            RaiseEvent OnProcessEnd(Error, I)
            Exit Sub
        End If
        Messages.Add clMessage
        RaiseEvent OnMessageAdded(I + 1, clMessage.Subject)
        
        Set clMessage = Nothing
        
        If frmStatus.bCancel Then
            bConfirmaFin = True
            RaiseEvent OnCancel(bConfirmaFin)
            If bConfirmaFin Then
                Exit For
            Else
                frmStatus.bCancel = False
                frmStatus.Visible = True
            End If
        End If
        
    Next I
            
    If bConfirmaFin Then
        RaiseEvent OnProcessEnd(Canceled, I + 1)
    Else
        RaiseEvent OnProcessEnd(Succesful, I + 1)
    End If
            
End Sub

Private Function MakeOutDir() As Boolean
    
    ' if dir already exists, ignores the error silently
    On Error GoTo Salir
    
    MkDir mvarOutDir
    
    On Error GoTo 0
    
Salir:
    If Not FileExists(mvarOutDir) Then
        RaiseEvent OnError(1200, "Mailbox:MakeOutDir", "Invalid OutDir: " & vbCrLf & mvarOutDir)
        MakeOutDir = False
    Else
        MakeOutDir = True
    End If
    
    Exit Function
End Function

Private Sub BuildOutDir()
    Dim sMboxDir As String
    Dim sMboxName As String
    
    ' decompose mailbox path
    If InStr(mvarPath, "\") > 0 Then
        sMboxDir = Left(mvarPath, InStrRev(mvarPath, "\") - 1) & "\"
        sMboxName = Mid(mvarPath, InStrRev(mvarPath, "\") + 1)
    Else
        If InStr(mvarPath, ":") > 0 Then
            sMboxDir = Left(mvarPath, 2)
            sMboxName = Mid(mvarPath, InStrRev(mvarPath, ":") + 1)
        Else
            sMboxDir = App.Path & "\"
            sMboxName = mvarPath
        End If
    End If

    mvarOutDir = sMboxDir & sMboxName & " mails\"
    
End Sub

Private Sub VbShellExecute(ByVal lpFile As String)

    On Error GoTo L_ERR

    Call ShellExecute(0, "open", lpFile, "", App.Path, 1)
    DoEvents
    Exit Sub

L_ERR:
    Err.Raise 1001, "Mailbox:vbShellExecute", Err.Description

End Sub

Private Function FileExists(sFileName) As Boolean
    On Error GoTo SalirPorFalse
    
    GetAttr sFileName
    FileExists = True
    Exit Function
SalirPorFalse:
    FileExists = False
    Exit Function
End Function

Public Sub OpenOutDir()
    Call VbShellExecute(mvarOutDir)
End Sub

Private Function TrimChr34(ByRef MyString As String) As String
    Do While Left(MyString, 1) = Chr(34)
        MyString = Mid(MyString, 2)
    Loop
    Do While Right(MyString, 1) = Chr(34)
        MyString = Mid(MyString, 1, Len(MyString) - 1)
    Loop
    
    TrimChr34 = MyString
End Function

Private Sub clMessage_OnError(Code As Long, Source As String, Description As String)

    RaiseEvent OnError(Code, Source, Description)
    
End Sub
