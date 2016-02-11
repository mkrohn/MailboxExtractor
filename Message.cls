VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarContent As String
Private mvarsDate As String
Private mvardDate As Date
Private mvarSize As Long
Private mvarBoundary As String
Private mvarSubject As String

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function SetFileTimeWrite Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByVal MullP As Long, ByVal NullP2 As Long, lpLastWriteTime As FILETIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Public Event OnError(Code As Long, Source As String, Description As String)

Public Property Let Subject(ByVal vData As String)
    mvarSubject = vData
End Property

Public Property Get Subject() As String
    Subject = mvarSubject
End Property

Public Property Let Boundary(ByVal vData As String)
    mvarBoundary = vData
End Property

Public Property Get Boundary() As String
    Boundary = mvarBoundary
End Property

Public Function Save(Path As String, Nro As Long, Total As Long) As Boolean
    Dim iFileOut As Long
    Dim sEmlName As String

    On Error GoTo ErrHandler
    ' open output file
    iFileOut = FreeFile
    sEmlName = MakeEMLName(Path, Nro, Total)
    Open sEmlName For Output As #iFileOut
            
    ' write it as a DOS file
    Print #iFileOut, Replace(mvarContent, vbLf, vbCrLf)
    Close #iFileOut
    Save = True
    
    ModifyFileTimeWrite (sEmlName)
    
    Exit Function
    
ErrHandler:
    RaiseEvent OnError(Err.Number, "Message:Save", Err.Description & vbCrLf & sEmlName)
    Save = False
    Exit Function
End Function

Public Property Let Size(ByVal vData As Long)
    mvarSize = vData
End Property

Public Property Get Size() As Long
    Size = mvarSize
End Property

Public Property Let Content(ByVal vData As String)
    mvarContent = vData
    
    ' strip leading linefeeds (this happens from the 2nd message)
    mvarContent = TrimLf(mvarContent)
    
    ' strip boundary from Message:
    mvarContent = Replace(mvarContent, TrimLf(mvarBoundary) & vbLf, "")
    
    mvarSize = Len(mvarContent)
End Property

Public Property Get Content() As String
Attribute Content.VB_UserMemId = 0
    Content = mvarContent
End Property

Private Function MakeEMLName(ByRef MboxDir As String, MailNum As Long, MailsTotal As Long) As String
    Dim reSubject As RegExp
    Dim mcSubjects As MatchCollection
    Dim iCtrLen As Integer
    
    iCtrLen = Len(CStr(MailsTotal))
    Set reSubject = New RegExp
    reSubject.Pattern = "Subject: .*"
    Set mcSubjects = reSubject.Execute(mvarContent)
    If mcSubjects.Count > 0 Then
        ' Use first subject found
        MakeEMLName = mcSubjects.Item(0).Value
        MakeEMLName = Replace(MakeEMLName, "Subject: ", "")
        ' decode subject
        MakeEMLName = DecodeSubject(MakeEMLName)
        mvarSubject = MakeEMLName
        ' check for invalid characters
        MakeEMLName = ValidFilename(MakeEMLName)
        ' add separator space
        MakeEMLName = " " & MakeEMLName
    End If
    
    ' add full path
    MakeEMLName = MboxDir & "\" & Right(String(iCtrLen, "0") & CStr(MailNum + 1), iCtrLen) & MakeEMLName & ".eml"
                    
    Set mcSubjects = Nothing
    Set reSubject = Nothing
    
End Function

Private Function DecodeSubject(ByRef MyString As String) As String
    Dim sSubject As String
    Dim lCodeStart As Long
    Dim lCodeEnd As Long
    Dim lStartEnc As Long
    
    
    On Error GoTo Errores
    lCodeStart = InStr(MyString, "=?")
    lCodeEnd = InStr(MyString, "?=")
    lStartEnc = InStr(MyString, "?Q?")
    
    If lStartEnc > 0 _
        And lCodeStart > 0 _
        And lCodeEnd > 0 Then
        sSubject = Mid(MyString, lStartEnc + 3, lCodeEnd - lStartEnc - 3)
        sSubject = MimeDecode(sSubject)
        sSubject = Mid(MyString, 1, lCodeStart - 1) & sSubject & Mid(MyString, lCodeEnd + 2)
        DecodeSubject = sSubject
    Else
        DecodeSubject = MyString
    End If

    Exit Function
    
Errores:
    DecodeSubject = MyString
    Exit Function
End Function

Private Function MimeDecode(ByRef MyString) As String
    Dim sCod As String
    Dim sAnt As String
    Dim sSig As String
    Dim sAux As String
    
    On Error GoTo Errores
    
    sAux = MyString
    
    Do While InStr(sAux, "=") > 0
        sCod = Mid(sAux, InStr(sAux, "=") + 1, 2)
        sAnt = Left(sAux, InStr(sAux, "=") - 1)
        sSig = Right(sAux, Len(sAux) - InStr(sAux, "=") - 2)
        sAux = sAnt & Chr(CLng("&h" & sCod)) & sSig
    Loop
    sAux = Replace(sAux, "_", " ")
    MimeDecode = sAux
    Exit Function
Errores:
    MimeDecode = MyString
    Exit Function
End Function

Private Function ValidFilename(sFileName) As String
    sFileName = Replace(sFileName, "\", "")
    sFileName = Replace(sFileName, "/", "")
    sFileName = Replace(sFileName, ":", "")
    sFileName = Replace(sFileName, "*", "")
    sFileName = Replace(sFileName, "?", "")
    sFileName = Replace(sFileName, Chr(34), "")
    sFileName = Replace(sFileName, "<", "")
    sFileName = Replace(sFileName, ">", "")
    sFileName = Replace(sFileName, "|", "")
    ValidFilename = sFileName
End Function

Private Function TrimLf(ByRef MyString As String) As String
    Do While Left(MyString, 1) = vbLf
        MyString = Mid(MyString, 2)
    Loop
    Do While Right(MyString, 1) = vbLf
        MyString = Mid(MyString, 1, Len(MyString) - 1)
    Loop
    TrimLf = MyString
End Function

Public Sub SearchDate()
    Dim reDate As RegExp
    Dim mcDates As MatchCollection
    
    Set reDate = New RegExp
    reDate.Pattern = "Date:( |\t)+(.*)"
    Set mcDates = reDate.Execute(mvarContent)
    
    If mcDates.Count > 0 Then
        mvarsDate = mcDates.Item(0).Value
        mvarsDate = reDate.Replace(mvarsDate, "$2")
    End If
    
    Set mcDates = Nothing
    Set reDate = Nothing
End Sub

Function ModifyFileTimeWrite(Filename As String) As Integer
    Dim X As Long
    Dim Handle As Long
    Dim System_Time As SYSTEMTIME
    Dim File_Time As FILETIME
    Dim Local_Time As FILETIME
    Dim TimeStamp As Variant
    Dim asFecha() As String
    Dim dFecha As Date
    Dim reDate As RegExp
    Dim mcDates As MatchCollection
    Dim sGmtSigno As String, sGmtHora As String, sGmtMinutos As String
    Dim bUTC As Boolean
    
    On Error GoTo ErrorModifyFileStamp
    
    asFecha = Split(mvarsDate, " ")
    
    bUTC = True
    
    If Mid(mvarsDate, 4, 2) = ", " Then
        ' there is day in spanish (lun, mar, mie, jue...), ignores it:
        TimeStamp = asFecha(3) & "-" & asFecha(2) & "-" & asFecha(1) & " " & asFecha(4)
    Else
        ' without weekday
        TimeStamp = asFecha(2) & "-" & asFecha(1) & "-" & asFecha(0) & " " & asFecha(3)
    End If
    
    TimeStamp = Replace(TimeStamp, "jan", "ene", 1, -1, vbTextCompare)
    TimeStamp = Replace(TimeStamp, "apr", "abr", 1, -1, vbTextCompare)
    TimeStamp = Replace(TimeStamp, "Aug", "ago", 1, -1, vbTextCompare)
    TimeStamp = Replace(TimeStamp, "dec", "dic", 1, -1, vbTextCompare)
    
    TimeStamp = CDate(TimeStamp)
    
    ' if it has GMT offset information
    Set reDate = New RegExp
    reDate.Pattern = "[\+\-][0-9]{4}"
    Set mcDates = reDate.Execute(mvarsDate)
    
    If mcDates.Count > 0 Then
        sGmtSigno = Left(mcDates.Item(0).Value, 1)
        sGmtHora = Mid(mcDates.Item(0).Value, 2, 2)
        sGmtMinutos = Mid(mcDates.Item(0).Value, 4, 2)
        TimeStamp = DateAdd("h", CInt(sGmtSigno & "1") * -1 * CInt(sGmtHora), TimeStamp)
        TimeStamp = DateAdd("m", CInt(sGmtSigno & "1") * -1 * CInt(sGmtMinutos), TimeStamp)
    Else
        reDate.Pattern = "\bGMT\b"
        Set mcDates = reDate.Execute(mvarsDate)
        
        If mcDates.Count > 0 Then
            sGmtSigno = "+"
            sGmtHora = "00"
            sGmtMinutos = "00"
            TimeStamp = DateAdd("h", CInt(sGmtSigno & "1") * -1 * CInt(sGmtHora), TimeStamp)
            TimeStamp = DateAdd("m", CInt(sGmtSigno & "1") * -1 * CInt(sGmtMinutos), TimeStamp)
        Else
            bUTC = False
        End If
    End If
                
    Set reDate = Nothing
    Set mcDates = Nothing
    
    mvardDate = TimeStamp
    
lblGetTime:
    System_Time.wYear = Year(TimeStamp)
    System_Time.wMonth = Month(TimeStamp)
    System_Time.wDay = Day(TimeStamp)
    System_Time.wDayOfWeek = Weekday(TimeStamp) - 1
    System_Time.wHour = Hour(TimeStamp)
    System_Time.wMinute = Minute(TimeStamp)
    System_Time.wSecond = Second(TimeStamp)
    System_Time.wMilliseconds = 0

    'convert the system time to a file time
    X = SystemTimeToFileTime(System_Time, Local_Time)

    'convert local file time to file time based on UTC
    X = LocalFileTimeToFileTime(Local_Time, File_Time)

    'open the file so we can get a file handle to the file
    Handle = CreateFile(Filename, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    'now change the file time and date stamp
    If bUTC Then
        X = SetFileTimeWrite(Handle, ByVal 0&, ByVal 0&, Local_Time)
    Else
        X = SetFileTimeWrite(Handle, ByVal 0&, ByVal 0&, File_Time)
    End If
    
    
    CloseHandle Handle
    
    Exit Function
    
ErrorModifyFileStamp:
    TimeStamp = Now
    bUTC = False
    Resume lblGetTime
End Function


