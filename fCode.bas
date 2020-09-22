Attribute VB_Name = "fCode"
Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Declare Function FileTimeToSystemTime Lib "kernel32.dll" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Declare Function SystemTimeToFileTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Declare Function SetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long

Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Const GENERIC_READ = &H80000000
Const GENERIC_WRITE = &H40000000
Const FILE_SHARE_READ = &H1
Const FILE_SHARE_WRITE = &H2
Const CREATE_ALWAYS = 2
Const CREATE_NEW = 1
Const OPEN_ALWAYS = 4
Const OPEN_EXISTING = 3
Const TRUNCATE_EXISTING = 5
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Const FILE_FLAG_NO_BUFFERING = &H20000000
Const FILE_FLAG_OVERLAPPED = &H40000000
Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Const FILE_FLAG_WRITE_THROUGH = &H80000000

Public Function GetFileAccessDate(fString As String) As Date
Dim hFile As Long
Dim ctime As FILETIME
Dim atime As FILETIME
Dim mtime As FILETIME
Dim thetime As SYSTEMTIME
Dim retval As Long
hFile = CreateFile(fString, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    If hFile = -1 Then
        MsgBox "Error Getting Date on " & fString
        Exit Function
    End If
retval = GetFileTime(hFile, ctime, atime, mtime)
retval = FileTimeToLocalFileTime(atime, atime)
retval = FileTimeToSystemTime(atime, thetime)
retval = CloseHandle(hFile)
GetFileAccessDate = CDate(thetime.wMonth & "/" & thetime.wDay & "/" & thetime.wYear)
End Function
Public Function GetFileModifyDate(fString As String) As Date
Dim hFile As Long
Dim ctime As FILETIME
Dim atime As FILETIME
Dim mtime As FILETIME
Dim thetime As SYSTEMTIME
Dim retval As Long
hFile = CreateFile(fString, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    If hFile = -1 Then
        MsgBox "Error Getting Date on " & fString
        Exit Function
    End If
retval = GetFileTime(hFile, ctime, atime, mtime)
retval = FileTimeToLocalFileTime(mtime, mtime)
retval = FileTimeToSystemTime(mtime, thetime)
retval = CloseHandle(hFile)
GetFileModifyDate = CDate(thetime.wMonth & "/" & thetime.wDay & "/" & thetime.wYear)
End Function
Public Function GetFileCreatedDate(fString As String) As Date
Dim hFile As Long
Dim ctime As FILETIME
Dim atime As FILETIME
Dim mtime As FILETIME
Dim thetime As SYSTEMTIME
Dim retval As Long
hFile = CreateFile(fString, GENERIC_READ, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    If hFile = -1 Then
        MsgBox "Error Getting Date on " & fString
        Exit Function
    End If
retval = GetFileTime(hFile, ctime, atime, mtime)
retval = FileTimeToLocalFileTime(ctime, ctime)
retval = FileTimeToSystemTime(ctime, thetime)
retval = CloseHandle(hFile)
GetFileCreatedDate = CDate(thetime.wMonth & "/" & thetime.wDay & "/" & thetime.wYear)
End Function

Public Function SetFileDates(fString As String, sCdate As Date, sAdate As Date, sMdate As Date) As Boolean
Dim hFile As Long
Dim cctime As SYSTEMTIME
Dim aatime As SYSTEMTIME
Dim mmtime As SYSTEMTIME
Dim ctime As FILETIME
Dim atime As FILETIME
Dim mtime As FILETIME
Dim retval As Long

cctime.wDay = CInt(Format(sCdate, "dd"))
cctime.wMonth = CInt(Format(sCdate, "mm"))
cctime.wYear = CInt(Format(sCdate, "yyyy"))

aatime.wDay = CInt(Format(sAdate, "dd"))
aatime.wMonth = CInt(Format(sAdate, "mm"))
aatime.wYear = CInt(Format(sAdate, "yyyy"))

mmtime.wDay = CInt(Format(sMdate, "dd"))
mmtime.wMonth = CInt(Format(sMdate, "mm"))
mmtime.wYear = CInt(Format(sMdate, "yyyy"))

retval = SystemTimeToFileTime(cctime, ctime)
retval = SystemTimeToFileTime(aatime, atime)
retval = SystemTimeToFileTime(mmtime, mtime)

hFile = CreateFile(fString, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ, ByVal CLng(0), OPEN_EXISTING, FILE_ATTRIBUTE_ARCHIVE, 0)
    If hFile = -1 Then
        MsgBox "Could Not Set File Dates On " & fString
        SetFileDates = False
        Exit Function
    End If
retval = SetFileTime(hFile, ctime, atime, mtime)
retval = CloseHandle(hFile)
SetFileDates = True
End Function
