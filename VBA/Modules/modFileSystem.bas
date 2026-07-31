Attribute VB_Name = "modFileSystem"
Option Explicit

'=========================================================
' USSP
'
' Module:
'   modFileSystem
'
' Purpose:
'   File and folder helper functions
'
'=========================================================

Public Function FileExists(ByVal FileName As String) As Boolean

    FileExists = (Len(Dir(FileName, vbNormal)) > 0)

End Function


Public Function FolderExists(ByVal FolderName As String) As Boolean

    FolderExists = False

    On Error Resume Next

    FolderExists = _
        (GetAttr(FolderName) And vbDirectory) = vbDirectory

    On Error GoTo 0

End Function


Public Function CreateFolder(ByVal FolderName As String) As Boolean

    On Error GoTo ErrorHandler

    If FolderExists(FolderName) Then

        CreateFolder = True
        Exit Function

    End If

    MkDir FolderName

    CreateFolder = True

    Exit Function

ErrorHandler:

    HandleError "CreateFolder", Err

    CreateFolder = False

End Function


Public Function DeleteFile(ByVal FileName As String) As Boolean

    On Error GoTo ErrorHandler

    If FileExists(FileName) Then

        Kill FileName

    End If

    DeleteFile = True

    Exit Function

ErrorHandler:

    HandleError "DeleteFile", Err

    DeleteFile = False

End Function

