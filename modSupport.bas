Attribute VB_Name = "modSupport"
Option Explicit

Private Const MODULE_NAME As String = "[FileAX:modSupport."

Public Function GetFileBinary(ByVal Filename As String) As String

    Dim lFile As Long, lLen As Long, lCnt As Long
    Dim arByte() As Byte
    Dim sStr As String
    
    On Error GoTo GetFileBinary_Error
    
    lFile = FreeFile
    
    lLen = FileLen(Filename)
    
    Open Filename For Binary As #lFile
    
    sStr = String(LOF(lFile), " ")
    
    Get #lFile, 1, sStr
    
    Close #lFile
    
    
    GetFileBinary = sStr
    
Exit_Properly:
    Exit Function

GetFileBinary_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "GetFileBinary]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error getting binary file"
    GoTo Exit_Properly

End Function

Public Function ReturnExt(ByVal Filename As String) As String

    Dim lPos As Long
    
    lPos = InStr(Filename, ".")
    
    ReturnExt = Mid$(Filename, lPos + 1)

End Function

Public Function SUBound(ByRef sArr() As String) As Long

    On Error GoTo err_handle
    
    SUBound = UBound(sArr)
    
    Exit Function
    
err_handle:
    SUBound = -1
    
End Function

Public Function GetPicIndex(ByRef oObj As Object) As Long

    'what type is this file?
    If TypeName(oObj) = "Folder" Then
        GetPicIndex = 1
    Else
        GetPicIndex = 2
    End If
    

End Function

Public Function GetDriveType(ByVal lType As Long) As String

    Select Case lType
        Case 0
            GetDriveType = "Unknown type"
        Case 1
            GetDriveType = "Removable Disk"
        Case 2
            GetDriveType = "Local Disk"
        Case 3
            GetDriveType = "Remote Disk"
        Case 4
            GetDriveType = "CD/DVD ROM"
        Case 5
            GetDriveType = "RAM Disk"
    End Select

End Function

Public Function GetDrivePicIndex(ByVal lType As Long) As Long

    Select Case lType
        Case 0
            GetDrivePicIndex = 6
        Case 1
            GetDrivePicIndex = 5
        Case 2
            GetDrivePicIndex = 2
        Case 3
            GetDrivePicIndex = 4
        Case 4
            GetDrivePicIndex = 1
        Case 5
            GetDrivePicIndex = 7
    End Select
    
End Function
Public Function GetFileFromPath(ByVal sPath As String) As String

    Dim lPos As Long
    Dim lAct As Long
    
    lPos = InStr(sPath, "\")
    Do While lPos <> 0
        lPos = InStr(lPos + 1, sPath, "\")
        If lPos <> 0 Then lAct = lPos
    Loop
    GetFileFromPath = Mid$(sPath, lAct + 1)
    
End Function

Public Function GetNumSelected(ByRef lv As ListView) As Long

    Dim oItem As ListItem
    Dim lCnt As Long
    
    On Error GoTo GetNumSelected_Error
    
    For Each oItem In lv.ListItems
        If oItem.Selected Then
            lCnt = lCnt + 1
        End If
    Next oItem
    
    GetNumSelected = lCnt
    
Exit_Properly:
    Set oItem = Nothing
    Exit Function
    
GetNumSelected_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "GetNumSelected]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error retrieving selected item count"
    GoTo Exit_Properly

End Function

Public Function ValidRefNo(ByVal strRefNo As String) As Boolean

    'check that there is actually a string
    If Trim$(strRefNo) = "" Then
        ValidRefNo = False
        Exit Function
    End If
    
    'is it a numeric
    If Not IsNumeric(strRefNo) Then
        ValidRefNo = False
        Exit Function
    End If
    
    ValidRefNo = True
    
End Function
