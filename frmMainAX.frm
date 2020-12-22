VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMainAX 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monica's DreamCasa Picture Uploader"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10605
   Icon            =   "frmMainAX.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPropRefNo 
      Height          =   285
      Left            =   2340
      TabIndex        =   13
      Top             =   345
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9330
      TabIndex        =   9
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Upload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9330
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   4
      Top             =   5490
      Width           =   1095
   End
   Begin VB.TextBox txtFilename 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4665
      TabIndex        =   3
      Top             =   5115
      Width           =   4215
   End
   Begin VB.ComboBox cboFiletype 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmMainAX.frx":0152
      Left            =   4665
      List            =   "frmMainAX.frx":015F
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5565
      Width           =   4230
   End
   Begin VB.CheckBox chkDesc 
      Height          =   435
      Left            =   8595
      TabIndex        =   0
      Top             =   6060
      Value           =   1  'Checked
      Width           =   240
   End
   Begin MSComctlLib.ImageList ilPvw 
      Left            =   9675
      Top             =   6660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   107
      ImageHeight     =   107
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":0185
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":069A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilDrives 
      Left            =   8625
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":0A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":1772
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":244C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":3126
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":3E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":4ADA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":A2CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo icboLookIn 
      Height          =   330
      Left            =   4680
      TabIndex        =   1
      Top             =   345
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageList       =   "ilDrives"
   End
   Begin MSComctlLib.ImageList ilFile 
      Left            =   7920
      Top             =   6420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":AB9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainAX.frx":B875
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   3825
      Left            =   4665
      TabIndex        =   6
      Top             =   1080
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   6747
      View            =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ilFile"
      SmallIcons      =   "ilFile"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Property reference number:"
      Height          =   285
      Left            =   180
      TabIndex        =   14
      Top             =   375
      Width           =   2115
   End
   Begin VB.Label Label5 
      Caption         =   "Enter descriptions for selected files:"
      Height          =   360
      Left            =   4665
      TabIndex        =   12
      Top             =   6165
      Width           =   3585
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Files of type:"
      Height          =   315
      Left            =   3690
      TabIndex        =   11
      Top             =   5610
      Width           =   960
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "File name:"
      Height          =   240
      Left            =   3780
      TabIndex        =   10
      Top             =   5145
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Preview"
      Height          =   225
      Left            =   360
      TabIndex        =   8
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Look in:"
      Height          =   255
      Left            =   3750
      TabIndex        =   7
      Top             =   375
      Width           =   900
   End
   Begin VB.Image imgUp 
      Height          =   480
      Left            =   8775
      MouseIcon       =   "frmMainAX.frx":11497
      MousePointer    =   99  'Custom
      Picture         =   "frmMainAX.frx":117A1
      ToolTipText     =   "Up one level"
      Top             =   255
      Width           =   480
   End
   Begin VB.Image imgList 
      Height          =   480
      Left            =   9285
      MouseIcon       =   "frmMainAX.frx":11943
      MousePointer    =   99  'Custom
      Picture         =   "frmMainAX.frx":11C4D
      ToolTipText     =   "List view"
      Top             =   255
      Width           =   480
   End
   Begin VB.Image imgDetail 
      Height          =   480
      Left            =   9885
      MouseIcon       =   "frmMainAX.frx":1785F
      MousePointer    =   99  'Custom
      Picture         =   "frmMainAX.frx":17B69
      Stretch         =   -1  'True
      ToolTipText     =   "Detail view"
      Top             =   255
      Width           =   480
   End
   Begin VB.Image imgPvw 
      BorderStyle     =   1  'Fixed Single
      Height          =   3780
      Left            =   255
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   4200
   End
End
Attribute VB_Name = "frmMainAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fso As FileSystemObject
Private arDrives() As String
Private oItem As ListItem
Private sCurrPath As String
Private mlPropRefNo As Long
Private arSelFiles() As String
Private arSelDesc() As String

Private Const MODULE_NAME As String = "[FileAX:ctlFileAX."

Private Function GetDrives() As Long

    'This function retrieves the drives of this machine
    Dim oDrive As Drive
    Dim oItem As ComboItem
    
    On Error GoTo GetDrives_Error
    
    ReDim arDrives(fso.Drives.Count - 1)
    
    For Each oDrive In fso.Drives
        Set oItem = icboLookIn.ComboItems.Add(, , GetDriveType(oDrive.DriveType) & " " & oDrive.Path, GetDrivePicIndex(oDrive.DriveType))
        arDrives(oItem.Index - 1) = oDrive.DriveLetter
    Next oDrive

    'init to first item
    icboLookIn.ComboItems.Item(1).Selected = True
    
Exit_Properly:
    Set oDrive = Nothing
    Set oItem = Nothing
    Exit Function
    
GetDrives_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "GetDrives]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error getting drives"
    GoTo Exit_Properly
    
End Function

Private Sub GetFiles(ByVal sDrive As String)

    'get all the files for this drive
    Dim oFolder As Folder
    Dim oFolders As Folders
    Dim oFiles As Files
    Dim oFile As File
    Dim oDrive As Drive
    Dim sName As String, sBuff As String
    Dim olvItem As ListItem
    Dim bShow As Boolean
    Dim dSize As Double
    Dim sType As String
    Dim dteMod As Date
    
    On Error GoTo GetFiles_Error
    
    Set oDrive = fso.GetDrive(sDrive)
    
    If Not oDrive.IsReady = True Then
        MsgBox "Drive '" & oDrive.DriveLetter & "' is not ready or there is no disk inserted!", vbExclamation + vbOKOnly, "Drive not ready"
        'set back to the first index(which is generally local disk c:)
        icboLookIn.ComboItems(1).Selected = True
        Exit Sub
    Else
        lvFiles.ListItems.Clear
    End If
    
    Set oFolders = oDrive.RootFolder.SubFolders
    Set oFiles = oDrive.RootFolder.Files
    
    For Each oFolder In oFolders
        Set olvItem = lvFiles.ListItems.Add(, , oFolder.Name, GetPicIndex(oFolder), GetPicIndex(oFolder))
        'store path in the tag
        olvItem.Tag = oFolder.Path
        'if it is in report mode, then add the extra info
        If lvFiles.View = lvwReport Then
            'get the info for this file
            'type
            sType = "File Folder"
            'date modified
            'getting the size of certain folders crashes the FSO (probably why Win Explorer don't do it!)
            dteMod = oFolder.DateLastModified
            olvItem.SubItems(2) = sType
            olvItem.SubItems(3) = dteMod
        End If
    Next oFolder
    
    For Each oFile In oFiles
        If oFile.Attributes <> 38 Then
            'filter the files
            sBuff = UCase$(Mid$(oFile.ShortName, InStr(oFile.ShortName, ".") + 1))
            'check the filetype given be the user (defaults to Both Types)
            bShow = False
            bShow = (cboFiletype.ListIndex = 2 And (sBuff = "JPG" Or sBuff = "GIF"))
            bShow = bShow Or (cboFiletype.ListIndex = 0 And sBuff = "GIF")
            bShow = bShow Or (cboFiletype.ListIndex = 1 And sBuff = "JPG")
            If bShow Then
                Set olvItem = lvFiles.ListItems.Add(, , oFile.Name, GetPicIndex(oFile), GetPicIndex(oFile))
                'and store the full path as a tag
                olvItem.Tag = oFile.Path
                'if it is in report mode, then add the extra info
                If lvFiles.View = lvwReport Then
                    'get the info for this file
                    'size
                    dSize = oFile.Size
                    'type
                    sType = "Image"
                    'date modified
                    dteMod = oFile.DateLastModified
                    olvItem.SubItems(1) = Format$(dSize, "###,###,###,###")
                    olvItem.SubItems(2) = sType
                    olvItem.SubItems(3) = dteMod
                End If
            End If
        End If
    Next oFile
    
Exit_Properly:
    Set oFolder = Nothing
    Set oFolders = Nothing
    Set oFiles = Nothing
    Set oFile = Nothing
    Set oDrive = Nothing
    Set olvItem = Nothing
    Exit Sub
    
GetFiles_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "GetFiles]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error getting files"
    GoTo Exit_Properly
    
End Sub

Private Sub GetFolderFiles(ByVal sFolder As String)

    'get all the files for this drive
    Dim oFolder As Folder
    Dim oFolders As Folders
    Dim oCurrFolder As Folder
    Dim oFiles As Files
    Dim oFile As File
    Dim oDrive As Drive
    Dim sName As String, sBuff As String
    Dim olvItem As ListItem
    Dim bShow As Boolean
    Dim dSize As Double
    Dim sType As String
    Dim dteMod As Date
    
    On Error GoTo GetFolderFiles_Error
    
    Set oCurrFolder = fso.GetFolder(sFolder)
    'get the subfolders for this folder
    Set oFolders = oCurrFolder.SubFolders
    'and get the files for this folder
    Set oFiles = oCurrFolder.Files
    
    For Each oFolder In oFolders
        Set olvItem = lvFiles.ListItems.Add(, , oFolder.Name, GetPicIndex(oFolder), GetPicIndex(oFolder))
        olvItem.Tag = oFolder.Path
        'if it is in report mode, then add the extra info
        If lvFiles.View = lvwReport Then
            'get the info for this file
            'type
            sType = "File Folder"
            'date modified
            dteMod = oFolder.DateLastModified
            olvItem.SubItems(2) = sType
            olvItem.SubItems(3) = dteMod
        End If
    Next oFolder
    
    For Each oFile In oFiles
        If oFile.Attributes <> 38 Then
            'filter the files
            sBuff = UCase$(Mid$(oFile.ShortName, InStr(oFile.ShortName, ".") + 1))
            'check the filetype given be the user (defaults to Both Types)
            bShow = False
            bShow = (cboFiletype.ListIndex = 2 And (sBuff = "JPG" Or sBuff = "GIF"))
            bShow = bShow Or (cboFiletype.ListIndex = 0 And sBuff = "GIF")
            bShow = bShow Or (cboFiletype.ListIndex = 1 And sBuff = "JPG")
            If bShow Then
                Set olvItem = lvFiles.ListItems.Add(, , oFile.Name, GetPicIndex(oFile), GetPicIndex(oFile))
                'and store the full path as a tag
                olvItem.Tag = oFile.Path
                'if it is in report mode, then add the extra info
                If lvFiles.View = lvwReport Then
                    'get the info for this file
                    'size
                    dSize = oFile.Size
                    'type
                    sType = "Image"
                    'date modified
                    dteMod = oFile.DateLastModified
                    olvItem.SubItems(1) = Format$(dSize, "###,###,###,###")
                    olvItem.SubItems(2) = sType
                    olvItem.SubItems(3) = dteMod
                End If
            End If
        End If
    Next oFile
    
Exit_Properly:
    Set oFolder = Nothing
    Set oFolders = Nothing
    Set oCurrFolder = Nothing
    Set oFiles = Nothing
    Set oFile = Nothing
    Set oDrive = Nothing
    Set olvItem = Nothing
    Exit Sub
    
GetFolderFiles_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "GetFolderFiles]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error getting folder files"
    GoTo Exit_Properly
    
    
End Sub

Private Sub cboFiletype_Click()

    On Error GoTo cboFiletype_Click_Error
    
    lvFiles.ListItems.Clear
    If sCurrPath = vbNullString Then
        GetFiles (arDrives(icboLookIn.SelectedItem.Index - 1))
    Else
        GetFolderFiles sCurrPath
    End If
    
Exit_Properly:
    Exit Sub
    
cboFiletype_Click_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "cboFiletype_Click]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error changing file type"
    GoTo Exit_Properly
    
End Sub

Private Sub cmdExit_Click()

    Dim lRet As Long
    
    lRet = MsgBox("Are you sure you want to exit, Monica?", vbYesNo + vbQuestion, "Quit DreamCasa picture uploader")
    If lRet = vbYes Then
        Unload Me
    End If
    
End Sub

Private Sub cmdOpen_Click()

    Dim oHTTP As WinHttpRequest
    Dim lSel As Long
    Dim lCnt As Long
    Dim arFiles() As String
    Dim sBody As String, sFile As String, sRefer As String
    Dim aPostBody() As Byte
    Dim ofDesc As frmDesc
    
    On Error GoTo cmdOpen_Click_Error
    
    If Not ValidRefNo(txtPropRefNo.Text) Then
        MsgBox "Please enter a valid Property reference number into the box provided!", vbCritical + vbOKOnly, "Missing or invalid Property reference number"
        txtPropRefNo.Text = ""
        txtPropRefNo.SetFocus
        GoTo Exit_Properly
    End If
    
    'set the proprefno to the box value
    mlPropRefNo = Val(txtPropRefNo.Text)
    
    'first get a list of the selected files
    arFiles = GetSelectedFiles
    lSel = SUBound(arFiles)
    
    'does the user want to do descriptions now?
    If chkDesc.Value = vbChecked Then
        'do we have any selected files?
        If SUBound(arFiles) > -1 Then
            'we do, so present the form
            Set ofDesc = New frmDesc
            ofDesc.FileList = arFiles
            ofDesc.Show vbModal
            If ofDesc.Cancelled Then
                MsgBox "You have cancelled picture descriptions!"
                Unload ofDesc
                Set ofDesc = Nothing
                Exit Sub
            End If
            arSelDesc = ofDesc.Descriptions
        End If
    Else
        'set up a dummy description array
        If lSel > -1 Then
            ReDim arSelDesc(lSel)
        End If
    End If
    
    'now we need -
    'the current http connection

    Set oHTTP = New WinHttpRequest
    
    'connect
    oHTTP.Open "POST", "http://www.dreamcasa.mercilessdevelopment.com/Admin/upld.php", False '"http://www.andaja.com/Admin/upld.php", False
    
    'cycle through the files, building a header
    
    If lSel = -1 Then
        MsgBox "You have not selected any files to upload.  Please select file(s) and try again.", vbExclamation + vbOKOnly, "No pictures selected"
        GoTo Exit_Properly
    End If
    
    'start my header
    With oHTTP
        'do the accept bit
        .SetRequestHeader "Accept", "image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, application/x-shockwave-flash, */*"
        'starting bounds
        .SetRequestHeader "Content-type", "multipart/form-data; boundary=MyBound"
        'set a form field for the proprefno
        sBody = "--MyBound" & vbCrLf & "Content-Disposition:form-data; name=""proprefno""" & vbCrLf & vbCrLf & _
        mlPropRefNo & vbCrLf
        'build the rest of the body for the files and their descriptions
        For lCnt = 0 To lSel
            'get the file string
            sFile = GetFileBinary(arFiles(lCnt))
            'build the body to send
            sBody = sBody & "--MyBound" & vbCrLf & _
            "Content-Disposition:form-data; name=""description[]""" & vbCrLf & vbCrLf & _
            arSelDesc(lCnt) & vbCrLf & _
            "--MyBound" & vbCrLf & _
            "Content-Disposition: form-data; name=file[]; filename=""" & arFiles(lCnt) & """" & _
            vbCrLf & _
            "Content-type: file" & vbCrLf & vbCrLf & _
            sFile & vbCrLf
        Next lCnt
        'close off the body
        sBody = sBody & "--MyBound--"
        aPostBody = StrConv(sBody, vbFromUnicode)
        'and send the header
        .Send aPostBody
    End With
    MsgBox "Your files have been successfully sent to DreamCasa.com!", vbExclamation + vbOKOnly, "DreamCasa.com has received your files"
    
Exit_Properly:
    Set oHTTP = Nothing
    If Not ofDesc Is Nothing Then
        Unload ofDesc
    End If
    Set ofDesc = Nothing
    Exit Sub
    
cmdOpen_Click_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "cmdOpen_Click]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error uploading files"
    GoTo Exit_Properly

End Sub

Private Sub Form_Load()

    On Error GoTo Form_Load_Error
    
    Set fso = New FileSystemObject
    
    
    'get the drives
    Call GetDrives
    
    'set the listindex BEFORE calling GetFiles, otherwise filter wont work
    cboFiletype.ListIndex = 2
    
    Call GetFiles(arDrives(icboLookIn.SelectedItem.Index - 1))
    
    
    'set the initial image in the preview
    Set imgPvw.Picture = ilPvw.ListImages(2).Picture
    
    'force a start up (for some reason the control gets locked sometimes on the web page)
    icboLookIn.Refresh
    
Exit_Properly:
    Exit Sub
    
Form_Load_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "UserControl_Show]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error showing FileAX"
    GoTo Exit_Properly


End Sub

Private Sub icboLookIn_Click()

    On Error GoTo icboLookIn_Click_Error
    
    'attempt to change to the drive that we have currently
    GetFiles arDrives(icboLookIn.SelectedItem.Index - 1)

    sCurrPath = vbNullString
    
Exit_Properly:
    Exit Sub
    
icboLookIn_Click_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "icboLookIn_Click]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error changing drive"
    GoTo Exit_Properly

End Sub


Private Sub imgDetail_Click()

    On Error GoTo imgDetail_Click_Error
    
    lvFiles.View = lvwReport
    'set the columns etc for the listview
    
    With lvFiles
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "Name", 4000, lvwColumnLeft
        .ColumnHeaders.Add , , "Size", 1000, lvwColumnRight
        .ColumnHeaders.Add , , "Type", 1000, lvwColumnLeft
        .ColumnHeaders.Add , , "Date Modified", 1300, lvwColumnLeft
    End With
    
    lvFiles.ListItems.Clear
    If sCurrPath = vbNullString Then
        GetFiles (arDrives(icboLookIn.SelectedItem.Index - 1))
    Else
        GetFolderFiles sCurrPath
    End If
        
Exit_Properly:
    Exit Sub
    
imgDetail_Click_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "imgDetail_Click]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error changing to detail view"
    GoTo Exit_Properly

End Sub

Private Sub imgList_Click()

    lvFiles.View = lvwList
    
End Sub

Private Sub imgUp_Click()

    'simply set the fso to the parent folder of the current folder
    'use the current path to access the parent
    Dim oFolder As Folder
    Dim oParent As Folder
    Dim oItem As ListItems
    
    On Error GoTo imgUp_Click_Error
    
    If sCurrPath <> "" Then
        Set oFolder = fso.GetFolder(sCurrPath)
        Set oParent = oFolder.ParentFolder
        If Not oParent Is Nothing Then
            lvFiles.ListItems.Clear
            GetFolderFiles oParent.Path
            'and reset the current folder
            sCurrPath = oParent.Path
            Set imgPvw.Picture = ilPvw.ListImages(2).Picture
        End If
    End If
    
Exit_Properly:
    Set oFolder = Nothing
    Set oParent = Nothing
    Exit Sub
    
imgUp_Click_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "imgUp_Click]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error moving up one level"
    GoTo Exit_Properly
    
End Sub

Private Sub lvFiles_DblClick()

    On Error GoTo lvFiles_DblClick_Error
    
    'if the item is set...
    If Not oItem Is Nothing Then
        'and if it is a directory...
        If oItem.Icon = 1 Then  '1 is the index for the folder picture
            'clear the list
            lvFiles.ListItems.Clear
            'get the files and subfolders for this directory - path is stored in the tag
            GetFolderFiles oItem.Tag
            'set the mod level var
            sCurrPath = oItem.Tag
        End If
    End If
            
Exit_Properly:
    Exit Sub
    
lvFiles_DblClick_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "lvFiles_DblClick]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error changing directory"
    GoTo Exit_Properly

End Sub

Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Dim oSel As ListItem
    Dim lCnt As Long
    Dim sFiles As String
    Dim lNumSel As Long
    Dim bMult As Boolean
    
    On Error GoTo lvFiles_ItemClick_Error
    
    Set oItem = Item
    
    'clear the selected files array
    Erase arSelFiles
    
    lCnt = 0
    
    'retrieve the selected count
    lNumSel = GetNumSelected(lvFiles)
    'display appropriate graphic for selection
    If lNumSel > 1 Then
        Set imgPvw.Picture = ilPvw.ListImages(1).Picture
        bMult = True
    End If
    'now add this/these selected items to the file textbox
    For Each oSel In lvFiles.ListItems
        'but only if a valid file
        If oSel.Icon <> 1 Then
            If oSel.Selected Then
                'add this item to the array
                ReDim Preserve arSelFiles(lCnt)
                arSelFiles(lCnt) = oSel.Tag
                'and to the string
                sFiles = sFiles & """" & GetFileFromPath(oSel.Tag) & """ "
                lCnt = lCnt + 1
                'and load the graphic into the preview
                If Not bMult Then
                    imgPvw.Picture = LoadPicture(oSel.Tag)
                End If
            End If
        Else
            'must be a folder, is it selected
            If oSel.Selected Then
                'change the pic to display nothing selected
                Set imgPvw.Picture = ilPvw.ListImages(2).Picture
            End If
        End If
    Next oSel
    
    'and put in the box
    txtFilename.Text = sFiles
            
Exit_Properly:
    Set oSel = Nothing
    Exit Sub
    
lvFiles_ItemClick_Error:
    MsgBox "The following error has occured in " & MODULE_NAME & "lvFiles_ItemClick]:" & vbCr & vbCr & _
    "Number: " & Err.Number & vbCr & _
    "Source: " & Err.Source & vbCr & _
    "Description: " & Err.Description, vbCritical + vbOKOnly, "Error selecting file"
    GoTo Exit_Properly

End Sub

Private Function GetSelectedFiles() As String()

    'retrieve the selected items of the list view
    Dim arFiles() As String
    Dim oItem As ListItem
    Dim lIndex As Long
    
    On Error GoTo GetSelectedFiles_Error
    
    lIndex = 0
    
    For Each oItem In lvFiles.ListItems
        If oItem.Selected Then
            If oItem.Icon <> 1 Then
                ReDim Preserve arFiles(lIndex)
                arFiles(lIndex) = oItem.Tag
            End If
            lIndex = lIndex + 1
        End If
    Next oItem
    
    GetSelectedFiles = arFiles
    
Exit_Properly:
    Set oItem = Nothing
    Exit Function
    
GetSelectedFiles_Error:
    GoTo Exit_Properly

End Function

Public Sub SetPropRefNo(ByVal PropRefNo As Long)

    mlPropRefNo = PropRefNo
    
End Sub


