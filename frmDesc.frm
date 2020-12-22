VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDesc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picture Descriptions"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDesc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   4335
      TabIndex        =   3
      Top             =   3195
      Width           =   3495
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   360
      Left            =   6630
      TabIndex        =   2
      Top             =   3720
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   360
      Left            =   5310
      TabIndex        =   1
      Top             =   3720
      Width           =   1200
   End
   Begin MSComctlLib.ImageList ilPics 
      Left            =   7050
      Top             =   2700
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvDesc 
      Height          =   2865
      Left            =   165
      TabIndex        =   0
      Top             =   675
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   5054
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Selected files:"
      Height          =   270
      Left            =   180
      TabIndex        =   6
      Top             =   360
      Width           =   2385
   End
   Begin VB.Label Label2 
      Caption         =   "Preview:"
      Height          =   285
      Left            =   4350
      TabIndex        =   5
      Top             =   375
      Width           =   1560
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   255
      Left            =   4335
      TabIndex        =   4
      Top             =   2940
      Width           =   1605
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2085
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   675
      Width           =   3510
   End
End
Attribute VB_Name = "frmDesc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private marDesc() As String
Private marFilelist() As String
Private mbCancelled As Boolean
Private mbChanging As Boolean

Private Const MODULE_NAME As String = "[FileAX:frmDesc."

Private Sub cmdCancel_Click()

    mbCancelled = True
    
    Me.Hide
    
End Sub

Public Property Get Descriptions() As String()

    Descriptions = marDesc
    
End Property

Public Property Get FileList() As String()

    FileList = marFilelist
    
End Property

Public Property Let FileList(ByRef NewFileList() As String)

    Dim lUList As Long
    
    marFilelist = NewFileList
    
    'and set up the file descriptions array
    lUList = SUBound(NewFileList)
    
    ReDim marDesc(lUList)
    
End Property

Private Sub cmdOK_Click()

    mbCancelled = False
    Me.Hide
    
End Sub

Private Sub Form_Load()

    Dim lUfiles As Long
    Dim lCnt As Long
    Dim oItem As ListItem
    
    'load the selected pictures into the imagelist
    lUfiles = SUBound(marFilelist)
    
    For lCnt = 0 To lUfiles
        ilPics.ListImages.Add , , LoadPicture(marFilelist(lCnt))
    Next lCnt
        
    'load the passed selected files into the listview
    For lCnt = 0 To lUfiles
        Set oItem = lvDesc.ListItems.Add(, , marFilelist(lCnt))
    Next lCnt

    'and select the first image
    lvDesc.ListItems(1).Selected = True
    Set Image1.Picture = ilPics.ListImages(1).Picture
    
End Sub

Private Sub lvDesc_ItemClick(ByVal Item As MSComctlLib.ListItem)

    'set the preview image
    Set Image1.Picture = ilPics.ListImages(Item.Index).Picture
    
    mbChanging = True
    txtDescription.Text = marDesc(lvDesc.SelectedItem.Index - 1)
    mbChanging = False
    
    txtDescription.SetFocus
    
End Sub

Public Property Get Cancelled() As Variant

    Cancelled = mbCancelled
    
End Property

Private Sub txtDescription_Change()

    If Not mbChanging Then
        marDesc(lvDesc.SelectedItem.Index - 1) = txtDescription.Text
    End If
    
End Sub
