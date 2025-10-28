VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_CommodityQualityParameter 
   Caption         =   "Commodity Quality Parameter"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1335
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Visible         =   0   'False
      Width           =   15015
      Begin VB.ComboBox CmbSQualityParameterID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmMst_CommodityQualityParameter.frx":0000
         Left            =   6600
         List            =   "FrmMst_CommodityQualityParameter.frx":0002
         TabIndex        =   15
         Top             =   172
         Width           =   8175
      End
      Begin VB.ComboBox CmbSCommodity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Sorted          =   -1  'True
         TabIndex        =   14
         Top             =   172
         Width           =   3150
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13755
         TabIndex        =   17
         Top             =   600
         Width           =   1000
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   12750
         TabIndex        =   16
         Top             =   600
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5460
         Left            =   120
         TabIndex        =   19
         Top             =   1050
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   9631
         _Version        =   393216
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label4 
         Caption         =   "Quality Parameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   39
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label11 
         Caption         =   "Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   22
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label Label23 
         Caption         =   "Total Records Found :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   21
         Top             =   648
         Width           =   2415
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9120
         TabIndex        =   20
         Top             =   655
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1080
      Left            =   120
      TabIndex        =   35
      Top             =   3720
      Width           =   7800
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6135
         TabIndex        =   9
         Top             =   585
         Width           =   1500
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4635
         TabIndex        =   10
         Top             =   585
         Width           =   1500
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3135
         TabIndex        =   11
         Top             =   585
         Width           =   1500
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1635
         TabIndex        =   12
         Top             =   585
         Width           =   1500
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   13
         Top             =   585
         Width           =   1500
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6135
         TabIndex        =   8
         Top             =   210
         Width           =   1500
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4635
         TabIndex        =   7
         Top             =   210
         Width           =   1500
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3135
         TabIndex        =   6
         Top             =   210
         Width           =   1500
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1635
         TabIndex        =   5
         Top             =   210
         Width           =   1500
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   0
         Top             =   210
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   7815
      Begin VB.CommandButton CmdQualityParameterID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   38
         Top             =   1200
         Width           =   375
      End
      Begin VB.ComboBox CmbCommodity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   4710
      End
      Begin VB.CommandButton CmdCommodity 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   36
         Top             =   713
         Width           =   375
      End
      Begin VB.TextBox TxtMinValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox TxtMaxValue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TxtCommodityQualityParameterID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CmbQualityParameterID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmMst_CommodityQualityParameter.frx":0004
         Left            =   2400
         List            =   "FrmMst_CommodityQualityParameter.frx":0006
         TabIndex        =   2
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4695
      End
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3120
         Width           =   4695
      End
      Begin VB.TextBox TxtQualityParameterID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   37
         Top             =   720
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   765
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Commodity Quality Parameter ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Minimum Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Maximum Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   2220
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Quality Parameter"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1253
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Created  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   2700
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Updated :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   3180
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_CommodityQualityParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCommodityID, mQualityParameterID As Variant
Dim mSCommodityID, mSQualityParameterID As Variant
Dim Edit_Flg As Variant

Private Sub CmbQualityParameterID_GotFocus()
tmp = CmbQualityParameterID.Text
CmbQualityParameterID.Clear
Call TableMst_QualityParameter
If RsMst_QualityParameter.RecordCount = 0 Then
   MsgBox "Quality Parameter Not Found."
   Exit Sub
End If
For I = 1 To RsMst_QualityParameter.RecordCount
    CmbQualityParameterID.AddItem RsMst_QualityParameter!QualityParameter
    RsMst_QualityParameter.MoveNext
Next
CmbQualityParameterID.Text = tmp
End Sub

Private Sub CmbQualityParameterID_LostFocus()
If CmbQualityParameterID.Text = "" Then Exit Sub
RsMst_QualityParameter.MoveFirst
RsMst_QualityParameter.Find "QualityParameter='" & CmbQualityParameterID.Text & "'"
If RsMst_QualityParameter.EOF Then
   MsgBox "Invalid Selection."
   Exit Sub
End If
mQualityParameterID = RsMst_QualityParameter!QualityParameterID
End Sub

Private Sub CmdCommodity_Click()
FrmMstCommodity.Show
End Sub

Private Sub CmdQualityParameterID_Click()
FrmMstQualityParameter.Show
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtCommodityQualityParameterID.Text = ""
CmbCommodity.Visible = True
TxtCommodity.Visible = False
CmbCommodity.Text = ""
CmbQualityParameterID.Visible = True
TxtQualityParameterID.Visible = False
CmbQualityParameterID.Text = ""
TxtMinValue.Text = ""
TxtMaxValue.Text = ""
CmdCommodity.Visible = True
CmdQualityParameterID.Visible = True
TxtCreated = ""
TxtUpdated = ""
Frame1.Enabled = True
If RsMst_CommodityQualityParameter.RecordCount > 0 Then
   CmbCommodity.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub


Private Sub CmdGo_Click()
Dim wc As String
wc = ""
Call Grid_Head
If CmbSCommodity.Text <> "" Then
   wc = " Where A.CommodityID=" & mSCommodityID & ""
End If

If CmbSQualityParameterID.Text <> "" Then
   If wc = "" Then
      wc = " Where A.QualityParameterID=" & mSQualityParameterID & ""
   Else
      wc = wc & " And A.QualityParameterID=" & mSQualityParameterID & ""
   End If
End If

tmp = "Select A.CommodityQualityParameterID,MC.Commodity,MQ.QualityParameter,A.MinValue,A.MaxValue "
tmp = tmp & " From Mst_CommodityQualityParameter A"
tmp = tmp & " Inner Join Mst_Commodity MC On MC.CommodityID=A.CommodityID"
tmp = tmp & " Inner Join Mst_QualityParameter MQ On A.QualityParameterID = MQ.QualityParameterID"
tmp = tmp & wc
tmp = tmp & " Order By MC.Commodity,MQ.QualityParameter"
Call TmpTable

For I = 1 To TmpRs.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = TmpRs!CommodityQualityParameterID
    MSHFlexGrid1.TextMatrix(I, 1) = TmpRs!Commodity
    MSHFlexGrid1.TextMatrix(I, 2) = TmpRs!QualityParameter
    MSHFlexGrid1.TextMatrix(I, 3) = TmpRs!MinValue
    MSHFlexGrid1.TextMatrix(I, 4) = TmpRs!MaxValue
    TmpRs.MoveNext
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
Next
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   
   RsMst_CommodityQualityParameter.Delete
   RsMst_CommodityQualityParameter.Update
   If RsMst_CommodityQualityParameter.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_CommodityQualityParameter.MoveNext
   If RsMst_CommodityQualityParameter.EOF() Then
      RsMst_CommodityQualityParameter.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_CommodityQualityParameter.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
Frame1.Enabled = True
CmbCommodity.Visible = True
TxtCommodity.Visible = False
CmbCommodity.Text = TxtCommodity.Text

CmbQualityParameterID.Visible = True
TxtQualityParameterID.Visible = False
CmbQualityParameterID.Text = TxtQualityParameterID.Text

CmdCommodity.Visible = True
CmdQualityParameterID.Visible = True
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_CommodityQualityParameter.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
CmbSCommodity.Text = ""
CmbSQualityParameterID.Text = ""
Label22.Caption = ""
Call Grid_Head
Frame1.Enabled = True
Call TableMst_CommodityQualityParameter(" Order by CommodityQualityParameterID ")
If RsMst_CommodityQualityParameter.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()
Frame1.Enabled = False
TxtCommodityQualityParameterID.Text = ""
CmbCommodity.Visible = False
TxtCommodity.Visible = True
TxtCommodity.Text = ""

CmbQualityParameterID.Visible = False
TxtQualityParameterID.Visible = True
TxtQualityParameterID.Text = ""
TxtMinValue.Text = ""
TxtMaxValue.Text = ""
CmdCommodity.Visible = False
CmdQualityParameterID.Visible = False

TxtCommodityQualityParameterID.Text = RsMst_CommodityQualityParameter!CommodityQualityParameterID
mCommodityID = RsMst_CommodityQualityParameter!CommodityID
TxtCommodity.Text = GetCommodityName(RsMst_CommodityQualityParameter!CommodityID)
mQualityParameterID = RsMst_CommodityQualityParameter!QualityParameterID
TxtQualityParameterID = GetQualityParameter(RsMst_CommodityQualityParameter!QualityParameterID)
TxtMinValue.Text = RsMst_CommodityQualityParameter!MinValue
TxtMaxValue.Text = RsMst_CommodityQualityParameter!MaxValue
TxtCreated = IIf(IsNull(RsMst_CommodityQualityParameter!CreatedBy), "", RsMst_CommodityQualityParameter!CreatedBy) & " :- " & IIf(IsNull(RsMst_CommodityQualityParameter!CreatedDate), "", RsMst_CommodityQualityParameter!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_CommodityQualityParameter!UpdatedBy), "", RsMst_CommodityQualityParameter!UpdatedBy) & " :- " & IIf(IsNull(RsMst_CommodityQualityParameter!UpdatedDate), "", RsMst_CommodityQualityParameter!UpdatedDate)
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_CommodityQualityParameter.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
If mCaId = "" Then Exit Sub
RsMst_CommodityQualityParameter.MoveFirst

If RsMst_CommodityQualityParameter.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_CommodityQualityParameter.MoveFirst
RsMst_CommodityQualityParameter.Find "CommodityQualityParameterID = " & Val(mCaId) & ""
If RsMst_CommodityQualityParameter.EOF Then
   RsMst_CommodityQualityParameter.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_CommodityQualityParameter.MoveNext
If RsMst_CommodityQualityParameter.EOF Then
   RsMst_CommodityQualityParameter.MoveLast
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
   NextCmd.Enabled = False
   LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True

End Sub

Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_CommodityQualityParameter.MovePrevious
If RsMst_CommodityQualityParameter.BOF Then
   RsMst_CommodityQualityParameter.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
   PreviousCmd.Enabled = False
   FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SaveCmd_Click()
If SaveCmd.Enabled = False Then
   Exit Sub
End If

If CmbCommodity.Text = "" Then
   MsgBox "Blank Commodity Not Allowed !!"
   CmbCommodity.SetFocus
   Exit Sub
End If

If CmbQualityParameterID.Text = "" Then
   MsgBox "Blank QualityParameter  not Allowed"
   CmbQualityParameterID.SetFocus
   Exit Sub
End If

If TxtMinValue.Text = "" Then
   MsgBox "Blank Minimum Value Not Allowed !!"
   TxtMinValue.SetFocus
   Exit Sub
End If

If TxtMaxValue.Text = "" Then
   MsgBox "Blank Maximum Value Not Allowed !!"
   TxtMaxValue.SetFocus
   Exit Sub
End If


If Val(TxtMinValue.Text) > Val(TxtMaxValue.Text) Then
   MsgBox "Max Value Should Be Greater Than Min Value !!"
   TxtMinValue.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_CommodityQualityParameter.RecordCount = 0 Then
      tmp = 1
   Else
      RsMst_CommodityQualityParameter.MoveFirst
      RsMst_CommodityQualityParameter.Filter = "CommodityID=" & mCommodityID & " And  QualityParameterID = " & mQualityParameterID & ""
      If RsMst_CommodityQualityParameter.RecordCount > 0 Then
         MsgBox "Duplicate Commodity Quality Parameter not Allowed"
         RsMst_CommodityQualityParameter.Filter = ""
         Exit Sub
      End If
      RsMst_CommodityQualityParameter.Filter = ""
      
      tmp = "select max(CommodityQualityParameterID) from Mst_CommodityQualityParameter"
      Call TmpTable
      tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
   End If
   RsMst_CommodityQualityParameter.AddNew
   RsMst_CommodityQualityParameter!CommodityQualityParameterID = tmp
   TxtCommodityQualityParameterID = tmp
Else
   RsMst_CommodityQualityParameter.Filter = "CommodityID=" & mCommodityID & " And QualityParameterID = " & mQualityParameterID & ""
   If RsMst_CommodityQualityParameter.RecordCount > 0 Then
      If RsMst_CommodityQualityParameter!CommodityQualityParameterID <> TxtCommodityQualityParameterID Then
         MsgBox "Duplicate Commodity Quality Parameter not Allowed"
         RsMst_CommodityQualityParameter.Filter = ""
         Exit Sub
      End If
   End If
   RsMst_CommodityQualityParameter.Filter = ""
   RsMst_CommodityQualityParameter.MoveFirst
   RsMst_CommodityQualityParameter.Find "CommodityQualityParameterID = " & TxtCommodityQualityParameterID & ""
End If

RsMst_CommodityQualityParameter!CommodityID = mCommodityID
RsMst_CommodityQualityParameter!QualityParameterID = mQualityParameterID
RsMst_CommodityQualityParameter!MinValue = Val(TxtMinValue.Text)
RsMst_CommodityQualityParameter!MaxValue = Val(TxtMaxValue.Text)

If IsNull(RsMst_CommodityQualityParameter!CreatedBy) = True Or RsMst_CommodityQualityParameter!CreatedBy = "" Then
   RsMst_CommodityQualityParameter!CreatedBy = Gen_UserLoginID
   RsMst_CommodityQualityParameter!CreatedDate = Now
Else
   RsMst_CommodityQualityParameter!UpdatedBy = Gen_UserLoginID
   RsMst_CommodityQualityParameter!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If

RsMst_CommodityQualityParameter.Update

Call GButtonLock(Me, True)
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_CommodityQualityParameter.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_CommodityQualityParameter.EOF Then
      RsMst_CommodityQualityParameter.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 800
Frame0.Width = Me.Width - 300
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top

Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1200
RsMst_CommodityQualityParameter.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'CmbSCommodity.Text = ""
'CmbSQualityParameterID.Text = ""
'Label22.Caption = ""
'Call Grid_Head
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub


Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.Font.Size = 10
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.RowHeight(0) = 500
MSHFlexGrid1.TextMatrix(0, 0) = "Commodity Quality Parameter ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 2) = "Quality Parameter"
MSHFlexGrid1.TextMatrix(0, 3) = "Min Value"
MSHFlexGrid1.TextMatrix(0, 4) = "Max Value"
MSHFlexGrid1.ColWidth(0) = 1800
MSHFlexGrid1.ColWidth(1) = 3500
MSHFlexGrid1.ColWidth(2) = 5500
MSHFlexGrid1.ColWidth(3) = 1600
MSHFlexGrid1.ColWidth(4) = 1600
End Sub
Private Sub TxtQualityParameter_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtQualityParameter")
End If
End Sub

Private Sub TxtQualityParameter_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtQualityParameter)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtQualityParameter.SetFocus
End If
End Sub






Private Sub TxtMaxValue_LostFocus()
If TxtMaxValue.Text = "" Then Exit Sub

If IsNumeric(TxtMaxValue.Text) = False Or Val(TxtMaxValue.Text) < 0 Then
   MsgBox "Invalid Numeric Format."
   TxtMaxValue.SetFocus
End If
End Sub

Private Sub TxtMinValue_LostFocus()
If TxtMinValue.Text = "" Then Exit Sub

If IsNumeric(TxtMinValue.Text) = False Or Val(TxtMinValue.Text) < 0 Then
   MsgBox "Invalid Numeric Format."
   TxtMinValue.SetFocus
End If

End Sub

Private Sub TxtQualityParameterID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtQualityParameterID")
End If
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_CommodityQualityParameter.xls")
Call Xls_Detail_Rpt
End Sub


Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & "Mst_CommodityQualityParameter.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub CmbCommodity_GotFocus()
tmp = CmbCommodity.Text
Call TableMst_Commodity("")

CmbCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Commodity found"
   CmdCommodity.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodity.Text = tmp
End Sub

Private Sub CmbCommodity_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "CmbCommodity")
End If
End Sub

Private Sub CmbCommodity_LostFocus()
If CmbCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodity.Text & "'"

If RsMst_Commodity.EOF Then
   MsgBox "Invalid Commodity selection "
   CmbCommodity.SetFocus
   Exit Sub
End If
mCommodityID = RsMst_Commodity.Fields(0)
End Sub
Private Sub CmbSQualityParameterID_GotFocus()
tmp = CmbSQualityParameterID.Text
CmbSQualityParameterID.Clear
Call TableMst_QualityParameter
If RsMst_QualityParameter.RecordCount = 0 Then
   MsgBox "Quality Parameter Not Found."
   Exit Sub
End If
For I = 1 To RsMst_QualityParameter.RecordCount
    CmbSQualityParameterID.AddItem RsMst_QualityParameter!QualityParameter
    RsMst_QualityParameter.MoveNext
Next
CmbSQualityParameterID.Text = tmp
End Sub

Private Sub CmbSQualityParameterID_LostFocus()
If CmbSQualityParameterID.Text = "" Then Exit Sub
RsMst_QualityParameter.MoveFirst
RsMst_QualityParameter.Find "QualityParameter='" & CmbSQualityParameterID.Text & "'"
If RsMst_QualityParameter.EOF Then
   MsgBox "Invalid Selection."
   Exit Sub
End If
mSQualityParameterID = RsMst_QualityParameter!QualityParameterID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity("")
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Commodity found"
   CmdSCommodity.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbSCommodity.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"

If RsMst_Commodity.EOF Then
   MsgBox "Invalid Commodity selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_Commodity.Fields(0)
End Sub

