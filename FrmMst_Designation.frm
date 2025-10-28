VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Designation 
   Caption         =   "Designation"
   ClientHeight    =   2580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2580
   ScaleWidth      =   6555
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   5100
      Visible         =   0   'False
      Width           =   15060
      Begin VB.ComboBox CmbSReportingDesignationID 
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
         Left            =   6540
         Sorted          =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   3225
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
         Left            =   10830
         TabIndex        =   32
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchDesignation 
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
         Left            =   1305
         TabIndex        =   31
         Top             =   240
         Width           =   3045
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
         Height          =   345
         Left            =   9810
         TabIndex        =   30
         Top             =   240
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label5 
         Caption         =   "Reporting Designation"
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
         Left            =   4440
         TabIndex        =   37
         Top             =   300
         Width           =   2040
      End
      Begin VB.Label Label19 
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   35
         Top             =   285
         Width           =   1215
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
         Left            =   11880
         TabIndex        =   34
         Top             =   285
         Width           =   2355
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   14280
         TabIndex        =   33
         Top             =   285
         Width           =   90
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   4545
      Width           =   7680
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
         Left            =   6060
         TabIndex        =   8
         Top             =   615
         Width           =   1455
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
         Left            =   4605
         TabIndex        =   22
         Top             =   615
         Width           =   1455
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
         Left            =   3030
         TabIndex        =   21
         Top             =   615
         Width           =   1575
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
         Left            =   1575
         TabIndex        =   20
         Top             =   615
         Width           =   1455
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
         Left            =   120
         TabIndex        =   19
         Top             =   615
         Width           =   1455
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
         Left            =   6060
         TabIndex        =   9
         Top             =   240
         Width           =   1455
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
         Left            =   4605
         TabIndex        =   18
         Top             =   240
         Width           =   1455
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
         Left            =   3030
         TabIndex        =   17
         Top             =   240
         Width           =   1575
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
         Left            =   1575
         TabIndex        =   7
         Top             =   240
         Width           =   1455
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4545
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   7695
      Begin VB.ComboBox CmbSecReportingDesignationID 
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
         ItemData        =   "FrmMst_Designation.frx":0000
         Left            =   2400
         List            =   "FrmMst_Designation.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   4860
      End
      Begin VB.ComboBox CmbMISFlag 
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
         ItemData        =   "FrmMst_Designation.frx":0004
         Left            =   2400
         List            =   "FrmMst_Designation.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   3090
         Width           =   1455
      End
      Begin VB.ComboBox CmbDirectLink 
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
         ItemData        =   "FrmMst_Designation.frx":001B
         Left            =   2400
         List            =   "FrmMst_Designation.frx":0025
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   2625
         Width           =   4860
      End
      Begin VB.CheckBox ChkSelfReporting 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   1200
         Width           =   240
      End
      Begin VB.ComboBox CmbReportingDesignationID 
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
         ItemData        =   "FrmMst_Designation.frx":0032
         Left            =   2400
         List            =   "FrmMst_Designation.frx":0034
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1680
         Width           =   4860
      End
      Begin VB.TextBox TxtReportingDesignationID 
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
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4815
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3570
         Width           =   4815
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4050
         Width           =   4815
      End
      Begin VB.TextBox TxtDesignation 
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
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox TxtDesignationID 
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
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtDierctLink 
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
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2625
         Width           =   4815
      End
      Begin VB.TextBox TxtMISFlag 
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
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   3105
         Width           =   1410
      End
      Begin VB.TextBox TxtSecReportingDesignationID 
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
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4815
      End
      Begin VB.Label Label8 
         Caption         =   "2nd Reporting Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   225
         TabIndex        =   43
         Top             =   2130
         Width           =   2070
      End
      Begin VB.Label Label7 
         Caption         =   "MIS Flag"
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
         Left            =   240
         TabIndex        =   41
         Top             =   3150
         Width           =   1350
      End
      Begin VB.Label Label6 
         Caption         =   "Direct Link"
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
         Left            =   240
         TabIndex        =   39
         Top             =   2685
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Self Reporting"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   29
         Top             =   1252
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "1st Reporting Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   225
         TabIndex        =   28
         Top             =   1605
         Width           =   2070
      End
      Begin VB.Label Label10 
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
         TabIndex        =   26
         Top             =   3630
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Height          =   270
         Left            =   240
         TabIndex        =   25
         Top             =   4095
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Designation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   14
         Top             =   772
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Designation ID"
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
         TabIndex        =   13
         Top             =   285
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_Designation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mRepostingDesignationID, mDirectLink, mMISFlag, mSecRepostingDesignationID As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtDesignation.Locked = False
TxtDesignation = ""
TxtDesignationID = ""

ChkSelfReporting.Enabled = True

CmbReportingDesignationID.Visible = True
TxtReportingDesignationID.Visible = False
CmbReportingDesignationID.Text = ""
mRepostingDesignationID = Null

CmbSecReportingDesignationID.Visible = True
TxtSecReportingDesignationID.Visible = False
CmbSecReportingDesignationID.Text = ""
mSecRepostingDesignationID = Null


CmbDirectLink.Visible = True
TxtDierctLink.Visible = False
CmbDirectLink.Text = ""
mDirectLink = Null

CmbMISFlag.Visible = True
TxtMISFlag.Visible = False
CmbMISFlag.Text = ""
mMISFlag = Null

Call ChkSelfReporting_Click

TxtCreated = ""
TxtUpdated = ""

If RsMst_Designation.RecordCount > 0 Then
   TxtDesignation.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub
Private Sub ChkSelfReporting_Click()
If ChkSelfReporting.Value = 0 Then
   CmbReportingDesignationID.Text = ""
   CmbReportingDesignationID.Enabled = True
   mRepostingDesignationID = Null
   CmbSecReportingDesignationID.Text = ""
   CmbSecReportingDesignationID.Enabled = True
   mSecRepostingDesignationID = Null
ElseIf ChkSelfReporting.Value = 1 Then
   CmbReportingDesignationID.Text = ""
   CmbReportingDesignationID.Enabled = False
   mRepostingDesignationID = Null
   CmbSecReportingDesignationID.Text = ""
   CmbSecReportingDesignationID.Enabled = False
   mSecRepostingDesignationID = Null
End If
End Sub

Private Sub CmbDirectLink_LostFocus()
If CmbDirectLink.Text = "" Then Exit Sub

If LCase(CmbDirectLink.Text) <> "no" And LCase(CmbDirectLink.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbDirectLink.SetFocus
   Exit Sub
End If

mDirectLink = Left(CmbDirectLink, 1)

End Sub

Private Sub CmbMISFlag_LostFocus()
If CmbMISFlag.Text = "" Then Exit Sub

If LCase(CmbMISFlag.Text) <> "no" And LCase(CmbMISFlag.Text) <> "yes" Then
   MsgBox "Invalid Selection!!!"
   CmbMISFlag.SetFocus
   Exit Sub
End If

mMISFlag = Left(CmbMISFlag, 1)
End Sub

Private Sub CmbSReportingDesignationID_GotFocus()
tmp1 = CmbSReportingDesignationID.Text
'Call TableMst_Designation("")

tmp = "Select * from Mst_Designation"
Call TmpTable
CmbSReportingDesignationID.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbSReportingDesignationID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbSReportingDesignationID.AddItem TmpRs!Designation
    TmpRs.MoveNext
Next
CmbSReportingDesignationID.Text = tmp1
End Sub

Private Sub CmbSReportingDesignationID_LostFocus()
If CmbSReportingDesignationID.Text = "" Then
   Exit Sub
End If

TmpRs.MoveFirst
TmpRs.Find "Designation = '" & CmbSReportingDesignationID.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSReportingDesignationID.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Designation.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Designation.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 2
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


Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   'tmp = "Select * from Mst_Employee Where (DesignationID = " & TxtDesignationID & ")"
   'Call Tmp1Table
   'If TmpRs1.RecordCount > 0 Then
   '   MsgBox "Child Record Present!!!"
   '   Exit Sub
   'End If
   RsMst_Designation.Delete
   RsMst_Designation.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_Designation.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Designation.MoveNext
   If RsMst_Designation.EOF() Then
      RsMst_Designation.MoveLast
   End If
   Call Indata
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_Designation.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtDesignation.Locked = False
ChkSelfReporting.Enabled = True

CmbReportingDesignationID.Visible = True
TxtReportingDesignationID.Visible = False
CmbReportingDesignationID.Text = TxtReportingDesignationID

CmbSecReportingDesignationID.Visible = True
TxtSecReportingDesignationID.Visible = False
CmbSecReportingDesignationID.Text = TxtSecReportingDesignationID

CmbDirectLink.Visible = True
TxtDierctLink.Visible = False
CmbDirectLink.Text = TxtDierctLink

CmbMISFlag.Visible = True
TxtMISFlag.Visible = False
CmbMISFlag.Text = TxtMISFlag


End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Designation.MoveFirst
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
   MsgBox "Access Denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
TxtSearchDesignation = ""
CmbSReportingDesignationID.Text = ""
Label22.Caption = ""
Call TableMst_Designation(" Order by DesignationID ")

If RsMst_Designation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()

TxtDesignationID = RsMst_Designation!DesignationID
TxtDesignation = RsMst_Designation!Designation

TxtDesignation.Locked = True

CmbReportingDesignationID.Visible = False
TxtReportingDesignationID.Visible = True

CmbSecReportingDesignationID.Visible = False
TxtSecReportingDesignationID.Visible = True

ChkSelfReporting.Enabled = False
ChkSelfReporting.Value = IIf(IsNull(RsMst_Designation!SelfReporting), 0, RsMst_Designation!SelfReporting)
mRepostingDesignationID = RsMst_Designation!ReportingDesignationID
If IsNull(mRepostingDesignationID) = False Then
   tmp = "Select Designation from Mst_Designation Where DesignationID = " & mRepostingDesignationID & " "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtReportingDesignationID = TmpRs!Designation
   End If
Else
   TxtReportingDesignationID = ""
End If

mSecRepostingDesignationID = RsMst_Designation!SecReportingDesignationID
If IsNull(mSecRepostingDesignationID) = False Then
   tmp = "Select Designation from Mst_Designation Where DesignationID = " & mSecRepostingDesignationID & " "
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      TxtSecReportingDesignationID = TmpRs!Designation
   End If
Else
   TxtSecReportingDesignationID = ""
End If


mDirectLink = IIf(IsNull(RsMst_Designation!DirectLink), "", RsMst_Designation!DirectLink)

mMISFlag = IIf(IsNull(RsMst_Designation!MISFlag), "", RsMst_Designation!MISFlag)

CmbDirectLink.Visible = False
TxtDierctLink.Visible = True

CmbMISFlag.Visible = False
TxtMISFlag.Visible = True

If LCase(mDirectLink) = "n" Then
   TxtDierctLink = "No"
ElseIf LCase(mDirectLink) = "y" Then
   TxtDierctLink = "Yes"
Else
   TxtDierctLink = ""
End If

If LCase(mMISFlag) = "n" Then
   TxtMISFlag = "No"
ElseIf LCase(mMISFlag) = "y" Then
   TxtMISFlag = "Yes"
Else
   TxtMISFlag = ""
End If

TxtCreated = IIf(IsNull(RsMst_Designation!CreatedBy), "", RsMst_Designation!CreatedBy) & " :- " & IIf(IsNull(RsMst_Designation!CreatedDate), "", RsMst_Designation!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Designation!UpdatedBy), "", RsMst_Designation!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Designation!UpdatedDate), "", RsMst_Designation!UpdatedDate)

Call GButtonLock(Me, True)

End Sub
Private Sub LastCmd_Click()
RsMst_Designation.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_Designation.MoveFirst

If RsMst_Designation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Designation.MoveFirst
RsMst_Designation.Find "DesignationID = " & Val(mCaId) & ""
If RsMst_Designation.EOF Then
   RsMst_Designation.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Designation.MoveNext
If RsMst_Designation.EOF Then
   RsMst_Designation.MoveLast
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
RsMst_Designation.MovePrevious
If RsMst_Designation.BOF Then
   RsMst_Designation.MoveFirst
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

If TxtDesignation = "" Then
   MsgBox "Blank Designation  not Allowed"
   TxtDesignation.SetFocus
   Exit Sub
End If

If ChkSelfReporting.Value = 0 Then
   If CmbReportingDesignationID.Text = "" Then
      MsgBox "Blank Reporting Designation  not allowed !!!"
      CmbReportingDesignationID.SetFocus
      Exit Sub
   End If
End If

If CmbDirectLink.Text = "" Then
   MsgBox "Blank Direct Link not allowed !!!"
   CmbDirectLink.SetFocus
   Exit Sub
End If

If CmbMISFlag.Text = "" Then
   MsgBox "Blank MIS Flag not allowed !!!"
   CmbMISFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Designation.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_Designation.MoveFirst
        RsMst_Designation.Find "Designation = '" & TxtDesignation & "'"
        If Not RsMst_Designation.EOF Then
           MsgBox "Duplicate Designation not Allowed"
           Exit Sub
        End If
        RsMst_Designation.MoveLast
        
        tmp = "select max(DesignationID) from Mst_Designation"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_Designation.AddNew
    RsMst_Designation!DesignationID = tmp
    TxtDesignationID = tmp
Else
    RsMst_Designation.MoveFirst
    RsMst_Designation.Find "Designation = '" & TxtDesignation & "'"
    If Not RsMst_Designation.EOF Then
       If RsMst_Designation!DesignationID <> TxtDesignationID Then
          MsgBox "Duplicate Designation not Allowed"
          Exit Sub
       End If
    End If
   RsMst_Designation.MoveFirst
   RsMst_Designation.Find "DesignationID = " & TxtDesignationID & ""
End If

RsMst_Designation!Designation = TxtDesignation
RsMst_Designation!SelfReporting = ChkSelfReporting.Value
RsMst_Designation!ReportingDesignationID = mRepostingDesignationID
RsMst_Designation!SecReportingDesignationID = mSecRepostingDesignationID
RsMst_Designation!DirectLink = mDirectLink
RsMst_Designation!MISFlag = mMISFlag

If IsNull(RsMst_Designation!CreatedBy) = True Or RsMst_Designation!CreatedBy = "" Then
   RsMst_Designation!CreatedBy = Gen_UserLoginID
   RsMst_Designation!CreatedDate = Now
Else
   RsMst_Designation!UpdatedBy = Gen_UserLoginID
   RsMst_Designation!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If

RsMst_Designation.Update
RsMst_Designation.Requery
RsMst_Designation.MoveFirst
RsMst_Designation.Find "DesignationID = " & TxtDesignationID & ""

Call GButtonLock(Me, True)
TxtDesignation.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_Designation.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_Designation.EOF Then
      RsMst_Designation.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top

Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 900

RsMst_Designation.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head
'TxtSearchDesignation = ""
'CmbSReportingDesignationID.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7

MSHFlexGrid1.TextMatrix(0, 0) = "Designation Id"
MSHFlexGrid1.TextMatrix(0, 1) = "Designation"
MSHFlexGrid1.TextMatrix(0, 2) = "Self Reporting"
MSHFlexGrid1.TextMatrix(0, 3) = "Reporting Designation"
MSHFlexGrid1.TextMatrix(0, 4) = "Second Reporting Designation"
MSHFlexGrid1.TextMatrix(0, 5) = "Direct Link"
MSHFlexGrid1.TextMatrix(0, 6) = "MIS Flag"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 3500
MSHFlexGrid1.ColWidth(2) = 1400
MSHFlexGrid1.ColWidth(3) = 3200
MSHFlexGrid1.ColWidth(4) = 3200
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000

End Sub
Private Sub TxtDesignation_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtDesignation")
End If
End Sub

Private Sub TxtDesignation_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtDesignation)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtDesignation.SetFocus
End If
End Sub


Private Sub TxtDesignationID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtDesignationID")
End If
End Sub

Private Sub CmbReportingDesignationID_GotFocus()
tmp1 = CmbReportingDesignationID.Text
'Call TableMst_Designation("")

tmp = "Select * from Mst_Designation"
Call TmpTable
CmbReportingDesignationID.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbReportingDesignationID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbReportingDesignationID.AddItem TmpRs!Designation
    TmpRs.MoveNext
Next
CmbReportingDesignationID.Text = tmp1
End Sub
Private Sub CmbReportingDesignationID_LostFocus()

If CmbReportingDesignationID.Text = "" Then
   mRepostingDesignationID = Null
   Exit Sub
End If

TmpRs.MoveFirst
TmpRs.Find "Designation = '" & CmbReportingDesignationID.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbReportingDesignationID.SetFocus
   Exit Sub
End If
mRepostingDesignationID = TmpRs!DesignationID

End Sub

Private Sub CmdGo_Click()

Call Grid_Head

tmp = "Select MD.DesignationID,MD.Designation,MD.SelfReporting,MRD.Designation,MRD1.Designation,MD.DirectLink,MD.MISFlag " & _
      " From Mst_Designation MD " & _
      " Left Join Mst_Designation MRD On MD.ReportingDesignationID = MRD.DesignationID " & _
      " Left Join Mst_Designation MRD1 On MD.SecReportingDesignationID = MRD1.DesignationID "
      
Dim wc As Variant
wc = ""
If TxtSearchDesignation <> "" Then
   wc = " Where MD.Designation Like '%" & TxtSearchDesignation & "%' "
End If

If CmbSReportingDesignationID.Text <> "" Then
   If TxtSearchDesignation = "" Then
      wc = " Where MRD.Designation = '" & CmbSReportingDesignationID.Text & "'"
   Else
      wc = wc & " And MRD.Designation = '" & CmbSReportingDesignationID.Text & "'"
   End If
End If
tmp = tmp & wc & " Order by MD.Designation"

Call TmpTable

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If MSHFlexGrid1.TextMatrix(iu, 2) = 0 Then
       MSHFlexGrid1.TextMatrix(iu, 2) = "No"
    Else
       MSHFlexGrid1.TextMatrix(iu, 2) = "Yes"
    End If
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 5)) = "n" Then
       MSHFlexGrid1.TextMatrix(iu, 5) = "No"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 5)) = "y" Then
       MSHFlexGrid1.TextMatrix(iu, 5) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 5) = ""
    End If
    
    If LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "n" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "No"
    ElseIf LCase(MSHFlexGrid1.TextMatrix(iu, 6)) = "y" Then
       MSHFlexGrid1.TextMatrix(iu, 6) = "Yes"
    Else
       MSHFlexGrid1.TextMatrix(iu, 6) = ""
    End If
    
    
    TmpRs.MoveNext
Next

Label22.Refresh
Label22.Caption = TmpRs.RecordCount



End Sub

Private Sub CmbSecReportingDesignationID_GotFocus()
tmp1 = CmbSecReportingDesignationID.Text
'Call TableMst_Designation("")

tmp = "Select * from Mst_Designation"
Call TmpTable
CmbSecReportingDesignationID.Clear
If TmpRs.RecordCount = 0 Then
   MsgBox "No Designation found"
   CmbSecReportingDesignationID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRs.RecordCount
    CmbSecReportingDesignationID.AddItem TmpRs!Designation
    TmpRs.MoveNext
Next
CmbSecReportingDesignationID.Text = tmp1
End Sub
Private Sub CmbSecReportingDesignationID_LostFocus()

If CmbSecReportingDesignationID.Text = "" Then
   mSecRepostingDesignationID = Null
   Exit Sub
End If

TmpRs.MoveFirst
TmpRs.Find "Designation = '" & CmbSecReportingDesignationID.Text & "'"

If TmpRs.EOF Then
   MsgBox "Invalid Designation Selection "
   CmbSecReportingDesignationID.SetFocus
   Exit Sub
End If
mSecRepostingDesignationID = TmpRs!DesignationID

End Sub
