VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_StateIndia 
   Caption         =   "State India"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   1845
      Left            =   120
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   15000
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
         Left            =   10215
         TabIndex        =   14
         Top             =   252
         Width           =   1000
      End
      Begin VB.TextBox TxtSearchState 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   7125
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
         Left            =   9210
         TabIndex        =   12
         Top             =   255
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   975
         Left            =   75
         TabIndex        =   15
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1720
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label19 
         Caption         =   "State India  Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   330
         Width           =   1680
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
         Left            =   11325
         TabIndex        =   17
         Top             =   300
         Width           =   2355
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
         Left            =   13725
         TabIndex        =   16
         Top             =   300
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   28
      Top             =   5880
      Width           =   7200
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
         Height          =   400
         Left            =   1500
         TabIndex        =   9
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   105
         TabIndex        =   0
         Top             =   240
         Width           =   1400
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
         Height          =   400
         Left            =   105
         TabIndex        =   35
         Top             =   645
         Width           =   1400
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
         Height          =   400
         Left            =   1500
         TabIndex        =   34
         Top             =   645
         Width           =   1400
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
         Height          =   400
         Left            =   2895
         TabIndex        =   33
         Top             =   240
         Width           =   1400
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
         Height          =   405
         Left            =   2895
         TabIndex        =   32
         Top             =   645
         Width           =   1400
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
         Height          =   400
         Left            =   4290
         TabIndex        =   31
         Top             =   240
         Width           =   1400
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
         Height          =   400
         Left            =   4290
         TabIndex        =   30
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   5685
         TabIndex        =   29
         Top             =   240
         Width           =   1400
      End
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
         Height          =   400
         Left            =   5685
         TabIndex        =   10
         Top             =   645
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   19
      Top             =   0
      Width           =   7200
      Begin VB.TextBox TxtStateIndiaCode 
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
         Left            =   5445
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1635
         Width           =   1635
      End
      Begin VB.TextBox TxtTINNo 
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
         Left            =   2085
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1635
         Width           =   1275
      End
      Begin VB.TextBox TxtStateIndiaOfficeAddress 
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
         Height          =   1800
         Left            =   2085
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3000
         Width           =   4995
      End
      Begin VB.TextBox TxtNAVCode 
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
         Left            =   2085
         MaxLength       =   25
         TabIndex        =   7
         Top             =   2565
         Width           =   4995
      End
      Begin VB.TextBox TxtGSTIN 
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
         Left            =   2085
         MaxLength       =   15
         TabIndex        =   6
         Top             =   2115
         Width           =   4995
      End
      Begin VB.TextBox TxtStateIndiaName 
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
         Left            =   2085
         MaxLength       =   100
         TabIndex        =   1
         Top             =   708
         Width           =   4995
      End
      Begin VB.TextBox TxtStateIndiaID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   1275
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   5370
         Width           =   4995
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
         Left            =   2085
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4890
         Width           =   5010
      End
      Begin VB.ComboBox CmbFlag 
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
         ItemData        =   "FrmMst_StateIndia.frx":0000
         Left            =   2085
         List            =   "FrmMst_StateIndia.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1176
         Width           =   4995
      End
      Begin VB.TextBox TxtFlag 
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
         Left            =   2085
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1176
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Label Label5 
         Caption         =   "State NAV Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3480
         TabIndex        =   40
         Top             =   1710
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "State TIN Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   39
         Top             =   1710
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "NCML - State Office Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   38
         Top             =   3675
         Width           =   1800
      End
      Begin VB.Label Label2 
         Caption         =   "NAV Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   37
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "NCML - GST No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   36
         Top             =   2190
         Width           =   1920
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   1206
         Width           =   600
      End
      Begin VB.Label LblStateName 
         Caption         =   "State India Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   26
         Top             =   750
         Width           =   1800
      End
      Begin VB.Label LblStateID 
         Caption         =   "State India ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   308
         Width           =   1800
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
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   5430
         Width           =   975
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
         Height          =   240
         Left            =   120
         TabIndex        =   23
         Top             =   4950
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmMst_StateIndia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mGSTINNo As Variant

Private Sub CmbFlag_LostFocus()
If LCase(CmbFlag.Text) = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
Call Grid_Head

tmp = "Select StateIndiaId,StateIndiaName,Flag,TINNo,StateIndiaCode,GSTIN,NAVCode,StateIndiaOfficeAddress from Mst_StateIndia "
wc = ""
If TxtSearchState <> "" Then
   wc = " Where StateIndiaName Like '%" & Trim(TxtSearchState) & "%' "
End If

tmp = tmp & wc & " Order By StateIndiaName "
Call TmpTable


For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head
TxtSearchState = ""
Label22.Caption = ""

Call TableMst_StateIndia
If RsMst_StateIndia.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_StateIndia.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook

Dim oXL As Object
Dim oWB As Object


Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
 
'Set oXL = New Excel.Application
Set oXL = CreateObject("Excel.Application")
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_StateIndia.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_StateIndia.MoveNext
If RsMst_StateIndia.EOF Then
   RsMst_StateIndia.MoveLast
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
RsMst_StateIndia.MovePrevious
If RsMst_StateIndia.BOF Then
   RsMst_StateIndia.MoveFirst
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
Private Sub FirstCmd_Click()
RsMst_StateIndia.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_StateIndia.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()


If SearchCmd.Caption = "Cancel" Then
   If RsMst_StateIndia.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_StateIndia.EOF Then
      RsMst_StateIndia.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Width = Me.Width - 300
Frame0.Height = Me.Height - 550
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 800
RsMst_StateIndia.MoveFirst
''============comment on 7 july 2012 for not clearing search grid==============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'TxtSearchState = ""
'Label22.Caption = ""
''============comment on 7 july 2012 for not clearing search grid==============

'MSHFlexGrid1.Rows = RsMst_StateIndia.RecordCount + 1
''MSHFlexGrid1.ColWidth(0) = 1000
'For i = 1 To RsMst_StateIndia.RecordCount
'    For c = 0 To MSHFlexGrid1.Cols - 1
'       MSHFlexGrid1.TextMatrix(i, c) = IIf(IsNull(RsMst_StateIndia.Fields(c)), "", RsMst_StateIndia.Fields(c))
'    Next
'    RsMst_StateIndia.MoveNext
'Next
End Sub
Public Sub Indata()

TxtStateIndiaID.Locked = True
TxtStateIndiaName.Locked = True
CmbFlag.Visible = False
TxtFlag.Visible = True
''-- 05 July 2017 Ullhas Patil
TxtTINNo.Locked = True
TxtStateIndiaCode.Locked = True
TxtGSTIN.Locked = True
TxtNAVCode.Locked = True
TxtStateIndiaOfficeAddress.Locked = True
''---------------------------------

TxtStateIndiaID = IIf(IsNull(RsMst_StateIndia!StateIndiaID), "", RsMst_StateIndia!StateIndiaID)
TxtStateIndiaName = IIf(IsNull(RsMst_StateIndia!StateIndiaName), "", RsMst_StateIndia!StateIndiaName)
TxtCreated = IIf(IsNull(RsMst_StateIndia!CreatedBy), "", RsMst_StateIndia!CreatedBy) & " :- " & IIf(IsNull(RsMst_StateIndia!CreatedDate), "", RsMst_StateIndia!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_StateIndia!UpdatedBy), "", RsMst_StateIndia!UpdatedBy) & " :- " & IIf(IsNull(RsMst_StateIndia!UpdatedDate), "", RsMst_StateIndia!UpdatedDate)
If LCase(RsMst_StateIndia!Flag) = "a" Then
   TxtFlag.Text = "Active"
ElseIf LCase(RsMst_StateIndia!Flag) = "d" Then
   TxtFlag.Text = "De-Active"
Else
   TxtFlag.Text = ""
End If
'TxtFlag = IIf(IsNull(RsMst_StateIndia!Flag), "", RsMst_StateIndia!Flag)
''-- 05 July 2017 Ullhas Patil
TxtTINNo = IIf(IsNull(RsMst_StateIndia!TINNo), "", RsMst_StateIndia!TINNo)
TxtStateIndiaCode = IIf(IsNull(RsMst_StateIndia!StateIndiaCode), "", RsMst_StateIndia!StateIndiaCode)
TxtGSTIN = IIf(IsNull(RsMst_StateIndia!GSTIN), "", RsMst_StateIndia!GSTIN)
TxtNAVCode = IIf(IsNull(RsMst_StateIndia!NavCode), "", RsMst_StateIndia!NavCode)
TxtStateIndiaOfficeAddress = IIf(IsNull(RsMst_StateIndia!StateIndiaOfficeAddress), "", RsMst_StateIndia!StateIndiaOfficeAddress)
''---------------------------------
If RsMst_StateIndia.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub


Private Sub TxtFlag_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtFlag")
End If
End Sub


Private Sub TxtStateIndiaID_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
   Call AssignHelpData(Me, "TxtStateIndiaID")
End If

End Sub

Private Sub TxtStateIndiaID_LostFocus()
If TxtStateIndiaID = "" Then Exit Sub
If IsNumeric(TxtStateIndiaID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtStateIndiaID.SetFocus
   Exit Sub
End If
TxtStateIndiaID = Val(TxtStateIndiaID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtStateIndiaID.Locked = True
TxtStateIndiaName.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = ""
TxtStateIndiaID = ""
TxtStateIndiaName = ""
TxtCreated = ""
TxtUpdated = ""
''-- 05 July 2017 Ullhas Patil
TxtTINNo.Locked = False
TxtStateIndiaCode.Locked = False
TxtGSTIN.Locked = False
TxtNAVCode.Locked = False
TxtStateIndiaOfficeAddress.Locked = False
''---------------------------------
If RsMst_StateIndia.RecordCount > 0 Then
   TxtStateIndiaName.SetFocus
End If
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtStateIndiaID.Locked = True
TxtStateIndiaName.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag.Text
''-- 05 July 2017 Ullhas Patil
TxtTINNo.Locked = False
TxtStateIndiaCode.Locked = False
TxtGSTIN.Locked = False
TxtNAVCode.Locked = False
TxtStateIndiaOfficeAddress.Locked = False
''---------------------------------
End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_StateIndia.Delete
   RsMst_StateIndia.Update
   If RsMst_StateIndia.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_StateIndia.MoveNext
   If RsMst_StateIndia.EOF() Then
      RsMst_StateIndia.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_StateIndia.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtStateIndiaName.Text = "" Then
   MsgBox "Blank State India Name Not Allowed !!! "
   TxtStateIndiaName.SetFocus
   Exit Sub
End If
If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed !!! "
   CmbFlag.SetFocus
   Exit Sub
End If

''-- Ullhas Patil 05 July 2017 for State India Wise GSTIN No
If Trim(TxtTINNo.Text) = "" Then
   MsgBox "Blank State TIN Code not allowed !!! "
   TxtTINNo.SetFocus
   Exit Sub
End If
If Trim(TxtStateIndiaCode.Text) = "" Then
   MsgBox "Blank State NAV Code not allowed !!! "
   TxtStateIndiaCode.SetFocus
   Exit Sub
End If
If Trim(TxtGSTIN.Text) = "" Then
   MsgBox "Blank NCML - GST No not allowed !!! "
   TxtGSTIN.SetFocus
   Exit Sub
End If
If Trim(TxtNAVCode.Text) = "" Then
   MsgBox "Blank NAV Code not allowed !!! "
   TxtNAVCode.SetFocus
   Exit Sub
End If
If Trim(TxtStateIndiaOfficeAddress.Text) = "" Then
   MsgBox "Blank NCML - State office address not allowed !!! "
   TxtStateIndiaOfficeAddress.SetFocus
   Exit Sub
End If
If Len(TxtGSTIN) > 0 Then
   ''-- Validation is with specific Error
   mSplitError = Split(ValidateGSTIN(TxtTINNo.Text, TxtGSTIN.Text), "~")
   If CBool(mSplitError(0)) = False Then
      MsgBox mSplitError(1)
      TxtGSTIN.SetFocus
      Exit Sub
   End If
End If
''-----------------------------

If Edit_Flg = "A" Then
   If RsMst_StateIndia.RecordCount > 0 Then
      RsMst_StateIndia.MoveFirst
      RsMst_StateIndia.Find "StateIndiaName= '" & TxtStateIndiaName.Text & "'"
      If Not RsMst_StateIndia.EOF Then
        MsgBox "Duplicate State Name Not Allowed !!! "
         TxtStateIndiaName.SetFocus
         Exit Sub
      End If
      ''-- Ullhas Patil 05 July 2017 for State India Wise GSTIN No
      RsMst_StateIndia.MoveFirst
      RsMst_StateIndia.Find "TINNo= '" & TxtTINNo.Text & "'"
      If Not RsMst_StateIndia.EOF Then
        MsgBox "Duplicate State TIN Number Not Allowed !!! "
         TxtTINNo.SetFocus
         Exit Sub
      End If
      RsMst_StateIndia.MoveFirst
      RsMst_StateIndia.Find "StateIndiaCode= '" & TxtStateIndiaCode.Text & "'"
      If Not RsMst_StateIndia.EOF Then
        MsgBox "Duplicate State Code Not Allowed !!! "
         TxtStateIndiaCode.SetFocus
         Exit Sub
      End If
      RsMst_StateIndia.MoveFirst
      RsMst_StateIndia.Find "GSTIN= '" & TxtGSTIN.Text & "'"
      If Not RsMst_StateIndia.EOF Then
        MsgBox "Duplicate GSTIN No Not Allowed !!! "
         TxtGSTIN.SetFocus
         Exit Sub
      End If
      ''------------------------------
   End If
   
    RsMst_StateIndia.AddNew
    RsMst_StateIndia!StateIndiaID = GetStateID
    RsMst_StateIndia!G_UID = GetGUID
    TxtStateIndiaID.Text = RsMst_StateIndia!StateIndiaID

Else

    RsMst_StateIndia.MoveFirst
    RsMst_StateIndia.Find "StateIndiaName= '" & TxtStateIndiaName.Text & "'"
    If Not RsMst_StateIndia.EOF Then
       If RsMst_StateIndia!StateIndiaID <> TxtStateIndiaID.Text Then
          MsgBox "Duplicate State Name Not Allowed !!! "
          TxtStateIndiaName.SetFocus
          Exit Sub
       End If
    End If
    ''-- Ullhas Patil 05 July 2017 for State India Wise GSTIN No
    RsMst_StateIndia.MoveFirst
    RsMst_StateIndia.Find "TINNo= '" & TxtTINNo.Text & "'"
    If Not RsMst_StateIndia.EOF Then
       If RsMst_StateIndia!StateIndiaID <> TxtStateIndiaID.Text Then
        MsgBox "Duplicate State TIN Number Not Allowed !!! "
        TxtTINNo.SetFocus
        Exit Sub
       End If
    End If
    RsMst_StateIndia.MoveFirst
    RsMst_StateIndia.Find "StateIndiaCode= '" & TxtStateIndiaCode.Text & "'"
    If Not RsMst_StateIndia.EOF Then
      If RsMst_StateIndia!StateIndiaID <> TxtStateIndiaID.Text Then
       MsgBox "Duplicate State Code Not Allowed !!! "
       TxtStateIndiaCode.SetFocus
       Exit Sub
      End If
    End If
    RsMst_StateIndia.MoveFirst
    RsMst_StateIndia.Find "GSTIN= '" & TxtGSTIN.Text & "'"
    If Not RsMst_StateIndia.EOF Then
      If RsMst_StateIndia!StateIndiaID <> TxtStateIndiaID.Text Then
       MsgBox "Duplicate GSTIN No Not Allowed !!! "
       TxtGSTIN.SetFocus
       Exit Sub
      End If
    End If
    ''------------------------------
    RsMst_StateIndia.MoveFirst
    RsMst_StateIndia.Find " StateIndiaID= " & TxtStateIndiaID.Text
End If

RsMst_StateIndia!StateIndiaID = TxtStateIndiaID.Text
RsMst_StateIndia!StateIndiaName = TxtStateIndiaName.Text
RsMst_StateIndia!Flag = Left(CmbFlag.Text, 1)
''-- Ullhas Patil 05 July 2017
RsMst_StateIndia!TINNo = Trim(TxtTINNo.Text)
RsMst_StateIndia!StateIndiaCode = Trim(TxtStateIndiaCode.Text)
RsMst_StateIndia!GSTIN = Trim(TxtGSTIN.Text)
RsMst_StateIndia!NavCode = Trim(TxtNAVCode.Text)
RsMst_StateIndia!StateIndiaOfficeAddress = Trim(TxtStateIndiaOfficeAddress.Text)
''------------------------------------------

If IsNull(RsMst_StateIndia!CreatedBy) = True Or RsMst_StateIndia!CreatedBy = "" Then
   RsMst_StateIndia!CreatedBy = Gen_UserLoginID
   RsMst_StateIndia!CreatedDate = Now
Else
   RsMst_StateIndia!UpdatedBy = Gen_UserLoginID
   RsMst_StateIndia!UpdatedDate = Now
End If

RsMst_StateIndia.Update
RsMst_StateIndia.Requery
RsMst_StateIndia.MoveFirst
RsMst_StateIndia.Find "StateIndiaID = " & TxtStateIndiaID & ""
Call Indata
End Sub

Private Function GetStateID() As Double
Dim Retval As Double
tmp = "Select max(StateIndiaID) from Mst_StateIndia"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetStateID = Retval
End Function

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 8
    .WordWrap = True
    .Font.Size = 10
    .TextMatrix(0, 0) = "State India ID":               .ColWidth(0) = 1300
    .TextMatrix(0, 1) = "State India Name":             .ColWidth(1) = 2500
    .TextMatrix(0, 2) = "Flag":                         .ColWidth(2) = 800
    .TextMatrix(0, 3) = "TIN Code":                     .ColWidth(3) = 1000
    .TextMatrix(0, 4) = "State Code":                   .ColWidth(4) = 1100
    .TextMatrix(0, 5) = "NCML - GST No":                .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "NAV Code":                     .ColWidth(6) = 1800
    .TextMatrix(0, 7) = "NCML - State Office Address":  .ColWidth(7) = 4000
    .RowHeight(0) = 600
End With
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_StateIndia.MoveFirst
RsMst_StateIndia.Find "StateIndiaID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_StateIndia.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

Private Sub TxtStateIndiaName_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtStateIndiaName")
End If
End Sub

Private Sub TxtTINNo_LostFocus()
If Len(TxtTINNo) > 0 Then
    If Len(TxtTINNo) <> 2 Then
        MsgBox "Invalid State TIN Code."
        'TxtTINNo.SetFocus
        Exit Sub
    End If
    If IsClean(TxtTINNo) = False Then
        MsgBox "Special character not Allowed."
        'TxtTINNo.SetFocus
        Exit Sub
    End If
    If IsNumeric(TxtTINNo) = False Then
        MsgBox "Invalid State TIN Code."
        'TxtTINNo.SetFocus
        Exit Sub
    End If
    If Val(TxtTINNo) = 0 Then
       MsgBox "Invalid State TIN Code."
       'TxtTINNo.SetFocus
       Exit Sub
    End If
End If
End Sub
Private Sub TxtStateIndiaCode_LostFocus()
If Len(TxtStateIndiaCode) > 0 Then
'    If Len(TxtStateIndiaCode) <> 2 Then
'        MsgBox "Invalid State Code."
'        'TxtStateIndiaCode.SetFocus
'        Exit Sub
'    End If
    If IsClean(TxtStateIndiaCode) = False Then
        MsgBox "Special character not Allowed."
        'TxtStateIndiaCode.SetFocus
        Exit Sub
    End If
End If
End Sub
Private Sub TxtGSTIN_LostFocus()
Dim mSplitError As Variant
If Len(TxtGSTIN) > 0 Then
   ''-- Validation is with specific Error
   mSplitError = Split(ValidateGSTIN(TxtTINNo.Text, TxtGSTIN.Text), "~")
   If CBool(mSplitError(0)) = False Then
      MsgBox CStr(mSplitError(1))
       Exit Sub
   End If
End If
End Sub

