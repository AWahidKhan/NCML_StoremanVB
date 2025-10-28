VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_EarthquakeZone 
   Caption         =   "Earthquake Zone"
   ClientHeight    =   5910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11175
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   11175
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   12480
      Begin VB.ComboBox CmbSFlag 
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
         ItemData        =   "FrmMst_EarthquakeZone.frx":0000
         Left            =   7770
         List            =   "FrmMst_EarthquakeZone.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   29
         Text            =   "Active"
         Top             =   150
         Width           =   1575
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
         Left            =   10140
         TabIndex        =   27
         Top             =   150
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchEQZone 
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
         Left            =   2745
         TabIndex        =   26
         Top             =   150
         Width           =   4065
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
         Left            =   11160
         TabIndex        =   3
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   4
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
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
      Begin VB.Label Label1 
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6930
         TabIndex        =   30
         Top             =   150
         Width           =   1800
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Earthquake Zone"
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
         Left            =   90
         TabIndex        =   28
         Top             =   180
         Width           =   2895
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   7245
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
         TabIndex        =   21
         Top             =   660
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   660
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
         TabIndex        =   18
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
         Height          =   400
         Left            =   2895
         TabIndex        =   17
         Top             =   660
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
         TabIndex        =   16
         Top             =   240
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
         TabIndex        =   15
         Top             =   660
         Width           =   1400
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
         Height          =   400
         Left            =   1500
         TabIndex        =   14
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
         Left            =   100
         TabIndex        =   13
         Top             =   660
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
         Left            =   100
         TabIndex        =   12
         Top             =   240
         Width           =   1400
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   7245
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1680
         Width           =   4335
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox TxtEQZone 
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   0
         Top             =   735
         Width           =   4320
      End
      Begin VB.TextBox TxtEQZoneID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   1560
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
         ItemData        =   "FrmMst_EarthquakeZone.frx":0021
         Left            =   2040
         List            =   "FrmMst_EarthquakeZone.frx":002B
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "Active"
         Top             =   1230
         Width           =   4335
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1230
         Visible         =   0   'False
         Width           =   4320
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
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label8 
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
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label LblFlag 
         Caption         =   "Status"
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
         TabIndex        =   10
         Top             =   1260
         Width           =   1800
      End
      Begin VB.Label LblZone 
         Caption         =   "Earthquake Zone"
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
         TabIndex        =   9
         Top             =   765
         Width           =   1800
      End
      Begin VB.Label LblZoneID 
         Caption         =   "Earthquake Zone ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   270
         Width           =   1800
      End
   End
End
Attribute VB_Name = "FrmMst_EarthquakeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub CmbFlag_LostFocus()
If CmbFlag.Text = "" Then Exit Sub

If CmbFlag.Text <> "Active" And CmbFlag.Text <> "De-Active" Then
   MsgBox "Invalid selection!!! "
   CmbFlag.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdGo_Click()

Dim wc As Variant
wc = ""
If TxtSearchEQZone <> "" Then
   wc = " Where EQZone Like '%" & TxtSearchEQZone & "%' "
End If
If CmbSFlag.Text = "Active" Then
   If wc <> "" Then
      wc = wc & " And Flag = 'A'"
   Else
      wc = " Where Flag = 'A'"
   End If
ElseIf CmbSFlag.Text = "De-Active" Then
   If wc <> "" Then
      wc = wc & " And Flag = 'D'"
   Else
      wc = " Where  Flag = 'D'"
   End If
End If

tmp = "Select * From Mst_EarthquakeZone"

tmp = tmp & wc

Call TmpTable
Call Grid_Head

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2

For I = 1 To TmpRs.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    If MSHFlexGrid1.TextMatrix(I, 2) = "A" Then
       MSHFlexGrid1.TextMatrix(I, 2) = "Active"
    Else
       MSHFlexGrid1.TextMatrix(I, 2) = "De-Active"
    End If
    TmpRs.MoveNext
Next

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
CmbSFlag.Text = ""
TxtSearchEQZone = ""

Call TableMst_EarthquakeZone
If RsMst_EarthquakeZone.RecordCount = 0 Then
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
Call Gen_Create_Xls("Mst_EarthquakeZone.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & "Mst_EarthquakeZone.xls")
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
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_EarthquakeZone.MoveNext
If RsMst_EarthquakeZone.EOF Then
   RsMst_EarthquakeZone.MoveLast
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
RsMst_EarthquakeZone.MovePrevious
If RsMst_EarthquakeZone.BOF Then
   RsMst_EarthquakeZone.MoveFirst
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
RsMst_EarthquakeZone.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_EarthquakeZone.MoveLast
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
   If RsMst_EarthquakeZone.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
'Call Grid_Head
'CmbSFlag.Text = ""
'TxtSearchEQZone = ""
''===========comment on 7 july 2012 for not clearing search grid=============


End Sub
Public Sub Indata()

TxtEQZoneID.Locked = True
TxtEQZone.Locked = True
TxtFlag.Locked = True

CmbFlag.Visible = False
TxtFlag.Visible = True


TxtEQZoneID = IIf(IsNull(RsMst_EarthquakeZone!EQZoneID), "", RsMst_EarthquakeZone!EQZoneID)
TxtEQZone = IIf(IsNull(RsMst_EarthquakeZone!EQZone), "", RsMst_EarthquakeZone!EQZone)
TxtFlag = IIf(IsNull(RsMst_EarthquakeZone!Flag), "", RsMst_EarthquakeZone!Flag)
If RsMst_EarthquakeZone!Flag = "A" Then
   TxtFlag = "Active" 'TxtFlag.Text
Else
   TxtFlag = "De-Active"  'TxtFlag.Text
End If

TxtCreated = IIf(IsNull(RsMst_EarthquakeZone!CreatedBy), "", RsMst_EarthquakeZone!CreatedBy) & " :- " & IIf(IsNull(RsMst_EarthquakeZone!CreatedDate), "", RsMst_EarthquakeZone!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_EarthquakeZone!UpdatedBy), "", RsMst_EarthquakeZone!UpdatedBy) & " :- " & IIf(IsNull(RsMst_EarthquakeZone!UpdatedDate), "", RsMst_EarthquakeZone!UpdatedDate)

If RsMst_EarthquakeZone.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtEQZone_LostFocus()
TxtEQZone = Trim(TxtEQZone)
End Sub

Private Sub TxtEQZoneID_LostFocus()
If TxtEQZoneID = "" Then Exit Sub
If IsNumeric(TxtEQZoneID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtEQZoneID.SetFocus
   Exit Sub
End If
TxtEQZoneID = Val(TxtEQZoneID)
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtEQZoneID.Locked = True
TxtEQZone.Locked = False
TxtFlag.Locked = False

TxtEQZoneID = ""

TxtEQZone = ""
CmbFlag.Visible = True
TxtFlag.Visible = False
TxtCreated = ""
TxtUpdated = ""
CmbFlag.Text = ""

If RsMst_EarthquakeZone.RecordCount > 0 Then
   TxtEQZone.SetFocus
End If


End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtEQZoneID.Locked = True
TxtEQZone.Locked = False
TxtFlag.Locked = False
CmbFlag.Visible = True
TxtFlag.Visible = False
CmbFlag.Text = TxtFlag.Text
End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_EarthquakeZone.Delete
   RsMst_EarthquakeZone.Update
   If RsMst_EarthquakeZone.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_EarthquakeZone.MoveNext
   If RsMst_EarthquakeZone.EOF() Then
      RsMst_EarthquakeZone.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_EarthquakeZone.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtEQZone.Text = "" Then
   MsgBox "Blank Earthquake Zone Not Allowed !!! "
   TxtEQZone.SetFocus
   Exit Sub
End If

If CmbFlag.Text <> "Active" And CmbFlag.Text <> "De-Active" Then
   MsgBox "Invalid selection!!! "
   CmbFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_EarthquakeZone.RecordCount > 0 Then
      RsMst_EarthquakeZone.MoveFirst
      RsMst_EarthquakeZone.Find "EQZone= '" & TxtEQZone.Text & "'"
      If Not RsMst_EarthquakeZone.EOF Then
        MsgBox "Duplicate Earthquake Zone Not Allowed !!! "
         TxtEQZone.SetFocus
         Exit Sub
      End If
   End If
      RsMst_EarthquakeZone.AddNew
      RsMst_EarthquakeZone!EQZoneID = GetEQZoneID
      RsMst_EarthquakeZone!G_UID = GetGUID
      TxtEQZoneID.Text = RsMst_EarthquakeZone!EQZoneID
Else
    RsMst_EarthquakeZone.MoveFirst
    RsMst_EarthquakeZone.Find "EQZone= '" & TxtEQZone.Text & "'"
    If Not RsMst_EarthquakeZone.EOF Then
       If RsMst_EarthquakeZone!EQZoneID <> TxtEQZoneID.Text Then
          MsgBox "Duplicate Earthquake Zone Not Allowed !!! "
          TxtEQZone.SetFocus
          Exit Sub
       End If
    End If
    RsMst_EarthquakeZone.MoveFirst
    RsMst_EarthquakeZone.Find " EQZoneID= " & TxtEQZoneID.Text
End If

RsMst_EarthquakeZone!EQZoneID = TxtEQZoneID.Text
RsMst_EarthquakeZone!EQZone = TxtEQZone.Text
If CmbFlag = "Active" Then
   RsMst_EarthquakeZone!Flag = "A" 'TxtFlag.Text
Else
   RsMst_EarthquakeZone!Flag = "D" 'TxtFlag.Text
End If

If IsNull(RsMst_EarthquakeZone!CreatedBy) = True Or RsMst_EarthquakeZone!CreatedBy = "" Then
   RsMst_EarthquakeZone!CreatedBy = Gen_UserLoginID
   RsMst_EarthquakeZone!CreatedDate = Now
Else
   RsMst_EarthquakeZone!UpdatedBy = Gen_UserLoginID
   RsMst_EarthquakeZone!UpdatedDate = Now
End If

RsMst_EarthquakeZone.Update
Call Indata
End Sub
Private Function GetEQZoneID() As Double
Dim Retval As Double
tmp = "Select max(EQZoneID) from Mst_EarthquakeZone"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetEQZoneID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.TextMatrix(0, 0) = "Earthquake Zone ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Earthquake Zone"
MSHFlexGrid1.TextMatrix(0, 2) = "Flag"

MSHFlexGrid1.ColWidth(1) = 4000
MSHFlexGrid1.ColWidth(0) = 1800
End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_EarthquakeZone.MoveFirst
RsMst_EarthquakeZone.Find "EQZoneID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_EarthquakeZone.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub

