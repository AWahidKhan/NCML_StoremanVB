VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_Godown 
   Caption         =   "Godown Master"
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7485
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   7485
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   9600
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   3235
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   80
         TabIndex        =   10
         Top             =   150
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   1
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   6390
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5085
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7245
      Begin VB.CommandButton CmdCategoryID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   45
         Top             =   1395
         Width           =   375
      End
      Begin VB.ComboBox CmbCategoryID 
         Height          =   315
         Left            =   2160
         TabIndex        =   42
         Top             =   1380
         Width           =   3975
      End
      Begin VB.TextBox TxtCategoryID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2175
         TabIndex        =   43
         Top             =   1395
         Width           =   3945
      End
      Begin VB.CommandButton CmdUnitTypeID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   41
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton CmdCMPID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         TabIndex        =   40
         Top             =   600
         Width           =   375
      End
      Begin VB.ComboBox CmbUnitTypeID 
         Height          =   315
         Left            =   2160
         TabIndex        =   20
         Top             =   990
         Width           =   4005
      End
      Begin VB.ComboBox CmbCMPID 
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Top             =   615
         Width           =   4000
      End
      Begin VB.TextBox TxtGodownID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   240
         Width           =   3960
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   16
         Top             =   615
         Width           =   3960
      End
      Begin VB.TextBox TxtUnitTypeID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   19
         Top             =   990
         Width           =   3960
      End
      Begin VB.TextBox TxtGodownNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   22
         Top             =   1785
         Width           =   3960
      End
      Begin VB.TextBox TxtBagsCapacity 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   24
         Top             =   3870
         Width           =   3960
      End
      Begin VB.TextBox TxtGodownCapacity 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         TabIndex        =   26
         Top             =   4260
         Width           =   3960
      End
      Begin VB.TextBox TxtGodownAddress 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   2160
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2205
         Width           =   3960
      End
      Begin VB.TextBox TxtGodownArea 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   100
         TabIndex        =   30
         Top             =   3090
         Width           =   3960
      End
      Begin VB.TextBox TxtGodownTelephoneNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   32
         Top             =   3480
         Width           =   3960
      End
      Begin MSComCtl2.DTPicker GodownStartDate 
         Height          =   360
         Left            =   2160
         TabIndex        =   34
         Top             =   4650
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   57081857
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker GodownCloseDate 
         Height          =   360
         Left            =   4815
         TabIndex        =   36
         Top             =   4665
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   57081857
         CurrentDate     =   36494
      End
      Begin VB.Label LblCategoryID 
         Caption         =   "Category"
         Height          =   300
         Left            =   225
         TabIndex        =   44
         Top             =   1380
         Width           =   1560
      End
      Begin VB.Label LblGodownID 
         Caption         =   "Godown ID"
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP "
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblUnitTypeID 
         Caption         =   "Unit Type"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   990
         Width           =   1800
      End
      Begin VB.Label LblGodownNo 
         Caption         =   "Godown No"
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1815
         Width           =   1800
      End
      Begin VB.Label LblBagsCapacity 
         Caption         =   "Godown Bags Capacity"
         Height          =   300
         Left            =   225
         TabIndex        =   23
         Top             =   3885
         Width           =   1800
      End
      Begin VB.Label LblGodownCapacity 
         Caption         =   "Godown Capacity in MT"
         Height          =   300
         Left            =   225
         TabIndex        =   25
         Top             =   4260
         Width           =   1800
      End
      Begin VB.Label LblAddress 
         Caption         =   "Godown Address"
         Height          =   300
         Left            =   225
         TabIndex        =   27
         Top             =   2220
         Width           =   1800
      End
      Begin VB.Label LblArea 
         Caption         =   "Godown Area"
         Height          =   300
         Left            =   240
         TabIndex        =   29
         Top             =   3090
         Width           =   1800
      End
      Begin VB.Label LblTelephoneNo 
         Caption         =   "Telephone No"
         Height          =   300
         Left            =   240
         TabIndex        =   31
         Top             =   3480
         Width           =   1800
      End
      Begin VB.Label LblStartDate 
         Caption         =   "Start Date"
         Height          =   300
         Left            =   240
         TabIndex        =   33
         Top             =   4650
         Width           =   1800
      End
      Begin VB.Label LblCloseDate 
         Caption         =   "Close Date"
         Height          =   300
         Left            =   3840
         TabIndex        =   35
         Top             =   4665
         Width           =   840
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   37
      Top             =   5160
      Width           =   7245
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   38
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   39
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   2
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   3
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2895
         TabIndex        =   4
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   5
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   6
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   7
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   120
         TabIndex        =   8
         Top             =   645
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   9
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "FrmMst_Godown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mCMPID, mUnitTypeID, mCategoryID As Double



Private Sub Form_Load()
Call TableMst_Godown
If RsMst_Godown.RecordCount = 0 Then
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
Call InData
End Sub
Private Sub CmdExcel_Click()
Call Gen_Create_Xls("Mst_Godown.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\Mst_Godown.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(i, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(i, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(i, C)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Godown.MoveNext
If RsMst_Godown.EOF Then
   RsMst_Godown.MoveLast
   LF_Flag = "Y"
End If
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Godown.MovePrevious
If RsMst_Godown.BOF Then
   RsMst_Godown.MoveFirst
   LF_Flag = "Y"
End If
Call InData
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_Godown.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_Godown.MoveLast
Call InData
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
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 350
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_Godown.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_Godown.RecordCount + 1
MSHFlexGrid1.ColWidth(0) = 1000
For i = 1 To RsMst_Godown.RecordCount
    For C = 0 To MSHFlexGrid1.Cols - 1
       MSHFlexGrid1.TextMatrix(i, C) = IIf(IsNull(RsMst_Godown.Fields(C)), "", RsMst_Godown.Fields(C))
    Next
    RsMst_Godown.MoveNext
Next
End Sub
Public Sub InData()

TxtGodownID.Locked = True
TxtCMPID.Locked = True
CmbCMPID.Visible = False
CmdCMPID.Visible = False
TxtCMPID.Visible = True
TxtUnitTypeID.Locked = True
CmbUnitTypeID.Visible = False
CmdUnitTypeID.Visible = False
TxtUnitTypeID.Visible = True
TxtCategoryID.Locked = True
CmbCategoryID.Visible = False
CmdCategoryID.Visible = False
TxtCategoryID.Visible = True
TxtGodownNo.Locked = True
TxtBagsCapacity.Locked = True
TxtGodownCapacity.Locked = True
TxtAddress.Locked = True
TxtArea.Locked = True
TxtTelephoneNo.Locked = True
StartDate.Enabled = False
CloseDate.Enabled = False

TxtGodownID = IIf(IsNull(RsMst_Godown!GodownID), "", RsMst_Godown!GodownID)
mCMPID = IIf(IsNull(RsMst_Godown!CMPID), "", RsMst_Godown!CMPID)
TxtCMPID = GetCMPName(RsMst_Godown!CMPID) '(IsNull(RsMst_Godown!CMPName), "", RsMst_Godown!CMPName)
mUnitTypeID = IIf(IsNull(RsMst_Godown!UnitTypeID), "", RsMst_Godown!UnitTypeID)
TxtUnitTypeID = GetUnitName(RsMst_Godown!UnitTypeID) '(IsNull(RsMst_Godown!UnitType), "", RsMst_Godown!UnitType)
mCategoryID = IIf(IsNull(RsMst_Godown!CategoryID), "", RsMst_Godown!CategoryID)
'TxtCategoryID = IIf(IsNull(RsMst_CMP!Category), "", RsMst_CMP!Category)
TxtCategoryID = GetCategoryName(RsMst_Godown!CategoryID) 'IIf(IsNull(RsMst_CMP!Category), "", RsMst_CMP!Category)
TxtGodownNo = IIf(IsNull(RsMst_Godown!GodownNo), "", RsMst_Godown!GodownNo)
TxtBagsCapacity = IIf(IsNull(RsMst_Godown!BagsCapacity), "", RsMst_Godown!BagsCapacity)
TxtGodownCapacity = IIf(IsNull(RsMst_Godown!GodownCapacity), "", RsMst_Godown!GodownCapacity)
TxtAddress = IIf(IsNull(RsMst_Godown!Address), "", RsMst_Godown!Address)
TxtArea = IIf(IsNull(RsMst_Godown!Area), "", RsMst_Godown!Area)
TxtTelephoneNo = IIf(IsNull(RsMst_Godown!TelephoneNo), "", RsMst_Godown!TelephoneNo)
StartDate.Value = IIf(IsNull(RsMst_Godown!StartDate), "", RsMst_Godown!StartDate)
CloseDate.Value = IIf(IsNull(RsMst_Godown!CloseDate), "", RsMst_Godown!CloseDate)
If RsMst_Godown.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub TxtGodownID_LostFocus()
If TxtGodownID = "" Then Exit Sub
If IsNumeric(TxtGodownID) = False Then
   MsgBox "Invalid number format !!!!"
   TxtGodownID.SetFocus
   Exit Sub
End If
TxtGodownID = Val(TxtGodownID)
End Sub
Private Sub TxtGodownCapacity_LostFocus()
If TxtGodownCapacity = "" Then Exit Sub
If IsNumeric(TxtGodownCapacity) = False Then
   MsgBox "Invalid number format !!!!"
   TxtGodownCapacity.SetFocus
   Exit Sub
End If
TxtGodownCapacity = Val(TxtGodownCapacity)
End Sub
Private Sub TxtStartDate_LostFocus()
End Sub
Private Sub TxtCloseDate_LostFocus()
End Sub
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtGodownID.Locked = True
TxtCMPID.Locked = False
CmbCMPID.Visible = True
CmdCMPID.Visible = True
TxtCMPID.Visible = False
TxtUnitTypeID.Locked = False
CmbUnitTypeID.Visible = True
CmdUnitTypeID.Visible = True
TxtUnitTypeID.Visible = False
TxtCategoryID.Locked = False
CmbCategoryID.Visible = True
CmdCategoryID.Visible = True
TxtCategoryID.Visible = False
TxtGodownNo.Locked = False
TxtBagsCapacity.Locked = False
TxtGodownCapacity.Locked = False
TxtAddress.Locked = False
TxtArea.Locked = False
TxtTelephoneNo.Locked = False
StartDate.Enabled = True
CloseDate.Enabled = True

TxtGodownID = ""

TxtCMPID = ""

TxtUnitTypeID = ""

TxtCategoryID = ""

TxtGodownNo = ""

TxtBagsCapacity = ""

TxtGodownCapacity = ""

TxtAddress = ""

TxtArea = ""

TxtTelephoneNo = ""

StartDate = Date

CloseDate.Value = Date
End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtGodownID.Locked = True
TxtCMPID.Locked = False
CmbCMPID.Visible = True
CmdCMPID.Visible = True
TxtCMPID.Visible = False
TxtUnitTypeID.Locked = False
CmbUnitTypeID.Visible = True
CmdUnitTypeID.Visible = True
TxtUnitTypeID.Visible = True
TxtCategoryID.Locked = False
CmbCategoryID.Visible = True
CmdCategoryID.Visible = True
TxtCategoryID.Visible = False
TxtGodownNo.Locked = False
TxtBagsCapacity.Locked = False
TxtGodownCapacity.Locked = False
TxtAddress.Locked = False
TxtArea.Locked = False
TxtTelephoneNo.Locked = False
StartDate.Enabled = True
CloseDate.Enabled = True

CmbCMPID.Text = TxtCMPID
CmbUnitTypeID.Text = TxtUnitTypeID
CmbCategoryID.Text = TxtCategoryID
End Sub
Private Sub DeleteCmd_Click()
Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If Ans = vbYes Then
   RsMst_Godown.Delete
   RsMst_Godown.Update
   If RsMst_Godown.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Godown.MoveNext
   If RsMst_Godown.EOF() Then
      RsMst_Godown.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_Godown.CancelUpdate
End Sub
Private Sub SaveCmd_Click()


If TxtGodownNo.Text = "" Then
   MsgBox "Blank GodownNo Not Allowed !!! "
   TxtGodownNo.SetFocus
   Exit Sub
End If


If TxtBagsCapacity.Text = "" Then
   MsgBox "Blank BagsCapacity Not Allowed !!! "
   TxtBagsCapacity.SetFocus
   Exit Sub
End If


If TxtGodownCapacity.Text = "" Then
   MsgBox "Blank GodownCapacity Not Allowed !!! "
   TxtGodownCapacity.SetFocus
   Exit Sub
End If


If TxtAddress.Text = "" Then
   MsgBox "Blank Address Not Allowed !!! "
   TxtAddress.SetFocus
   Exit Sub
End If


If TxtArea.Text = "" Then
   MsgBox "Blank Area Not Allowed !!! "
   TxtArea.SetFocus
   Exit Sub
End If


If TxtTelephoneNo.Text = "" Then
   MsgBox "Blank TelephoneNo Not Allowed !!! "
   TxtTelephoneNo.SetFocus
   Exit Sub
End If


If StartDate.Value = "" Then
   MsgBox "Blank StartDate Not Allowed !!! "
   StartDate.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Godown.RecordCount > 0 Then
      RsMst_Godown.MoveFirst
      RsMst_Godown.Find "GodownNo= '" & TxtGodownNo.Text & "'"
      If Not RsMst_Godown.EOF Then
        MsgBox "Duplicate Godown No Not Allowed !!! "
         Txt.SetFocus
         Exit Sub
      End If
   End If
      RsMst_Godown.AddNew
      RsMst_Godown!GodownID = GetGodownID
      TxtGodownID.Text = RsMst_Godown!GodownID
Else
    RsMst_Godown.MoveFirst
    RsMst_Godown.Find "GodownNo= '" & TxtGodownNo.Text & "'"
    If Not RsMst_Godown.EOF Then
       If RsMst_Godown!GodownID <> TxtGodownID.Text Then
          MsgBox "Duplicate Godown No Not Allowed !!! "
          Txt.SetFocus
          Exit Sub
       End If
    End If
    RsMst_Godown.MoveFirst
    RsMst_Godown.Find " GodownID= " & TxtGodownID.Text
End If

RsMst_Godown!GodownID = TxtGodownID.Text
RsMst_Godown!CMPID = mCMPID
RsMst_Godown!UnitTypeID = mUnitTypeID
RsMst_Godown!CategoryID = mCategoryID
RsMst_Godown!GodownNo = TxtGodownNo.Text
RsMst_Godown!BagsCapacity = TxtBagsCapacity.Text
RsMst_Godown!GodownCapacity = TxtGodownCapacity.Text
RsMst_Godown!Address = TxtAddress.Text
RsMst_Godown!Area = TxtArea.Text
RsMst_Godown!TelephoneNo = TxtTelephoneNo.Text
RsMst_Godown!StartDate = StartDate.Value
RsMst_Godown!CloseDate = CloseDate.Value
RsMst_Godown.Update
Call InData
End Sub
Private Function GetGodownID() As Double
Dim RetVal As Double
tmp = "Select max(GodownID) from Mst_Godown"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetGodownID = RetVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12
MSHFlexGrid1.TextMatrix(0, 0) = "Godown ID"
MSHFlexGrid1.TextMatrix(0, 1) = "CMP "
MSHFlexGrid1.TextMatrix(0, 2) = "Unit Type"
MSHFlexGrid1.TextMatrix(0, 3) = "Category Type"
MSHFlexGrid1.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 5) = "Godown Bags Capacity"
MSHFlexGrid1.TextMatrix(0, 6) = "Godown Capacity in MT"
MSHFlexGrid1.TextMatrix(0, 7) = "Godown Address"
MSHFlexGrid1.TextMatrix(0, 8) = "Godown Area"
MSHFlexGrid1.TextMatrix(0, 9) = "Telephone No"
MSHFlexGrid1.TextMatrix(0, 10) = "Start Date"
MSHFlexGrid1.TextMatrix(0, 11) = "Close Date"
End Sub
Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_Godown.MoveFirst
RsMst_Godown.Find "GodownID = " & MSHFlexGrid1.TextMatrix(i, 0)
If Not RsMst_Godown.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCMPID_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
Call TableMst_CMP
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub
'----
Private Sub CmdUnitTypeID_Click()
    FrmMst_UnitType.Show
End Sub
'----
Private Sub CmbUnitTypeID_GotFocus()
tmp = CmbUnitTypeID.Text
Call TableMst_UnitType
CmbUnitTypeID.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_UnitType.RecordCount
    CmbUnitTypeID.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbUnitTypeID.Text = tmp
End Sub
'----
Private Sub CmbUnitTypeID_LostFocus()
If CmbUnitTypeID.Text = "" Then
   Exit Sub
End If
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbUnitTypeID.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid  selection "
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
mUnitTypeID = RsMst_UnitType!UnitTypeID
End Sub
'----
Private Sub CmdCategoryID_Click()
    FrmMst_Category.Show
End Sub
'----
Private Sub CmbCategoryID_GotFocus()
tmp = CmbCategoryID.Text
Call TableMst_Category
CmbCategoryID.Clear
If RsMst_Category.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Category.RecordCount
    CmbCategoryID.AddItem RsMst_Category!Category
    RsMst_Category.MoveNext
Next
CmbCategoryID.Text = tmp
End Sub
'----
Private Sub CmbCategoryID_LostFocus()
If CmbCategoryID.Text = "" Then
   Exit Sub
End If
RsMst_Category.MoveFirst
RsMst_Category.Find "Category = '" & CmbCategoryID.Text & "'"
If RsMst_Category.EOF Then
   MsgBox "Invalid  selection "
   CmbCategoryID.SetFocus
   Exit Sub
End If
mCategoryID = RsMst_Category!CategoryID
End Sub
'----

