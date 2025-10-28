VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Branch 
   Caption         =   "Branch Master"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   3645
      Left            =   120
      TabIndex        =   36
      Top             =   7440
      Visible         =   0   'False
      Width           =   15000
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   10320
         TabIndex        =   40
         Top             =   150
         Width           =   1000
      End
      Begin VB.ComboBox CmbSBank 
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
         Left            =   6480
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   165
         Width           =   2670
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   9315
         TabIndex        =   38
         Top             =   150
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchBranch 
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
         TabIndex        =   37
         Top             =   165
         Width           =   3330
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   70
         TabIndex        =   41
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
      Begin VB.Label Label11 
         Caption         =   "Bank"
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
         Left            =   5955
         TabIndex        =   45
         Top             =   240
         Width           =   510
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
         Left            =   14190
         TabIndex        =   44
         Top             =   165
         Width           =   75
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
         Left            =   11790
         TabIndex        =   43
         Top             =   165
         Width           =   2355
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Branch Name"
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
         Left            =   120
         TabIndex        =   42
         Top             =   195
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   14070
      Begin VB.TextBox TxtIFSCCode 
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
         MaxLength       =   100
         TabIndex        =   7
         Top             =   4890
         Width           =   3375
      End
      Begin VB.ComboBox CmbHRBranchID 
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
         TabIndex        =   8
         Top             =   5280
         Width           =   11535
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
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   6240
         Width           =   11535
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5760
         Width           =   11535
      End
      Begin VB.CommandButton CmdBank 
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
         Height          =   330
         Left            =   13545
         TabIndex        =   20
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox TxtBranchID 
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtBranchName 
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
         TabIndex        =   0
         Top             =   720
         Width           =   11535
      End
      Begin VB.TextBox TxtContactPerson 
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
         TabIndex        =   2
         Top             =   1680
         Width           =   11535
      End
      Begin VB.TextBox TxtAddress 
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
         Height          =   1215
         Left            =   2400
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2160
         Width           =   11535
      End
      Begin VB.TextBox TxtTelephone1 
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
         Top             =   3480
         Width           =   11535
      End
      Begin VB.TextBox TxtTelephone2 
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
         TabIndex        =   5
         Top             =   3960
         Width           =   11535
      End
      Begin VB.TextBox TxtEmailId 
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
         TabIndex        =   6
         Top             =   4440
         Width           =   11535
      End
      Begin VB.ComboBox CmbBankName 
         Height          =   315
         Left            =   2400
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1200
         Width           =   11085
      End
      Begin VB.TextBox TxtBankName 
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
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   11550
      End
      Begin VB.TextBox TxtHRBranchID 
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
         TabIndex        =   46
         Top             =   5280
         Width           =   11550
      End
      Begin VB.Label Label25 
         Caption         =   "IFSC Code"
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
         TabIndex        =   48
         Top             =   4943
         Width           =   1695
      End
      Begin VB.Label Label36 
         Caption         =   "HR Branch"
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
         TabIndex        =   47
         Top             =   5340
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
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   6240
         Width           =   1695
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
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   5760
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Contact Person"
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
         Left            =   240
         TabIndex        =   17
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Branch Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Branch ID"
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
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile"
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
         Left            =   240
         TabIndex        =   13
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Telephone 2"
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
         Left            =   240
         TabIndex        =   12
         Top             =   3960
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Email Id"
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
         Left            =   240
         TabIndex        =   11
         Top             =   4440
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Width           =   14070
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   11100
         TabIndex        =   31
         Top             =   615
         Width           =   2655
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   8445
         TabIndex        =   30
         Top             =   615
         Width           =   2655
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   5670
         TabIndex        =   29
         Top             =   615
         Width           =   2775
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   3015
         TabIndex        =   28
         Top             =   615
         Width           =   2655
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   615
         Width           =   2655
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   11100
         TabIndex        =   26
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   8445
         TabIndex        =   25
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   5670
         TabIndex        =   24
         Top             =   240
         Width           =   2775
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   3015
         TabIndex        =   23
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmMst_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mBankID As Variant
Dim mHRBankID, mHRBranchID As Variant


Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtBranchName.Locked = False
CmdBank.Visible = True
Frame1.Enabled = True
TxtBranchID = ""
TxtBranchName = ""
TxtContactPerson = ""
TxtAddress = ""
TxtTelephone1 = ""
TxtTelephone2 = ""
TxtEmailId = ""

TxtCreated = ""
TxtUpdated = ""

TxtBankName.Visible = False
CmbBankName.Visible = True
CmbBankName.Text = ""

TxtHRBranchID.Visible = False
CmbHRBranchID.Visible = True
CmbHRBranchID.Text = ""

If RsMst_Branch.RecordCount > 0 Then
   TxtBranchName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbBankName_GotFocus()
tmp = CmbBankName.Text
Call TableMst_Bank("")

CmbBankName.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbBankName.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbBankName.AddItem RsMst_Bank!Bankname
    RsMst_Bank.MoveNext
Next
CmbBankName.Text = tmp
End Sub

Private Sub CmbBankName_LostFocus()
If CmbBankName.Text = "" Then
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbBankName.Text & "'"

If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection "
   CmbBankName.SetFocus
   Exit Sub
End If
mBankID = RsMst_Bank!BankID
mHRBankID = IIf(IsNull(RsMst_Bank!HRBankID), 0, RsMst_Bank!HRBankID)
If mHRBankID = 0 Then
   MsgBox "No Universal Bank Mapped to this Bank in Bank Master . Please map same!!!"
   CmbBankName.SetFocus
   Exit Sub
End If
End Sub



Private Sub CmbSBank_GotFocus()
tmp = CmbSBank.Text
Call TableMst_Bank("")

CmbSBank.Clear
If RsMst_Bank.RecordCount = 0 Then
   MsgBox "No Bank found"
   CmbSBank.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_Bank.RecordCount
    CmbSBank.AddItem RsMst_Bank!Bankname
    RsMst_Bank.MoveNext
Next
CmbSBank.Text = tmp

End Sub

Private Sub CmbSBank_LostFocus()
If CmbSBank.Text = "" Then
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankName = '" & CmbSBank.Text & "'"

If RsMst_Bank.EOF Then
   MsgBox "Invalid Bank selection "
   CmbSBank.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmdBank_Click()
FrmMst_Bank.Show
End Sub



Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Branch.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Branch.xls")
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

Private Sub CmdGo_Click()
tmp = "Select MBranch.BranchID,MBranch.BranchName,MBank.BankName,MBranch.ContactPerson,MBranch.Address,MBranch.Tel1,MBranch.Tel2, " & _
      " MBranch.EmailID,isnull(MBranch.HRBranchID,0) 'HRBranchID' " & _
      " From Mst_Branch MBranch " & _
      " Inner Join Mst_Bank MBank On MBranch.BankID = MBank.BankID"

Dim wc As Variant
wc = ""
If TxtSearchBranch <> "" Then
   wc = " Where MBranch.BranchName Like '%" & TxtSearchBranch & "%' "
End If

If CmbSBank.Text <> "" Then
   If TxtSearchBranch = "" Then
      wc = " Where MBank.BankName = '" & CmbSBank.Text & "'"
   Else
      wc = wc & " And MBank.BankName = '" & CmbSBank.Text & "'"
   End If
End If
tmp = tmp & wc & " Order by MBank.BankName,MBranch.BranchName"

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    MSHFlexGrid1.TextMatrix(iu, 8) = GetHRBranch(MSHFlexGrid1.TextMatrix(iu, 8))
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   RsMst_Branch.Delete
   RsMst_Branch.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_Branch.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Branch.MoveNext
   If RsMst_Branch.EOF() Then
      RsMst_Branch.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Branch.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
TxtBankName.Visible = False
CmbBankName.Visible = True
CmbBankName.Text = TxtBankName
Call CmbBankName_GotFocus
Call CmbBankName_LostFocus

TxtHRBranchID.Visible = False
CmbHRBranchID.Visible = True
CmbHRBranchID.Text = TxtHRBranchID

Call GButtonLock(Me, False)
TxtBranchName.Locked = False
CmdBank.Visible = True

TxtBranchName.SetFocus
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Branch.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
TxtSearchBranch = ""
CmbSBank.Text = ""
Label22.Caption = ""

Call TableMst_Branch("  Order by BankID,BranchName ")
If RsMst_Branch.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
Frame1.Enabled = False
TxtBranchID.Locked = True
TxtBranchID = RsMst_Branch!BranchID
TxtBranchName = RsMst_Branch!BranchName
TxtHRBranchID.Visible = True
CmbHRBranchID.Visible = False
mBankID = RsMst_Branch!BankID
TxtContactPerson = IIf(IsNull(RsMst_Branch!ContactPerson), "", RsMst_Branch!ContactPerson)
TxtAddress = IIf(IsNull(RsMst_Branch!Address), "", RsMst_Branch!Address)
TxtTelephone1 = IIf(IsNull(RsMst_Branch!Tel1), "", RsMst_Branch!Tel1)
TxtTelephone2 = IIf(IsNull(RsMst_Branch!Tel2), "", RsMst_Branch!Tel2)
TxtEmailId = IIf(IsNull(RsMst_Branch!EmailID), "", RsMst_Branch!EmailID)
TxtBankName = GetBankName(RsMst_Branch!BankID, "N")
mHRBranchID = IIf(IsNull(RsMst_Branch!HRBranchID), 0, RsMst_Branch!HRBranchID)
TxtHRBranchID = GetHRBranch(mHRBranchID)
TxtIFSCCode = GetHRBranchIFSC(mHRBranchID)
TxtCreated = IIf(IsNull(RsMst_Branch!CreatedBy), "", RsMst_Branch!CreatedBy) & " :- " & IIf(IsNull(RsMst_Branch!CreatedDate), "", RsMst_Branch!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Branch!UpdatedBy), "", RsMst_Branch!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Branch!UpdatedDate), "", RsMst_Branch!UpdatedDate)


TxtBankName.Visible = True
CmbBankName.Visible = False
CmdBank.Visible = False

Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_Branch.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_Branch.MoveFirst

If RsMst_Branch.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchID = " & Val(mCaId) & ""
If RsMst_Branch.EOF Then
   RsMst_Branch.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Branch.MoveNext
If RsMst_Branch.EOF Then
   RsMst_Branch.MoveLast
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
RsMst_Branch.MovePrevious
If RsMst_Branch.BOF Then
   RsMst_Branch.MoveFirst
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

If TxtBranchName = "" Then
   MsgBox "Blank Branch Name not Allowed"
   TxtBranchName.SetFocus
   Exit Sub
End If
If CmbBankName.Text = "" Then
   MsgBox "Blank Branch Name not Allowed"
   CmbBankName.SetFocus
   Exit Sub
End If
If TxtAddress = "" Then
   MsgBox "Blank Branch Address not Allowed"
   TxtAddress.SetFocus
   Exit Sub
End If

If CmbHRBranchID.Text = "" Then
   MsgBox "Blank HR Branch not Allowed"
   CmbHRBranchID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Branch.RecordCount > 0 Then
      RsMst_Branch.MoveFirst
      RsMst_Branch.Find "BranchName = '" & TxtBranchName & "'"
      If Not RsMst_Branch.EOF Then
         MsgBox "Duplicate Branch Name not Allowed"
         Exit Sub
      End If
   End If
   RsMst_Branch.AddNew
   RsMst_Branch!BranchID = GetBranchID 'tmp
   RsMst_Branch!G_UID = GetGUID
  'RsMst_Branch!BranchID = tmp
   TxtBranchID = RsMst_Branch!BranchID
Else
   RsMst_Branch.MoveFirst
   RsMst_Branch.Find "BranchName = '" & TxtBranchName & "'"
   If Not RsMst_Branch.EOF Then
      If RsMst_Branch!BranchID <> TxtBranchID Then
         MsgBox "Duplicate Branch Name not Allowed"
         Exit Sub
      End If
   End If
   RsMst_Branch.MoveFirst
   RsMst_Branch.Find "BranchID = " & TxtBranchID & ""
End If

RsMst_Branch!BranchName = TxtBranchName
RsMst_Branch!BankID = mBankID

RsMst_Branch!ContactPerson = TxtContactPerson
RsMst_Branch!Address = TxtAddress
RsMst_Branch!Tel1 = TxtTelephone1
RsMst_Branch!Tel2 = TxtTelephone2
RsMst_Branch!EmailID = TxtEmailId
RsMst_Branch!HRBranchID = mHRBranchID
If IsNull(RsMst_Branch!CreatedBy) = True Or RsMst_Branch!CreatedBy = "" Then
   RsMst_Branch!CreatedBy = Gen_UserLoginID
   RsMst_Branch!CreatedDate = Now
Else
   RsMst_Branch!UpdatedBy = Gen_UserLoginID
   RsMst_Branch!UpdatedDate = Now
End If


RsMst_Branch.Update
RsMst_Branch.Requery

RsMst_Branch.MoveFirst
RsMst_Branch.Find "BranchID = " & TxtBranchID & ""
If RsMst_Branch.EOF Then
   RsMst_Branch.MoveFirst
End If
Call GButtonLock(Me, True)
TxtBranchID.Locked = True
Call Indata
End Sub
Private Function GetBranchID() As Double
Dim Retval As Double
tmp = "Select max(BranchID) from Mst_Branch"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetBranchID = Retval
End Function
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_Branch.EOF Then
      RsMst_Branch.MoveFirst
   End If
   Call Indata
   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top '120
Frame0.Width = Me.Width - 150
Frame0.Visible = True

MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

RsMst_Branch.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head
'TxtSearchBranch = ""
'CmbSBank.Text = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9 '8


MSHFlexGrid1.TextMatrix(0, 0) = "Branch ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Branch Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Bank Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Contact Person"
MSHFlexGrid1.TextMatrix(0, 4) = "Address"
MSHFlexGrid1.TextMatrix(0, 5) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 6) = "Telephone No"
MSHFlexGrid1.TextMatrix(0, 7) = "EmailID"
MSHFlexGrid1.TextMatrix(0, 8) = "HR BRanch Name"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 2800
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 2000
MSHFlexGrid1.ColWidth(7) = 1800


End Sub

Private Sub TxtBranchID_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtBranchID)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtBranchID.SetFocus
End If
End Sub

Private Sub TxtBranchName_Lostfocus()
Dim ans As Boolean
ans = IsClean(TxtBranchName)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtBranchName.SetFocus
End If

End Sub

Private Sub CmbHRBranchID_GotFocus()

If CmbBankName.Text = "" Then
   MsgBox "Select Bank!!!"
   CmbBankName.SetFocus
   Exit Sub
End If

tmp1 = CmbHRBranchID.Text
'tmp = "select * from Mst_Branch WHere BankID = " & mHRBankID & " "

If Trim(TxtIFSCCode) = "" Then
   tmp = "select * from Mst_Branch WHere BankID = " & mHRBankID & " "
   Call TmpTableHRPro
Else
   tmp = "Select * from Mst_Branch WHere IFSCCode = '" & Trim(TxtIFSCCode) & "'  "
   Call TmpTableHRPro
   If TmpRsHRPro.RecordCount = 0 Then
      tmp = "select * from Mst_Branch WHere BankID = " & mHRBankID & "  "
      Call TmpTableHRPro
   End If
End If

'Call TmpTableHRPro

CmbHRBranchID.Clear
If TmpRsHRPro.RecordCount = 0 Then
   MsgBox "No HR Branch found!!!"
   CmbHRBranchID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRsHRPro.RecordCount
    CmbHRBranchID.AddItem TmpRsHRPro!BranchName
    TmpRsHRPro.MoveNext
Next
CmbHRBranchID.Text = tmp1
End Sub

Private Sub CmbHRBranchID_LostFocus()
If CmbHRBranchID.Text = "" Then
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "BranchName = '" & CmbHRBranchID.Text & "'"

If TmpRsHRPro.EOF Then
   MsgBox "Invalid HR Branch selection "
   CmbHRBranchID.SetFocus
   Exit Sub
End If
mHRBranchID = TmpRsHRPro!BranchID
End Sub
