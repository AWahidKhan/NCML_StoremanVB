VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_FAGBank 
   Caption         =   "FAG Bank Master"
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
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   6510
      Visible         =   0   'False
      Width           =   15015
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
         Left            =   7035
         TabIndex        =   10
         Top             =   270
         Width           =   1005
      End
      Begin VB.TextBox TxtSearchBank 
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
         Left            =   2265
         TabIndex        =   9
         Top             =   240
         Width           =   4605
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
         Left            =   8040
         TabIndex        =   8
         Top             =   270
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   120
         TabIndex        =   11
         Top             =   705
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
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
         Left            =   11790
         TabIndex        =   14
         Top             =   285
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
         Left            =   9390
         TabIndex        =   13
         Top             =   318
         Width           =   2355
      End
      Begin VB.Label Label19 
         Caption         =   "Search By Bank Name"
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
         TabIndex        =   12
         Top             =   285
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   8415
      Begin VB.TextBox TxtPaymentLimitation 
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
         MaxLength       =   25
         TabIndex        =   36
         Top             =   3960
         Width           =   5415
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
         Height          =   375
         Left            =   2400
         MaxLength       =   25
         TabIndex        =   5
         Top             =   3480
         Width           =   5415
      End
      Begin VB.TextBox TxtAccountNo 
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
         Top             =   1200
         Width           =   5415
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
         TabIndex        =   4
         Top             =   2160
         Width           =   5415
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
         TabIndex        =   3
         Top             =   1680
         Width           =   5415
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
         TabIndex        =   1
         Top             =   720
         Width           =   5415
      End
      Begin VB.TextBox TxtBankID 
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
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4440
         Width           =   5415
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
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   4920
         Width           =   5415
      End
      Begin VB.Label Label7 
         Caption         =   "Payment Limitaion"
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
         TabIndex        =   37
         Top             =   4005
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "NAV Code"
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
         TabIndex        =   35
         Top             =   3525
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Account No"
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
         TabIndex        =   34
         Top             =   1260
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Address"
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
         TabIndex        =   33
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   270
         Left            =   240
         TabIndex        =   32
         Top             =   1725
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Bank ID"
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
         Left            =   240
         TabIndex        =   31
         Top             =   307
         Width           =   1695
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   780
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
         TabIndex        =   29
         Top             =   4440
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
         TabIndex        =   28
         Top             =   4920
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   8415
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
         Left            =   6540
         TabIndex        =   7
         Top             =   615
         Width           =   1575
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
         Left            =   4965
         TabIndex        =   23
         Top             =   615
         Width           =   1575
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
         Left            =   3270
         TabIndex        =   22
         Top             =   615
         Width           =   1695
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
         Left            =   1695
         TabIndex        =   21
         Top             =   615
         Width           =   1575
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
         TabIndex        =   20
         Top             =   615
         Width           =   1575
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
         Left            =   6540
         TabIndex        =   19
         Top             =   240
         Width           =   1575
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
         Left            =   4965
         TabIndex        =   18
         Top             =   240
         Width           =   1575
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
         Left            =   3270
         TabIndex        =   17
         Top             =   240
         Width           =   1695
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
         Left            =   1695
         TabIndex        =   6
         Top             =   240
         Width           =   1575
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
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmMst_FAGBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtCreated = ""
TxtUpdated = ""
TxtNAVCode = ""
TxtBankID = ""
TxtBankName = ""
TxtAccountNo = ""
TxtContactPerson = ""
TxtAddress = ""
TxtPaymentLimitation = ""
Frame1.Enabled = True
If RsMst_FAGBank.RecordCount > 0 Then
   TxtBankName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_FAGBank.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_FAGBank.xls")
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
tmp = "Select FAGBankId,BankName,AccountNo,ContactPerson,Address,NAVCode,Flag,Remarks,PaymentLimitation From Mst_FAGBank "

Dim Wc As Variant
Wc = ""
If TxtSearchBank <> "" Then
   Wc = " Where BankName Like '%" & TxtSearchBank & "%' "
End If

tmp = tmp & Wc & " Order by BankName"

Call TmpTable
Call Grid_Head
For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 9

MSHFlexGrid1.TextMatrix(I, 0) = "FAG Bank Id"
MSHFlexGrid1.TextMatrix(I, 1) = "Bank Name"
MSHFlexGrid1.TextMatrix(I, 2) = "Account No."
MSHFlexGrid1.TextMatrix(I, 3) = "Contact Person"
MSHFlexGrid1.TextMatrix(I, 4) = "Address"
MSHFlexGrid1.TextMatrix(I, 5) = "NAV Code"
MSHFlexGrid1.TextMatrix(I, 6) = "Flag"
MSHFlexGrid1.TextMatrix(I, 7) = "Remarks"
MSHFlexGrid1.TextMatrix(I, 8) = "Payment Limitation"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3600
MSHFlexGrid1.ColWidth(2) = 3600
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 1000

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_FAGBank.Delete
   RsMst_FAGBank.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_FAGBank.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_FAGBank.MoveNext
   If RsMst_FAGBank.EOF() Then
      RsMst_FAGBank.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_FAGBank.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
Call GButtonLock(Me, False)
''TxtBankID.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_FAGBank.MoveFirst
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
TxtSearchBank = ""
Label22.Caption = ""

Call TableMst_FAGBank("  Order by BankName ")
If RsMst_FAGBank.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()


TxtBankID.Locked = True
TxtBankID = RsMst_FAGBank!FAGBankId
TxtBankName = RsMst_FAGBank!BankName
TxtAccountNo = RsMst_FAGBank!AccountNo
TxtContactPerson = IIf(IsNull(RsMst_FAGBank!ContactPerson), "", RsMst_FAGBank!ContactPerson)
TxtAddress = IIf(IsNull(RsMst_FAGBank!Address), "", RsMst_FAGBank!Address)
TxtNAVCode = IIf(IsNull(RsMst_FAGBank!NAVCode), "", RsMst_FAGBank!NAVCode)
TxtPaymentLimitation = IIf(IsNull(RsMst_FAGBank!PaymentLimitation), 0, RsMst_FAGBank!PaymentLimitation)
TxtCreated = IIf(IsNull(RsMst_FAGBank!CreatedBy), "", RsMst_FAGBank!CreatedBy) & " :- " & IIf(IsNull(RsMst_FAGBank!CreatedDate), "", RsMst_FAGBank!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_FAGBank!UpdatedBy), "", RsMst_FAGBank!UpdatedBy) & " :- " & IIf(IsNull(RsMst_FAGBank!UpdatedDate), "", RsMst_FAGBank!UpdatedDate)

Frame1.Enabled = False
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_FAGBank.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_FAGBank.MoveFirst

If RsMst_FAGBank.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "FAGBankID = " & Val(mCaId) & ""
If RsMst_FAGBank.EOF Then
   RsMst_FAGBank.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_FAGBank.MoveNext
If RsMst_FAGBank.EOF Then
   RsMst_FAGBank.MoveLast
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
RsMst_FAGBank.MovePrevious
If RsMst_FAGBank.BOF Then
   RsMst_FAGBank.MoveFirst
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

If TxtBankName = "" Then
   MsgBox "Blank Bank Name not Allowed"
   TxtBankName.SetFocus
   Exit Sub
End If

If TxtAccountNo = "" Then
   MsgBox "Blank Account No. not Allowed"
   TxtAccountNo.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_FAGBank.RecordCount > 0 Then
      RsMst_FAGBank.MoveFirst
      RsMst_FAGBank.Find "AccountNo = '" & TxtAccountNo & "'"
      If Not RsMst_FAGBank.EOF Then
         MsgBox "Duplicate Account No. not allowed. Already exist for " & RsMst_FAGBank!BankName & " !!!"
         TxtAccountNo.SetFocus
         Exit Sub
      End If
      RsMst_FAGBank.MoveLast
   End If
   RsMst_FAGBank.AddNew
   RsMst_FAGBank!FAGBankId = GetMaxFAGBankID
   RsMst_FAGBank!G_UID = GetGUID
   TxtBankID = RsMst_FAGBank!FAGBankId
    
Else
   RsMst_FAGBank.MoveFirst
   RsMst_FAGBank.Find "AccountNo = '" & TxtAccountNo & "'"
   If Not RsMst_FAGBank.EOF Then
       If RsMst_FAGBank!FAGBankId <> TxtBankID Then
          MsgBox "Duplicate Account No. not allowed. Already exist for " & RsMst_FAGBank!BankName & " !!!"
          TxtAccountNo.SetFocus
          Exit Sub
       End If
   End If
   RsMst_FAGBank.MoveFirst
   RsMst_FAGBank.Find "FAGBankID = " & TxtBankID & ""
End If
RsMst_FAGBank!BankName = TxtBankName
RsMst_FAGBank!AccountNo = TxtAccountNo
RsMst_FAGBank!ContactPerson = TxtContactPerson
RsMst_FAGBank!Address = TxtAddress
RsMst_FAGBank!NAVCode = TxtNAVCode
RsMst_FAGBank!PaymentLimitation = Val(TxtPaymentLimitation)
If IsNull(RsMst_FAGBank!CreatedBy) = True Or RsMst_FAGBank!CreatedBy = "" Then
   RsMst_FAGBank!CreatedBy = Gen_UserLoginID
   RsMst_FAGBank!CreatedDate = Now
Else
   RsMst_FAGBank!UpdatedBy = Gen_UserLoginID
   RsMst_FAGBank!UpdatedDate = Now
End If

RsMst_FAGBank.Update
RsMst_FAGBank.Requery
RsMst_FAGBank.MoveFirst
RsMst_FAGBank.Find "FAGBankID = " & TxtBankID & ""
If RsMst_FAGBank.EOF Then
   RsMst_FAGBank.MoveFirst
End If
Call GButtonLock(Me, True)
Call Indata
End Sub
Private Function GetMaxFAGBankID() As Double
Dim Retval As Double
tmp = "Select max(FAGBankID) from Mst_FAGBank"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxFAGBankID = Retval
End Function
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_FAGBank.EOF Then
      RsMst_FAGBank.MoveFirst
   End If
   Call Indata

   Exit Sub
End If

Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 300
Frame0.Visible = True


MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000

RsMst_FAGBank.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'Call Grid_Head
'TxtSearchBank = ""
'Label22.Caption = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub
Private Sub TxtBankID_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtBankID)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtBankID.SetFocus
End If
End Sub

Private Sub TxtBankName_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtBankName)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtBankName.SetFocus
End If
End Sub

Private Sub TxtPaymentLimitation_LostFocus()
If TxtPaymentLimitation = "" Then Exit Sub

If IsNumeric(TxtPaymentLimitation) = False Or Val(TxtPaymentLimitation) < 0 Then
   MsgBox "Invalid format!!!"
   TxtPaymentLimitation.SetFocus
   Exit Sub
End If


End Sub
