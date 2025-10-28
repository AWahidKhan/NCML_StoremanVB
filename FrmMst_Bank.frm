VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_Bank 
   Caption         =   "Bank Master"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   4215
      Left            =   240
      TabIndex        =   13
      Top             =   8040
      Visible         =   0   'False
      Width           =   15015
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   7080
         TabIndex        =   37
         Top             =   270
         Width           =   1000
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
         TabIndex        =   36
         Top             =   240
         Width           =   3645
      End
      Begin VB.CommandButton CmdGo 
         Caption         =   "Go"
         Height          =   345
         Left            =   6075
         TabIndex        =   35
         Top             =   270
         Width           =   1005
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2220
         Left            =   75
         TabIndex        =   14
         Top             =   660
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   3916
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         TabIndex        =   40
         Top             =   285
         Width           =   2895
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
         TabIndex        =   39
         Top             =   285
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
         Left            =   11790
         TabIndex        =   38
         Top             =   285
         Width           =   75
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   6360
      Width           =   10020
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2055
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3990
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   6045
         TabIndex        =   29
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7980
         TabIndex        =   11
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2055
         TabIndex        =   27
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3990
         TabIndex        =   26
         Top             =   615
         Width           =   2055
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   6045
         TabIndex        =   25
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7980
         TabIndex        =   12
         Top             =   615
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   10035
      Begin VB.CheckBox ChkMktYard 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8760
         TabIndex        =   45
         Top             =   4470
         Width           =   195
      End
      Begin VB.ComboBox CmbHRBankID 
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
         Left            =   2370
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   4920
         Width           =   7590
      End
      Begin VB.TextBox TxtHRBankID 
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
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   4920
         Width           =   7575
      End
      Begin VB.CheckBox ChkNFin 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5400
         TabIndex        =   8
         Top             =   4470
         Width           =   195
      End
      Begin VB.CheckBox Check1 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   4470
         Width           =   195
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
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   5880
         Width           =   3735
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
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   5400
         Width           =   3735
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
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
         Width           =   7575
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
         Top             =   1200
         Width           =   7575
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
         Top             =   1680
         Width           =   7575
      End
      Begin VB.TextBox TxtMobile 
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
         Top             =   3000
         Width           =   7575
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
         Top             =   3480
         Width           =   7575
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
         Top             =   3960
         Width           =   7575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Market Yard"
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
         Left            =   6600
         TabIndex        =   46
         Top             =   4470
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "HR Bank"
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
         TabIndex        =   44
         Top             =   4980
         Width           =   1695
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "NCML Finance"
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
         Left            =   3240
         TabIndex        =   42
         Top             =   4470
         Width           =   1320
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CM Agreement"
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
         TabIndex        =   41
         Top             =   4470
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   34
         Top             =   5940
         Width           =   885
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Left            =   240
         TabIndex        =   33
         Top             =   5460
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   787
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         TabIndex        =   21
         Top             =   307
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1267
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   2167
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   3067
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   3547
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   4027
         Width           =   720
      End
   End
End
Attribute VB_Name = "FrmMst_Bank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mHRBankID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtCreated = ""
TxtUpdated = ""

TxtBankID = ""
TxtBankName = ""
TxtContactPerson = ""
TxtAddress = ""
TxtMobile = ""
TxtTelephone2 = ""
TxtEmailId = ""
Frame1.Enabled = True
Check1.Value = 0
ChkNFin.Value = 0
ChkMktYard.Value = 0
ChkNFin.Enabled = False
ChkMktYard.Enabled = False
TxtHRBankID.Visible = False
CmbHRBankID.Visible = True
CmbHRBankID.Text = ""

If LCase(Gen_UserRole) = "power" Then
   ChkNFin.Enabled = True
   ChkMktYard.Enabled = True
End If
If RsMst_Bank.RecordCount > 0 Then
   TxtBankName.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_Bank.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_Bank.xls")
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
tmp = "Select BankId,BankName,ContactPerson,Address,Tel1,Tel2,EmailID,Flag,Remarks, CMAgreement,NCMLFinance,isnull(HRBankID,0)'HRBankID',MarketYard From Mst_Bank "
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
    If MSHFlexGrid1.TextMatrix(iu, 9) = "1" Then
        MSHFlexGrid1.TextMatrix(iu, 9) = "Yes"
    Else
        MSHFlexGrid1.TextMatrix(iu, 9) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 10) = "1" Then
        MSHFlexGrid1.TextMatrix(iu, 10) = "Yes"
    Else
        MSHFlexGrid1.TextMatrix(iu, 10) = "No"
    End If
    If MSHFlexGrid1.TextMatrix(iu, 12) = "1" Then
        MSHFlexGrid1.TextMatrix(iu, 12) = "Yes"
    Else
        MSHFlexGrid1.TextMatrix(iu, 12) = "No"
    End If
    MSHFlexGrid1.TextMatrix(iu, 11) = GetHRBank(MSHFlexGrid1.TextMatrix(iu, 11))
    TmpRs.MoveNext
Next
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
End Sub
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13

MSHFlexGrid1.TextMatrix(I, 0) = "BankId"
MSHFlexGrid1.TextMatrix(I, 1) = "BankName"
MSHFlexGrid1.TextMatrix(I, 2) = "ContactPerson"
MSHFlexGrid1.TextMatrix(I, 3) = "Address"
MSHFlexGrid1.TextMatrix(I, 4) = "Tel1"
MSHFlexGrid1.TextMatrix(I, 5) = "Tel2"
MSHFlexGrid1.TextMatrix(I, 6) = "EmailID"
MSHFlexGrid1.TextMatrix(I, 7) = "Flag"
MSHFlexGrid1.TextMatrix(I, 8) = "Remarks"
MSHFlexGrid1.TextMatrix(I, 9) = "CM Agreement"
MSHFlexGrid1.TextMatrix(I, 10) = "NCML Finance"
MSHFlexGrid1.TextMatrix(I, 11) = "HR Bank Name"
MSHFlexGrid1.TextMatrix(I, 12) = "Market Yard Flag"

MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3600
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 1000
MSHFlexGrid1.ColWidth(7) = 800
MSHFlexGrid1.ColWidth(8) = 1000
MSHFlexGrid1.ColWidth(9) = 1000
MSHFlexGrid1.ColWidth(10) = 1000

End Sub
Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_Bank.Delete
   RsMst_Bank.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_Bank.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_Bank.MoveNext
   If RsMst_Bank.EOF() Then
      RsMst_Bank.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Bank.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame1.Enabled = True
ChkNFin.Enabled = False
ChkMktYard.Enabled = False
If LCase(Gen_UserRole) = "power" Then
   ChkNFin.Enabled = True
   ChkMktYard.Enabled = True
End If
TxtHRBankID.Visible = False
CmbHRBankID.Visible = True
CmbHRBankID.Text = TxtHRBankID

Call GButtonLock(Me, False)
TxtBankID.Locked = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_Bank.MoveFirst
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

Call TableMst_Bank("  Order by BankName ")
If RsMst_Bank.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)

End Sub

Public Sub Indata()

TxtHRBankID.Visible = True
CmbHRBankID.Visible = False

TxtBankID.Locked = True
TxtBankID = RsMst_Bank!BankID
TxtBankName = RsMst_Bank!BankName
TxtContactPerson = IIf(IsNull(RsMst_Bank!ContactPerson), "", RsMst_Bank!ContactPerson)
TxtAddress = IIf(IsNull(RsMst_Bank!Address), "", RsMst_Bank!Address)
TxtMobile = IIf(IsNull(RsMst_Bank!Tel1), "", RsMst_Bank!Tel1)
TxtTelephone2 = IIf(IsNull(RsMst_Bank!Tel2), "", RsMst_Bank!Tel2)
TxtEmailId = IIf(IsNull(RsMst_Bank!EmailID), "", RsMst_Bank!EmailID)
Check1.Value = IIf(IsNull(RsMst_Bank!CMAgreement), 0, RsMst_Bank!CMAgreement)
ChkNFin.Value = IIf(IsNull(RsMst_Bank!NCMLFinance), 0, RsMst_Bank!NCMLFinance)
ChkMktYard.Value = IIf(IsNull(RsMst_Bank!MarketYard), 0, RsMst_Bank!MarketYard)
TxtCreated = IIf(IsNull(RsMst_Bank!CreatedBy), "", RsMst_Bank!CreatedBy) & " :- " & IIf(IsNull(RsMst_Bank!CreatedDate), "", RsMst_Bank!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_Bank!UpdatedBy), "", RsMst_Bank!UpdatedBy) & " :- " & IIf(IsNull(RsMst_Bank!UpdatedDate), "", RsMst_Bank!UpdatedDate)
mHRBankID = IIf(IsNull(RsMst_Bank!HRBankID), 0, RsMst_Bank!HRBankID)
TxtHRBankID = GetHRBank(mHRBankID)


Frame1.Enabled = False
Call GButtonLock(Me, True)
End Sub

Private Sub LastCmd_Click()
RsMst_Bank.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel

mCaId = MSHFlexGrid1.TextMatrix(x, 0)

RsMst_Bank.MoveFirst

If RsMst_Bank.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankID = " & Val(mCaId) & ""
If RsMst_Bank.EOF Then
   RsMst_Bank.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_Bank.MoveNext
If RsMst_Bank.EOF Then
   RsMst_Bank.MoveLast
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
RsMst_Bank.MovePrevious
If RsMst_Bank.BOF Then
   RsMst_Bank.MoveFirst
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
'If TxtAddress = "" Then
'   MsgBox "Blank Bank Address not Allowed"
'   TxtAddress.SetFocus
'   Exit Sub
'End If
If CmbHRBankID.Text = "" Then
   MsgBox "Blank HR Bank Not Allowed !!!"
   If CmbHRBankID.Visible = True Then CmbHRBankID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_Bank.RecordCount > 0 Then
      RsMst_Bank.MoveFirst
      RsMst_Bank.Find "BankName = '" & TxtBankName & "'"
      If Not RsMst_Bank.EOF Then
         MsgBox "Duplicate Bank Name not Allowed"
         Exit Sub
      End If
      RsMst_Bank.MoveLast
   End If
   RsMst_Bank.AddNew
   RsMst_Bank!BankID = GetBankID
   RsMst_Bank!G_UID = GetGUID
   TxtBankID = RsMst_Bank!BankID
    
Else
   RsMst_Bank.MoveFirst
   RsMst_Bank.Find "BankName = '" & TxtBankName & "'"
   If Not RsMst_Bank.EOF Then
       If RsMst_Bank!BankID <> TxtBankID Then
          MsgBox "Duplicate Bank Name not Allowed"
          Exit Sub
       End If
   End If
   RsMst_Bank.MoveFirst
   RsMst_Bank.Find "BankID = " & TxtBankID & ""
End If
RsMst_Bank!BankName = TxtBankName
RsMst_Bank!ContactPerson = TxtContactPerson
RsMst_Bank!Address = TxtAddress
RsMst_Bank!Tel1 = TxtMobile
RsMst_Bank!Tel2 = TxtTelephone2
RsMst_Bank!EmailID = TxtEmailId
RsMst_Bank!CMAgreement = Check1.Value
RsMst_Bank!NCMLFinance = ChkNFin.Value
RsMst_Bank!MarketYard = ChkMktYard.Value
RsMst_Bank!HRBankID = mHRBankID
If IsNull(RsMst_Bank!CreatedBy) = True Or RsMst_Bank!CreatedBy = "" Then
   RsMst_Bank!CreatedBy = Gen_UserLoginID
   RsMst_Bank!CreatedDate = Now
Else
   RsMst_Bank!UpdatedBy = Gen_UserLoginID
   RsMst_Bank!UpdatedDate = Now
End If

RsMst_Bank.Update
RsMst_Bank.Requery
RsMst_Bank.MoveFirst
RsMst_Bank.Find "BankID = " & TxtBankID & ""
If RsMst_Bank.EOF Then
   RsMst_Bank.MoveFirst
End If
Call GButtonLock(Me, True)
TxtBankID.Locked = True
Call Indata
End Sub
Private Function GetBankID() As Double
Dim Retval As Double
tmp = "Select max(BankID) from Mst_Bank"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetBankID = Retval
End Function
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   Call GButtonLock(Me, True)
   If RsMst_Bank.EOF Then
      RsMst_Bank.MoveFirst
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

RsMst_Bank.MoveFirst

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


Private Sub CmbHRBankID_GotFocus()

tmp1 = CmbHRBankID.Text
tmp = "select * from Mst_Bank "

Call TmpTableHRPro

CmbHRBankID.Clear
If TmpRsHRPro.RecordCount = 0 Then
   MsgBox "No HR Bank found!!!"
   CmbHRBankID.SetFocus
   Exit Sub
End If
For I = 1 To TmpRsHRPro.RecordCount
    CmbHRBankID.AddItem TmpRsHRPro!BankName
    TmpRsHRPro.MoveNext
Next
CmbHRBankID.Text = tmp1
End Sub

Private Sub CmbHRBankID_LostFocus()
If CmbHRBankID.Text = "" Then
   Exit Sub
End If
TmpRsHRPro.MoveFirst
TmpRsHRPro.Find "BankName = '" & CmbHRBankID.Text & "'"

If TmpRsHRPro.EOF Then
   MsgBox "Invalid HR Bank selection "
   CmbHRBankID.SetFocus
   Exit Sub
End If
mHRBankID = TmpRsHRPro!BankID
End Sub

