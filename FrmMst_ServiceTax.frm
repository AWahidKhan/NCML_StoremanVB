VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ServiceTax 
   Caption         =   "GST Group Code"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   6750
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2415
      Left            =   6840
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   150
         TabIndex        =   18
         Top             =   195
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1740
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3069
         _Version        =   393216
         Rows            =   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6495
      Begin VB.TextBox TxtInvoiceType 
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
         MaxLength       =   20
         TabIndex        =   4
         Top             =   2160
         Width           =   3735
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
         ItemData        =   "FrmMst_ServiceTax.frx":0000
         Left            =   2400
         List            =   "FrmMst_ServiceTax.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1680
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
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   2640
         Width           =   3735
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
         Top             =   3120
         Width           =   3735
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
         TabIndex        =   2
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox TxtServiceTaxID 
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtSerTaxDescription 
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
         Width           =   3735
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label Label4 
         Caption         =   "Invoice Type"
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
         Left            =   240
         TabIndex        =   29
         Top             =   2190
         Width           =   1695
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
         Left            =   240
         TabIndex        =   28
         Top             =   1710
         Width           =   600
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
         TabIndex        =   26
         Top             =   2700
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
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3180
         Width           =   1695
      End
      Begin VB.Label Label3 
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
         Height          =   300
         Left            =   240
         TabIndex        =   14
         Top             =   1237
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "GST Group Code ID"
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
         TabIndex        =   12
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Description"
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
         TabIndex        =   10
         Top             =   780
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   6480
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   1335
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   2550
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3885
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   5100
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   1335
         TabIndex        =   17
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   2550
         TabIndex        =   19
         Top             =   615
         Width           =   1335
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   3885
         TabIndex        =   21
         Top             =   615
         Width           =   1215
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   5100
         TabIndex        =   22
         Top             =   615
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmMst_ServiceTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Private Sub Form_Load()
'TxtExpenseGroupID
'TxtServiceTaxID
'
'TxtSerTaxDescription
'TxtExpenseGroup
'End Sub
Dim Edit_Flg As Variant
Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtSerTaxDescription.Locked = False
TxtNAVCode.Locked = False
TxtSerTaxDescription = ""
TxtServiceTaxID = ""
TxtNAVCode = ""
TxtCreated = ""
TxtUpdated = ""
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = ""
TxtInvoiceType.Locked = False
TxtInvoiceType = ""


If RsMst_ServiceTax.RecordCount > 0 Then
   TxtSerTaxDescription.SetFocus
Else
   SearchCmd.Enabled = False
End If
End Sub

Private Sub CmbFlag_LostFocus()
If LCase(CmbFlag.Text) = "" Then Exit Sub

If LCase(CmbFlag.Text) <> "active" And LCase(CmbFlag.Text) <> "de-active" Then
   MsgBox "Invalid Selection"
   CmbFlag.SetFocus
   Exit Sub
End If

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_GSTGroupCode.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_GSTGroupCode.xls")
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
Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record ", vbYesNo)
If ans = vbYes Then
   'tmp = "Select * from Mst_Employee Where (ServiceTaxID = " & TxtServiceTaxID & ")"
   'Call Tmp1Table
   'If TmpRs1.RecordCount > 0 Then
   '   MsgBox "Child Record Present!!!"
   '   Exit Sub
   'End If
   RsMst_ServiceTax.Delete
   RsMst_ServiceTax.Update
   'MsgBox "Record Successfully Removed"
   If RsMst_ServiceTax.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ServiceTax.MoveNext
   If RsMst_ServiceTax.EOF() Then
      RsMst_ServiceTax.MoveLast
   End If
   Call Indata
   'TmpRs1.Close
End If

Exit Sub
msgError:
MsgBox "Child Record Present "
RsMst_ServiceTax.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtSerTaxDescription.Locked = False
TxtNAVCode.Locked = False
TxtFlag.Visible = False
CmbFlag.Visible = True
CmbFlag.Text = TxtFlag.Text
TxtInvoiceType.Locked = False



TxtSerTaxDescription.SetFocus
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ServiceTax.MoveFirst
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

Call TableMst_ServiceTax(" Order by ServiceTaxID ")
Frame1.Enabled = True
If RsMst_ServiceTax.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
Call Indata
Call GButtonLock(Me, True)


End Sub

Public Sub Indata()
TxtSerTaxDescription.Locked = True
TxtNAVCode.Locked = True
TxtInvoiceType.Locked = True
TxtServiceTaxID = RsMst_ServiceTax!ServiceTaxID
TxtSerTaxDescription = RsMst_ServiceTax!Description
TxtNAVCode = IIf(IsNull(RsMst_ServiceTax!NavCode), "", RsMst_ServiceTax!NavCode)
TxtInvoiceType = IIf(IsNull(RsMst_ServiceTax!InvoiceType), "", RsMst_ServiceTax!InvoiceType)

CmbFlag.Visible = False
TxtFlag.Visible = True
If LCase(RsMst_ServiceTax!Flag) = "a" Then
   TxtFlag.Text = "Active"
ElseIf LCase(RsMst_ServiceTax!Flag) = "d" Then
   TxtFlag.Text = "De-Active"
Else
   TxtFlag.Text = ""
End If

TxtCreated = IIf(IsNull(RsMst_ServiceTax!CreatedBy), "", RsMst_ServiceTax!CreatedBy) & " :- " & IIf(IsNull(RsMst_ServiceTax!CreatedDate), "", RsMst_ServiceTax!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ServiceTax!UpdatedBy), "", RsMst_ServiceTax!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ServiceTax!UpdatedDate), "", RsMst_ServiceTax!UpdatedDate)
Call GButtonLock(Me, True)
End Sub



Private Sub LastCmd_Click()
RsMst_ServiceTax.MoveLast
Call Indata
End Sub

Private Sub MSHFlexGrid1_Click()
Dim x As Variant
Dim mCaId As Variant

x = MSHFlexGrid1.RowSel
mCaId = MSHFlexGrid1.TextMatrix(x, 0)
RsMst_ServiceTax.MoveFirst

If RsMst_ServiceTax.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "ServiceTaxID = " & Val(mCaId) & ""
If RsMst_ServiceTax.EOF Then
   RsMst_ServiceTax.MoveLast
End If
Frame0.Visible = False
Call Indata

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ServiceTax.MoveNext
If RsMst_ServiceTax.EOF Then
   RsMst_ServiceTax.MoveLast
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
RsMst_ServiceTax.MovePrevious
If RsMst_ServiceTax.BOF Then
   RsMst_ServiceTax.MoveFirst
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

If TxtSerTaxDescription = "" Then
   MsgBox "Blank GST Group Code Description  not allowed !!!"
   TxtSerTaxDescription.SetFocus
   Exit Sub
End If

If TxtInvoiceType = "" Then
   MsgBox "Blank Invoice Type  not allowed !!!"
   TxtInvoiceType.SetFocus
   Exit Sub
End If

If CmbFlag.Text = "" Then
   MsgBox "Blank Flag Not Allowed !!! "
   CmbFlag.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_ServiceTax.RecordCount = 0 Then
      tmp = 1
   Else
        RsMst_ServiceTax.MoveFirst
        RsMst_ServiceTax.Find "Description = '" & TxtSerTaxDescription & "'"
        If Not RsMst_ServiceTax.EOF Then
           MsgBox "Duplicate Description not Allowed"
           Exit Sub
        End If
        RsMst_ServiceTax.MoveLast
        
        tmp = "select max(ServiceTaxID) from Mst_ServiceTax"
        Call TmpTable
        tmp = IIf(IsNull(TmpRs.Fields(0)), 1, TmpRs.Fields(0) + 1)
        
    End If
    RsMst_ServiceTax.AddNew
    RsMst_ServiceTax!ServiceTaxID = tmp
    RsMst_ServiceTax!G_UID = GetGUID
    TxtServiceTaxID = tmp
Else
    RsMst_ServiceTax.MoveFirst
    RsMst_ServiceTax.Find "Description = '" & TxtSerTaxDescription & "'"
    If Not RsMst_ServiceTax.EOF Then
       If RsMst_ServiceTax!ServiceTaxID <> TxtServiceTaxID Then
          MsgBox "Duplicate GST Group Code not Allowed"
          Exit Sub
       End If
    End If
   RsMst_ServiceTax.MoveFirst
   RsMst_ServiceTax.Find "ServiceTaxID = " & TxtServiceTaxID & ""
End If
RsMst_ServiceTax!Description = TxtSerTaxDescription
RsMst_ServiceTax!NavCode = TxtNAVCode
RsMst_ServiceTax!Flag = Left(CmbFlag.Text, 1)
RsMst_ServiceTax!InvoiceType = Trim(TxtInvoiceType)

If IsNull(RsMst_ServiceTax!CreatedBy) = True Or RsMst_ServiceTax!CreatedBy = "" Then
   RsMst_ServiceTax!CreatedBy = Gen_UserLoginID
   RsMst_ServiceTax!CreatedDate = Now
Else
   RsMst_ServiceTax!UpdatedBy = Gen_UserLoginID
   RsMst_ServiceTax!UpdatedDate = Now 'Format(Date, "yyyy-mm-dd hh:mm:ss")
End If
RsMst_ServiceTax.Update
RsMst_ServiceTax.Requery
RsMst_ServiceTax.MoveFirst
RsMst_ServiceTax.Find "ServiceTaxID = " & TxtServiceTaxID & ""

Call GButtonLock(Me, True)
TxtSerTaxDescription.Locked = True
TxtNAVCode.Locked = True
Call Indata
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_ServiceTax.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   If RsMst_ServiceTax.EOF Then
      RsMst_ServiceTax.MoveFirst
   End If
   Call Indata
   Exit Sub
End If


Frame0.Height = Me.Height - 700
Frame0.Width = Me.Width - 500
Frame0.Left = 120
Frame0.Top = 120

Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.Width = Frame0.Width - 500
MSHFlexGrid1.Height = Frame0.Height - 700

RsMst_ServiceTax.MoveFirst
MSHFlexGrid1.Rows = RsMst_ServiceTax.RecordCount + 1
MSHFlexGrid1.TextMatrix(0, 0) = "GST Group Code ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Description"
MSHFlexGrid1.TextMatrix(0, 2) = "NAV Code"
MSHFlexGrid1.TextMatrix(0, 3) = "Flag"

MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 5500
MSHFlexGrid1.ColWidth(2) = 3500
MSHFlexGrid1.ColWidth(3) = 2500

For I = 1 To RsMst_ServiceTax.RecordCount
    MSHFlexGrid1.TextMatrix(I, 0) = RsMst_ServiceTax!ServiceTaxID
    MSHFlexGrid1.TextMatrix(I, 1) = RsMst_ServiceTax!Description
    MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(RsMst_ServiceTax!NavCode), "", RsMst_ServiceTax!NavCode)
    MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(RsMst_ServiceTax!Flag), "", RsMst_ServiceTax!Flag)
    RsMst_ServiceTax.MoveNext
Next

End Sub

Private Sub TxtSerTaxDescription_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtSerTaxDescription")
End If
End Sub

Private Sub TxtSerTaxDescription_LostFocus()
Dim ans As Boolean
ans = IsClean(TxtSerTaxDescription)
If ans = False Then
   MsgBox "Special character not Allowed  !!!!!!!! "
   TxtSerTaxDescription.SetFocus
End If
End Sub


Private Sub TxtServiceTaxID_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtServiceTaxID")
End If
End Sub
Private Sub TxtNAVCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 4 Then
   Call AssignHelpData(Me, "TxtNAVCode")
End If
End Sub


