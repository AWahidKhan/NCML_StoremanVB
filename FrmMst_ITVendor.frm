VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMst_ITVendor 
   Caption         =   "IT Vendor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2535
      Left            =   10320
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1380
         Left            =   105
         TabIndex        =   14
         Top             =   660
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2434
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   22
      Top             =   30
      Width           =   9615
      Begin VB.TextBox TxtVendorID 
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
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TxtVendorName 
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
         Left            =   2415
         MaxLength       =   200
         TabIndex        =   1
         Top             =   660
         Width           =   6600
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
         Left            =   2415
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1080
         Width           =   6615
      End
      Begin VB.TextBox TxtMobileNo 
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
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   6
         Top             =   3615
         Width           =   6615
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
         Left            =   2415
         MaxLength       =   50
         TabIndex        =   8
         Top             =   4455
         Width           =   6615
      End
      Begin VB.TextBox TxtFax 
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
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   7
         Top             =   4035
         Width           =   6615
      End
      Begin VB.TextBox TxtContactNo 
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
         Left            =   2415
         MaxLength       =   100
         TabIndex        =   5
         Top             =   3195
         Width           =   6615
      End
      Begin VB.CommandButton CmdLocationID 
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
         Height          =   315
         Left            =   9105
         TabIndex        =   26
         Top             =   2370
         Width           =   375
      End
      Begin VB.ComboBox CmbLocationID 
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
         Left            =   2415
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2355
         Width           =   6645
      End
      Begin VB.CommandButton CmdTDSID 
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
         Height          =   315
         Left            =   9120
         TabIndex        =   25
         Top             =   2775
         Width           =   375
      End
      Begin VB.ComboBox CmbTDSID 
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
         Left            =   2415
         TabIndex        =   4
         Top             =   2775
         Width           =   6630
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
         Top             =   5400
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
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4920
         Width           =   3735
      End
      Begin VB.TextBox TxtLocationID 
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
         Left            =   2415
         TabIndex        =   28
         Top             =   2355
         Width           =   6615
      End
      Begin VB.TextBox TxtTDSID 
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
         Left            =   2415
         MaxLength       =   25
         TabIndex        =   29
         Top             =   2775
         Width           =   6615
      End
      Begin VB.Label LblVendorName 
         Caption         =   "Vendor Name"
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
         TabIndex        =   40
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label LblVendorID 
         Caption         =   "Vendor ID"
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
         TabIndex        =   39
         Top             =   292
         Width           =   1695
      End
      Begin VB.Label LblAdess 
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
         Height          =   240
         Left            =   255
         TabIndex        =   38
         Top             =   1567
         Width           =   1695
      End
      Begin VB.Label LblMobileNo 
         Caption         =   "Mobile No"
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
         TabIndex        =   37
         Top             =   3682
         Width           =   1695
      End
      Begin VB.Label LblEmailId 
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
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   4515
         Width           =   1695
      End
      Begin VB.Label LblFax 
         Caption         =   "Fax"
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
         TabIndex        =   35
         Top             =   4095
         Width           =   1695
      End
      Begin VB.Label LblContactNo 
         Caption         =   "Contact No"
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
         TabIndex        =   34
         Top             =   3255
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Location"
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
         TabIndex        =   33
         Top             =   2415
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "TDS Description"
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
         TabIndex        =   32
         Top             =   2835
         Width           =   2115
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
         TabIndex        =   31
         Top             =   5460
         Width           =   1695
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
         TabIndex        =   30
         Top             =   4980
         Width           =   1695
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1095
      Left            =   135
      TabIndex        =   15
      Top             =   5910
      Width           =   9600
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   375
         Left            =   2055
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   375
         Left            =   3870
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5805
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
         Height          =   375
         Left            =   7620
         TabIndex        =   11
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   375
         Left            =   2055
         TabIndex        =   18
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   375
         Left            =   3870
         TabIndex        =   17
         Top             =   615
         Width           =   1935
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   375
         Left            =   5805
         TabIndex        =   16
         Top             =   615
         Width           =   1815
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   375
         Left            =   7620
         TabIndex        =   10
         Top             =   615
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmMst_ITVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mTDSID, mLocationID As Variant

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtVendorName.Locked = False
TxtAddress.Locked = False
TxtLocationID.Locked = True
TxtTDSID.Locked = True
TxtContactNo.Locked = False
TxtMobileNo.Locked = False
TxtFax.Locked = False
TxtEmailID.Locked = False

TxtLocationID.Visible = False
CmbLocationID.Visible = True

TxtTDSID.Visible = False
CmbTDSID.Visible = True

TxtCreated = ""
TxtUpdated = ""

TxtVendorID.Text = ""
TxtVendorName.Text = ""
TxtAddress.Text = ""
TxtContactNo.Text = ""
TxtLocationID.Text = ""
TxtTDSID.Text = ""
TxtMobileNo.Text = ""
TxtFax.Text = ""
TxtEmailID.Text = ""
CmbLocationID.Text = ""
CmbTDSID.Text = ""

If RsMst_ITVendor.RecordCount > 0 Then
   TxtVendorName.SetFocus
End If

End Sub
Private Sub CmbTDSID_GotFocus()
tmp = CmbTDSID.Text
Call TableMst_TDS
CmbTDSID.Clear
If RsMst_TDS.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_TDS.RecordCount
    CmbTDSID.AddItem RsMst_TDS!Description
    RsMst_TDS.MoveNext
Next
CmbTDSID.Text = tmp
End Sub

Private Sub CmbTDSID_LostFocus()
If CmbTDSID.Text = "" Then
   Exit Sub
End If
RsMst_TDS.MoveFirst
RsMst_TDS.Find "Description = '" & CmbTDSID.Text & "'"
If RsMst_TDS.EOF Then
   MsgBox "Invalid selection "
   CmbTDSID.SetFocus
   Exit Sub
End If
mTDSID = RsMst_TDS!TDSID
End Sub
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ITVendor.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ITVendor.xls")
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

Private Sub CmdLocationID_Click()
FrmMst_Location.Show
End Sub

Private Sub CmdTDSID_Click()
FrmMst_TDS.Show
End Sub

Private Sub DeleteCmd_Click()
Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
If ans = vbYes Then
   
   RsMst_ITVendor.Delete
   RsMst_ITVendor.Update
   If RsMst_ITVendor.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ITVendor.MoveNext
   If RsMst_ITVendor.EOF() Then
      RsMst_ITVendor.MoveLast
   End If
   Call Indata
End If

Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_Agent.CancelUpdate


End Sub

Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtVendorName.Locked = False
TxtAddress.Locked = False
TxtLocationID.Visible = False
CmbLocationID.Visible = True
CmbLocationID.Text = TxtLocationID.Text
CmdLocationID.Visible = True

TxtTDSID.Visible = False
CmbTDSID.Visible = True
CmdTDSID.Visible = True
CmbTDSID.Text = TxtTDSID.Text

TxtEmailID.Locked = False
TxtMobileNo.Locked = False
TxtContactNo.Locked = False
TxtFax.Locked = False
TxtVendorName.SetFocus
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub FirstCmd_Click()
RsMst_ITVendor.MoveFirst
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
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_ITVendor

Frame1.Enabled = True
If RsMst_ITVendor.RecordCount = 0 Then
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
'Call Grid_Head1
End Sub

Private Sub LastCmd_Click()
RsMst_ITVendor.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_ITVendor.MoveFirst
RsMst_ITVendor.Find "VendorID = " & MSHFlexGrid1.TextMatrix(I, 0)
If Not RsMst_ITVendor.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ITVendor.MoveNext
If RsMst_ITVendor.EOF Then
   RsMst_ITVendor.MoveLast
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
RsMst_ITVendor.MovePrevious
If RsMst_ITVendor.BOF Then
   RsMst_ITVendor.MoveFirst
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

If TxtVendorName.Text = "" Then
    MsgBox "Blank Vendor Name Not Allowed!!!"
    TxtVendorName.SetFocus
    Exit Sub
End If
 
If TxtAddress.Text = "" Then
    MsgBox "Blank Address Not Allowed!!!"
    TxtAddress.SetFocus
    Exit Sub
End If

If CmbLocationID.Text = "" Then
    MsgBox "Blank Location Name Not Allowed!!!"
    CmbLocationID.SetFocus
    Exit Sub
End If

If CmbTDSID.Text = "" Then
   MsgBox "Blank TDS Description Not Allowed!!!"
   CmbTDSID.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
    If RsMst_ITVendor.RecordCount > 0 Then
      RsMst_ITVendor.MoveFirst
      RsMst_ITVendor.Find "VendorName= '" & TxtVendorName.Text & "'"
      If Not RsMst_ITVendor.EOF Then
         MsgBox "Duplicate Vendor Name Not Allowed !!! "
         TxtVendorName.SetFocus
         Exit Sub
      End If
      RsMst_ITVendor.MoveFirst
   End If
   
   RsMst_ITVendor.AddNew
   RsMst_ITVendor!VendorID = GetVendorID
   RsMst_ITVendor!G_UID = GetGUID
   TxtVendorID.Text = RsMst_ITVendor!VendorID
Else
    RsMst_ITVendor.MoveFirst
    RsMst_ITVendor.Find "VendorName='" & TxtVendorName.Text & "'  "
    If Not RsMst_ITVendor.EOF Then
         If RsMst_ITVendor!VendorID <> TxtVendorID.Text Then
          MsgBox "Duplicate Vendor Name Not Allowed !!! "
          TxtVendorName.SetFocus
          Exit Sub
       End If
    End If
    RsMst_ITVendor.MoveFirst
    RsMst_ITVendor.Find " VendorID= " & TxtVendorID.Text
End If

RsMst_ITVendor!VendorID = TxtVendorID.Text
RsMst_ITVendor!VendorName = TxtVendorName.Text
RsMst_ITVendor!VendorAddress = TxtAddress.Text
RsMst_ITVendor!LocationID = mLocationID
RsMst_ITVendor!ContactNo = TxtContactNo.Text
RsMst_ITVendor!MobileNo = TxtMobileNo.Text
RsMst_ITVendor!Fax = TxtFax.Text
RsMst_ITVendor!EmailID = TxtEmailID.Text
RsMst_ITVendor!TDSID = mTDSID
RsMst_ITVendor!Flag = "0" 'TxtFlag.Text
If IsNull(RsMst_ITVendor!CreatedBy) = True Or RsMst_ITVendor!CreatedBy = "" Then
   RsMst_ITVendor!CreatedBy = Gen_UserLoginID
   RsMst_ITVendor!CreatedDate = Now
Else
   RsMst_ITVendor!UpdatedBy = Gen_UserLoginID
   RsMst_ITVendor!UpdatedDate = Now
End If


RsMst_ITVendor.Update
Call Indata
End Sub
Private Function GetVendorID() As Double
Dim Retval As Double
tmp = "Select max(VendorID) from Mst_ITVendor"
Call TmpTable

Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))

If Retval = 0 Then
   
End If
Retval = Retval + 1
GetVendorID = Retval
End Function

Public Sub Indata()

TxtVendorID.Locked = True
TxtVendorName.Locked = True
TxtLocationID.Locked = True
CmbLocationID.Visible = False
CmdLocationID.Visible = False
TxtLocationID.Visible = True

TxtTDSID.Locked = True
CmbTDSID.Visible = False
CmdTDSID.Visible = False
TxtTDSID.Visible = True

TxtAddress.Locked = True
TxtContactNo.Locked = True
TxtMobileNo.Locked = True
TxtFax.Locked = True
TxtEmailID.Locked = True

mLocationID = RsMst_ITVendor!LocationID
mTDSID = RsMst_ITVendor!TDSID

TxtVendorID = IIf(IsNull(RsMst_ITVendor!VendorID), "", RsMst_ITVendor!VendorID)
TxtVendorName = IIf(IsNull(RsMst_ITVendor!VendorName), "", RsMst_ITVendor!VendorName)
TxtLocationID = IIf(IsNull(RsMst_ITVendor!LocationID), "", GetLocationName(RsMst_ITVendor!LocationID))
TxtAddress = IIf(IsNull(RsMst_ITVendor!VendorAddress), "", RsMst_ITVendor!VendorAddress)
TxtContactNo = IIf(IsNull(RsMst_ITVendor!ContactNo), "", RsMst_ITVendor!ContactNo)
TxtMobileNo = IIf(IsNull(RsMst_ITVendor!MobileNo), "", RsMst_ITVendor!MobileNo)
TxtFax = IIf(IsNull(RsMst_ITVendor!Fax), "", RsMst_ITVendor!Fax)
TxtEmailID = IIf(IsNull(RsMst_ITVendor!EmailID), "", RsMst_ITVendor!EmailID)
TxtTDSID.Text = GetTDSDescription(RsMst_ITVendor!TDSID)
TxtCreated = IIf(IsNull(RsMst_ITVendor!CreatedBy), "", RsMst_ITVendor!CreatedBy) & " :- " & IIf(IsNull(RsMst_ITVendor!CreatedDate), "", RsMst_ITVendor!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ITVendor!UpdatedBy), "", RsMst_ITVendor!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ITVendor!UpdatedDate), "", RsMst_ITVendor!UpdatedDate)


If RsMst_ITVendor.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.TextMatrix(0, 0) = "Vendor ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Vendor Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "TDS"
MSHFlexGrid1.TextMatrix(0, 4) = "Contact No"
MSHFlexGrid1.TextMatrix(0, 5) = "Mobile No"
MSHFlexGrid1.TextMatrix(0, 6) = "Fax No"
MSHFlexGrid1.TextMatrix(0, 7) = "Email ID"


MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3400
MSHFlexGrid1.ColWidth(2) = 1800
MSHFlexGrid1.ColWidth(3) = 1800
MSHFlexGrid1.ColWidth(4) = 1600
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1800
MSHFlexGrid1.ColWidth(7) = 1800
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_ITVendor.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 250
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1000
RsMst_ITVendor.MoveFirst
Call Grid_Head
MSHFlexGrid1.Rows = RsMst_ITVendor.RecordCount + 1

For I = 1 To RsMst_ITVendor.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(I, 0) = RsMst_ITVendor!VendorID
        MSHFlexGrid1.TextMatrix(I, 1) = RsMst_ITVendor!VendorName
        MSHFlexGrid1.TextMatrix(I, 2) = GetLocationName(RsMst_ITVendor!LocationID)
        MSHFlexGrid1.TextMatrix(I, 3) = GetTDSDescription(RsMst_ITVendor!TDSID)
        MSHFlexGrid1.TextMatrix(I, 4) = RsMst_ITVendor!ContactNo
        MSHFlexGrid1.TextMatrix(I, 5) = RsMst_ITVendor!MobileNo
        MSHFlexGrid1.TextMatrix(I, 6) = RsMst_ITVendor!Fax
        MSHFlexGrid1.TextMatrix(I, 7) = RsMst_ITVendor!EmailID
    Next
    RsMst_ITVendor.MoveNext
Next
End Sub

