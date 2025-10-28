VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMst_NAVCodeMapping 
   Caption         =   "Navision Code Mapping"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8010
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   8010
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15015
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
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
         Left            =   11670
         TabIndex        =   14
         Top             =   240
         Width           =   1260
      End
      Begin VB.TextBox TxtSearchID 
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
         TabIndex        =   10
         Top             =   720
         Width           =   5280
      End
      Begin VB.CommandButton CmdSave2DB 
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
         Left            =   13905
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Left            =   12915
         TabIndex        =   8
         Top             =   240
         Width           =   975
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
         Height          =   375
         Left            =   10695
         TabIndex        =   7
         Top             =   240
         Width           =   975
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
         Height          =   375
         Left            =   9720
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pending"
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
         Left            =   8040
         TabIndex        =   5
         Top             =   293
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All"
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
         Left            =   6960
         TabIndex        =   4
         Top             =   293
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7695
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   13573
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
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "FrmMst_NAVCodeMapping.frx":0000
         Left            =   1440
         List            =   "FrmMst_NAVCodeMapping.frx":0049
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   660
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
         Left            =   12360
         TabIndex        =   13
         Top             =   780
         Width           =   90
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
         Left            =   9840
         TabIndex        =   12
         Top             =   773
         Width           =   2415
      End
      Begin VB.Label LblID 
         Caption         =   "Client ID"
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
         TabIndex        =   11
         Top             =   773
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Account Type"
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
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmMst_NAVCodeMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
Dim mSave2DB, FieldName, mCol, Acc_Type As Variant
Dim mTANFlag As Boolean

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then
   MSHFlexGrid1.Clear
   CmdSave2DB.Enabled = False
   Exit Sub
End If
If LCase(CmbType.Text) = "expense type" Or LCase(CmbType.Text) = "client - ncdex" Or LCase(CmbType.Text) = "client - non ncdex" Or LCase(CmbType.Text) = "client - spot" Or LCase(CmbType.Text) = "location" Or LCase(CmbType.Text) = "commodity" Or LCase(CmbType.Text) = "employee" Or LCase(CmbType.Text) = "agent" Or LCase(CmbType.Text) = "vendor" Or LCase(CmbType.Text) = "dp - ncdex" Or LCase(CmbType.Text) = "dp - non ncdex" Or LCase(CmbType.Text) = "dp - spot" Or LCase(CmbType.Text) = "it vendor" Or LCase(CmbType.Text) = "reserved client - non ncdex" Or LCase(CmbType.Text) = "security agency" Then
   mTANFlag = False
''--- 10th Aug 2017 Ullhas Patil
ElseIf LCase(CmbType.Text) = LCase("Client - NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("Client - Non NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("Client - SPOT (TAN No)") Or LCase(CmbType.Text) = LCase("DP - NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("DP - Non NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("DP - SPOT (TAN No)") Or LCase(CmbType.Text) = LCase("Security Agency (TAN No)") Or LCase(CmbType.Text) = LCase("Vendor (TAN No)") Then
   mTANFlag = True
Else
   MsgBox "Invalid Selection!!!"
   CmbType.SetFocus
   Exit Sub
End If

If LCase(CmbType.Text) = "client - ncdex" Or LCase(CmbType.Text) = "client - non ncdex" Or LCase(CmbType.Text) = "client - spot" Or LCase(CmbType.Text) = "dp - ncdex" Or LCase(CmbType.Text) = "dp - non ncdex" Or LCase(CmbType.Text) = "dp - spot" Or LCase(CmbType.Text) = "reserved client - non ncdex" Then
   TxtSearchID.Locked = False
Else
   TxtSearchID = ""
   TxtSearchID.Locked = True
End If
If LCase(CmbType.Text) = "dp - ncdex" Or LCase(CmbType.Text) = "dp - non ncdex" Or LCase(CmbType.Text) = "dp - spot" Then
   LblID.Caption = "DP ID"
Else
   LblID.Caption = "Client ID"
End If

Call Grid_Head

End Sub

Private Sub CmdGetFile_Click()
On Error GoTo errortrap

With CommonDialog1
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls)|*.xls" '|Text files(*.txt)|*.txt"
     .MaxFileSize = &H7FFF ' 32KB filename buffer
    'same as .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNLongNames or cdlOFNExplorer --Or cdlOFNAllowMultiselect
     .Flags = &H4 Or &H800 Or &H40000 Or &H80000 Or &H200
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = CommonDialog1.FileName
mFilenames = CommonDialog1.FileName
If LblFile.Caption <> Path Then
 '  Call Clear
End If
LblFile.Caption = Path
If LblFile.Caption <> "" Then
   CmdImport.Enabled = True
End If
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Sub CmdGo_Click()
Dim wc As Variant
If CmbType.Text = "" Then
   MsgBox "Please select Account type!!!"
   CmbType.SetFocus
   Exit Sub
End If

wc = ""
Call Grid_Head
mSave2DB = ""
Call TableConectionQry

If Option1.Value = True Then
   wc = ""
   If LCase(CmbType.Text) = LCase("Security Agency (TAN No)") Then
        wc = " Order By SecurityAgencyID "
   ElseIf LCase(CmbType.Text) = LCase("Vendor (TAN No)") Then
        wc = " Order By MV.VendorID "
    '''-----------------------------------------------------------
   End If
ElseIf Option2.Value = True Then
   If LCase(CmbType.Text) = "client - ncdex" Or LCase(CmbType.Text) = "client - non ncdex" Or LCase(CmbType.Text) = "client - spot" Then
      wc = " And (MC.NAVCode is Null or MC.NAVCode = '')"
   ElseIf LCase(CmbType.Text) = "dp - ncdex" Or LCase(CmbType.Text) = "dp - non ncdex" Or LCase(CmbType.Text) = "dp - spot" Then
      wc = " And (NAVCode is Null or NAVCode = '')"
   ElseIf LCase(CmbType.Text) = "vendor" Or LCase(CmbType.Text) = "it vendor" Then
      wc = "Where MV.NAVCode is Null or MV.NAVCode = ''"
   ElseIf LCase(CmbType.Text) = "reserved client - non ncdex" Then
      wc = " And (MC.RNAVCode is Null or MC.RNAVCode = '')"
   ''--- 10th Aug 2017 Ullhas Patil
   ElseIf LCase(CmbType.Text) = LCase("Client - NCDEX (TAN No)") Then
        wc = " And (MC.TANNO is Null or MC.TANNO = '')"
   ElseIf LCase(CmbType.Text) = LCase("Client - Non NCDEX (TAN No)") Then
        wc = " And (MC.TANNO is Null or MC.TANNO = '')"
   ElseIf LCase(CmbType.Text) = LCase("Client - SPOT (TAN No)") Then
        wc = " And (MC.TANNO is Null or MC.TANNO = '')"
   ElseIf LCase(CmbType.Text) = LCase("DP - NCDEX (TAN No)") Then
        wc = " And (TANNO is Null or TANNO = '')"
   ElseIf LCase(CmbType.Text) = LCase("DP - Non NCDEX (TAN No)") Then
        wc = " And (TANNO is Null or TANNO = '')"
   ElseIf LCase(CmbType.Text) = LCase("DP - SPOT (TAN No)") Then
        wc = " And (TANNO is Null or TANNO = '') "
   ElseIf LCase(CmbType.Text) = LCase("Security Agency (TAN No)") Then
        wc = "Where TANNO is Null or TANNO = '' Order By SecurityAgencyID "
   ElseIf LCase(CmbType.Text) = LCase("Vendor (TAN No)") Then
        wc = "Where MV.TANNO is Null or MV.TANNO = '' Order By MV.VendorID "
    '''-----------------------------------------------------------
   Else
      wc = " Where NAVCode is Null or NAVCode = ''"
   End If
End If

If LCase(CmbType.Text) = "client - ncdex" Or LCase(CmbType.Text) = "client - non ncdex" Or LCase(CmbType.Text) = "client - spot" Or LCase(CmbType.Text) = "reserved client - non ncdex" Then
   If TxtSearchID <> "" Then
      wc = wc & " And ClientID = '" & TxtSearchID & "' "
   End If
End If

If LCase(CmbType.Text) = LCase("Client - NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("Client - Non NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("Client - SPOT (TAN No)") Then
   If TxtSearchID <> "" Then
      wc = wc & " And ClientID = '" & TxtSearchID & "'"
   End If
End If

If LCase(CmbType.Text) = "dp - ncdex" Or LCase(CmbType.Text) = "dp - non ncdex" Or LCase(CmbType.Text) = "dp - spot" Then
   If TxtSearchID.Text <> "" Then
      wc = wc & " And DepositoryID = '" & TxtSearchID.Text & "' "
   End If
End If
If LCase(CmbType.Text) = LCase("DP - NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("DP - Non NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("DP - SPOT (TAN No)") Then
   If TxtSearchID.Text <> "" Then
      wc = wc & " And DepositoryID = '" & TxtSearchID.Text & "' Order By DepositoryName"
   End If
End If

tmp = tmp & wc
Call TmpTable
Call dispdata
Label22.Refresh
Label22.Caption = TmpRs.RecordCount
CmdExcel.Enabled = True

End Sub

Private Sub dispdata()
Dim x As Double
x = 0
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       x = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
       
       TmpRs.MoveNext
   Next
Else
   MsgBox "No Record Found!!!"
End If
End Sub
Private Sub TableConectionQry()
tmp = ""
mCol = ""
FieldName = ""
If LCase(CmbType.Text) = "expense type" Then
   tmp = "Select ExpenseTypeID,ExpenseType,NAVCode,FAGUpdated,FAGUpdatedDate from Mst_ExpenseType "
   If mSave2DB = "s" Then
      FieldName = "ExpenseTypeID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = "client - ncdex" Then
   tmp = "Select MC.ClientIDRow, MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.NAVCode,MC.NAVClientName,MC.FAGUpdated,MC.FAGUpdatedDate " & _
         "From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID  " & _
         "Where MC.ExchangeTypeID = 2 "
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
ElseIf LCase(CmbType.Text) = "client - non ncdex" Then
   tmp = "Select MC.ClientIDRow ,MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.NAVCode,MC.NAVClientName,MC.FAGUpdated,MC.FAGUpdatedDate " & _
         "From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID " & _
         "Where MC.ExchangeTypeID = 1"
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
ElseIf LCase(CmbType.Text) = "reserved client - non ncdex" Then
    tmp = "Select MC.ClientIDRow ,MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.RNAVCode,MC.NAVClientName,MC.FAGUpdated,MC.FAGUpdatedDate " & _
         "From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID " & _
         "Where MC.ExchangeTypeID = 1"
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
ElseIf LCase(CmbType.Text) = "client - spot" Then
   tmp = "Select MC.ClientIDRow,MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.NAVCode,MC.NAVClientName,MC.FAGUpdated,MC.FAGUpdatedDate " & _
         "From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID " & _
         "Where MC.ExchangeTypeID = 3"
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
'=========Code For DP
ElseIf LCase(CmbType.Text) = "dp - ncdex" Then
   tmp = "Select DepositoryID,DepositoryName,NavCode,NAVDepositoryName,FAGUpdated,FAGUpdatedDate from Mst_Depository " & _
         "Where ExchangeTypeID = 2"
   If mSave2DB = "s" Then
      FieldName = "DepositoryID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = "dp - non ncdex" Then
   tmp = "Select DepositoryID,DepositoryName,NavCode,NAVDepositoryName,FAGUpdated,FAGUpdatedDate from Mst_Depository " & _
         "Where ExchangeTypeID = 1"
   If mSave2DB = "s" Then
      FieldName = "DepositoryID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = "dp - spot" Then
   tmp = "Select DepositoryID,DepositoryName,NavCode,NAVDepositoryName,FAGUpdated,FAGUpdatedDate from Mst_Depository " & _
         "Where ExchangeTypeID = 3"
   If mSave2DB = "s" Then
      FieldName = "DepositoryID"
      mCol = 2
   End If
'===============
ElseIf LCase(CmbType.Text) = "location" Then
   tmp = "Select LocationID,LocationName,SM_Prefix,NAVCode,FAGUpdated,FAGUpdatedDate from Mst_Location "
   If mSave2DB = "s" Then
      FieldName = "LocationID"
      mCol = 3
   End If
ElseIf LCase(CmbType.Text) = "commodity" Then
   tmp = "Select CommodityGroupID,CommodityGroup,NavCode,FAGUpdated,FAGUpdatedDate from Mst_CommodityGroup"
   If mSave2DB = "s" Then
      FieldName = "CommodityGroupID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = "employee" Then
   tmp = "Select EmployeeID,EmployeeNo,EmployeeName,DesignationID,NAVCode,FAGUpdated,FAGUpdatedDate from Mst_Employee "
   If mSave2DB = "s" Then
      FieldName = "EmployeeID"
      mCol = 4
   End If
ElseIf LCase(CmbType.Text) = "agent" Then
   tmp = "Select AgentID,AgentName,NAVCode,FAGUpdated,FAGUpdatedDate from Mst_Agent "
   If mSave2DB = "s" Then
      FieldName = "AgentID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = "vendor" Then
   tmp = "Select MV.VendorID,MV.VendorName,ML.LocationName,MT.Description,MV.NAVCode,MV.FAGUpdated,MV.FAGUpdatedDate " & _
         " From Mst_Vendor MV " & _
         " Inner join Mst_Location ML On MV.LocationID = ML.LocationID " & _
         " Inner join Mst_TDS MT On MV.TDSID = MT.TDSID "
   If mSave2DB = "s" Then
      FieldName = "VendorID"
      mCol = 4
   End If
ElseIf LCase(CmbType.Text) = "it vendor" Then
   tmp = "Select MV.VendorID,MV.VendorName,ML.LocationName,MT.Description,MV.NAVCode,MV.FAGUpdated,MV.FAGUpdatedDate " & _
         " From Mst_ITVendor MV " & _
         " Inner join Mst_Location ML On MV.LocationID = ML.LocationID " & _
         " Inner join Mst_TDS MT On MV.TDSID = MT.TDSID "
   If mSave2DB = "s" Then
      FieldName = "VendorID"
      mCol = 4
   End If
ElseIf LCase(CmbType.Text) = "security agency" Then
   tmp = "Select SecurityAgencyID,AgencyName,NAVCode,FAGUpdated,FAGUpdatedDate from Mst_SecurityAgency "
   If mSave2DB = "s" Then
      FieldName = "SecurityAgencyID"
      mCol = 2
   End If
''--- 10th Aug 2017 Ullhas Patil
ElseIf LCase(CmbType.Text) = LCase("Client - NCDEX (TAN No)") Then
   tmp = "Select MC.ClientIDRow, MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.TANNo,MC.UpdatedBy,MC.UpdatedDate "
   tmp = tmp & " From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID "
   tmp = tmp & " Where MC.ExchangeTypeID = 2 "
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
ElseIf LCase(CmbType.Text) = LCase("Client - Non NCDEX (TAN No)") Then
   tmp = "Select MC.ClientIDRow, MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.TANNo,MC.UpdatedBy,MC.UpdatedDate "
   tmp = tmp & " From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID "
   tmp = tmp & " Where MC.ExchangeTypeID = 1 "
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
ElseIf LCase(CmbType.Text) = LCase("Client - SPOT (TAN No)") Then
   tmp = "Select MC.ClientIDRow, MC.ClientID,MC.ClientName,DP.DepositoryID,DP.DepositoryName,MC.TANNo,MC.UpdatedBy,MC.UpdatedDate "
   tmp = tmp & " From Mst_Client MC Left Join Mst_Depository DP ON MC.DepositoryID = DP.DepositoryID "
   tmp = tmp & " Where MC.ExchangeTypeID = 3 "
   If mSave2DB = "s" Then
      FieldName = "ClientIDRow"
      mCol = 5
   End If
ElseIf LCase(CmbType.Text) = LCase("DP - NCDEX (TAN No)") Then
   tmp = "Select DepositoryID,DepositoryName,TANNO,UpdatedBy,UpdatedDate from Mst_Depository "
   tmp = tmp & " Where ExchangeTypeID = 2"
   If mSave2DB = "s" Then
      FieldName = "DepositoryID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = LCase("DP - Non NCDEX (TAN No)") Then
   tmp = "Select DepositoryID,DepositoryName,TANNO,UpdatedBy,UpdatedDate from Mst_Depository "
   tmp = tmp & " Where ExchangeTypeID = 1"
   If mSave2DB = "s" Then
      FieldName = "DepositoryID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = LCase("DP - SPOT (TAN No)") Then
   tmp = "Select DepositoryID,DepositoryName,TANNO,UpdatedBy,UpdatedDate from Mst_Depository "
   tmp = tmp & " Where ExchangeTypeID = 3"
   If mSave2DB = "s" Then
      FieldName = "DepositoryID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = LCase("Security Agency (TAN No)") Then
   tmp = "Select SecurityAgencyID,AgencyName,TANNO,UpdatedBy,UpdatedDate from Mst_SecurityAgency "
   If mSave2DB = "s" Then
      FieldName = "SecurityAgencyID"
      mCol = 2
   End If
ElseIf LCase(CmbType.Text) = LCase("Vendor (TAN No)") Then
   tmp = "Select MV.VendorID,MV.VendorName,ML.LocationName,MV.TANNO,MV.UpdatedBy,MV.UpdatedDate "
   tmp = tmp & " From Mst_Vendor MV "
   tmp = tmp & " Inner join Mst_Location ML On MV.LocationID = ML.LocationID "
   If mSave2DB = "s" Then
      FieldName = "VendorID"
      mCol = 3
   End If
'''-----------------------------------------------------------
End If

End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear

If CmbType.Text = "" Then Exit Sub

If LCase(CmbType.Text) = "expense type" Then
   Call Grid_Head1
ElseIf LCase(CmbType.Text) = "client - ncdex" Then
   Call Grid_Head2
ElseIf LCase(CmbType.Text) = "client - non ncdex" Or LCase(CmbType.Text) = "reserved client - non ncdex" Then
   Call Grid_Head2
ElseIf LCase(CmbType.Text) = "client - spot" Then
   Call Grid_Head2
ElseIf LCase(CmbType.Text) = "location" Then
   Call Grid_Head3
ElseIf LCase(CmbType.Text) = "commodity" Then
   Call Grid_Head4
ElseIf LCase(CmbType.Text) = "employee" Then
   Call Grid_Head5
ElseIf LCase(CmbType.Text) = "agent" Then
   Call Grid_Head6
ElseIf LCase(CmbType.Text) = "vendor" Then
   Call Grid_Head7
ElseIf LCase(CmbType.Text) = "dp - ncdex" Then
   Call Grid_Head8
ElseIf LCase(CmbType.Text) = "dp - non ncdex" Then
   Call Grid_Head8
ElseIf LCase(CmbType.Text) = "dp - spot" Then
   Call Grid_Head8
ElseIf LCase(CmbType.Text) = "it vendor" Then
   Call Grid_Head9
ElseIf LCase(CmbType.Text) = "security agency" Then
   Call Grid_Head10
''--- 10th Aug 2017 Ullhas Patil
ElseIf LCase(CmbType.Text) = LCase("Client - NCDEX (TAN No)") Then
    Call Grid_Head11
ElseIf LCase(CmbType.Text) = LCase("Client - Non NCDEX (TAN No)") Then
    Call Grid_Head11
ElseIf LCase(CmbType.Text) = LCase("Client - SPOT (TAN No)") Then
    Call Grid_Head11
ElseIf LCase(CmbType.Text) = LCase("DP - NCDEX (TAN No)") Then
    Call Grid_Head12
ElseIf LCase(CmbType.Text) = LCase("DP - Non NCDEX (TAN No)") Then
    Call Grid_Head12
ElseIf LCase(CmbType.Text) = LCase("DP - SPOT (TAN No)") Then
    Call Grid_Head12
ElseIf LCase(CmbType.Text) = LCase("Vendor (TAN No)") Then
    Call Grid_Head13
ElseIf LCase(CmbType.Text) = LCase("Security Agency (TAN No)") Then
    Call Grid_Head14
'''-----------------------------------------------------------
End If
End Sub
Private Sub Grid_Head1()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "Expense Type ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Expense Type"
MSHFlexGrid1.TextMatrix(0, 2) = "NAV Code"

MSHFlexGrid1.ColWidth(0) = 2000
MSHFlexGrid1.ColWidth(1) = 3600
MSHFlexGrid1.ColWidth(2) = 3000

End Sub

Private Sub Grid_Head2()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 7
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "ClientIDRow"
MSHFlexGrid1.TextMatrix(0, 1) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 2) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 3) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 4) = "DP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "NAV Code"
MSHFlexGrid1.TextMatrix(0, 6) = "NAV Client Name"
MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 1600
MSHFlexGrid1.ColWidth(2) = 2600
MSHFlexGrid1.ColWidth(3) = 2200
MSHFlexGrid1.ColWidth(4) = 2600
MSHFlexGrid1.ColWidth(5) = 2000
MSHFlexGrid1.ColWidth(6) = 3000

End Sub

Private Sub Grid_Head3()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.WordWrap = True

MSHFlexGrid1.TextMatrix(0, 0) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 3) = "NAV Code"
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1200
MSHFlexGrid1.ColWidth(3) = 1000
End Sub

Private Sub Grid_Head4()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Commodity Group ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Commodity Group"
MSHFlexGrid1.TextMatrix(0, 2) = "Nav Code"
MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 3400
MSHFlexGrid1.ColWidth(2) = 2400
End Sub

Private Sub Grid_Head5()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Employee ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Employee No"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Designation ID"
MSHFlexGrid1.TextMatrix(0, 4) = "NAV Code"
MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 2400
MSHFlexGrid1.ColWidth(2) = 3000
MSHFlexGrid1.ColWidth(3) = 1400
MSHFlexGrid1.ColWidth(4) = 2000
End Sub
Private Sub Grid_Head6()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Agent ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Agent Name"
MSHFlexGrid1.TextMatrix(0, 2) = "NAV Code"
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 4500
MSHFlexGrid1.ColWidth(2) = 2000

End Sub
Private Sub Grid_Head7()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Vendor ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Vendor Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "TDS Description"
MSHFlexGrid1.TextMatrix(0, 4) = "NAV Code"
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.ColWidth(3) = 2500
MSHFlexGrid1.ColWidth(4) = 2000

End Sub
Private Sub Grid_Head8()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 4
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Depository ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Depository Name"
MSHFlexGrid1.TextMatrix(0, 2) = "NAV Code"
MSHFlexGrid1.TextMatrix(0, 3) = "NAV Depository Name"
MSHFlexGrid1.ColWidth(0) = 1400
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 3000

End Sub
Private Sub Grid_Head9()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 5
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Vendor ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Vendor Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Location"
MSHFlexGrid1.TextMatrix(0, 3) = "TDS Description"
MSHFlexGrid1.TextMatrix(0, 4) = "NAV Code"
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 3000
MSHFlexGrid1.ColWidth(2) = 2500
MSHFlexGrid1.ColWidth(3) = 2500
MSHFlexGrid1.ColWidth(4) = 2000

End Sub

Private Sub Grid_Head10()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Agency ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Agency Name"
MSHFlexGrid1.TextMatrix(0, 2) = "NAV Code"
MSHFlexGrid1.ColWidth(0) = 1000
MSHFlexGrid1.ColWidth(1) = 4500
MSHFlexGrid1.ColWidth(2) = 2000

End Sub

''--- 10th Aug 2017 Ullhas Patil
Private Sub Grid_Head11()
'' --- Client (NCDEX/Non-NCDEX/SPOT)
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 6
    .WordWrap = True
    .TextMatrix(0, 0) = "ClientIDRow":      .ColWidth(0) = 1600
    .TextMatrix(0, 1) = "Client ID":        .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "Client Name":      .ColWidth(2) = 2500
    .TextMatrix(0, 3) = "DP ID":            .ColWidth(3) = 1800
    .TextMatrix(0, 4) = "DP Name":          .ColWidth(4) = 2500
    .TextMatrix(0, 5) = "TAN No":           .ColWidth(5) = 1600
End With
End Sub
Private Sub Grid_Head12()
'' --- DP (NCDEX/Non-NCDEX/SPOT)
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 3
    .WordWrap = True
    .TextMatrix(0, 0) = "Depository ID":      .ColWidth(0) = 2500
    .TextMatrix(0, 1) = "Depository Name":    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "TAN No":             .ColWidth(2) = 2000
End With
End Sub
Private Sub Grid_Head13()
'' --- Vendor
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 4
    .WordWrap = True
    .TextMatrix(0, 0) = "Vendor ID":      .ColWidth(0) = 1600
    .TextMatrix(0, 1) = "Vendor Name":    .ColWidth(1) = 3500
    .TextMatrix(0, 2) = "Location":       .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "TAN No":         .ColWidth(3) = 1600
End With
End Sub
Private Sub Grid_Head14()
'' --- Security Agency
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 3
    .WordWrap = True
    .TextMatrix(0, 0) = "Agency ID":      .ColWidth(0) = 2000
    .TextMatrix(0, 1) = "Agency Name":    .ColWidth(1) = 4000
    .TextMatrix(0, 2) = "TAN No":         .ColWidth(2) = 2000
End With
End Sub

''-------------------------------

Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows < 3 Then Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & " " & CmbType.Text & " " & "NAVCodeMapping.xls")
'Gen_mTimeStamp = GetTimeStamp
'Call Gen_Create_Xls(Gen_mTimeStamp & "Mst_QualityParameter.xls")
Call Xls_Detail_Rpt
CmdImport.Enabled = True
End Sub

Public Sub Xls_Detail_Rpt()

MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & " " & CmbType.Text & " " & "NAVCodeMapping.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub

Private Sub CmdImport_Click()
On Error GoTo msgError
If CmbType.Text = "" Then
   MsgBox "Please select Account type!!!"
   CmbType.SetFocus
   Exit Sub
End If
Call Grid_Head
CmdSave2DB.Enabled = True
CmdExcel.Enabled = False
CmdImport.Enabled = False
If LblFile.Caption = "" Then
   Call Grid_Head
   Exit Sub
End If
'Dim oXL As Excel.Application
'Dim oWB As Excel.Workbook
'Dim mRow As Integer
'Dim mCol As Integer
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
'Pat = File1.Path & "\" & File1.FileName
Pat = App.Path '& "\Mst_CMPLocation.xls"
'Set oWB = oXL.Workbooks.Open(Pat)
'Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & CmbType.Text & "NAVCodeMapping.xls")
'Set oWS = oWB.Worksheets("Sheet1")
'oWS.ClearArrows
'mRow = 1

'MsgBox "Wait till next message!!!"
Set oXL = New Excel.Application
Set oWB = oXL.Workbooks.Open(LblFile.Caption)
Set oWS = oWB.Worksheets(1)

If ChkImportFile = False Then
   MsgBox "Invalid file!!!"
   oXL.Visible = True
   CmdImport.Enabled = True
   Exit Sub
End If

Call Grid_Head

Dim I As Integer
I = 1
'MSHFlexGrid1.Cols = 3
'MSHFlexGrid1.Rows = 2
Do 'For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    'MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    If oWS.Cells(mRow, 1) <> "" Then
       For C = 0 To MSHFlexGrid1.Cols - 1
           MSHFlexGrid1.TextMatrix(I, C) = oWS.Cells(mRow, C + 1)
       Next
    Else
       Exit Do
    End If
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    I = I + 1
Loop Until False
oWB.Save
'MsgBox ("Done  !!! ")
'oXL.Visible = True
oWB.Close

Label22.Refresh
Label22.Caption = mRow - 2

Exit Sub
msgError:
MsgBox "Excel File Missing!!!"
End Sub

Private Sub CmdSave2DB_Click()
Dim totrows As Variant

mSave2DB = "s"
Call TableConectionQry
Call TmpTable
If LCase(CmbType.Text) = "employee" Then
   Call TableMst_EmpLedger
End If
totrows = 0
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) <> "" Then
       If MSHFlexGrid1.TextMatrix(I, mCol) <> "" Then
          If TmpRs.RecordCount > 0 Then
             TmpRs.MoveFirst
          Else
             Exit Sub
          End If
          If FieldName = "DepositoryID" Then
             TmpRs.Find FieldName & " = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
          Else
             TmpRs.Find FieldName & " = " & MSHFlexGrid1.TextMatrix(I, 0) & ""
          End If
'          If Not TmpRs.EOF Then
'             If TmpRs.Fields(1) = MSHFlexGrid1.TextMatrix(i, 1) Then
'                TmpRs!NAVCode = MSHFlexGrid1.TextMatrix(i, mCol)
'                TmpRs.Update
'                If LCase(CmbType.Text) = "employee" Then
'                   tmp = "Update Mst_EmpLedger Set EmployeeID = " & MSHFlexGrid1.TextMatrix(i, 0) & "  Where VendorCode = '" & MSHFlexGrid1.TextMatrix(i, 4) & "' "
'                   Call Tmp2Table
'                End If
'                totrows = totrows + 1
'             End If
'          End If
           If Not TmpRs.EOF Then
              If FieldName = "DepositoryID" Then
                 If LCase(TmpRs.Fields(0)) = LCase(MSHFlexGrid1.TextMatrix(I, 0)) Then
                 
                    If mTANFlag = False Then
                        TmpRs!NavCode = MSHFlexGrid1.TextMatrix(I, mCol)
                        If MSHFlexGrid1.TextMatrix(I, mCol + 1) <> "" Then
                           TmpRs!NAVDepositoryName = MSHFlexGrid1.TextMatrix(I, mCol + 1)
                        End If
                        TmpRs!FAGUpdated = Gen_UserLoginID
                        TmpRs!FAGUpdatedDate = Now
                        TmpRs.Update
                        totrows = totrows + 1
                    ElseIf mTANFlag = True Then
                        If LCase(CmbType.Text) = LCase("DP - NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("DP - Non NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("DP - SPOT (TAN No)") Then
                            ''-- 09 Aug 2017 Ullhas Patil
                            If MSHFlexGrid1.TextMatrix(I, mCol) <> "" Then
                               If ValidateTANNo(UCase(MSHFlexGrid1.TextMatrix(I, mCol))) = True Then
                                 TmpRs!TANNO = MSHFlexGrid1.TextMatrix(I, mCol)
                                 TmpRs!UpdatedBy = Gen_UserLoginID
                                 TmpRs!UpdatedDate = Now
                                 TmpRs.Update
                                 totrows = totrows + 1
                               End If
                            End If
                            ''----------------------------
                        End If
                    End If
                    
                 End If
              Else
                 If LCase(FieldName) = "clientidrow" Then
                    If TmpRs.Fields(0) = MSHFlexGrid1.TextMatrix(I, 0) Then  '
                          
                       If mTANFlag = False Then
                            
                            If LCase(CmbType.Text) = "reserved client - non ncdex" Then
                               TmpRs!RNAVCode = MSHFlexGrid1.TextMatrix(I, mCol)
                            Else
                               TmpRs!NavCode = MSHFlexGrid1.TextMatrix(I, mCol)
                            End If
                            If MSHFlexGrid1.TextMatrix(I, mCol + 1) <> "" Then
                               TmpRs!NAVClientName = MSHFlexGrid1.TextMatrix(I, mCol + 1)
                            End If
                            TmpRs!FAGUpdated = Gen_UserLoginID
                            TmpRs!FAGUpdatedDate = Now
                            TmpRs.Update
                            totrows = totrows + 1
                       ElseIf mTANFlag = True Then
                       
                            If LCase(CmbType.Text) = LCase("Client - NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("Client - Non NCDEX (TAN No)") Or LCase(CmbType.Text) = LCase("Client - SPOT (TAN No)") Then
                                ''-- 09 Aug 2017 Ullhas Patil
                                If MSHFlexGrid1.TextMatrix(I, mCol) <> "" Then
                                   If ValidateTANNo(UCase(MSHFlexGrid1.TextMatrix(I, mCol))) = True Then
                                     TmpRs!TANNO = MSHFlexGrid1.TextMatrix(I, mCol)
                                     TmpRs!UpdatedBy = Gen_UserLoginID
                                     TmpRs!UpdatedDate = Now
                                     TmpRs.Update
                                     totrows = totrows + 1
                                   End If
                                End If
                                ''----------------------------
                            End If
                       
                       End If
                       If LCase(CmbType.Text) = "employee" Then
                          tmp = "Update Mst_EmpLedger Set EmployeeID = " & MSHFlexGrid1.TextMatrix(I, 0) & "  Where VendorCode = '" & MSHFlexGrid1.TextMatrix(I, 4) & "' "
                          Call Tmp2Table
                          totrows = totrows + 1
                       End If
                    End If
                 ElseIf LCase(CmbType.Text) = "vendor" Then
                    tmp = "Update Mst_Vendor Set NAVCode = '" & MSHFlexGrid1.TextMatrix(I, mCol) & "',FAGUpdated = '" & Gen_UserLoginID & "',FAGUpdatedDate = '" & Now & "' Where VendorID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "' "
                    Call Tmp2Table
                    totrows = totrows + 1
                 ElseIf LCase(CmbType.Text) = "it vendor" Then
                    tmp = "Update Mst_ITVendor Set NAVCode = '" & MSHFlexGrid1.TextMatrix(I, mCol) & "',FAGUpdated = '" & Gen_UserLoginID & "',FAGUpdatedDate = '" & Now & "'  Where VendorID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "' "
                    Call Tmp2Table
                    totrows = totrows + 1
                 ElseIf LCase(CmbType.Text) = LCase("Security Agency (TAN No)") Then
                    If mTANFlag = True Then
                        ''-- 09 Aug 2017 Ullhas Patil
                        If Trim(MSHFlexGrid1.TextMatrix(I, mCol)) <> "" Then
                            If ValidateTANNo(UCase(Trim(MSHFlexGrid1.TextMatrix(I, mCol)))) = True Then
                                tmp = "Update Mst_SecurityAgency Set TANNO = '" & MSHFlexGrid1.TextMatrix(I, mCol) & "', UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = '" & Now & "' Where SecurityAgencyID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "' "
                                Call Tmp2Table
                                totrows = totrows + 1
                            End If
                        End If
                        ''--------------------------------
                    End If
                 ElseIf LCase(CmbType.Text) = LCase("Vendor (TAN No)") Then
                    If mTANFlag = True Then
                        ''-- 09 Aug 2017 Ullhas Patil
                        If Trim(MSHFlexGrid1.TextMatrix(I, mCol)) <> "" Then
                            If ValidateTANNo(UCase(Trim(MSHFlexGrid1.TextMatrix(I, mCol)))) = True Then
                                tmp = "Update Mst_Vendor Set TANNO = '" & MSHFlexGrid1.TextMatrix(I, mCol) & "',UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = '" & Now & "' Where VendorID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "' "
                                Call Tmp2Table
                                totrows = totrows + 1
                            End If
                        End If
                        ''--------------------------------
                    End If
                 Else
                    If TmpRs.Fields(1) = MSHFlexGrid1.TextMatrix(I, 1) Then  '
                       TmpRs!NavCode = MSHFlexGrid1.TextMatrix(I, mCol)
                       TmpRs!FAGUpdated = Gen_UserLoginID
                       TmpRs!FAGUpdatedDate = Now
                       TmpRs.Update
                       totrows = totrows + 1
                       If LCase(CmbType.Text) = "employee" Then
                          tmp = "Update Mst_EmpLedger Set EmployeeID = " & MSHFlexGrid1.TextMatrix(I, 0) & "  Where VendorCode = '" & MSHFlexGrid1.TextMatrix(I, 4) & "' "
                          Call Tmp2Table
                          totrows = totrows + 1
                       End If
                    End If
                 
                 End If

              End If
              
           End If
       End If
    End If
Next
CmdSave2DB.Enabled = False
CmdExcel.Enabled = True
CmdImport.Enabled = True
MsgBox "Data Saved for " & totrows & " rows !!!"
LblFile.Caption = ""
'Call CmdGo_Click
End Sub

Private Function ChkImportFile() As Boolean
Dim Retval As Boolean

Retval = True

If LCase(CmbType.Text) = "expense type" Then
   If ChkImportFile1 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "client - ncdex" Then
'   Call ChkImportFile2
   If ChkImportFile2 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "client - non ncdex" Then
'   Call ChkImportFile2
   If ChkImportFile2 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "client - spot" Then
'   Call ChkImportFile2
   If ChkImportFile2 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "location" Then
   'Call ChkImportFile3
   If ChkImportFile3 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "commodity" Then
   'Call ChkImportFile4
   If ChkImportFile4 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "employee" Then
   'Call ChkImportFile5
   If ChkImportFile5 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "agent" Then
   'Call ChkImportFile6
   If ChkImportFile6 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "tds" Then
   'Call ChkImportFile6
   If ChkImportFile7 = False Then
      Retval = False
   End If
ElseIf LCase(CmbType.Text) = "service tax" Then
   'Call ChkImportFile6
   If ChkImportFile8 = False Then
      Retval = False
   End If

End If

ChkImportFile = Retval
End Function
Private Function ChkImportFile1() As Boolean
Dim Retval As Boolean
Retval = True
If oWS.Cells(1, 1) <> "Expense Type ID" Then
   MsgBox "Invalid Column Name No 'Expense Type ID' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Expense Type" Then
   MsgBox "Invalid Column Name No 'Expense Type': 2"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 3) <> "NAV Code" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 3"
   'oXL.Visible = True
   Retval = False
End If

ChkImportFile1 = Retval
End Function

Private Function ChkImportFile2() As Boolean

Dim Retval As Boolean
Retval = True

If oWS.Cells(1, 1) <> "ClientIDRow" Then
   MsgBox "Invalid Column Name No 'ClientIDRow' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Client ID" Then
   MsgBox "Invalid Column Name No 'Client ID' : 2"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 3) <> "Client Name" Then
   MsgBox "Invalid Column Name No 'Client Name': 3"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 4) <> "DP ID" Then
   MsgBox "Invalid Column Name No 'DP ID' : 4"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 5) <> "DP Name" Then
   MsgBox "Invalid Column Name No 'DP Name' : 5"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 6) <> "NAV Code" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 6"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 7) <> "NAV Client Name" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 6"
   'oXL.Visible = True
   Retval = False
End If

ChkImportFile2 = Retval
End Function

Private Function ChkImportFile3() As Boolean
Dim Retval As Boolean
Retval = True

If oWS.Cells(1, 1) <> "Location ID" Then
   MsgBox "Invalid Column Name No 'Location ID' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Location Name" Then
   MsgBox "Invalid Column Name No 'Location Name': 2"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 3) <> "SM_Prefix" Then
   MsgBox "Invalid Column Name No 'SM_Prefix' : 3"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 4) <> "NAV Code" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 4"
   'oXL.Visible = True
   Retval = False
End If
ChkImportFile3 = Retval
End Function

Private Function ChkImportFile4() As Boolean
Dim Retval As Boolean
Retval = True

If oWS.Cells(1, 1) <> "Commodity Group ID" Then
   MsgBox "Invalid Column Name No 'Commodity Group ID' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Commodity Group" Then
   MsgBox "Invalid Column Name No 'Commodity Group': 2"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 3) <> "Nav Code" Then
   MsgBox "Invalid Column Name No 'Nav Code' : 3"
   'oXL.Visible = True
   Retval = False
End If

ChkImportFile4 = Retval
End Function

Private Function ChkImportFile5() As Boolean
Dim Retval As Boolean
Retval = True
If oWS.Cells(1, 1) <> "Employee ID" Then
   MsgBox "Invalid Column Name No 'Employee ID' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Employee No" Then
   MsgBox "Invalid Column Name No 'Employee No': 2"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 3) <> "Employee Name" Then
   MsgBox "Invalid Column Name No 'Employee Name' : 3"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 4) <> "Designation ID" Then
   MsgBox "Invalid Column Name No 'Designation ID' : 4"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 5) <> "NAV Code" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 5"
   'oXL.Visible = True
   Retval = False
End If

ChkImportFile5 = Retval
End Function

Private Function ChkImportFile6() As Boolean
Dim Retval As Boolean
Retval = True
If oWS.Cells(1, 1) <> "Agent ID" Then
   MsgBox "Invalid Column Name No 'Agent ID' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Agent Name" Then
   MsgBox "Invalid Column Name No 'Agent Name': 2"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 3) <> "NAV Code" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 3"
   'oXL.Visible = True
   Retval = False
End If

ChkImportFile6 = Retval
End Function
Private Function ChkImportFile7() As Boolean
Dim Retval As Boolean
Retval = True
If oWS.Cells(1, 1) <> "Vendor ID" Then
   MsgBox "Invalid Column Name No 'Vendor ID' : 1"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 2) <> "Vendor Name" Then
   MsgBox "Invalid Column Name No 'Vendor Name': 2"
   'oXL.Visible = True
   Retval = False
End If
If oWS.Cells(1, 3) <> "Location Name" Then
   MsgBox "Invalid Column Name No 'Location Name' : 3"
   'oXL.Visible = True
   Retval = False
End If

If oWS.Cells(1, 4) <> "TDS Description" Then
   MsgBox "Invalid Column Name No 'TDS Description' : 4"
   'oXL.Visible = True
   Retval = False
End If


If oWS.Cells(1, 5) <> "NAV Code" Then
   MsgBox "Invalid Column Name No 'NAV Code' : 5"
   'oXL.Visible = True
   Retval = False
End If


ChkImportFile7 = Retval
End Function



Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   'Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
'Frame2.Enabled = True
CmbType.Text = ""
Call Grid_Head
Label22.Caption = ""
End Sub


