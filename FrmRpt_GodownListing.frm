VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_GodownListing 
   Caption         =   "Godown Listing Report"
   ClientHeight    =   9675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   7650
      Left            =   75
      TabIndex        =   11
      Top             =   2400
      Width           =   15030
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   7275
         Width           =   14790
         _ExtentX        =   26088
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   6975
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12303
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   -60
      Width           =   15015
      Begin VB.CheckBox ChkCorporateStock 
         Caption         =   "Corporate Stock"
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
         Left            =   5640
         TabIndex        =   27
         Top             =   1680
         Width           =   1875
      End
      Begin VB.CheckBox ChkRetailStock 
         Caption         =   "Retail Stock"
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
         Left            =   3840
         TabIndex        =   26
         Top             =   1680
         Width           =   1410
      End
      Begin VB.CheckBox ChkReservationStk 
         Caption         =   "Reservation Mapped and Stock"
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
         Width           =   3210
      End
      Begin VB.CheckBox ChkTotalStock 
         Caption         =   "Total Stock"
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
         Left            =   13560
         TabIndex        =   24
         Top             =   1193
         Width           =   1380
      End
      Begin VB.CheckBox ChkExchangeStk 
         Caption         =   "Exchange wise Stock"
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
         Left            =   11160
         TabIndex        =   23
         Top             =   1193
         Width           =   2235
      End
      Begin VB.ComboBox CmbStateID 
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
         Left            =   105
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   435
         Width           =   3675
      End
      Begin VB.ComboBox CmbStock 
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
         ItemData        =   "FrmRpt_GodownListing.frx":0000
         Left            =   11160
         List            =   "FrmRpt_GodownListing.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   435
         Width           =   3615
      End
      Begin VB.CommandButton CmdClear 
         Caption         =   "Clear"
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
         Left            =   13695
         TabIndex        =   10
         Top             =   1935
         Width           =   1215
      End
      Begin VB.ComboBox CmbStatus 
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
         ItemData        =   "FrmRpt_GodownListing.frx":001E
         Left            =   7440
         List            =   "FrmRpt_GodownListing.frx":0028
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   1200
         Width           =   3615
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
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   435
         Width           =   3480
      End
      Begin VB.ComboBox CmbCategoryID 
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
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   3645
      End
      Begin VB.ComboBox CmbUnitTypeID 
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
         Left            =   3840
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   3525
      End
      Begin VB.ComboBox CmbCMPName 
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
         Left            =   7410
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   435
         Width           =   3615
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
         Left            =   12480
         TabIndex        =   9
         Top             =   1935
         Width           =   1215
      End
      Begin VB.CommandButton CmdGetReport 
         Caption         =   "Get Report"
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
         Left            =   11265
         TabIndex        =   8
         Top             =   1935
         Width           =   1215
      End
      Begin VB.Label LblStateID 
         AutoSize        =   -1  'True
         Caption         =   "State Name"
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
         Left            =   1410
         TabIndex        =   21
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Balance Stock"
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
         Left            =   12315
         TabIndex        =   20
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblCMP 
         AutoSize        =   -1  'True
         Caption         =   "CMP Name"
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
         Left            =   8700
         TabIndex        =   19
         Top             =   135
         Width           =   1035
      End
      Begin VB.Label LblStatus 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   8970
         TabIndex        =   18
         Top             =   870
         Width           =   555
      End
      Begin VB.Label LblLocationID 
         AutoSize        =   -1  'True
         Caption         =   "Location"
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
         Left            =   5198
         TabIndex        =   17
         Top             =   135
         Width           =   765
      End
      Begin VB.Label LblCategoryID 
         AutoSize        =   -1  'True
         Caption         =   "Category"
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
         Left            =   1530
         TabIndex        =   16
         Top             =   870
         Width           =   825
      End
      Begin VB.Label LblUnitTypeID 
         AutoSize        =   -1  'True
         Caption         =   "Unit Type"
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
         Left            =   5167
         TabIndex        =   15
         Top             =   870
         Width           =   870
      End
      Begin VB.Label Label2 
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
         Left            =   5880
         TabIndex        =   14
         Top             =   2070
         Width           =   90
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record Found :-"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   2070
         Width           =   2295
      End
   End
End
Attribute VB_Name = "FrmRpt_GodownListing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID, mUnitTypeID, mCategoryID, mCMPID As Double

Private Sub CmbCategoryID_GotFocus()
tmp = CmbCategoryID.Text
Call TableMst_Category
CmbCategoryID.Clear
If RsMst_Category.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Category.RecordCount
    CmbCategoryID.AddItem RsMst_Category!Category
    RsMst_Category.MoveNext
Next
CmbCategoryID.Text = tmp
End Sub

Private Sub CmbCategoryID_LostFocus()
If CmbCategoryID.Text = "" Then
   Exit Sub
End If
RsMst_Category.MoveFirst
RsMst_Category.Find "Category = '" & CmbCategoryID.Text & "'"
If RsMst_Category.EOF Then
   MsgBox "Invalid selection "
   CmbCategoryID.SetFocus
   Exit Sub
End If
mCategoryID = RsMst_Category!CategoryID
End Sub

Private Sub CmbCMPName_GotFocus()
tmp = CmbCMPName.Text
Call TableMst_CMP(" where LocationID=" & mLocationID & " ")
CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "First Select Location!!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub

Private Sub CmbCMPName_LostFocus()
If CmbCMPName.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPName.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPName.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationID <> 0 And StateID= " & mStateID & "")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "First Select State  !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   CmbCMPName.Text = ""
   Exit Sub
End If
If tmp <> CmbLocationID.Text Then
   CmbCMPName.Text = ""
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

Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
Call TableMst_State
CmbStateID.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbStateID.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbStateID.Text = tmp
End Sub

Private Sub CmbStateID_LostFocus()
If CmbStateID.Text = "" Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   Exit Sub
End If
If tmp <> CmbStateID.Text Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

mStateID = RsMst_State!StateID
End Sub

Private Sub CmbStock_LostFocus()
If Trim(CmbStock.Text) = "" Then Exit Sub
If LCase(CmbStock.Text) <> "non zero" And LCase(CmbStock.Text) <> "zero" Then
   MsgBox "Invalid Selection!!!"
   CmbStock.SetFocus
   Exit Sub
End If
End Sub

Private Sub CmbUnitTypeID_GotFocus()
tmp = CmbUnitTypeID.Text
Call TableMst_UnitType
CmbUnitTypeID.Clear
If RsMst_UnitType.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_UnitType.RecordCount
    CmbUnitTypeID.AddItem RsMst_UnitType!UnitType
    RsMst_UnitType.MoveNext
Next
CmbUnitTypeID.Text = tmp
End Sub

Private Sub CmbUnitTypeID_LostFocus()
If CmbUnitTypeID.Text = "" Then
   Exit Sub
End If
RsMst_UnitType.MoveFirst
RsMst_UnitType.Find "UnitType = '" & CmbUnitTypeID.Text & "'"
If RsMst_UnitType.EOF Then
   MsgBox "Invalid selection "
   CmbUnitTypeID.SetFocus
   Exit Sub
End If
mUnitTypeID = RsMst_UnitType!UnitTypeID
End Sub

Private Sub CmdClear_Click()
mStateID = 0
mLocationID = 0
mUnitTypeID = 0
mCategoryID = 0
mCMPID = 0
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbUnitTypeID.Text = ""
CmbCMPName.Text = ""
CmbCategoryID.Text = ""
CmbStatus.Text = ""
Label2.Caption = ""
ChkExchangeStk.Value = 1
ChkTotalStock.Value = 1
ChkReservationStk.Value = 1
Call Grid_Head
End Sub

Private Sub CmdGetReport_Click()
Dim wc As String
MsgBox "Wait till next message"

Call Grid_Head
wc = ""
If CmbStateID.Text <> "" Then
   wc = GenWc(wc, mStateID, "l.StateID ", "N", "N")
  ' wc = wc & "and s.StateName=" & CmbStateID.Text & ""
End If
If CmbLocationID.Text <> "" Then
   wc = GenWc(wc, mLocationID, "c.LocationID ", "N", "N")
'   wc = wc & "and l.LocationName=" & CmbLocationID.Text & ""
End If
If CmbCMPName.Text <> "" Then
   wc = GenWc(wc, mCMPID, "g.CMPID", "N", "N")
'   wc = wc & "and c.CMPName=" & CmbCMPName.Text & ""
End If
If CmbCategoryID.Text <> "" Then
   wc = GenWc(wc, mCategoryID, "g.CategoryID", "N", "N")
'   wc = wc & "and Cat.Category=" & CmbCategoryID.Text & ""
End If
If CmbUnitTypeID.Text <> "" Then
   wc = GenWc(wc, mUnitTypeID, "g.UnitTypeID", "N", "N")
 '  wc = wc & "and u.UnitType=" & CmbUnitTypeID.Text & ""
End If
If CmbStatus.Text = "Open" Then
   'wc = GenWc(wc, "Null", "g.CloseDate", "N", "N")
   If wc = "" Then
      wc = " Where g.CloseDate Is Null"
   Else
      wc = wc & " and g.CloseDate Is Null"
   End If
ElseIf CmbStatus.Text = "Close" Then
   If wc = "" Then
      wc = " Where g.CloseDate Is Not Null"
   Else
      wc = wc & " and g.CloseDate IS Not Null"
   End If
End If

tmp = "Select s.StateName 'State',l.LocationName 'Location',C.CMPID, c.CMPName,"
tmp = tmp & " g.GodownNo ,g.Address, u.UnitType, Cat.Category,Ls.LicenseNo, g.Amount,g.AdvAmount,"
tmp = tmp & " g.GodownId,g.GodownCapacity , g.NCDEXAccreditedQty, g.AgreementEndDate, g.StartDate, g.CloseDate,c.StartDate as 'CMPStartDate',c.CurrentDate,ls.LicenseDate,ls.LicenseExpiryDate,g.TelephoneNo,g.LeasedUptoDate "
tmp = tmp & " ,g.CPEN, g.CPED, g.COPEN, g.COPED"
tmp = tmp & " from Mst_Godown g"
tmp = tmp & " Inner Join Mst_CMP C on g.CMPID = C.CMPID"
tmp = tmp & " Inner Join Mst_Location l on c.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_UnitType u on g.UnitTypeID = u.UnitTypeID"
tmp = tmp & " Inner Join Mst_Category cat on g.CategoryID = Cat.CategoryID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_License Ls on Ls.LicenseID = g.LicenseID "
If Trim(CmbStock.Text) <> "" Then
   tmp = tmp & " left JOIN Mst_GSL MGSL ON g.GodownID = MGSL.GodownID and mgsl.Status <> 'E' "
End If
If wc <> "" Then
   tmp = tmp & wc
End If
If Trim(CmbStock.Text) <> "" Then
   tmp = tmp & " group by s.StateName ,l.LocationName ,C.CMPID, c.CMPName, g.GodownNo ,g.Address, u.UnitType, Cat.Category,Ls.LicenseNo, g.Amount,g.AdvAmount, g.GodownId,g.GodownCapacity , g.NCDEXAccreditedQty, g.AgreementEndDate, g.StartDate, g.CloseDate,c.StartDate ,c.CurrentDate,ls.LicenseDate,ls.LicenseExpiryDate,g.TelephoneNo,g.LeasedUptoDate,g.CPEN, g.CPED, g.COPEN, g.COPED"
   If LCase(CmbStock.Text) = "zero" Then
      tmp = tmp & " Having Sum(IsNull(MGSL.DepositeWeight, 0) - IsNull(MGSL.CDTFWeight, 0)) = 0"
   ElseIf LCase(CmbStock.Text) = "non zero" Then
      tmp = tmp & " Having Sum(IsNull(MGSL.DepositeWeight, 0) - IsNull(MGSL.CDTFWeight, 0)) <> 0"
   End If
End If

'tmp = "SELECT Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock"
'tmp = tmp & " FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID"
'tmp = tmp & " Where Mst_GSL.flag <> 'E' And Mst_Godown.GodownID = " & mgodownID_ & ""

Call TmpTable


      
tmp = " Select TGA.GodownID ,Max(TA.ExpiryDate) 'ExpDate' from Txn_GodownLessorAgreeDetail TGA "
tmp = tmp & " Inner Join Txn_LessorAgreement TA On TGA.LAID = TA.LAID "
tmp = tmp & " Where TA.Flag not in ('R') "
tmp = tmp & " Group BY TGA.GodownID "
      
Call Tmp3Table

Call dispdata

If MSHFlexGrid1.Rows > 1 Then
   MSHFlexGrid1.FixedRows = 1
End If

MsgBox "Done!!!"

End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
'Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownListing.xls")
'Call Xls_Detail_Rpt


tmp = ""
ans = MsgBox("Do you want to Genearate CSV File", vbYesNo)
If ans = vbNo Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_GodownListing.xls")
   Call Xls_Detail_Rpt
Else
   Path = App.Path
   strFilName = "\Excel\" & Gen_mTimeStamp & "Rpt_GodownListing" & Gen_UserName & ".csv"
   strFilLocation = Path + strFilName
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
   Open strFilLocation For Output As #1
   For I = 0 To MSHFlexGrid1.Rows - 1
       tmp = ""
       For C = 0 To MSHFlexGrid1.Cols - 1
          If C = 0 Then
             tmp = tmp & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          ElseIf C = 6 Then
             tmp = tmp & ","
          Else
             tmp = tmp & "," & Replace(Replace(MSHFlexGrid1.TextMatrix(I, C), vbNewLine, ""), ",", "~")
          End If
       Next
       Print #1, tmp
   Next
   Close #1
   MsgBox "CSV Generated!!!"
End If

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
Set oWB = oXL.Workbooks.Open(Pat & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_GodownListing.xls")
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
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   CmdGetReport.Enabled = False
   CmdExcel.Enabled = False
   Exit Sub
End If
mStateID = 0
mLocationID = 0
mUnitTypeID = 0
mCategoryID = 0
mCMPID = 0
Call Grid_Head
CmdExcel.Enabled = False
ChkExchangeStk.Value = 1
ChkTotalStock.Value = 1
ChkReservationStk.Value = 1
End Sub
Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 43 '35

MSHFlexGrid1.TextMatrix(0, 0) = "Sr No"
MSHFlexGrid1.TextMatrix(0, 1) = "State "
MSHFlexGrid1.TextMatrix(0, 2) = "Location "
MSHFlexGrid1.TextMatrix(0, 3) = "CMP ID "
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name  "
MSHFlexGrid1.TextMatrix(0, 5) = "Godown No "
MSHFlexGrid1.TextMatrix(0, 6) = "Godown Address "
MSHFlexGrid1.TextMatrix(0, 7) = "Unit Type"
MSHFlexGrid1.TextMatrix(0, 8) = "Category"
MSHFlexGrid1.TextMatrix(0, 9) = "LicenseNo "
MSHFlexGrid1.TextMatrix(0, 10) = "Amount"
MSHFlexGrid1.TextMatrix(0, 11) = "Advance Amount"

MSHFlexGrid1.TextMatrix(0, 12) = "NCDEX Stock"
MSHFlexGrid1.TextMatrix(0, 13) = "Non NCDEX Stock"
MSHFlexGrid1.TextMatrix(0, 14) = "SPOT Stock"
MSHFlexGrid1.TextMatrix(0, 15) = "Stock"

MSHFlexGrid1.TextMatrix(0, 16) = "Godown Capacity (in MT)"
MSHFlexGrid1.TextMatrix(0, 17) = "Utilization"
MSHFlexGrid1.TextMatrix(0, 18) = "NCDEX Accredited Qty (in MT)"
MSHFlexGrid1.TextMatrix(0, 19) = "Agreement End Date"
MSHFlexGrid1.TextMatrix(0, 20) = "Start Date"
MSHFlexGrid1.TextMatrix(0, 21) = "Close Date"
MSHFlexGrid1.TextMatrix(0, 22) = "CMPM"
MSHFlexGrid1.TextMatrix(0, 23) = "Status"
MSHFlexGrid1.TextMatrix(0, 24) = "CMP BOD Date"

MSHFlexGrid1.TextMatrix(0, 25) = "Reservation"
MSHFlexGrid1.TextMatrix(0, 26) = "Reservation Stock"

MSHFlexGrid1.TextMatrix(0, 27) = "License Start Date"
MSHFlexGrid1.TextMatrix(0, 28) = "License Expiry Date"
MSHFlexGrid1.TextMatrix(0, 29) = "Godown Telephone No"
MSHFlexGrid1.TextMatrix(0, 30) = "Leased Upto"

MSHFlexGrid1.TextMatrix(0, 31) = "Cotton Endorsement no"
MSHFlexGrid1.TextMatrix(0, 32) = "Cotton endorsed Date"
MSHFlexGrid1.TextMatrix(0, 33) = "Commodity Endorement no"
MSHFlexGrid1.TextMatrix(0, 34) = "Commodity wise endorsed Date"


MSHFlexGrid1.TextMatrix(0, 35) = "Retail Stock"
MSHFlexGrid1.TextMatrix(0, 36) = "Retail Reservation Stock"
MSHFlexGrid1.TextMatrix(0, 37) = "Total Retail Stock"
MSHFlexGrid1.TextMatrix(0, 38) = "Corporate Stock"
MSHFlexGrid1.TextMatrix(0, 39) = "Corporate Reservation Stock"
MSHFlexGrid1.TextMatrix(0, 40) = "Total Corporate Stock"
MSHFlexGrid1.TextMatrix(0, 41) = "Retail Reservation Qty"
MSHFlexGrid1.TextMatrix(0, 42) = "Corporate Reservation Qty"


MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1800
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 0
MSHFlexGrid1.ColWidth(4) = 2600
MSHFlexGrid1.ColWidth(5) = 1000
MSHFlexGrid1.ColWidth(6) = 3600
MSHFlexGrid1.ColWidth(7) = 1800
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1400
MSHFlexGrid1.ColWidth(10) = 1200
MSHFlexGrid1.ColWidth(11) = 1200

MSHFlexGrid1.ColWidth(12) = 1200
MSHFlexGrid1.ColWidth(13) = 1200
MSHFlexGrid1.ColWidth(14) = 1200

MSHFlexGrid1.ColWidth(15) = 1200
MSHFlexGrid1.ColWidth(16) = 2000
MSHFlexGrid1.ColWidth(17) = 1000
MSHFlexGrid1.ColWidth(18) = 1000
MSHFlexGrid1.ColWidth(19) = 2000
MSHFlexGrid1.ColWidth(20) = 1200
MSHFlexGrid1.ColWidth(21) = 1200
MSHFlexGrid1.ColWidth(22) = 2600
MSHFlexGrid1.ColWidth(23) = 1000
MSHFlexGrid1.ColWidth(24) = 1200
MSHFlexGrid1.ColWidth(25) = 1500
MSHFlexGrid1.ColWidth(26) = 1500
MSHFlexGrid1.ColWidth(27) = 1500
MSHFlexGrid1.ColWidth(28) = 1500
MSHFlexGrid1.ColWidth(29) = 1500
MSHFlexGrid1.ColWidth(30) = 1500

MSHFlexGrid1.RowHeight(0) = 600
MSHFlexGrid1.WordWrap = True


End Sub

Public Sub dispdata()
Dim x As Double
x = 0
ProgressBar1.Value = 0
ProgressBar1.Max = TmpRs.RecordCount + 2
If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       x = MSHFlexGrid1.Rows - 1
       MSHFlexGrid1.TextMatrix(x, 0) = I
       For C = 1 To 11
           MSHFlexGrid1.TextMatrix(x, C) = IIf(IsNull(TmpRs.Fields(C - 1)), "", TmpRs.Fields(C - 1))
       Next
       MSHFlexGrid1.TextMatrix(x, 5) = IIf(IsNull(TmpRs!GodownNo), "", "'" & TmpRs!GodownNo)
       MSHFlexGrid1.TextMatrix(x, 16) = IIf(IsNull(TmpRs!GodownCapacity), "", TmpRs!GodownCapacity)
       MSHFlexGrid1.TextMatrix(x, 18) = IIf(IsNull(TmpRs!NCDEXAccreditedQty), "", TmpRs!NCDEXAccreditedQty)
       MSHFlexGrid1.TextMatrix(x, 19) = IIf(IsNull(TmpRs!AgreementEndDate), "", Format(TmpRs!AgreementEndDate, "dd-mmm-yyyy"))
       
       If LCase(MSHFlexGrid1.TextMatrix(x, 7)) = "leased" Then
          If TmpRs3.RecordCount > 0 Then
             TmpRs3.MoveFirst
             TmpRs3.Find "GodownID = " & TmpRs!GodownID & ""
             If TmpRs3.EOF = False Then
                MSHFlexGrid1.TextMatrix(x, 19) = IIf(IsNull(TmpRs3!ExpDate), "", Format(TmpRs3!ExpDate, "dd-mmm-yyyy"))
             Else
                MSHFlexGrid1.TextMatrix(x, 19) = ""
             End If
          End If
       End If
       
       MSHFlexGrid1.TextMatrix(x, 20) = IIf(IsNull(TmpRs!StartDate), "", Format(TmpRs!StartDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 21) = IIf(IsNull(TmpRs!CloseDate), "", Format(TmpRs!CloseDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 24) = IIf(IsNull(TmpRs!CurrentDate), Format(TmpRs!CMPStartDate, "dd-mmm-yyyy"), Format(TmpRs!CurrentDate, "dd-mmm-yyyy"))
       
       MSHFlexGrid1.TextMatrix(x, 27) = IIf(IsNull(TmpRs!LicenseDate), "", Format(TmpRs!LicenseDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 28) = IIf(IsNull(TmpRs!LicenseExpiryDate), "", Format(TmpRs!LicenseExpiryDate, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 29) = IIf(IsNull(TmpRs!TelephoneNo), "", TmpRs!TelephoneNo)
       MSHFlexGrid1.TextMatrix(x, 30) = IIf(IsNull(TmpRs!LeasedUptoDate), "", Format(TmpRs!LeasedUptoDate, "dd-mmm-yyyy"))
       
       MSHFlexGrid1.TextMatrix(x, 31) = IIf(IsNull(TmpRs!CPEN), "", TmpRs!CPEN)
       MSHFlexGrid1.TextMatrix(x, 32) = IIf(IsNull(TmpRs!CPED), "", Format(TmpRs!CPED, "dd-mmm-yyyy"))
       MSHFlexGrid1.TextMatrix(x, 33) = IIf(IsNull(TmpRs!COPEN), "", TmpRs!COPEN)
       MSHFlexGrid1.TextMatrix(x, 34) = IIf(IsNull(TmpRs!COPED), "", Format(TmpRs!COPED, "dd-mmm-yyyy"))
       If IsNull(TmpRs!CloseDate) = True Then
          MSHFlexGrid1.TextMatrix(x, 23) = "Open"
       Else
          MSHFlexGrid1.TextMatrix(x, 23) = "Close"
       End If
       
       If ChkExchangeStk.Value = 1 Then
          Call AddExchangeWiseStockData(TmpRs!GodownID, x)
       End If
       
       If ChkTotalStock.Value = 1 Then
          Call AddStock(TmpRs!GodownID, x)
          MSHFlexGrid1.TextMatrix(x, 17) = Round(Format((Val(MSHFlexGrid1.TextMatrix(x, 15)) / Val(MSHFlexGrid1.TextMatrix(x, 16)) * 100), "0.000"), 0)
       End If
       
       Call AddGData(TmpRs!CMPID, x)
      
       If ChkReservationStk.Value = 1 Then
          Call AddReservation(TmpRs!GodownID, x)
          MSHFlexGrid1.TextMatrix(x, 25) = Val(MSHFlexGrid1.TextMatrix(x, 41)) + Val(MSHFlexGrid1.TextMatrix(x, 42))
          Call AddReservationStock(TmpRs!GodownID, x)
       End If
       
       If ChkRetailStock.Value = 1 Then
          Call AddRetailCorporateStockData(TmpRs!GodownID, x, "R")
          MSHFlexGrid1.TextMatrix(x, 37) = Val(MSHFlexGrid1.TextMatrix(x, 37)) + Val(MSHFlexGrid1.TextMatrix(x, 35)) + Val(MSHFlexGrid1.TextMatrix(x, 36))
       End If
       
       If ChkCorporateStock.Value = 1 Then
          Call AddRetailCorporateStockData(TmpRs!GodownID, x, "C")
          MSHFlexGrid1.TextMatrix(x, 40) = Val(MSHFlexGrid1.TextMatrix(x, 40)) + Val(MSHFlexGrid1.TextMatrix(x, 38)) + Val(MSHFlexGrid1.TextMatrix(x, 39))
       End If
       
       x = MSHFlexGrid1.Rows - 1

       TmpRs.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
   Next
   ProgressBar1.Value = ProgressBar1.Max
   Label2.Refresh
   Label2.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
Else
   Call Grid_Head
End If
   Label2.Caption = TmpRs.RecordCount

End Sub
Sub AddExchangeWiseStockData(mGodownID_ As Double, x1 As Double)


tmp = "SELECT Mst_ExchangeType.ExchangeType,Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock " & _
      " From Mst_Godown " & _
      " INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID " & _
      " Left Join Mst_ExchangeType On Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID " & _
      " Where Mst_GSL.Status <> 'E' And Mst_Godown.GodownID = " & mGodownID_ & " " & _
      " Group By Mst_ExchangeType.ExchangeType " & _
      " Order By Mst_ExchangeType.ExchangeType "

Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
'    For c1 = 12 To 14
'        MSHFlexGrid1.TextMatrix(X1, c1) = IIf(IsNull(TmpRs1.Fields(c1 - 12)), "", TmpRs1.Fields(c1 - 12))
'    Next
    If LCase(TmpRs1!ExchangeType) = "ncdex" Then
       MSHFlexGrid1.TextMatrix(x1, 12) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    ElseIf LCase(TmpRs1!ExchangeType) = "non ncdex" Then
       MSHFlexGrid1.TextMatrix(x1, 13) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    ElseIf LCase(TmpRs1!ExchangeType) = "spot" Then
       MSHFlexGrid1.TextMatrix(x1, 14) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    ElseIf IsNull(TmpRs1!ExchangeType) = True Then
       MSHFlexGrid1.TextMatrix(x1, 13) = Val(MSHFlexGrid1.TextMatrix(x1, 13)) + IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    End If
    TmpRs1.MoveNext
'    If TmpRs1.RecordCount > 1 Then
'       X1 = MSHFlexGrid1.Rows
'       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    End If
Next
'If TmpRs1.RecordCount > 1 Then
'   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
'End If

TmpRs1.Close
End Sub

Sub AddGData(mCMPID_ As Double, x1 As Double)

tmp = "Select EMP.EmployeeName + '~' + EMP.EmployeeNo as CMPM  from Mst_EmpCMPMapping ECM " & _
      " Inner Join Mst_Employee EMP On ECM.EmployeeID = EMP.EmployeeID " & _
      " Where CMPID = " & mCMPID_ & " And EMP.flag = 'Y'"
Call Tmp2Table

If TmpRs2.RecordCount > 0 Then
'For i1 = 1 To TmpRs1.RecordCount
   For c1 = 22 To 22
       MSHFlexGrid1.TextMatrix(x1, c1) = IIf(IsNull(TmpRs2.Fields(c1 - 22)), "", TmpRs2.Fields(c1 - 22))
   Next
   'TmpRs1.MoveNext
'    If TmpRs1.RecordCount > 1 Then
'       X1 = MSHFlexGrid1.Rows
'       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
'    End If
'Next
'If TmpRs1.RecordCount > 1 Then
'   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows - 1
'End If
End If
TmpRs2.Close
End Sub
Sub AddStock(mGodownID_ As Double, x1 As Double)

If x1 = 1 Then
   tmp = "SELECT Mst_GSL.GodownID,Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock"
   tmp = tmp & " FROM Mst_GSL Where Mst_GSL.Status <> 'E' Group By Mst_GSL.GodownID "

   Call Tmp1Table
Else
   If TmpRs1.State = "0" Then
      tmp = "SELECT Mst_GSL.GodownID,Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock"
      tmp = tmp & " FROM Mst_GSL Where Mst_GSL.Status <> 'E' Group By Mst_GSL.GodownID "

      Call Tmp1Table
   End If

End If

'For i1 = 1 To TmpRs1.RecordCount
If TmpRs1.RecordCount > 0 Then TmpRs1.MoveFirst
TmpRs1.Find "GodownID = " & mGodownID_ & " "
'    For c1 = 15 To 15
c1 = 15
If TmpRs1.EOF = False Then
   MSHFlexGrid1.TextMatrix(x1, c1) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
Else
  MSHFlexGrid1.TextMatrix(x1, c1) = 0
End If
'    Next
'    TmpRs1.MoveNext
'Next

'TmpRs1.Close

End Sub

Private Sub AddReservation(mGodownID_ As Double, x1 As Double)

'tmp = " Select GodownID,Sum(ReservationQty) as 'ReservationQty' From  Txn_SRGodownMapDetails Where GodownID = " & mGodownID_ & " " & _
'      " Group by GodownID "
      


tmp = " Select MC.CorpClient,Sum(ReservationQty) as 'ReservationQty' From  Txn_SRGodownMapDetails TSRGM"
tmp = tmp & " Inner Join Txn_SRGodownMapping TSR On TSRGM.SRGodownMappingID = TSR.SRGodownMappingID"
tmp = tmp & " Inner Join Mst_StorageContract MSC On TSR.ContractID = MSC.ContractID"
tmp = tmp & " Inner Join Mst_Client MC On MSC.ClientID = MC.ClientID"
tmp = tmp & " Where  GodownID = " & mGodownID_ & ""
tmp = tmp & " Group By MC.CorpClient"
Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    If TmpRs1!CorpClient = "0" Then
       MSHFlexGrid1.TextMatrix(x1, 41) = IIf(IsNull(TmpRs1!ReservationQty), 0, TmpRs1!ReservationQty)
    ElseIf TmpRs1!CorpClient = "1" Then
       MSHFlexGrid1.TextMatrix(x1, 42) = IIf(IsNull(TmpRs1!ReservationQty), 0, TmpRs1!ReservationQty)
    End If
    TmpRs1.MoveNext
Next

TmpRs1.Close

End Sub

Private Sub AddReservationStock(mGodownID_ As Double, x1 As Double)
'
'tmp = " Select MG.ContractID, Sum(MG.balanceWeight) " & _
'      " From Mst_GSL MG " & _
'      " Inner Join Txn_SRGodownMapping TSM On MG.ContractID = TSM.ContractID  " & _
'      " Inner Join Txn_SRGodownMapDetails TSD On TSM.SRGodownMappingID = TSD.SRGodownMappingID  " & _
'      " Where MG.GodownID = " & mGodownID_ & " And (MG.ContractID <> 0 And MG.ContractID Is Not Null)  " & _
'      " Group by MG.ContractID "


tmp = "select GodownID,sum(BalanceWeight) 'R-Stock' from mst_gsl " & _
"where GodownID = " & mGodownID_ & " And (Status = 'O' or Status = 'J') and ContractID <> 0 and ContractID is not null " & _
"group by GodownID order by GodownID"

      
Call Tmp1Table


For i1 = 1 To TmpRs1.RecordCount
    For c1 = 26 To 26
        MSHFlexGrid1.TextMatrix(x1, c1) = IIf(IsNull(TmpRs1.Fields(c1 - 25)), 0, TmpRs1.Fields(c1 - 25))
    Next
    TmpRs1.MoveNext
Next

TmpRs1.Close

End Sub

Private Sub AddRetailCorporateStockData(mGodownID_ As Double, x1 As Double, mFlg_ As Variant)

tmp = "SELECT Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock "
tmp = tmp & " From Mst_Godown "
tmp = tmp & " INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
tmp = tmp & " INNER JOIN Mst_CLient ON Mst_GSL.ClientID = Mst_Client.ClientID "
tmp = tmp & " Where (Mst_GSL.ContractID = 0 Or Mst_GSL.ContractID is null) And "

If LCase(mFlg_) = "r" Then
   tmp = tmp & " Mst_Client.CorpClient = '0' "
ElseIf LCase(mFlg_) = "c" Then
   tmp = tmp & " Mst_Client.CorpClient = '1' "
End If

tmp = tmp & " And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.Status <> 'E' And Mst_Godown.GodownID = " & mGodownID_ & " "
tmp = tmp & " Group By Mst_CLient.CorpClient "
tmp = tmp & " Order By Mst_CLient.CorpClient "

Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    If LCase(mFlg_) = "r" Then
       MSHFlexGrid1.TextMatrix(x1, 35) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    ElseIf LCase(mFlg_) = "c" Then
       MSHFlexGrid1.TextMatrix(x1, 38) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    End If
    TmpRs1.MoveNext
Next

TmpRs1.Close

tmp = "SELECT Sum(Mst_GSL.DepositeWeight)-Sum(Mst_GSL.CDTFWeight) as Stock "
tmp = tmp & " From Mst_Godown "
tmp = tmp & " INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
tmp = tmp & " INNER JOIN Mst_CLient MC ON Mst_GSL.ClientID = MC.ClientID "
tmp = tmp & " Where (Mst_GSL.ContractID <> 0 And Mst_GSL.ContractID is not null) And "
If LCase(mFlg_) = "r" Then
   tmp = tmp & " MC.CorpClient = '0' And"
ElseIf LCase(mFlg_) = "c" Then
   tmp = tmp & " MC.CorpClient = '1' And"
End If
tmp = tmp & " Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.Status <> 'E' And Mst_Godown.GodownID = " & mGodownID_ & " "
tmp = tmp & " Group By MC.CorpClient "
tmp = tmp & " Order By MC.CorpClient "

Call Tmp1Table

For i1 = 1 To TmpRs1.RecordCount
    If LCase(mFlg_) = "r" Then
       MSHFlexGrid1.TextMatrix(x1, 36) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    ElseIf LCase(mFlg_) = "c" Then
       MSHFlexGrid1.TextMatrix(x1, 39) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
    End If
    TmpRs1.MoveNext
Next

'If LCase(mFlg_) = "r" Then
'   MSHFlexGrid1.TextMatrix(x1, 36) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
'ElseIf LCase(mFlg_) = "c" Then
'   MSHFlexGrid1.TextMatrix(x1, 39) = IIf(IsNull(TmpRs1!Stock), 0, TmpRs1!Stock)
'End If


TmpRs1.Close


End Sub

