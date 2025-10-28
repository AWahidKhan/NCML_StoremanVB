VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRpt_ISINRevalidation 
   Caption         =   "ISIN Revalidation Report"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   10950
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      Begin VB.CommandButton CmdGo 
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
         Left            =   150
         TabIndex        =   4
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
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
         Height          =   375
         Left            =   1500
         TabIndex        =   3
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
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
         Height          =   375
         Left            =   2850
         TabIndex        =   2
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave2DB 
         Caption         =   "Save"
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
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   285
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   8295
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   14631
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
      Begin VB.Label LblFile 
         AutoSize        =   -1  'True
         Caption         =   "FileName"
         Height          =   195
         Left            =   6480
         TabIndex        =   8
         Top             =   600
         Width           =   660
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
         Left            =   6480
         TabIndex        =   7
         Top             =   255
         Width           =   2415
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
         Left            =   9000
         TabIndex        =   6
         Top             =   255
         Width           =   90
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
Attribute VB_Name = "FrmRpt_ISINRevalidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Private Sub CmdExcel_Click()
If MSHFlexGrid1.Rows < 3 Then Exit Sub
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_ISINRevalidation.xls")
Call Xls_Detail_Rpt
CmdImport.Enabled = True
End Sub

Public Sub Xls_Detail_Rpt()

MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_ISINRevalidation.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        If InStr(MSHFlexGrid1.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(I, c), "mm/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        End If
    Next
Next
oWB.Save
Set oWB = Nothing
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


Private Sub CmdGo_Click()
Label22.Caption = ""
lblFile.Caption = ""
Call Grid_Head

tmp = "Select TI.NIRID,TI.RevalDate,TI.ARNo,TI.ARDate,TI.Grade,TI.OldISIN,ML.LocationName,MC1.CMPName,MCd.Commodity,"
tmp = tmp & " MCL.ClientID,MCl.ClientName,MD.DepositoryID,MD.DepositoryName,TNI.NIGRID,TNI.GSLID,MG.SM_Prefix,MC.CMPName,MGd.GodownNo,"
tmp = tmp & " MG.StackNo , MG.LotNo, TNI.BalanceBags, TNI.BalanceWeight, TNI.NewISIN, TNI.DRNo, TNI.DRDate, TNI.OldQty, TNI.NewQty"
tmp = tmp & " From Txn_NewISINGSLRevalidation TNI"
tmp = tmp & " inner join Txn_NewISINRevalidation TI on TNI.NIRID=TI.NIRID"
tmp = tmp & " inner join Mst_ISIN MI on TI.OldISIN=MI.ISINID"
tmp = tmp & " inner join Mst_Location ML on MI.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_CMP MC1 on MI.CMPID=MC1.CMPID"
tmp = tmp & " inner join Mst_Commodity MCd on MI.CommodityID=MCd.CommodityID"
tmp = tmp & " inner join Mst_Client MCl on TI.ClientIDRow=MCl.ClientIDRow"
tmp = tmp & " inner join Mst_Depository MD on MCl.DepositoryID=MD.DepositoryID"
tmp = tmp & " inner join Mst_GSL MG on TNI.GSLID=MG.GSLID And TNI.SM_Prefix=MG.SM_Prefix"
tmp = tmp & " inner join Mst_Godown MGd on MG.GodownID=MGd.GodownID"
tmp = tmp & " inner join Mst_CMP MC on MGd.CMPID=MC.CMPID"
tmp = tmp & " Where TNI.NewISIN='' Or TNI.NewISIN is Null"
tmp = tmp & " Order By TI.NIRID,TNI.NIGRID,MG.SM_Prefix,MC.CMPName,MGd.GodownNo,MG.StackNo,MG.LotNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For c = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       TmpRs.MoveNext
   Next
End If
Label22.Caption = TmpRs.RecordCount
MsgBox ("Done!!")
CmdExcel.Enabled = True
End Sub

Private Sub CmdImport_Click()
On Error GoTo msgError

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
lblFile.Caption = CommonDialog1.FileName
Label22.Caption = ""
MsgBox ("Wait till Next Message")

Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(lblFile.Caption)  'Pat & "\excel\" & "Rpt_ISINRevalidation.xls")
Set oWS = oWB.Worksheets("Sheet1")
If ChkImportFile = False Then
   MsgBox "Invalid file!!!"
   oXL.Visible = True
   CmdImport.Enabled = True
   Exit Sub
End If

Call Grid_Head
Dim I As Integer

I = 1
mRow = 1

tmp = "Select TI.NIRID,TI.RevalDate,TI.ARNo,TI.ARDate,TI.Grade,TI.OldISIN,ML.LocationName,MC1.CMPName,MCd.Commodity,"
tmp = tmp & " MCL.ClientID,MCl.ClientName,MD.DepositoryID,MD.DepositoryName,TNI.NIGRID,TNI.GSLID,MG.SM_Prefix,MC.CMPName,MGd.GodownNo,"
tmp = tmp & " MG.StackNo , MG.LotNo, TNI.BalanceBags, TNI.BalanceWeight, TNI.NewISIN, TNI.DRNo, TNI.DRDate, TNI.OldQty, TNI.NewQty"
tmp = tmp & " From Txn_NewISINGSLRevalidation TNI"
tmp = tmp & " inner join Txn_NewISINRevalidation TI on TNI.NIRID=TI.NIRID"
tmp = tmp & " inner join Mst_ISIN MI on TI.OldISIN=MI.ISINID"
tmp = tmp & " inner join Mst_Location ML on MI.LocationID=ML.LocationID"
tmp = tmp & " inner join Mst_CMP MC1 on MI.CMPID=MC1.CMPID"
tmp = tmp & " inner join Mst_Commodity MCd on MI.CommodityID=MCd.CommodityID"
tmp = tmp & " inner join Mst_Client MCl on TI.ClientIDRow=MCl.ClientIDRow"
tmp = tmp & " inner join Mst_Depository MD on MCl.DepositoryID=MD.DepositoryID"
tmp = tmp & " inner join Mst_GSL MG on TNI.GSLID=MG.GSLID And TNI.SM_Prefix=MG.SM_Prefix"
tmp = tmp & " inner join Mst_Godown MGd on MG.GodownID=MGd.GodownID"
tmp = tmp & " inner join Mst_CMP MC on MGd.CMPID=MC.CMPID"
tmp = tmp & " Where TNI.NewISIN='' Or TNI.NewISIN is Null"
tmp = tmp & " Order By TI.NIRID,TNI.NIGRID,MG.SM_Prefix,MC.CMPName,MGd.GodownNo,MG.StackNo,MG.LotNo"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   TmpRs.MoveFirst
End If

Do
    mRow = mRow + 1 ' + 6
    If oWS.Cells(mRow, 1) <> "" Then
       TmpRs.Filter = "NIRID=" & Val(oWS.Cells(mRow, 1)) & " And OldISIN='" & Trim(oWS.Cells(mRow, 6)) & "' And NIGRID=" & Val(oWS.Cells(mRow, 14)) & " And GSLID=" & Val(oWS.Cells(mRow, 15)) & " And SM_Prefix='" & Trim(oWS.Cells(mRow, 16)) & "' "
       If TmpRs.RecordCount > 0 Then
          MSHFlexGrid1.TextMatrix(I, 0) = Val(oWS.Cells(mRow, 1))
          MSHFlexGrid1.TextMatrix(I, 1) = oWS.Cells(mRow, 2)
          MSHFlexGrid1.TextMatrix(I, 2) = Trim(oWS.Cells(mRow, 3))
          MSHFlexGrid1.TextMatrix(I, 3) = oWS.Cells(mRow, 4)
          MSHFlexGrid1.TextMatrix(I, 4) = Trim(oWS.Cells(mRow, 5))
          MSHFlexGrid1.TextMatrix(I, 5) = Trim(oWS.Cells(mRow, 6))
          MSHFlexGrid1.TextMatrix(I, 6) = Trim(oWS.Cells(mRow, 7))
          MSHFlexGrid1.TextMatrix(I, 7) = Trim(oWS.Cells(mRow, 8))
          MSHFlexGrid1.TextMatrix(I, 8) = Trim(oWS.Cells(mRow, 9))
          MSHFlexGrid1.TextMatrix(I, 9) = Trim(oWS.Cells(mRow, 10))
          MSHFlexGrid1.TextMatrix(I, 10) = Trim(oWS.Cells(mRow, 11))
          MSHFlexGrid1.TextMatrix(I, 11) = Trim(oWS.Cells(mRow, 12))
          MSHFlexGrid1.TextMatrix(I, 12) = Trim(oWS.Cells(mRow, 13))
          MSHFlexGrid1.TextMatrix(I, 13) = Val(oWS.Cells(mRow, 14))
          MSHFlexGrid1.TextMatrix(I, 14) = Val(oWS.Cells(mRow, 15))
          MSHFlexGrid1.TextMatrix(I, 15) = Trim(oWS.Cells(mRow, 16))
          MSHFlexGrid1.TextMatrix(I, 16) = Trim(oWS.Cells(mRow, 17))
          MSHFlexGrid1.TextMatrix(I, 17) = Trim(oWS.Cells(mRow, 18))
          MSHFlexGrid1.TextMatrix(I, 18) = Trim(oWS.Cells(mRow, 19))
          MSHFlexGrid1.TextMatrix(I, 19) = Trim(oWS.Cells(mRow, 20))
          MSHFlexGrid1.TextMatrix(I, 20) = Val(oWS.Cells(mRow, 21))
          MSHFlexGrid1.TextMatrix(I, 21) = Val(oWS.Cells(mRow, 22))
          MSHFlexGrid1.TextMatrix(I, 22) = Trim(oWS.Cells(mRow, 23))
          MSHFlexGrid1.TextMatrix(I, 23) = oWS.Cells(mRow, 24)
          MSHFlexGrid1.TextMatrix(I, 24) = oWS.Cells(mRow, 25)
          MSHFlexGrid1.TextMatrix(I, 25) = Val(oWS.Cells(mRow, 26))
          MSHFlexGrid1.TextMatrix(I, 26) = Val(oWS.Cells(mRow, 27))
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          I = I + 1
       End If
       TmpRs.Filter = ""
    Else
       Exit Do
    End If
Loop Until False
oWB.Save
oWB.Close
CmdExcel.Enabled = False
CmdImport.Enabled = False
Label22.Refresh
Label22.Caption = mRow - 2
CmdSave2DB.Enabled = True
Exit Sub
msgError:
MsgBox "Excel File Missing!!!"
End Sub
Private Function ChkImportFile() As Boolean
Dim Retval As Boolean
Retval = True


If LCase(oWS.Cells(1, 1)) <> "revalidation id" Then
   MsgBox "Invalid Column Name No 'Revalidation ID' : 1"
   Retval = False
End If

If LCase(oWS.Cells(1, 2)) <> "revalidation date" Then
   MsgBox "Invalid Column Name No 'Revalidation Date': 2"
   Retval = False
End If
If LCase(oWS.Cells(1, 3)) <> "ar no" Then
   MsgBox "Invalid Column Name No 'AR No' : 3"
   Retval = False
End If
If LCase(oWS.Cells(1, 4)) <> "ar date" Then
   MsgBox "Invalid Column Name No 'AR Date' : 4"
   Retval = False
End If
If LCase(oWS.Cells(1, 5)) <> "grade" Then
   MsgBox "Invalid Column Name No 'Grade' : 5"
   Retval = False
End If
If LCase(oWS.Cells(1, 6)) <> "old isin" Then
   MsgBox "Invalid Column Name No 'Old ISIN' : 6"
   Retval = False
End If
If LCase(oWS.Cells(1, 7)) <> "location" Then
   MsgBox "Invalid Column Name No 'Location' : 7"
   Retval = False
End If
If LCase(oWS.Cells(1, 8)) <> "cmp name" Then
   MsgBox "Invalid Column Name No 'CMP Name' : 8"
   Retval = False
End If
If LCase(oWS.Cells(1, 9)) <> "commodity" Then
   MsgBox "Invalid Column Name No 'Commodity' : 9"
   Retval = False
End If
If LCase(oWS.Cells(1, 10)) <> "client id" Then
   MsgBox "Invalid Column Name No 'Client ID' : 10"
   Retval = False
End If
If LCase(oWS.Cells(1, 11)) <> "client name" Then
   MsgBox "Invalid Column Name No 'Client Name' : 11"
   Retval = False
End If
If LCase(oWS.Cells(1, 12)) <> "dp id" Then
   MsgBox "Invalid Column Name No 'DP ID' : 12"
   Retval = False
End If
If LCase(oWS.Cells(1, 13)) <> "dp name" Then
   MsgBox "Invalid Column Name No 'DP ID' : 13"
   Retval = False
End If
If LCase(oWS.Cells(1, 14)) <> "revdetail id" Then
   MsgBox "Invalid Column Name No 'RevDetail ID' : 14"
   Retval = False
End If
If LCase(oWS.Cells(1, 15)) <> "gsl id" Then
   MsgBox "Invalid Column Name No 'GSL ID' : 15"
   Retval = False
End If
If LCase(oWS.Cells(1, 16)) <> "location id" Then
   MsgBox "Invalid Column Name No 'Location ID' : 16"
   Retval = False
End If
If LCase(oWS.Cells(1, 17)) <> "cmp name" Then
   MsgBox "Invalid Column Name No 'CMP Name' : 17"
   Retval = False
End If
If LCase(oWS.Cells(1, 18)) <> "godown no" Then
   MsgBox "Invalid Column Name No 'Godown No' : 18"
   Retval = False
End If
If LCase(oWS.Cells(1, 19)) <> "stack no" Then
   MsgBox "Invalid Column Name No 'Stack No' : 19"
   Retval = False
End If
If LCase(oWS.Cells(1, 20)) <> "lot no" Then
   MsgBox "Invalid Column Name No 'Lot No' : 20"
   Retval = False
End If
If LCase(oWS.Cells(1, 21)) <> "no of bags" Then
   MsgBox "Invalid Column Name No 'No of Bags' : 21"
   Retval = False
End If
If LCase(oWS.Cells(1, 22)) <> "balance qty" Then
   MsgBox "Invalid Column Name No 'Balance Qty' : 22"
   Retval = False
End If
If LCase(oWS.Cells(1, 23)) <> "new isin" Then
   MsgBox "Invalid Column Name No 'New ISIN' : 23"
   Retval = False
End If
If LCase(oWS.Cells(1, 24)) <> "dr no" Then
   MsgBox "Invalid Column Name No 'DR No' : 24"
   Retval = False
End If
If LCase(oWS.Cells(1, 25)) <> "dr date" Then
   MsgBox "Invalid Column Name No 'DR Date' : 25"
   Retval = False
End If
If LCase(oWS.Cells(1, 26)) <> "old qty" Then
   MsgBox "Invalid Column Name No 'Old Qty' : 26"
   Retval = False
End If
If LCase(oWS.Cells(1, 27)) <> "new qty" Then
   MsgBox "Invalid Column Name No 'New Qty' : 27"
   Retval = False
End If

ChkImportFile = Retval
End Function

Private Sub CmdSave2DB_Click()
If MSHFlexGrid1.Rows < 3 Then Exit Sub
Dim totrows As Variant
totrows = 0
Call TableMst_GSL("")
Call TableTxn_NewISINGSLRevalidation("")
For I = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(I, 22) <> "" Then
       If RsTxn_NewISINGSLRevalidation.RecordCount > 0 Then
          RsTxn_NewISINGSLRevalidation.MoveFirst
       End If
       RsTxn_NewISINGSLRevalidation.Find ("NIGRID = " & Val(MSHFlexGrid1.TextMatrix(I, 13)))
       If Not RsTxn_NewISINGSLRevalidation.EOF Then
          If IsNull(RsTxn_NewISINGSLRevalidation!NewISIN) = True Or RsTxn_NewISINGSLRevalidation!NewISIN = "" Then
             If RsMst_GSL.RecordCount > 0 Then
                RsMst_GSL.MoveFirst
             End If
             RsMst_GSL.Filter = "GSLID=" & Val(MSHFlexGrid1.TextMatrix(I, 14)) & " And SM_Prefix='" & MSHFlexGrid1.TextMatrix(I, 15) & "'"
             If RsMst_GSL.RecordCount > 0 Then
                RsMst_GSL!ISINID = Trim(MSHFlexGrid1.TextMatrix(I, 22))
                RsMst_GSL.Update
                
                RsTxn_NewISINGSLRevalidation!NewISIN = MSHFlexGrid1.TextMatrix(I, 22)
                RsTxn_NewISINGSLRevalidation!DRNo = MSHFlexGrid1.TextMatrix(I, 23)
                If IsDate(MSHFlexGrid1.TextMatrix(I, 24)) = True Then
                   RsTxn_NewISINGSLRevalidation!DRDate = Format(MSHFlexGrid1.TextMatrix(I, 24), Gen_DatFormat)
                End If
                RsTxn_NewISINGSLRevalidation!OldQty = Val(MSHFlexGrid1.TextMatrix(I, 25))
                RsTxn_NewISINGSLRevalidation!NewQty = Val(MSHFlexGrid1.TextMatrix(I, 26))
                RsTxn_NewISINGSLRevalidation.Update
                totrows = totrows + 1
             End If
             RsMst_GSL.Filter = ""
          End If
       End If
    End If
Next
CmdSave2DB.Enabled = False
CmdExcel.Enabled = True
CmdImport.Enabled = True

MsgBox "Data Saved for " & totrows & " rows !!!"

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Call Grid_Head
Label22.Caption = ""
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 27
MSHFlexGrid1.WordWrap = True
MSHFlexGrid1.TextMatrix(0, 0) = "Revalidation ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Revalidation Date"
MSHFlexGrid1.TextMatrix(0, 2) = "AR No"
MSHFlexGrid1.TextMatrix(0, 3) = "AR Date"
MSHFlexGrid1.TextMatrix(0, 4) = "Grade"
MSHFlexGrid1.TextMatrix(0, 5) = "Old ISIN"
MSHFlexGrid1.TextMatrix(0, 6) = "Location"
MSHFlexGrid1.TextMatrix(0, 7) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 8) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 9) = "Client ID"
MSHFlexGrid1.TextMatrix(0, 10) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 11) = "DP ID"
MSHFlexGrid1.TextMatrix(0, 12) = "DP Name"
MSHFlexGrid1.TextMatrix(0, 13) = "RevDetail ID"
MSHFlexGrid1.TextMatrix(0, 14) = "GSL ID"
MSHFlexGrid1.TextMatrix(0, 15) = "Location ID"
MSHFlexGrid1.TextMatrix(0, 16) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 17) = "Godown No"
MSHFlexGrid1.TextMatrix(0, 18) = "Stack No"
MSHFlexGrid1.TextMatrix(0, 19) = "Lot No"
MSHFlexGrid1.TextMatrix(0, 20) = "No of Bags"
MSHFlexGrid1.TextMatrix(0, 21) = "Balance Qty"
MSHFlexGrid1.TextMatrix(0, 22) = "New ISIN"
MSHFlexGrid1.TextMatrix(0, 23) = "DR No"
MSHFlexGrid1.TextMatrix(0, 24) = "DR Date"
MSHFlexGrid1.TextMatrix(0, 25) = "Old Qty"
MSHFlexGrid1.TextMatrix(0, 26) = "New Qty"

MSHFlexGrid1.ColWidth(0) = 1300
MSHFlexGrid1.ColWidth(1) = 1300
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1200
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1200
MSHFlexGrid1.ColWidth(10) = 2000
MSHFlexGrid1.ColWidth(11) = 1200
MSHFlexGrid1.ColWidth(12) = 2000
MSHFlexGrid1.ColWidth(13) = 0
MSHFlexGrid1.ColWidth(14) = 1000
MSHFlexGrid1.ColWidth(15) = 1200
MSHFlexGrid1.ColWidth(16) = 1600
MSHFlexGrid1.ColWidth(17) = 1200
MSHFlexGrid1.ColWidth(18) = 1200
MSHFlexGrid1.ColWidth(19) = 1200
MSHFlexGrid1.ColWidth(20) = 1200
MSHFlexGrid1.ColWidth(21) = 1200
MSHFlexGrid1.ColWidth(22) = 1400
MSHFlexGrid1.ColWidth(23) = 1200
MSHFlexGrid1.ColWidth(24) = 1200
MSHFlexGrid1.ColWidth(25) = 1200
MSHFlexGrid1.ColWidth(26) = 1200

MSHFlexGrid1.RowHeight(0) = 600
End Sub


