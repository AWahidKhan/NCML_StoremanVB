VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Frm_DRUpload 
   Caption         =   "DR Upload"
   ClientHeight    =   9240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   12735
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      Begin VB.TextBox TxtError 
         Height          =   3615
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2640
         Width           =   10575
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   5610
         TabIndex        =   6
         Top             =   360
         Width           =   1245
      End
      Begin VB.CommandButton CmdGetFile 
         Caption         =   "Get File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   3120
         TabIndex        =   5
         Top             =   360
         Width           =   1245
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   415
         Left            =   4365
         TabIndex        =   3
         Top             =   360
         Width           =   1245
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   390
         Left            =   1560
         TabIndex        =   4
         Top             =   915
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         Caption         =   "Error : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   4200
         Width           =   825
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label lblProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1245
         Width           =   6375
      End
      Begin VB.Label lblProcess 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1605
         Width           =   6375
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_DRUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strFilename As String
Dim fn, CFlag As Variant
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet
Dim mRow As Integer
Dim mCol As Integer
Dim strSplitCmpId As Variant
Dim mCMPID As Double
Dim mCommodityID As Double
Dim mLocationID As Double
Private Sub CmdExport_Click()
Call Import_Data
End Sub
Private Sub Import_Data()
Ans = Check_File
If Ans = True Then
   MsgBox "Import Successfull !!!! "
Else
   MsgBox "Error in Import !!!! "
End If
   oXL.Quit
End Sub
Function Check_File() As Boolean

fn = GetFileFromPath(strFilename)
Dim objFileSystem As Object
Dim FileToCheck As String
Set objFileSystem = CreateObject("Scripting.FileSystemObject")
CFlag = True
FileToCheck = strFilename 'App.Path & "\Excel\" & fn & ""

If objFileSystem.FileExists(FileToCheck) Then

Else
   MsgBox "File not found !!!! "
   CFlag = False
End If
    
If CFlag = True Then
   Set oXL = New Excel.Application
   Dim Pat As String
   fn = GetFileFromPath(strFilename)
   Pat = App.Path
   Set oWB = oXL.Workbooks.Open(strFilename) '(Pat & "\excel\" & fn & "", , , , "")
   Set oWS = oWB.Worksheets(1)
   If oWS.Name <> "Sheet1" Then
      CFlag = False
   Else
      If CheckFile = False Then
         MsgBox "Please Select Proper File!!!"
         CmdImport.Enabled = False
         CFlag = False
      End If
   End If
   Check_File = CFlag
End If
Dim DAdd, CString As Variant
Dim mStatus As Boolean
If CFlag = True Then
   Dim i As Variant
   'DR Master Import
   Call TableMst_DR
   Call TableMst_ISIN
   'Call TableMst_CMP
   Call TableMst_Commodity
   Call TableMst_GSL(" Where ExchangeTypeID= 2 And FLag not in ('Z') ")
   
   Set oWS = oWB.Worksheets(1)
   i = 2
   ProgressBar1.Value = 0
   ProgressBar1.Max = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).Row
   Do
     If oWS.Cells(i, 1) = "" Then
        Exit Do
     End If
     If oWS.Cells(i, 3) = "N" Then
        If RsMst_DR.RecordCount > 0 Then
           RsMst_DR.MoveFirst
           RsMst_DR.Find "DRNo = " & Val(oWS.Cells(i, 1))
        End If
'        If oWS.Cells(i, 1) = 205789 Or _
'           oWS.Cells(i, 1) = 205790 Or _
'           oWS.Cells(i, 1) = 205791 Or _
'           oWS.Cells(i, 1) = 211549 Or _
'           oWS.Cells(i, 1) = 216779 Or _
'           oWS.Cells(i, 1) = 216780 Or _
'           oWS.Cells(i, 1) = 219282 Or _
'           oWS.Cells(i, 1) = 219283 Or _
'           oWS.Cells(i, 1) = 219284 Then
'           tmp = tmp
'        End If
        If RsMst_DR.EOF Then
           ' Get KR Detail from GSL Master
           'RsMst_GSL.MoveFirst
           RsMst_GSL.Filter = "krno = " & Val(oWS.Cells(i, 1))
           
           If RsMst_GSL.EOF Then
              ' KR not found in GSL Master
              TxtError = TxtError & i & " KR Not found in GSL Master " & oWS.Cells(i, 1) & Chr(13)
           Else
              ' Check for more than one GSL
              If RsMst_GSL.RecordCount > 1 Then
                 CString = RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID
                 For z = 1 To RsMst_GSL.RecordCount
                     If CString <> RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID Then
                        MsgBox "KR No : " & RsMst_GSL!KRNO & "Commodity or CMP Mismatch please contact technology!!!"
                        End
                     End If
                     RsMst_GSL.MoveNext
                 Next
                 RsMst_GSL.MoveFirst
              End If
              
              If oWS.Cells(i, 2) = "" Then
                 ' Rejected Lot
                 If oWS.Cells(i, 20) = "Rejected By Assayer" Then
                    For z = 1 To RsMst_GSL.RecordCount
                        RsMst_GSL!Flag = "J"
                        RsMst_GSL.Update
                        RsMst_GSL.MoveNext
                    Next
                    RsMst_GSL.MoveFirst
                 End If
              Else
                    ' Check for ISIN in master in case not found Add in master
                    RsMst_ISIN.MoveFirst
                    RsMst_ISIN.Find "ISINID =  '" & oWS.Cells(i, 2) & "' "
                    If RsMst_ISIN.EOF Then
                       Call SaveISIN(oWS.Cells(i, 2), RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID & "~" & oWS.Cells(i, 8), RsMst_GSL!CMPID, oWS.Cells(i, 10), oWS.Cells(i, 11), GetLocationIDFromCMP(RsMst_GSL!CMPID), RsMst_GSL!CommodityID, 2, Trim(oWS.Cells(i, 9)), "0")
                    Else
                       Call TableMst_CMP
                       RsMst_CMP.MoveFirst
                       RsMst_CMP.Find "Cmpid = " & RsMst_GSL!CMPID
                       If RsMst_ISIN!LocationID <> RsMst_CMP!LocationID Then
                          MsgBox "Critical error ISIN Location & GSL Location missmatch. Process Terminated. ISIN No=" & oWS.Cells(i, 2) & ""
                          End
                       End If
                       
                    End If
                    If oWS.Cells(i, 20) = "Successful - Credited" Then
                       'insert Data in Mst_DR Table
                       RsMst_DR.AddNew
                       RsMst_DR!DRID = GetDRID
                       RsMst_DR!DRNo = oWS.Cells(i, 1)
                       RsMst_DR!DRDate = Date
                       RsMst_DR!Txn_Flag = "D"
                       RsMst_DR!NewISINID = oWS.Cells(i, 2)
                       RsMst_DR!ClientIDRow = RsMst_GSL!ClientIDRow 'GetClientIDRow(oWS.Cells(i, 13))
                       RsMst_DR!Net_Weight_e = oWS.Cells(i, 16)
                       RsMst_DR!Remark = oWS.Cells(i, 18)
                       RsMst_DR!Status = oWS.Cells(i, 19) 'Left(oWS.Cells(I, 19), 200)
                       RsMst_DR!CommtrackFlag = "0"
                       'RsMst_DR!OldISINID = ""
                       'RsMst_DR!Net_Weight_d=
                       RsMst_DR.Update
                       If RsMst_GSL.RecordCount > 1 Then
                          CString = RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID
                          For z = 1 To RsMst_GSL.RecordCount
                              RsMst_GSL!ISINID = oWS.Cells(i, 2)
                              RsMst_GSL.Update
                              RsMst_GSL.MoveNext
                           Next
                       Else
                          RsMst_GSL!ISINID = oWS.Cells(i, 2)
                          RsMst_GSL.Update
                       End If
                       DAdd = DAdd + 1
                    End If
              End If
           End If
        End If
           
     ElseIf oWS.Cells(i, 3) = "Y" Then
        RsMst_ISIN.MoveFirst
        RsMst_ISIN.Find "ISINID = '" & oWS.Cells(i, 2) & "'"
        If RsMst_ISIN.EOF = True Then
           RsMst_ISIN.MoveFirst
           RsMst_ISIN.Find "ISINID = '" & oWS.Cells(i, 4) & "'"
           If Not RsMst_ISIN.EOF = True Then
              mCommodityID = RsMst_ISIN!CommodityID
              mCMPID = RsMst_ISIN!CMPID
              mLocationID = RsMst_ISIN!LocationID
              Call SaveISIN(oWS.Cells(i, 2), GetCMPName(mCMPID) & "~" & GetCommodityName(mCommodityID) & "~" & Trim(oWS.Cells(i, 9)) & "~ NCMSL Upload ", mCMPID, oWS.Cells(i, 10), oWS.Cells(i, 11), mLocationID, mCommodityID, 2, Trim(oWS.Cells(i, 9)), "0")
           Else
              TxtError = TxtError & i & " ISIN not in ISIN Master " & oWS.Cells(i, 4) & vbCrLf 'Chr(13)
           End If
        Else
           TxtError = TxtError & i & " ISIN Already in ISIN Master " & oWS.Cells(i, 2) & vbCrLf 'Chr(13)
        End If
     End If
     i = i + 1
     ProgressBar1.Value = ProgressBar1.Value + 1
     lblProgress.Caption = "Completed : " & Int(ProgressBar1.Value * 100 / ProgressBar1.Max) & "%"
     lblProcess.Caption = "Buffering " & oWS.Name & " Records"
     lblProgress.Refresh
     lblProcess.Refresh

   Loop Until False
   ProgressBar1.Value = ProgressBar1.Max
   lblProgress.Caption = "Completed : " & Int(ProgressBar1.Value * 100 / ProgressBar1.Max) & "%"
   lblProcess.Caption = "Buffering " & oWS.Name & " Records"
   lblProgress.Refresh
   lblProcess.Refresh
End If
End Function
Private Sub SaveISIN(mISINID_ As Variant, mISINDescription_ As Variant, mCMPID_ As Double, mNE_Date_ As Date, mFED_ As Date, mLocationID_ As Double, mCommodityID_ As Double, mExchangeTypeID_ As Double, mGrade_ As Variant, mFlag_ As Variant)

RsMst_ISIN.AddNew
RsMst_ISIN!ISINID = mISINID_
RsMst_ISIN!ISINDescription = mISINDescription_
RsMst_ISIN!CMPID = mCMPID_
RsMst_ISIN!NE_Date = mNE_Date_
RsMst_ISIN!FED = mFED_
RsMst_ISIN!LocationID = mLocationID_
RsMst_ISIN!CommodityID = mCommodityID_
RsMst_ISIN!ExchangeTypeID = mExchangeTypeID_
RsMst_ISIN!Grade = mGrade_
RsMst_ISIN!Flag = mFlag_
RsMst_ISIN!G_UID = GetGUID
RsMst_ISIN!CreatedBy = Gen_UserLoginID
RsMst_ISIN!CreatedDate = Now
RsMst_ISIN.Update

End Sub

Private Function GetDRID() As Double
Dim Retval As Double
tmp = "Select max(DRID) from Mst_DR"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetDRID = Retval
End Function

Private Sub CmdCancel_Click()
TxtError = ""
CmdImport.Enabled = False
CmdGetFile.Enabled = True
LblFile.Refresh
LblFile.Caption = ""
ProgressBar1.Value = 0
lblProgress.Caption = ""
lblProcess.Caption = ""
End Sub

Private Sub CmdGetFile_Click()
TxtError = ""
With cdialog
   .InitDir = App.Path
   .DialogTitle = "Select File To Be Upload"
   .ShowOpen
End With
If cdialog.CancelError Then
   LblFile.Refresh
   LblFile.Caption = ""
   CmdImport.Enabled = False
Else
   strFilename = cdialog.FileName
   LblFile.Refresh
   LblFile.Caption = strFilename 'GetFileFromPath(strFilename)
   If strFilename <> "" Then
      CmdImport.Enabled = True
   Else
      CmdImport.Enabled = False
   End If
End If
'If Left(GetFileFromPath(strFilename), 9) = "BASE FILE" Then
'    CmdGetFile.Enabled = False

'Else
'    lblFile.Caption = "Please Select the correct file "
'    CmdGetFile.Enabled = True
'    CmdImport.Enabled = False
'End If
End Sub

Private Sub CmdImport_Click()
TxtError = ""
Call Import_Data
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
Frame1.Enabled = True
CmdImport.Enabled = False
lblProcess.Caption = ""
lblProgress.Caption = ""
End Sub
Function GetFileFromPath(vPath As String)
    Dim Items() As String
    Items = Split(vPath, "\")
    If UBound(Items) = -1 Then Exit Function
    GetFileFromPath = Items(UBound(Items))
End Function
Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True

If oWS.Cells(1, 1) <> "Receipt No" Then
   Retval = False
ElseIf oWS.Cells(1, 2) <> "ISIN" Then
   Retval = False
ElseIf oWS.Cells(1, 3) <> "Is Revalidation" Then
   Retval = False
ElseIf oWS.Cells(1, 4) <> "Old ISIN" Then
   Retval = False
ElseIf oWS.Cells(1, 5) <> "Old Grade" Then
   Retval = False
ElseIf oWS.Cells(1, 6) <> "Warehouse" Then
   Retval = False
ElseIf oWS.Cells(1, 7) <> "WarehouseAddress" Then
   Retval = False
ElseIf oWS.Cells(1, 8) <> "Commodity" Then
   Retval = False
ElseIf oWS.Cells(1, 9) <> "Grade" Then
   Retval = False
ElseIf oWS.Cells(1, 10) <> "Validity" Then
   Retval = False
ElseIf oWS.Cells(1, 11) <> "FED" Then
   Retval = False
ElseIf oWS.Cells(1, 12) <> "Deposit Date" Then
   Retval = False
ElseIf oWS.Cells(1, 13) <> "Client DP" Then
   Retval = False
ElseIf oWS.Cells(1, 14) <> "Client ID" Then
   Retval = False
ElseIf oWS.Cells(1, 15) <> "Client Name" Then
   Retval = False
ElseIf oWS.Cells(1, 16) <> "Credit Qty" Then
   Retval = False
ElseIf oWS.Cells(1, 17) <> "Debit Qty" Then
   Retval = False
ElseIf oWS.Cells(1, 18) <> "Assayer Ref.No" Then
   Retval = False
ElseIf oWS.Cells(1, 19) <> "Remarks" Then
   Retval = False
ElseIf oWS.Cells(1, 20) <> "Present Status" Then
   Retval = False
ElseIf oWS.Cells(1, 21) <> "Lot Number" Then
   Retval = False
End If

CheckFile = Retval

End Function
