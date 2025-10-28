VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmTxn_CMSEUpload 
   Caption         =   "CMS-E Upload"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   5850
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   12135
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3270
         Left            =   135
         TabIndex        =   10
         Top             =   1995
         Visible         =   0   'False
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   5768
         _Version        =   393216
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
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
         Left            =   9525
         TabIndex        =   4
         Top             =   600
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
         Left            =   8280
         TabIndex        =   3
         Top             =   600
         Width           =   1245
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
         Left            =   10770
         TabIndex        =   2
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox TxtError 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   6060
         Visible         =   0   'False
         Width           =   11085
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   390
         Left            =   75
         TabIndex        =   5
         Top             =   5355
         Width           =   12000
         _ExtentX        =   21167
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblProcess 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   6375
      End
      Begin VB.Label lblProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1245
         Width           =   6375
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackColor       =   &H00000080&
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   135
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
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   60
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmTxn_CMSEUpload"
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
Dim mError As Variant

Private Sub CmdExport_Click()
Call Import_Data
End Sub
Private Sub Import_Data()
ans = Check_File
If ans = True Then
   MsgBox "Import Successfull !!!! "
Else
   MsgBox "Error in Import !!!! "
End If
TxtError.Text = mError
oXL.Quit
'Dim ofs As New FileSystemObject
'Dim mFileName As String
'Gen_mTimeStamp = GetTimeStamp
'mFileName = App.Path & "\excel\" & Gen_mTimeStamp & Gen_UserName & "CMSEUpload.txt"
'
'ofs.CreateTextFile mFileName, True
'Open mFileName For Output As #1
'    Print #1, mError
'Close #1
'Shell "notepad.exe " & mFileName, vbMaximizedFocus

Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "CMSEUpload.xls")
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "CMSEUpload.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    oWS.Cells(mRow, 1) = MSHFlexGrid1.TextMatrix(I, 0)
Next
oWB.Save
'MsgBox ("Excel file created !!! ")
oXL.Visible = True
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
      MsgBox "Keep Sheet Caption as 'Sheet1'"
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
   Dim I As Variant
   'DR Master Import
   Call TableMst_DR
   Call TableMst_ISIN
   'Call TableMst_CMP
   Call TableMst_Commodity
   Call TableMst_GSL(" Where ExchangeTypeID= 2 And (Status not in ('B','E','Z')) ")
   Call TableMst_Godown
   Call TableMst_Client
   
   Set oWS = oWB.Worksheets(1)
   I = 2
   ProgressBar1.Value = 0
   ProgressBar1.Max = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
   Do
     If oWS.Cells(I, 1) = "" Then
        Exit Do
     End If
     
     If oWS.Cells(I, 4) <> "" And oWS.Cells(I, 4) <> "0" Then  'If oWS.Cells(I, 3) = "N" Then
          
     ''added for checking that if client, commodity and godown code not blank
       If oWS.Cells(I, 2) = "" Or oWS.Cells(I, 5) = "" Or oWS.Cells(I, 6) = "" Or oWS.Cells(I, 14) = "" Then
         
          MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Blank Warehouse Code Or Commodity Code Or DR Date Or ClientID" & oWS.Cells(I, 3) '& vbCrLf '& Chr(13)
          'mError = mError & I & " Blank Warehouse Code, Commodity Code Or ClientID" & oWS.Cells(I, 3) '& vbCrLf '& Chr(13)
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
          GoTo NextRow
       End If
             
             
        If RsMst_DR.RecordCount > 0 Then
           RsMst_DR.MoveFirst
           RsMst_DR.Find "DRNo = " & Val(oWS.Cells(I, 4))
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
           RsMst_GSL.Filter = "krno = " & Val(oWS.Cells(I, 3))
           
           If RsMst_GSL.EOF Then
              ' KR not found in GSL Master
              MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " CMS-P Not found in GSL Master " & oWS.Cells(I, 3) & vbCrLf ''& vbCrLf '& Chr(13)
              'mError = mError & I & " CMS-P Not found in GSL Master " & oWS.Cells(I, 3) & vbCrLf ''& vbCrLf '& Chr(13)
              MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           Else
               ''''''checking for godown, commodity commtrack ncdex code and clientid'''''''''''''''''''''''''''''
               RsMst_Godown.MoveFirst
               RsMst_Godown.Find "GodownID = " & RsMst_GSL!GodownID
               If Not RsMst_Godown.EOF Then
                  If UCase(IIf(IsNull(RsMst_Godown!CommtrackNCDEXCode), "", RsMst_Godown!CommtrackNCDEXCode)) <> UCase(oWS.Cells(I, 2)) Then
                     MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Godown NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                     'mError = mError & I & " Godown NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                     MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                     GoTo NextRow
                  End If
               End If
               RsMst_Commodity.MoveFirst
               RsMst_Commodity.Find "CommodityID = " & RsMst_GSL!CommodityID
               If Not RsMst_Commodity.EOF Then
                  If UCase(IIf(IsNull(RsMst_Commodity!CommtrackNCDEXCode), "", RsMst_Commodity!CommtrackNCDEXCode)) <> UCase(oWS.Cells(I, 5)) Then
                     MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Commodity NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                     'mError = mError & I & " Commodity NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                     MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                     GoTo NextRow
                  End If
               End If
               If LCase(Gen_UserRole) <> "power" Then
                  RsMst_Client.MoveFirst
                  RsMst_Client.Find "ClientID = " & RsMst_GSL!ClientID
                  ''RsMst_Client.Find "ClientIDRow = " & RsMst_GSL!ClientIDRow
                  If Not RsMst_Client.EOF Then
                     If UCase(IIf(IsNull(RsMst_Client!CMSPCode), "", RsMst_Client!CMSPCode)) <> UCase(oWS.Cells(I, 6)) Then
                        MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Client CMSP Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                        'mError = mError & I & " Client CMSP Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                        GoTo NextRow
                     End If
                  End If
               End If
              ' Check for more than one GSL
              If RsMst_GSL.RecordCount > 1 Then
                 CString = RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID
                 For z = 1 To RsMst_GSL.RecordCount
                     If CString <> RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID Then
                        MsgBox "CMS-P : " & RsMst_GSL!KRNo & "Commodity or CMP Mismatch please contact technology!!!"
                        End
                     End If
                     
                     ''''''checking for godown, commodity commtrack ncdex code and clientid'''''''''''''''''''''''''''''
                     RsMst_Godown.MoveFirst
                     RsMst_Godown.Find "GodownID = " & RsMst_GSL!GodownID
                     If Not RsMst_Godown.EOF Then
                        If UCase(IIf(IsNull(RsMst_Godown!CommtrackNCDEXCode), "", RsMst_Godown!CommtrackNCDEXCode)) <> UCase(oWS.Cells(I, 2)) Then
                           MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Godown NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                           'mError = mError & I & " Godown NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                           GoTo NextRow
                        End If
                     End If
                     RsMst_Commodity.MoveFirst
                     RsMst_Commodity.Find "CommodityID = " & RsMst_GSL!CommodityID
                     If Not RsMst_Commodity.EOF Then
                        If UCase(IIf(IsNull(RsMst_Commodity!CommtrackNCDEXCode), "", RsMst_Commodity!CommtrackNCDEXCode)) <> UCase(oWS.Cells(I, 5)) Then
                           MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Commodity NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                           'mError = mError & I & " Commodity NCDEX Commtrack Code Mismatch" & vbCrLf ''& vbCrLf '& Chr(13)
                           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                           GoTo NextRow
                        End If
                     End If
                     If LCase(Gen_UserRole) <> "power" Then
                        RsMst_Client.MoveFirst
                        RsMst_Client.Find "ClientID = " & RsMst_GSL!ClientID
                        ''RsMst_Client.Find "ClientIDRow = " & RsMst_GSL!ClientIDRow
                        If Not RsMst_Client.EOF Then
                           If UCase(IIf(IsNull(RsMst_Client!CMSPCode), "", RsMst_Client!CMSPCode)) <> UCase(oWS.Cells(I, 6)) Then
                              MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Client CMSP Code Mismatch" '& vbCrLf '& Chr(13)
                              'mError = mError & I & " Client CMSP Code Mismatch" '& vbCrLf '& Chr(13)
                              MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                              GoTo NextRow
                           End If
                        End If
                     End If
                     RsMst_GSL.MoveNext
                 Next
                 RsMst_GSL.MoveFirst
              End If
              
'              If oWS.Cells(I, 2) = "" Then
'                 ' Rejected Lot
'                 If oWS.Cells(I, 20) = "Rejected By Assayer" Then
'                    For z = 1 To RsMst_GSL.RecordCount
'                        RsMst_GSL!Flag = "J"
'                        RsMst_GSL.Update
'                        RsMst_GSL.MoveNext
'                    Next
'                    RsMst_GSL.MoveFirst
'                 End If
'              Else
                    ' Check for ISIN in master in case not found Add in master
                    RsMst_ISIN.MoveFirst
                    RsMst_ISIN.Find "ISINID =  '" & oWS.Cells(I, 4) & "' "
                    If RsMst_ISIN.EOF Then
                       Call SaveISIN(oWS.Cells(I, 4), RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID & "~" & oWS.Cells(I, 8), RsMst_GSL!CMPID, oWS.Cells(I, 11), oWS.Cells(I, 11), GetLocationIDFromCMP(RsMst_GSL!CMPID), RsMst_GSL!CommodityID, 2, Trim(oWS.Cells(I, 8)), "0")
                    Else
                       Call TableMst_CMP
                       RsMst_CMP.MoveFirst
                       RsMst_CMP.Find "Cmpid = " & RsMst_GSL!CMPID
                       If RsMst_ISIN!LocationID <> RsMst_CMP!LocationID Then
                          MsgBox "Critical error ISIN Location & GSL Location missmatch. Process Terminated. ISIN No=" & oWS.Cells(I, 2) & ""
                          End
                       End If
                       
                    End If
                    'If oWS.Cells(I, 20) = "Successful - Credited" Then
                       'insert Data in Mst_DR Table
                       RsMst_DR.AddNew
                       ''RsMst_DR!DRID = GetDRID
                       RsMst_DR!DRNo = oWS.Cells(I, 4)
                       'RsMst_DR!DRDate = Date
'Commented on 12 Nov 2014 by Parag after mail from Arush as COmmtrack file changed date format
'                       RsMst_DR!DRDate = CDate(oWS.Cells(i, 14))
                                            
                       If InStr(Trim(oWS.Cells(I, 14).Value), "-") > 0 Then
                          mDate = Split(Trim(oWS.Cells(I, 14).Value), "-")
                          If mDate(0) > 12 Then
                             mDate = Format(Trim(oWS.Cells(I, 14).Value), "mm/dd/yyyy")
                          Else
                             mDate = Format(Trim(oWS.Cells(I, 14).Value), "dd/mm/yyyy")
                          End If
                          RsMst_DR!DRDate = CDate(mDate) ' Trim(oWS.Cells(mRow, c + 1).Value) 'CDate(StrDate)  '
                       Else
                          RsMst_DR!DRDate = CDate(oWS.Cells(I, 14))
                       End If

                       RsMst_DR!Txn_Flag = "D"
                       RsMst_DR!NewISINID = oWS.Cells(I, 4)
                       RsMst_DR!ClientID = RsMst_GSL!ClientID
                       ''RsMst_DR!ClientIDRow = RsMst_GSL!ClientIDRow 'GetClientIDRow(oWS.Cells(i, 13))
                       RsMst_DR!Net_Weight_e = oWS.Cells(I, 10)
                       RsMst_DR!Remark = oWS.Cells(I, 12)
                       RsMst_DR!Status = oWS.Cells(I, 13) 'Left(oWS.Cells(I, 19), 200)
                       RsMst_DR!CommTrackFlag = "1"
                       'RsMst_DR!OldISINID = ""
                       'RsMst_DR!Net_Weight_d=
                       RsMst_DR.Update
                       If RsMst_GSL.RecordCount > 1 Then
                          CString = RsMst_GSL!CMPID & "~" & RsMst_GSL!CommodityID
                          For z = 1 To RsMst_GSL.RecordCount
                              RsMst_GSL!ISINID = oWS.Cells(I, 4)
                              RsMst_GSL.Update
                              RsMst_GSL.MoveNext
                           Next
                       Else
                          RsMst_GSL!ISINID = oWS.Cells(I, 4)
                          RsMst_GSL.Update
                       End If
                       DAdd = DAdd + 1
                    'End If
'              End If
           End If
        Else
           MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " CMS-E No (" & oWS.Cells(I, 4) & ") already exists" '& vbCrLf '& Chr(13)
           'mError = mError & I & " DR No (" & oWS.Cells(I, 4) & ") already exists" '& vbCrLf '& Chr(13)
           MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
           
        End If
      ElseIf oWS.Cells(I, 4) = "" And LCase(oWS.Cells(I, 8)) = "nonncdex" Then
        RsMst_GSL.Filter = "krno = " & Val(oWS.Cells(I, 3))
        For z = 1 To RsMst_GSL.RecordCount
        
            ''''''checking for godown, commodity commtrack ncdex code and clientid'''''''''''''''''''''''''''''
            RsMst_Godown.MoveFirst
            RsMst_Godown.Find "GodownID = " & RsMst_GSL!GodownID
            If Not RsMst_Godown.EOF Then
               If UCase(IIf(IsNull(RsMst_Godown!CommtrackNCDEXCode), "", RsMst_Godown!CommtrackNCDEXCode)) <> UCase(oWS.Cells(I, 2)) Then
                  MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Godown NCDEX Commtrack Code Mismatch" '& vbCrLf '& Chr(13)
                  'mError = mError & I & " Godown NCDEX Commtrack Code Mismatch" '& vbCrLf '& Chr(13)
                  MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                  GoTo NextRow
               End If
            End If
            RsMst_Commodity.MoveFirst
            RsMst_Commodity.Find "CommodityID = " & RsMst_GSL!CommodityID
            If Not RsMst_Commodity.EOF Then
               If UCase(IIf(IsNull(RsMst_Commodity!CommtrackNCDEXCode), "", RsMst_Commodity!CommtrackNCDEXCode)) <> UCase(oWS.Cells(I, 5)) Then
                  MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Commodity NCDEX Commtrack Code Mismatch" '& vbCrLf '& Chr(13)
                  'mError = mError & I & " Commodity NCDEX Commtrack Code Mismatch" '& vbCrLf '& Chr(13)
                  MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                  GoTo NextRow
               End If
            End If
            
            If LCase(Gen_UserRole) <> "power" Then
               RsMst_Client.MoveFirst
               RsMst_Client.Find "ClientID = " & RsMst_GSL!ClientID
               ''RsMst_Client.Find "ClientIDRow = " & RsMst_GSL!ClientIDRow
               If Not RsMst_Client.EOF Then
                  If UCase(IIf(IsNull(RsMst_Client!CMSPCode), "", RsMst_Client!CMSPCode)) <> UCase(oWS.Cells(I, 6)) Then
                     MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 0) = I & " Client CMSP Code Mismatch" '& vbCrLf '& Chr(13)
                     'mError = mError & I & " Client CMSP Code Mismatch" '& vbCrLf '& Chr(13)
                     MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
                     GoTo NextRow
                  End If
               End If
            End If
            RsMst_GSL!Flag1 = RsMst_GSL!Flag
            RsMst_GSL!Status = "O"
            RsMst_GSL!State = "J" ''-- Added
            RsMst_GSL!ISINID = "CMSP-" & oWS.Cells(I, 3)
            RsMst_GSL.Update
            RsMst_GSL.MoveNext
        Next
        If RsMst_GSL.RecordCount > 0 Then
           RsMst_GSL.MoveFirst
        End If
'     ElseIf oWS.Cells(I, 3) = "Y" Then
'        RsMst_ISIN.MoveFirst
'        RsMst_ISIN.Find "ISINID = '" & oWS.Cells(I, 2) & "'"
'        If RsMst_ISIN.EOF = True Then
'           RsMst_ISIN.MoveFirst
'           RsMst_ISIN.Find "ISINID = '" & oWS.Cells(I, 4) & "'"
'           If Not RsMst_ISIN.EOF = True Then
'              mCommodityID = RsMst_ISIN!CommodityID
'              mCMPID = RsMst_ISIN!CMPID
'              mLocationID = RsMst_ISIN!LocationID
'              Call SaveISIN(oWS.Cells(I, 2), GetCMPName(mCMPID) & "~" & GetCommodityName(mCommodityID) & "~" & Trim(oWS.Cells(I, 9)) & "~ NCMSL Upload ", mCMPID, oWS.Cells(I, 10), oWS.Cells(I, 11), mLocationID, mCommodityID, 2, Trim(oWS.Cells(I, 9)), "0")
'           Else
'              mError = mError & I & " ISIN not in ISIN Master " & oWS.Cells(I, 4) & vbCrLf 'Chr(13)
'           End If
'        Else
'           mError = mError & I & " ISIN Already in ISIN Master " & oWS.Cells(I, 2) & vbCrLf 'Chr(13)
'        End If
     End If
NextRow:
     I = I + 1
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

'                   mDate = Split(Trim(ps(c)), "/")
'                   'dd/nn/yy
'                   'StrDate = mdate(0) & "/" & mdate(1) & "/" & "20" & mdate(2)
'                   StrDate = "20" & mDate(2) & "/" & mDate(1) & "/" & mDate(0)

RsMst_ISIN.AddNew
RsMst_ISIN!ISINID = mISINID_
RsMst_ISIN!ISINDescription = mISINDescription_
RsMst_ISIN!CMPID = mCMPID_
mDate = Split(mNE_Date_, "/")
StrDate = mDate(2) & "/" & mDate(1) & "/" & mDate(0)
RsMst_ISIN!NE_Date = StrDate 'mNE_Date_
mDate = Split(mFED_, "/")
StrDate = mDate(2) & "/" & mDate(1) & "/" & mDate(0)
RsMst_ISIN!FED = StrDate 'mFED_
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
Call Grid_Head
TxtError = ""
mError = ""
CmdImport.Enabled = False
CmdGetFile.Enabled = True
lblFile.Refresh
lblFile.Caption = ""
ProgressBar1.Value = 0
lblProgress.Caption = ""
lblProcess.Caption = ""
End Sub

Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 1
MSHFlexGrid1.Cols = 1
End Sub

Private Sub CmdGetFile_Click()
Call Grid_Head
TxtError = ""
mError = ""
With cdialog
   .InitDir = App.Path
   .DialogTitle = "Select File To Be Upload"
   .ShowOpen
End With
If cdialog.CancelError Then
   lblFile.Refresh
   lblFile.Caption = ""
   CmdImport.Enabled = False
Else
   strFilename = cdialog.FileName
   lblFile.Refresh
   lblFile.Caption = strFilename 'GetFileFromPath(strFilename)
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
Call Grid_Head
TxtError = ""
mError = ""
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

If LCase(oWS.Cells(1, 1)) <> "s.no." Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 2)) <> "warehouse code" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 3)) <> "lot id" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 4)) <> "cmse lot id" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 5)) <> "commodity code" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 6)) <> "client id" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 7)) <> "client name" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 8)) <> "actual grade" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 9)) <> "nominal grade" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 10)) <> "actual quantity" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 11)) <> "fed" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 12)) <> "txn_ref_no" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 13)) <> "is_extinguish" Then
   Retval = False
ElseIf LCase(oWS.Cells(1, 14)) <> "dr date" Then
   Retval = False
End If

CheckFile = Retval

End Function


