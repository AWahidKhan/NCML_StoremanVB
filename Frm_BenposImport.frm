VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_BenposImport 
   Caption         =   "Upload Benpos"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   6495
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdMissingGSL 
      Caption         =   "Execute Queries on Azure DB For Missing GSL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton CmdImportQueries 
      Caption         =   "Import Queries"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton CmdExecQueries 
      Caption         =   "Execute Queries on Azure DB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame2 
      Height          =   5775
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Visible         =   0   'False
      Width           =   18975
      Begin VB.CheckBox Check1 
         Caption         =   "Select All"
         Height          =   255
         Left            =   4410
         TabIndex        =   20
         Top             =   180
         Width           =   1215
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5175
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   18615
         _ExtentX        =   32835
         _ExtentY        =   9128
         _Version        =   393216
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
      Begin VB.Label LblFrameGrid 
         AutoSize        =   -1  'True
         Caption         =   "Godown Details"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1320
      End
   End
   Begin VB.CommandButton CmdExcel 
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   11640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton CmdEmployeeData 
      Caption         =   "Get Employee Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.CommandButton cmdGetfile 
         Caption         =   "Get File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1815
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4005
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2910
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPImportDate 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   180
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   112066561
         CurrentDate     =   39989
      End
      Begin VB.Label LblDate 
         Caption         =   "Import Date"
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
         Left            =   840
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   10
         Top             =   1800
         Width           =   60
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Status :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   8
         Top             =   1440
         Width           =   60
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   345
         TabIndex        =   7
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label lblLines 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   6
         Top             =   2160
         Width           =   60
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Lines :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   570
      End
   End
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   5760
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Frm_BenposImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstbenpos As New ADODB.Recordset
Dim objFS As Scripting.FileSystemObject
Dim objFile As File
Dim objTS As TextStream
Dim objFScp As Scripting.FileSystemObject
Dim objFilecp As File
Dim objTScp As TextStream
Dim sFileLine As String
Dim sFileLineNo As Long
Dim temp As Variant
Dim arecordset As String
Dim strFilename As String
Dim strsql As String
Dim strdata As String
Dim t_line As String
Dim strfile As String
Dim strqry As String
Dim sArray(0 To 10) As String
Dim EmpID As Variant

Const strChecked = "þ"
Const strUnChecked = "q"

Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim oWS As Excel.Worksheet

Dim mQueryXL As Boolean

Private Sub UpdateWHRDetails()

tmp = "ALTER TABLE [Txn_WHR_Detail] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

tmp = "SET IDENTITY_INSERT [Txn_WHR_Detail] OFF"
Call TmpTableAzureStoreMan

tmp = "Select * from Txn_WHR_Detail Where CreatedDate >= '" & Date - 20 & "' " 'And WHRNo = 2115000132 "
Call TmpTable

tmp = "ALTER TABLE [Txn_WHR_Detail] NoCHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

tmp = "SET IDENTITY_INSERT [Txn_WHR_Detail] ON"
Call TmpTableAzureStoreMan

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "Select * from Txn_WHR_Detail WHere WHRNo = '" & TmpRs!WHRNo & "'"
       Call TmpTableAzureStoreMan
       If TmpRsAzureStoreMan.RecordCount = 0 Then
          TmpRsAzureStoreMan.AddNew
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
          TmpRsAzureStoreMan!WHRDate = TmpRs!WHRDate
          TmpRsAzureStoreMan!ConditionID = TmpRs!ConditionID
          TmpRsAzureStoreMan!StorageFrom = TmpRs!StorageFrom
          TmpRsAzureStoreMan!StorageTo = TmpRs!StorageTo
          TmpRsAzureStoreMan!Grade = TmpRs!Grade
          TmpRsAzureStoreMan!Weigher = TmpRs!Weigher
          TmpRsAzureStoreMan!Assayer = TmpRs!Assayer
          TmpRsAzureStoreMan!GradeByAssayer = TmpRs!GradeByAssayer
          TmpRsAzureStoreMan!GradeValidTill = TmpRs!GradeValidTill
          TmpRsAzureStoreMan!StorageCharges = TmpRs!StorageCharges
          TmpRsAzureStoreMan!StorageAmount = TmpRs!StorageAmount
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!LicenseID = TmpRs!LicenseID
          TmpRsAzureStoreMan!WHRNoOfBags = TmpRs!WHRNoOfBags
          TmpRsAzureStoreMan!WHRQuantity = TmpRs!WHRQuantity
          TmpRsAzureStoreMan!PrintFlag = TmpRs!PrintFlag
          TmpRsAzureStoreMan!Remark = TmpRs!Remark
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!PledgeID = TmpRs!PledgeID
          TmpRsAzureStoreMan!oldLicenseID = TmpRs!oldLicenseID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!DeliveredQuantity = TmpRs!DeliveredQuantity
          TmpRsAzureStoreMan!DeliveredNoOfBags = TmpRs!DeliveredNoOfBags
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!HologramNo = TmpRs!HologramNo
          TmpRsAzureStoreMan!PolicyNo = TmpRs!PolicyNo
          TmpRsAzureStoreMan!CoverNoteNo = TmpRs!CoverNoteNo
          TmpRsAzureStoreMan!SaidToContain = TmpRs!SaidToContain
          TmpRsAzureStoreMan!EstimatedWeight = TmpRs!EstimatedWeight
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!SRAsPerBag = TmpRs!SRAsPerBag
          TmpRsAzureStoreMan!CARNo = TmpRs!CARNo
          TmpRsAzureStoreMan!ConvertedClient = TmpRs!ConvertedClient
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!SendToMktYard = TmpRs!SendToMktYard
          TmpRsAzureStoreMan!SendToMktYardOn = TmpRs!SendToMktYardOn
          TmpRsAzureStoreMan!LotWiseWHR = TmpRs!LotWiseWHR
          TmpRsAzureStoreMan.Update
       Else
            TmpRsAzureStoreMan!Weigher = TmpRs!Weigher
            TmpRsAzureStoreMan!Assayer = TmpRs!Assayer
            TmpRsAzureStoreMan!StorageCharges = TmpRs!StorageCharges
            TmpRsAzureStoreMan!StorageAmount = TmpRs!StorageAmount
            TmpRsAzureStoreMan!LicenseID = TmpRs!LicenseID
            TmpRsAzureStoreMan!PrintFlag = TmpRs!PrintFlag
            TmpRsAzureStoreMan!Flag = TmpRs!Flag
            TmpRsAzureStoreMan!PledgeID = TmpRs!PledgeID
            TmpRsAzureStoreMan!DeliveredQuantity = TmpRs!DeliveredQuantity
            TmpRsAzureStoreMan!DeliveredNoOfBags = TmpRs!DeliveredNoOfBags
            TmpRsAzureStoreMan!HologramNo = TmpRs!HologramNo
            TmpRsAzureStoreMan!PolicyNo = TmpRs!PolicyNo
            TmpRsAzureStoreMan!CoverNoteNo = TmpRs!CoverNoteNo
            TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
            TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
'            TmpRsAzureStoreMan!CARNo = TmpRs!CARNo
            TmpRsAzureStoreMan!Rate = TmpRs!Rate
            TmpRsAzureStoreMan!SendToMktYard = TmpRs!SendToMktYard
            TmpRsAzureStoreMan!SendToMktYardOn = TmpRs!SendToMktYardOn
            TmpRsAzureStoreMan.Update
       End If
       TmpRs.MoveNext
   Next
End If


tmp = "ALTER TABLE [Txn_WHR_Detail] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

tmp = "SET IDENTITY_INSERT [Txn_WHR_Detail] OFF"
Call TmpTableAzureStoreMan


End Sub

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
   For I = 1 To MSHFlexGrid1.Rows - 2
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Wingdings"
      MSHFlexGrid1.CellFontSize = 12
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid1.Text = strChecked
   Next
Else
   For I = 1 To MSHFlexGrid1.Rows - 2
      MSHFlexGrid1.row = I
      MSHFlexGrid1.Col = 0
      MSHFlexGrid1.CellFontName = "Wingdings"
      MSHFlexGrid1.CellFontSize = 12
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
      MSHFlexGrid1.Text = strUnChecked
   Next
End If
End Sub

Private Sub CmdCancel_Click()
LblFile.Caption = ""
lblStatus.Caption = ""
lblLines.Caption = ""
strFilename = ""
CmdGetFile.Enabled = True
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdEmployeeData_Click()

Call Grid_Head
mQueryXL = False
tmp = "Select a.*,b.Designation from Mst_Employee a Inner Join Mst_Designation b On a.DesignationID = b.DesignationID where a.Flag = 'Y' And a.DesignationID not in (1,7,14)  Order By a.DesignationID "

Call Tmp3Table

tmp = "Delete from Mst_EmpGeoMap"
Call Tmp5Table

a = 1
For x = 1 To TmpRs3.RecordCount

   EmpID = TmpRs3!EmployeeID
   
   Call GetLinkedGodownForEmployee(EmpID)
   
   wc = " Where (EMP.EmployeeID = " & EmpID & " Or EMP1.EmployeeID = " & EmpID & " Or EMP2.EmployeeID =" & EmpID & " Or"
   wc = wc & " EMP3.EmployeeID = " & EmpID & "  Or EMP4.EmployeeID = " & EmpID & " Or"
   wc = wc & " EMP6.EmployeeID = " & EmpID & " )"
   

    tmp = "Select distinct EGM.CMPID"
    tmp = tmp & " From Mst_EmpCMPMapping EGM"
    tmp = tmp & " Right Outer Join Mst_Employee EMP On EGM.EmployeeID = EMP.EmployeeID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP1 On EMP.DirectReportingTo = EMP1.EmployeeID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP2 On EMP1.DirectReportingTo = EMP2.EmployeeID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP3 On EMP2.DirectReportingTo = EMP3.EmployeeID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP4 On EMP3.DirectReportingTo = EMP4.EmployeeID"
    tmp = tmp & " Left Outer Join Mst_Employee EMP6 On EMP.ConnectedToRAG = EMP6.EmployeeID"
    If wc <> "" Then
       tmp = tmp & wc & "  And CMPID is not Null"
    Else
       tmp = tmp & " where CMPID is not Null"
    End If
    tmp = tmp & " Order By EGM.CMPID"
    
    Call Tmp4Table
    
    If TmpRs4.RecordCount = 0 Then
       Gen_WcCMP = "CMPID=0"
    Else
       Gen_WcCMP = "CMPID in ("
       For I = 1 To TmpRs4.RecordCount
           If I = 1 Then
              Gen_WcCMP = Gen_WcCMP & TmpRs4.Fields(0) & " "
           Else
              Gen_WcCMP = Gen_WcCMP & "," & TmpRs4.Fields(0) & " "
           End If
           TmpRs4.MoveNext
       Next
       Gen_WcCMP = Gen_WcCMP & " )"
    End If

    tmp = " Select Distinct '" & TmpRs3!EmployeeID & "' 'EmployeeID','" & TmpRs3!EmployeeName & "','" & TmpRs3!EmployeeNo & "',MC.CMPID,MC.CMPName,ML.LocationName,ML.LocationID,MS.StateName,MS.StateID,MR.Region,MR.RegionID,'" & TmpRs3!Designation & "','" & TmpRs3!DesignationID & "' 'DesignationID'"
    tmp = tmp & " From Mst_CMP MC"
    tmp = tmp & " Inner Join Mst_Location ML On MC.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_STate MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " WHere MC." & Gen_WcCMP
    
    Call Tmp2Table
    
    tmp = "Select * from Mst_DesignationBasicGeoType"
    Call Tmp4Table
    
    For xy = 1 To TmpRs2.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        
        For C = 0 To MSHFlexGrid1.Cols - 1
            MSHFlexGrid1.TextMatrix(a, C) = IIf(IsNull(TmpRs2.Fields(C)), "", TmpRs2.Fields(C))
            If InStr(MSHFlexGrid1.TextMatrix(a, C), "/") > 0 Then
               If IsDate(MSHFlexGrid1.TextMatrix(a, C)) Then
                  MSHFlexGrid1.TextMatrix(a, C) = Format(MSHFlexGrid1.TextMatrix(a, C), "dd-mmm-yyyy")
               Else
                  MSHFlexGrid1.TextMatrix(a, C) = MSHFlexGrid1.TextMatrix(a, C)
               End If
            Else
              MSHFlexGrid1.TextMatrix(a, C) = MSHFlexGrid1.TextMatrix(a, C)
            End If
        Next
        If TmpRs4.RecordCount > 0 Then
           TmpRs4.MoveFirst
           TmpRs4.Find "DesignationID = " & TmpRs2!DesignationID & ""
           If TmpRs4.EOF = False Then
              mGeoTypeFlag = TmpRs4!GeoTypeFlag
              For C = 0 To MSHFlexGrid1.Cols - 1
                  If LCase(TmpRs4!ColumnName) = LCase(MSHFlexGrid1.TextMatrix(0, C)) Then
                     mGeoTypeID = Val(MSHFlexGrid1.TextMatrix(a, C))
                     Exit For
                  End If
              Next
             
              tmp = "Select * from Mst_EmpGeoMap Where EmployeeID = " & TmpRs2!EmployeeID & " And GeoType = '" & mGeoTypeFlag & "' And GeoID = '" & mGeoTypeID & "' "
              Call Tmp5Table
              If TmpRs5.RecordCount = 0 Then
                 TmpRs5.AddNew
                 TmpRs5!EmployeeID = TmpRs2!EmployeeID
                 TmpRs5!GeoType = mGeoTypeFlag
                 TmpRs5!GeoID = Val(mGeoTypeID)
                 TmpRs5!Status = "A"
                 TmpRs5!G_UID = GetGUID
                 TmpRs5!CreatedBy = "2047"
                 TmpRs5!CreatedDate = Now
                 TmpRs5.Update
              End If
           End If
        End If
        
        
        a = a + 1
        TmpRs2.MoveNext
    Next
    TmpRs3.MoveNext
Next


tmp = "Select a.*,b.Designation from Mst_Employee a Inner Join Mst_Designation b On a.DesignationID = b.DesignationID where a.Flag = 'Y' And a.DesignationID in (1,7)  Order By a.DesignationID "
Call Tmp3Table

a = MSHFlexGrid1.Rows

For x = 1 To TmpRs3.RecordCount

   EmpID = TmpRs3!EmployeeID
   
   tmp = "Select *,'" & TmpRs3!DesignationID & "' 'DesignationID','" & TmpRs3!EmployeeID & "' 'EmployeeID' from Mst_Region"
   Call Tmp2Table

    
    tmp = "Select * from Mst_DesignationBasicGeoType Where DesignationID in (1,7) "
    Call Tmp4Table
    
    For xy = 1 To TmpRs2.RecordCount
        If TmpRs4.RecordCount > 0 Then
           TmpRs4.MoveFirst
           TmpRs4.Find "DesignationID = " & TmpRs2!DesignationID & ""
           If TmpRs4.EOF = False Then
              mGeoTypeFlag = TmpRs4!GeoTypeFlag
              mGeoTypeID = TmpRs2!RegionID
              'mKey = "" & mGeoTypeFlag & "_" & Val(mGeoTypeID) & ""
              
              tmp = "Select * from Mst_EmpGeoMap Where EmployeeID = " & TmpRs2!EmployeeID & " And GeoType = '" & mGeoTypeFlag & "' And GeoID = '" & mGeoTypeID & "' "
              Call Tmp5Table
              If TmpRs5.RecordCount = 0 Then
                 TmpRs5.AddNew
                 TmpRs5!EmployeeID = TmpRs2!EmployeeID
                 TmpRs5!GeoType = mGeoTypeFlag
                 TmpRs5!GeoID = Val(mGeoTypeID)
                 TmpRs5!Status = "A"
                 TmpRs5!G_UID = GetGUID
                 TmpRs5!CreatedBy = "2047"
                 TmpRs5!CreatedDate = Now
                 TmpRs5.Update
              End If
           End If
        End If
        a = a + 1
        TmpRs2.MoveNext
    Next
    TmpRs3.MoveNext
Next


End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp

If mQueryXL = False Then
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "EMPCMPMapping.xls")
Else
   FileName = App.Path & "\Excel\" & Gen_mTimeStamp & Gen_UserName & "ExecutedOnAuzre.xls"
   Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "ExecutedOnAuzre.xls")
End If
   
Call Xls_Detail_RptExcel

If mQueryXL = True Then
   strFilLocation = FileName
   Call Gen_AutoMailReport
   If Len(Dir(strFilLocation)) > 1 Then Kill strFilLocation
'   MsgBox ("Excel Report send on mail!!!")
End If

End Sub
Public Sub Gen_AutoMailReport()

TOADD = Gen_UserEmailID
TOCC = "mayur.d@ncml.com,ullhas.p@ncml.com,parag.r@ncml.com"

mMsg = ""

mMsg = "Dear Sir/Madam, "
mMsg = mMsg & "<br><br>"
mMsg = mMsg & "Please find attached excel "
mMsg = mMsg & "<br><br>"

mMsgSubject = "Executed Queries on " & Now & " "

Call SendMailMIS(TOADD, TOCC, mMsgSubject, mMsg, FileName)

End Sub

Private Function SendMailMIS(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mFileName_ As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession

servertxt = "mail.ncml.com"

htmlStringhead = "<html><body><font size= 3 face= Zurich BT color = Blue><table border = 1>"
htmlStringtail = "</table></Font></body></html>"
   

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") '"mail.ncml.com" '"mail.ncml.com"
objSMTP.EnableLogging = True
objSMTP.LogFilePath = "C:\temp\PAF_LOG.txt"
objSMTP.FromAddr = Gen_UserEmailID
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "vaishali.k@ncml.com"
End If
objSMTP.ToAddr = ToADD_
objSMTP.AddAttachment mFileName_
objSMTP.Subject = MsgSubject_ & " (from StoreMan)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If


End Function

Public Sub Xls_Detail_RptExcel()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer
Dim mCol As Integer
If mQueryXL = False Then
   MsgBox ("Wait till next message")
End If
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
If mQueryXL = False Then
   Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "EMPCMPMapping.xls")
Else
   Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "ExecutedOnAuzre.xls")
End If
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
If mQueryXL = False Then
   MsgBox ("Excel file created !!! ")
   oXL.Visible = True
Else
   oXL.Quit
End If

End Sub

Private Sub GetLinkedGodownForEmployee(EmpID_ As Variant)


tmp = " Select ME.EmployeeID,ME.DirectReportingTo,MD.DirectLink from Mst_Employee  ME "
tmp = tmp & " Inner Join Mst_Designation MD On ME.DesignationID = MD.DesignationID "
tmp = tmp & " Where Me.EmployeeID = " & EmpID_ & " "

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   If LCase(TmpRs1!DirectLink) = "n" Then
      EmpID = TmpRs1!DirectReportingTo
      Call GetLinkedGodownForEmployee(EmpID)
   ElseIf LCase(TmpRs1!DirectLink) = "y" Then
      EmpID = TmpRs1!EmployeeID
      Exit Sub
   End If
End If

End Sub


Private Sub CmdExecQueries_Click()
mQueryXL = True
On Error GoTo Error
DoEvents
For x = 1 To MSHFlexGrid1.Rows - 1
    If MSHFlexGrid1.TextMatrix(x, 1) = "" Then Exit For
    If MSHFlexGrid1.TextMatrix(x, 0) = strChecked Then
       tmp = Trim(MSHFlexGrid1.TextMatrix(x, 1))
       Call TmpTableAzureStoreMan
       MSHFlexGrid1.TextMatrix(x, 2) = "Done"
    Else
       MSHFlexGrid1.TextMatrix(x, 2) = "Not selected for execution"
    End If

Error: If MSHFlexGrid1.TextMatrix(x, 2) = "" Then MSHFlexGrid1.TextMatrix(x, 2) = "Some Error" 'Error.Description

Next

Call CmdExcel_Click
Call Grid_HeadQueries
MsgBox "Done!!!"

End Sub

Private Sub CmdGetFile_Click()
CmdGetFile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = False
CmdCancel.Enabled = True
lblStatus.Refresh
lblStatus.Caption = ""
LblFile.Refresh
LblFile.Caption = ""
With cdialog
    .InitDir = App.Path
    .DialogTitle = "FILE UPDATION"
    .Filter = "Tx files(*.txt)|*.txt"
    .ShowOpen
    .CancelError = True
End With
strFilename = cdialog.FileName
LblFile.Refresh
LblFile.Caption = strFilename
If cdialog.CancelError Then
    lblStatus.Refresh
    lblStatus.Caption = ""
Else
    lblStatus.Refresh
    lblStatus.Caption = "File Dectected..." & GetFileFromPath(strFilename)
'    If Left(GetFileFromPath(strFilename), 11) = "Benpos File" Then
'        cmdUpdate.Enabled = True
'    Else
'        lblStatus.Caption = "Please Select the correct file "
'        cmdGetfile.Enabled = True
'    End If
End If
CmdUpdate.Enabled = True
End Sub
Function GetFileFromPath(vPath As String)
    Dim Items() As String
    Items = Split(vPath, "\")
    If UBound(Items) = -1 Then Exit Function
    GetFileFromPath = Items(UBound(Items))
End Function

Private Sub CmdImportQueries_Click()
LblFrameGrid.Caption = "Queries"
'On Error GoTo errortrap

With cdialog
     .DialogTitle = "Open Excel File..."
     .Filter = "Excel files(*.xls;*.xlsx)|*.xls;*.xlsx" '|Text files(*.txt)|*.txt"
     .MaxFileSize = &H7FFF ' 32KB filename buffer
    'same as .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNLongNames or cdlOFNExplorer --Or cdlOFNAllowMultiselect
     .Flags = &H4 Or &H800 Or &H40000 Or &H80000 Or &H200
     .CancelError = True
procReOpen:
     .ShowOpen
End With
Path = cdialog.FileName
mFilenames = cdialog.FileName
If LblFile.Caption <> Path Then
 '  Call Clear
End If
LblFile.Caption = Path
If LblFile.Caption <> "" Then
   MsgBox "Wait till next message!!!"
   Status = 0
   RowCount = 0
   Set oXL = New Excel.Application
   Set oWB = oXL.Workbooks.Open(LblFile.Caption)
   Set oWS = oWB.Worksheets(1)
   Call Grid_HeadQueries
   If CheckQueryFile = False Then
      MsgBox "Please select proper file !!!"
      oXL.Quit
      Exit Sub
   End If
   Call FillData
   oXL.Quit
End If
Exit Sub
errortrap:   MsgBox Err.Description, vbExclamation, "Error"


End Sub
Private Sub FillData()
oWS.ClearArrows
Call Grid_HeadQueries
RowCount = oWS.Cells(oWS.Rows.Count, "A").End(xlUp).row
lblLines.Caption = RowCount - 1
   For l = 1 To RowCount
       If oWS.Cells(l + 1, 1) <> "" Then
          For C = 0 To MSHFlexGrid1.Cols - 2
              MSHFlexGrid1.TextMatrix(l, C + 1) = Trim(oWS.Cells(l + 1, C + 1).Value)
          Next
          MSHFlexGrid1.row = l
          MSHFlexGrid1.Col = 0
          MSHFlexGrid1.CellFontName = "Wingdings"
          MSHFlexGrid1.CellFontSize = 12
          MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
          MSHFlexGrid1.TextMatrix(l, 0) = strChecked
          MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       End If
   Next
End Sub



Private Function CheckQueryFile() As Boolean
Dim Retval As Boolean
Retval = True
'Set oWS = oWB.Worksheets("Sheet1")
If LCase(Trim(oWS.Cells(1, 1))) <> LCase("Queries for Destination Store-Man DB") Then
   Retval = False
   MsgBox " Please Check Column Heading for Column 1 "
End If

CheckQueryFile = Retval

End Function

Private Sub CmdMissingGSL_Click()

Call UpdateWHRDetails

'----Client Mapping
'----User Master Data Mapping
tmp = "Select ClientID 'ClientIDRow',ClientName,ExchangeTypeID,DepositoryID,Flag,Address,Organization,EmailID,TelephoneNo,MobileNo,ClientIDRow 'ClientID',G_UID,NAVCode,Concurrency,CreatedBy,CreatedDate,UpdatedDate,UpdatedBy,ClientGroupID,FAGUpdated,FAGUpdatedDate,RNAVCode,NAVClientName,ChequeReturn,Status,Converted,ConvertedClientIDRow 'ConvertedClientID',SCMFlag,CMSPCode,KYCFlag,RiskType,CorpClient,CRFID,TANNO from Mst_CLient Where CreatedDate >= '" & Date - 40 & "' And CreatedDate < '" & Date + 1 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Mst_CLient] ON"
       Call TmpTableAzureStoreMan1
       
       tmp = "Select * from Mst_CLient WHere ClientID = '" & TmpRs!ClientID & "' "
       Call TmpTableAzureStoreMan
       
       If TmpRsAzureStoreMan.RecordCount = 0 Then
          
          tmp = "ALTER TABLE [Mst_CLient] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          
          mStrQuery = GetQueyForWebservice("Mst_CLient", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
'            mStrQuery = Replace(mStrQuery, ",ClientID,", ",ClientIDRow-,")
'            mStrQuery = Replace(mStrQuery, ",CLientIDRow,", ",ClientID,")
'            mStrQuery = Replace(mStrQuery, ",ClientIDRow-,", ",ClientIDRow,")
            '
'             mStrQuery = Replace(mStrQuery, ",ConvertedClientIDRow,", ",ConvertedClientID,")
            tmp = "SET IDENTITY_INSERT [Mst_CLient] ON " & mStrQuery
            Call TmpTableAzureStoreMan
          End If
       Else
'          tmp = "ALTER TABLE [Mst_CLient] NoCHECK CONSTRAINT ALL"
'          Call TmpTableAzureStoreMan
'          tmp = ""
'          mStrQuery = GetQueyForWebservice("Mst_CLient", "E", TmpRs, "Where ClientID = '" & TmpRs!ClientIDRow & "'", tmp, "ClientID")
'          If mStrQuery <> "" Then
'            mStrQuery = Replace(mStrQuery, ",ClientID,", ",ClientIDRow-,")
'            mStrQuery = Replace(mStrQuery, ",CLientIDRow,", ",ClientID,")
'            mStrQuery = Replace(mStrQuery, ",ClientIDRow-,", ",ClientIDRow,")
'            tmp = mStrQuery
'            Call TmpTableAzureStoreMan
'          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Mst_CLient] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Mst_CLient] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan



'----GSL Mapping
tmp = "Select * from Mst_GSL Where (CreatedDate >= '" & Date - 5 & "' And CreatedDate < '" & Date + 1 & "') " ' Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo And a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType WHere b.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_Spillage a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.OldGSLID)+'-'+a.SM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.NewGSLID)+'-'+a.NewSM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) "
Call TmpTable


If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount

       tmp = "ALTER TABLE [Mst_GSL] NoCHECK CONSTRAINT ALL"
       Call TmpTableAzureStoreMan

       tmp = "Select * from Mst_GSL WHere SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = '" & TmpRs!GSLID & "'"
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          TmpRsAzureStoreMan.AddNew
          TmpRsAzureStoreMan!GSLID = TmpRs!GSLID
          TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update
       Else
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update

       End If
       TmpRs.MoveNext
   Next
End If


'--For CDTF and CITF Transaction GSL Update


tmp = "Select a.GSLID,a.SM_Prefix  from Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo And a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType WHere b.Concurrency >= '" & Date - 5 & "'"
Call Tmp1Table


If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount

       tmp = "Select * from Mst_GSL Where GSLID = " & TmpRs1!GSLID & " And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
       Call TmpTable


       tmp = "ALTER TABLE [Mst_GSL] NoCHECK CONSTRAINT ALL"
       Call TmpTableAzureStoreMan

       tmp = "Select * from Mst_GSL WHere SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = '" & TmpRs!GSLID & "'"
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          TmpRsAzureStoreMan.AddNew
          TmpRsAzureStoreMan!GSLID = TmpRs!GSLID
          TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update
       Else
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update

       End If
       TmpRs1.MoveNext
   Next
End If




'--For Spillage and Madeup Transaction GSL Update
'tmp = "Select * from Mst_GSL Where (CreatedDate >= '" & Date - 20 & "' And CreatedDate < '" & Date & "') Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo And a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType WHere b.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_Spillage a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.OldGSLID)+'-'+a.SM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.NewGSLID)+'-'+a.NewSM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) "

tmp = "Select a.GSLID,a.SM_Prefix  from Txn_Spillage a Where  a.Concurrency >= '" & Date - 5 & "'"
Call Tmp1Table


If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount

       tmp = "Select * from Mst_GSL Where GSLID = " & TmpRs1!GSLID & " And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
       Call TmpTable


       tmp = "ALTER TABLE [Mst_GSL] NoCHECK CONSTRAINT ALL"
       Call TmpTableAzureStoreMan

       tmp = "Select * from Mst_GSL WHere SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = '" & TmpRs!GSLID & "'"
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          TmpRsAzureStoreMan.AddNew
          TmpRsAzureStoreMan!GSLID = TmpRs!GSLID
          TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update
       Else
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update

       End If
       TmpRs1.MoveNext
   Next
End If



'--For GSL Transfer Old GSL ID Transaction GSL Update
'tmp = "Select * from Mst_GSL Where (CreatedDate >= '" & Date - 20 & "' And CreatedDate < '" & Date & "') Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo And a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType WHere b.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_Spillage a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.OldGSLID)+'-'+a.SM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.NewGSLID)+'-'+a.NewSM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) "

tmp = "Select a.OldGSLID,a.SM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 5 & "'"
Call Tmp1Table


If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount

       tmp = "Select * from Mst_GSL Where GSLID = " & TmpRs1!OldGSLID & " And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
       Call TmpTable


       tmp = "ALTER TABLE [Mst_GSL] NoCHECK CONSTRAINT ALL"
       Call TmpTableAzureStoreMan

       tmp = "Select * from Mst_GSL WHere SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = '" & TmpRs!GSLID & "'"
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          TmpRsAzureStoreMan.AddNew
          TmpRsAzureStoreMan!GSLID = TmpRs!GSLID
          TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update
       Else
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update

       End If
       TmpRs1.MoveNext
   Next
End If

'--For GSL Transfer New GSL ID Transaction GSL Update
'tmp = "Select * from Mst_GSL Where (CreatedDate >= '" & Date - 20 & "' And CreatedDate < '" & Date & "') Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b on a.CxTFNo = b.CxTFNo And a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType WHere b.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix  from Txn_Spillage a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.OldGSLID)+'-'+a.SM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) Or (Convert(Varchar,GSLID)+'-'+SM_Prefix in (Select Convert(Varchar,a.NewGSLID)+'-'+a.NewSM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 20 & "')) "

tmp = "Select a.NewGSLID,a.NewSM_Prefix  from  Txn_GSL_Transfer a Where  a.Concurrency >= '" & Date - 5 & "'"
Call Tmp1Table


If TmpRs1.RecordCount > 0 Then
   For I = 1 To TmpRs1.RecordCount
       tmp = "Select * from Mst_GSL Where GSLID = " & TmpRs1!NewGSLID & " And SM_Prefix = '" & TmpRs1!NewSM_Prefix & "'"
       Call TmpTable
       tmp = "ALTER TABLE [Mst_GSL] NoCHECK CONSTRAINT ALL"
       Call TmpTableAzureStoreMan
       tmp = "Select * from Mst_GSL WHere SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = '" & TmpRs!GSLID & "'"
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          TmpRsAzureStoreMan.AddNew
          TmpRsAzureStoreMan!GSLID = TmpRs!GSLID
          TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update
       Else
          TmpRsAzureStoreMan!CMPID = TmpRs!CMPID
          TmpRsAzureStoreMan!CommodityID = TmpRs!CommodityID
          TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
          TmpRsAzureStoreMan!StackNo = TmpRs!StackNo
          TmpRsAzureStoreMan!LotNo = TmpRs!LotNo
          TmpRsAzureStoreMan!DepositeWeight = TmpRs!DepositeWeight
          TmpRsAzureStoreMan!DepositeBags = TmpRs!DepositeBags
          TmpRsAzureStoreMan!WHRBags = TmpRs!WHRBags
          TmpRsAzureStoreMan!CDTFBags = TmpRs!CDTFBags
          TmpRsAzureStoreMan!BalanceBags = TmpRs!BalanceBags
          TmpRsAzureStoreMan!WHRWeight = TmpRs!WHRWeight
          TmpRsAzureStoreMan!CDTFWeight = TmpRs!CDTFWeight
          TmpRsAzureStoreMan!BalanceWeight = TmpRs!BalanceWeight
          TmpRsAzureStoreMan!ExchangeTypeID = TmpRs!ExchangeTypeID
          TmpRsAzureStoreMan!ClientID = TmpRs!ClientIDRow
          TmpRsAzureStoreMan!PledgeNo = TmpRs!PledgeNo
          TmpRsAzureStoreMan!Flag = TmpRs!Flag
          TmpRsAzureStoreMan!ISINID = TmpRs!ISINID
          TmpRsAzureStoreMan!CdtfNo = TmpRs!CdtfNo
          TmpRsAzureStoreMan!CDTFDate = TmpRs!CDTFDate
          TmpRsAzureStoreMan!FirstISINID = TmpRs!FirstISINID
          TmpRsAzureStoreMan!PledgeBags = TmpRs!PledgeBags
          TmpRsAzureStoreMan!PledgeWeight = TmpRs!PledgeWeight
          TmpRsAzureStoreMan!SpillageFlag = TmpRs!SpillageFlag
          TmpRsAzureStoreMan!DepositAdjWt = TmpRs!DepositAdjWt
          TmpRsAzureStoreMan!WithdrawalAdjWt = TmpRs!WithdrawalAdjWt
          TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
          TmpRsAzureStoreMan!Flag1 = TmpRs!Flag1
          TmpRsAzureStoreMan!LCDTF_Date = TmpRs!LCDTF_Date
          TmpRsAzureStoreMan!GrossWt = TmpRs!GrossWt
          TmpRsAzureStoreMan!KRNo = TmpRs!KRNo
          TmpRsAzureStoreMan!KRNoBatchID = TmpRs!KRNoBatchID
          TmpRsAzureStoreMan!KRDate = TmpRs!KRDate
          TmpRsAzureStoreMan!BGSLID = TmpRs!BGSLID
          TmpRsAzureStoreMan!FGSLID = TmpRs!FGSLID
          TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
          TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
          TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
          TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
          TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
          TmpRsAzureStoreMan!ContractID = TmpRs!ContractID
          TmpRsAzureStoreMan!BagSize = TmpRs!BagSize
          TmpRsAzureStoreMan!BagEquivalent = TmpRs!BagEquivalent
          TmpRsAzureStoreMan!WtQty = TmpRs!WtQty
          TmpRsAzureStoreMan!CRAmount = TmpRs!CRAmount
          TmpRsAzureStoreMan!CRDate = TmpRs!CRDate
          TmpRsAzureStoreMan!NContractID = TmpRs!NContractID
          TmpRsAzureStoreMan!SCMFlag = TmpRs!SCMFlag
          TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
          TmpRsAzureStoreMan!MadeupFlag = TmpRs!MadeupFlag
          TmpRsAzureStoreMan!ConnectedGSLID = TmpRs!ConnectedGSLID
          TmpRsAzureStoreMan!BSM_Prefix = TmpRs!BSM_Prefix
          TmpRsAzureStoreMan!FSM_Prefix = TmpRs!FSM_Prefix
          TmpRsAzureStoreMan!CMSPBatchID = TmpRs!CMSPBatchID
          TmpRsAzureStoreMan!CMSPBatchGeneratedBy = TmpRs!CMSPBatchGeneratedBy
          TmpRsAzureStoreMan!CMSPBatchGeneratedDate = TmpRs!CMSPBatchGeneratedDate
          TmpRsAzureStoreMan!CMSPUpdatedBy = TmpRs!CMSPUpdatedBy
          TmpRsAzureStoreMan!CMSPUpdatedDate = TmpRs!CMSPUpdatedDate
          TmpRsAzureStoreMan!TransactionType = TmpRs!TransactionType
          TmpRsAzureStoreMan!TransactionID = TmpRs!TransactionID
          TmpRsAzureStoreMan!TransactionDate = TmpRs!TransactionDate
          TmpRsAzureStoreMan!MoisturePer = TmpRs!MoisturePer
          TmpRsAzureStoreMan!CITF_MoisturePer = TmpRs!CITF_MoisturePer
          TmpRsAzureStoreMan!MoistureValue = TmpRs!MoistureValue
          TmpRsAzureStoreMan!MAWWt = TmpRs!MAWWt
          TmpRsAzureStoreMan!OriginalDepositFlag = TmpRs!OriginalDepositFlag
          TmpRsAzureStoreMan!WHRNo = TmpRs!WHRNo
          TmpRsAzureStoreMan!NFinFlag = TmpRs!NFinFlag
          TmpRsAzureStoreMan!RateCardType = TmpRs!RateCardType
          TmpRsAzureStoreMan!BillingCycleID = TmpRs!BillingCycleID
          TmpRsAzureStoreMan!BillingUnit = TmpRs!BillingUnit
          TmpRsAzureStoreMan!Rate = TmpRs!Rate
          TmpRsAzureStoreMan!RateForDailyPnL = TmpRs!RateForDailyPnL
          TmpRsAzureStoreMan!DepConDID = TmpRs!DepConDID
          TmpRsAzureStoreMan!NoOfDays = TmpRs!NoOfDays
          TmpRsAzureStoreMan.Update

       End If
       TmpRs1.MoveNext
   Next
End If

tmp = "ALTER TABLE [Mst_GSL] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

'----Employee Mapping
tmp = "Select * from Mst_Employee Where CreatedDate >= '" & Date - 5 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Mst_Employee] ON"
       Call TmpTableAzureStoreMan1
       
       tmp = "Select * from Mst_Employee WHere EmployeeNo = '" & TmpRs!EmployeeNo & "' "
       Call TmpTableAzureStoreMan
       If TmpRsAzureStoreMan.RecordCount = 0 Then
          tmp = "ALTER TABLE [Mst_Employee] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          mStrQuery = GetQueyForWebservice("Mst_Employee", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
            tmp = "SET IDENTITY_INSERT [Mst_Employee] ON " & mStrQuery
            Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Mst_Employee] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Mst_Employee", "E", TmpRs, "Where EmployeeID = '" & TmpRs!EmployeeID & "'", tmp, "EmployeeID")
          If mStrQuery <> "" Then
             tmp = mStrQuery
             Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Mst_Employee] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Mst_Employee] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

'----Employee Allocation Details Mapping
tmp = "Select * from Mst_EmployeeAllocationDetails Where Concurrency >=  '" & Date - 5 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Mst_EmployeeAllocationDetails] ON"
       Call TmpTableAzureStoreMan1
       
       tmp = "Select * from Mst_EmployeeAllocationDetails WHere EmpAllDetailID = '" & TmpRs!EmpAllDetailID & "' "
       Call TmpTableAzureStoreMan
       
       If TmpRsAzureStoreMan.RecordCount = 0 Then
          
          tmp = "ALTER TABLE [Mst_EmployeeAllocationDetails] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          
          mStrQuery = GetQueyForWebservice("Mst_EmployeeAllocationDetails", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
             tmp = "SET IDENTITY_INSERT [Mst_EmployeeAllocationDetails] ON " & mStrQuery
             Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Mst_EmployeeAllocationDetails] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Mst_EmployeeAllocationDetails", "E", TmpRs, "Where EmpAllDetailID = '" & TmpRs!EmpAllDetailID & "'", tmp, "EmpAllDetailID")
          If mStrQuery <> "" Then
             tmp = mStrQuery
             Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Mst_EmployeeAllocationDetails] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Mst_EmployeeAllocationDetails] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

'----User Master Data Mapping
tmp = "Select * from Mst_Users Where CreatedDate >= '" & Date - 5 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Mst_Users] ON"
       Call TmpTableAzureStoreMan1
       
       tmp = "Select * from Mst_Users WHere UsersID = '" & TmpRs!UsersID & "' "
       Call TmpTableAzureStoreMan
       
       If TmpRsAzureStoreMan.RecordCount = 0 Then
          
          tmp = "ALTER TABLE [Mst_Users] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          
          mStrQuery = GetQueyForWebservice("Mst_Users", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
             mStrQuery = Replace(mStrQuery, ",EmployeeNo,", ",EmployeeNo,")
             mStrQuery = Replace(mStrQuery, ",EmployeeName,", ",EmployeeName,")
             mStrQuery = Replace(mStrQuery, ",UserGroupID,", ",UserRoleID,")
             tmp = "SET IDENTITY_INSERT [Mst_Users] ON " & mStrQuery
             Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Mst_Users] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Mst_Users", "E", TmpRs, "Where UsersID = '" & TmpRs!UsersID & "'", tmp, "UsersID")
          If mStrQuery <> "" Then
             mStrQuery = Replace(mStrQuery, "EmployeeNo=", "EmployeeNo=")
             mStrQuery = Replace(mStrQuery, "EmployeeName=", "EmployeeName=")
             mStrQuery = Replace(mStrQuery, "UserGroupID=", "UserRoleID=")
             tmp = mStrQuery
             Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Mst_Users] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Mst_Users] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan


tmp = "Select * from Txn_ImportNcdexLotPleadge Where CreatedDate >= '" & Date - 30 & "'"
Call TmpTable

tmp = "ALTER TABLE [Txn_ImportNcdexLotPleadge] NoCHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

tmp = "SET IDENTITY_INSERT [Txn_ImportNcdexLotPleadge] ON"
Call TmpTableAzureStoreMan

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount

       tmp = "Select * from Txn_ImportNcdexLotPleadge WHere CMSELotID = '" & TmpRs!CMSELotId & "' And TransNo = '" & TmpRs!TransNo & "'"
       Call TmpTableAzureStoreMan
       If TmpRsAzureStoreMan.RecordCount = 0 Then
            TmpRsAzureStoreMan.AddNew
'            TmpRsAzureStoreMan!SrNo = TmpRs!SrNo
            TmpRsAzureStoreMan!NCDEXCODE = TmpRs!NCDEXCODE
            TmpRsAzureStoreMan!CMSELotId = TmpRs!CMSELotId
            TmpRsAzureStoreMan!TransNo = TmpRs!TransNo
            TmpRsAzureStoreMan!CMSEClientID = TmpRs!CMSEClientID
            TmpRsAzureStoreMan!CMSEClientName = TmpRs!CMSEClientName
            TmpRsAzureStoreMan!CMSPLotId = TmpRs!CMSPLotId
            TmpRsAzureStoreMan!ActualQuantity = TmpRs!ActualQuantity
            TmpRsAzureStoreMan!UOM = TmpRs!UOM
            TmpRsAzureStoreMan!FED = TmpRs!FED
            TmpRsAzureStoreMan!Remarks = TmpRs!Remarks
            TmpRsAzureStoreMan!Type = TmpRs!Type
            TmpRsAzureStoreMan!BatchId1 = TmpRs!BatchId1
            TmpRsAzureStoreMan!BatchID2 = TmpRs!BatchID2
            TmpRsAzureStoreMan!Flag = TmpRs!Flag
            TmpRsAzureStoreMan!ErrorMsg = TmpRs!ErrorMsg
            TmpRsAzureStoreMan!RatePerUOM = TmpRs!RatePerUOM
            TmpRsAzureStoreMan!ValueCommodity = TmpRs!ValueCommodity
            TmpRsAzureStoreMan!ValidityDate = TmpRs!ValidityDate
            TmpRsAzureStoreMan!ACTUALGRADE = TmpRs!ACTUALGRADE
            TmpRsAzureStoreMan!GradePremium = TmpRs!GradePremium
            TmpRsAzureStoreMan!LotRefNo = TmpRs!LotRefNo
            TmpRsAzureStoreMan!PreviousColumnsDet = TmpRs!PreviousColumnsDet
            TmpRsAzureStoreMan!Pledge = TmpRs!Pledge
            TmpRsAzureStoreMan!Code = TmpRs!Code
            TmpRsAzureStoreMan!SerialNo = TmpRs!SerialNo
            TmpRsAzureStoreMan!SM_Prefix = TmpRs!SM_Prefix
            TmpRsAzureStoreMan!PledgeID = TmpRs!PledgeID
            TmpRsAzureStoreMan!CARNo = TmpRs!CARNo
            TmpRsAzureStoreMan!GSLID = TmpRs!GSLID
            TmpRsAzureStoreMan!Bags = TmpRs!Bags
            TmpRsAzureStoreMan!Quantity = TmpRs!Quantity
            TmpRsAzureStoreMan!G_UID = TmpRs!G_UID
            TmpRsAzureStoreMan!Concurrency = TmpRs!Concurrency
            TmpRsAzureStoreMan!CreatedBy = TmpRs!CreatedBy
            TmpRsAzureStoreMan!CreatedDate = TmpRs!CreatedDate
            TmpRsAzureStoreMan!UpdatedBy = TmpRs!UpdatedBy
            TmpRsAzureStoreMan!UpdatedDate = TmpRs!UpdatedDate
            TmpRsAzureStoreMan!ImportDate = TmpRs!ImportDate
            TmpRsAzureStoreMan!GodownID = TmpRs!GodownID
            TmpRsAzureStoreMan!BranchID = TmpRs!BranchID
            TmpRsAzureStoreMan.Update
       End If
       TmpRs.MoveNext
   Next
End If



''----CMP Master Data Mapping
tmp = "Select * from Mst_CMP Where CreatedDate >= '" & Date - 5 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Mst_CMP] ON"
       Call TmpTableAzureStoreMan1

       tmp = "Select * from Mst_CMP WHere CMPID = '" & TmpRs!CMPID & "' "
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          tmp = "ALTER TABLE [Mst_CMP] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          mStrQuery = GetQueyForWebservice("Mst_CMP", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
            tmp = "SET IDENTITY_INSERT [Mst_CMP] ON " & mStrQuery
            Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Mst_CMP] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Mst_CMP", "E", TmpRs, "Where CMPID = '" & TmpRs!CMPID & "'", tmp, "CMPID")
          If mStrQuery <> "" Then
            tmp = mStrQuery
            Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Mst_CMP] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Mst_CMP] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

'----Godown Master Data Mapping
tmp = "Select * from Mst_Godown Where CreatedDate >= '" & Date - 5 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Mst_Godown] ON"
       Call TmpTableAzureStoreMan1

       tmp = "Select * from Mst_Godown WHere GodownID = '" & TmpRs!GodownID & "' "
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then
          tmp = "ALTER TABLE [Mst_Godown] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          mStrQuery = GetQueyForWebservice("Mst_Godown", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
            tmp = "SET IDENTITY_INSERT [Mst_Godown] ON " & mStrQuery
            Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Mst_Godown] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Mst_Godown", "E", TmpRs, "Where GodownID = '" & TmpRs!GodownID & "'", tmp, "GodownID")
          If mStrQuery <> "" Then
            tmp = mStrQuery
            Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Mst_Godown] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Mst_Godown] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

'----Pledge Creation
tmp = "Select * from Txn_PledgeCreation Where CreatedDate >= '" & Date - 5 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Txn_PledgeCreation] ON"
       Call TmpTableAzureStoreMan1

       tmp = "Select * from Txn_PledgeCreation WHere PledgeID = '" & TmpRs!PledgeID & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' "
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then

          tmp = "ALTER TABLE [Txn_PledgeCreation] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan

          mStrQuery = GetQueyForWebservice("Txn_PledgeCreation", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
             mStrQuery = Replace(mStrQuery, "ClientIDRow,", "ClientID,")
            tmp = "SET IDENTITY_INSERT [Txn_PledgeCreation] OFF " & mStrQuery
            Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Txn_PledgeCreation] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Txn_PledgeCreation", "E", TmpRs, "Where PledgeID = '" & TmpRs!PledgeID & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "'", tmp, "PledgeID,SM_Prefix")
          If mStrQuery <> "" Then
            mStrQuery = Replace(mStrQuery, "ClientIDRow=", "ClientID=")
            tmp = mStrQuery
            Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Txn_PledgeCreation] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Txn_PledgeCreation] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan

'----Pledge Creation GSL
tmp = "Select * from Txn_PledgeCreation_GSL WHere Convert(Varchar,PledgeID) +'-'+SM_Prefix in (Select Convert(Varchar,PledgeID) +'-'+SM_Prefix From Txn_PledgeCreation Where CreatedDate >= '" & Date - 5 & "') " 'Where CreatedDate >= '" & Date - 30 & "' "
Call TmpTable

If TmpRs.RecordCount > 0 Then
   For I = 1 To TmpRs.RecordCount
       tmp = "SET IDENTITY_INSERT [Txn_PledgeCreation_GSL] ON"
       Call TmpTableAzureStoreMan1

       tmp = "Select * from Txn_PledgeCreation_GSL WHere PledgeID = '" & TmpRs!PledgeID & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = " & TmpRs!GSLID & " "
       Call TmpTableAzureStoreMan

       If TmpRsAzureStoreMan.RecordCount = 0 Then

          tmp = "ALTER TABLE [Txn_PledgeCreation_GSL] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan

          mStrQuery = GetQueyForWebservice("Txn_PledgeCreation_GSL", "A", TmpRs, "", tmp, "")
          If mStrQuery <> "" Then
            tmp = "SET IDENTITY_INSERT [Txn_PledgeCreation_GSL] Off " & mStrQuery
            Call TmpTableAzureStoreMan
          End If
       Else
          tmp = "ALTER TABLE [Txn_PledgeCreation_GSL] NoCHECK CONSTRAINT ALL"
          Call TmpTableAzureStoreMan
          tmp = ""
          mStrQuery = GetQueyForWebservice("Txn_PledgeCreation_GSL", "E", TmpRs, "Where PledgeID = '" & TmpRs!PledgeID & "' And SM_Prefix = '" & TmpRs!SM_Prefix & "' And GSLID = " & TmpRs!GSLID & "", tmp, "PledgeID,SM_Prefix")
          If mStrQuery <> "" Then
            tmp = mStrQuery
            Call TmpTableAzureStoreMan
          End If
       End If
       TmpRs.MoveNext
   Next
End If

tmp = "SET IDENTITY_INSERT [Txn_PledgeCreation] OFF"
Call TmpTableAzureStoreMan

tmp = "ALTER TABLE [Txn_PledgeCreation] CHECK CONSTRAINT ALL"
Call TmpTableAzureStoreMan
'----NCDEX Lot Pledge Release From Azure to VB

MsgBox "Done"

End Sub

Private Sub CmdUpdate_Click()

If LblFile.Caption = "" Then
   MsgBox "Please select file!!!"
   Call CmdCancel_Click
   Exit Sub
End If


If DTPImportDate.Value > Date Then
   MsgBox "Future import date not allowed!!!"
   DTPImportDate.SetFocus
   Exit Sub
End If


CmdGetFile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = False
CmdCancel.Enabled = False
lblStatus.Refresh
lblStatus.Refresh
ReadFiles (strFilename)
LblFile.Caption = ""
lblStatus.Caption = ""
lblLines.Caption = ""
strFilename = ""
CmdGetFile.Enabled = True
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
 
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
If LCase(Gen_UserRole) = "power" Then
   CmdEmployeeData.Visible = True
   CmdExcel.Visible = True
   CmdImportQueries.Visible = True
   CmdExecQueries.Visible = True
   Frame2.Visible = True
   LblFrameGrid.Caption = "Godown Details"
   mQueryXL = False
   CmdMissingGSL.Visible = True
End If

Frame1.Enabled = True

LblFile.Caption = ""
lblStatus.Caption = ""
CmdGetFile.Enabled = True
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = False
DTPImportDate.Value = Date

End Sub


Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 13 '11

MSHFlexGrid1.TextMatrix(0, 0) = "Employee ID"
MSHFlexGrid1.TextMatrix(0, 1) = "Employee Name"
MSHFlexGrid1.TextMatrix(0, 2) = "Employee No"
MSHFlexGrid1.TextMatrix(0, 3) = "CMPID"
MSHFlexGrid1.TextMatrix(0, 4) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 5) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 6) = "LocationID"
MSHFlexGrid1.TextMatrix(0, 7) = "State Name"
MSHFlexGrid1.TextMatrix(0, 8) = "StateID"
MSHFlexGrid1.TextMatrix(0, 9) = "Region"
MSHFlexGrid1.TextMatrix(0, 10) = "RegionID"
MSHFlexGrid1.TextMatrix(0, 11) = "Designation"
MSHFlexGrid1.TextMatrix(0, 12) = "Designation ID"


MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 1500
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1400
MSHFlexGrid1.ColWidth(5) = 1400
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1000

MSHFlexGrid1.RowHeight(0) = 600
    
End Sub
Private Sub Grid_HeadQueries()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 3 '11
MSHFlexGrid1.FixedCols = 0

MSHFlexGrid1.TextMatrix(0, 0) = "Select"
MSHFlexGrid1.TextMatrix(0, 1) = "Query"
MSHFlexGrid1.TextMatrix(0, 2) = "Result/ErrorLog"

MSHFlexGrid1.ColWidth(0) = 800
MSHFlexGrid1.ColWidth(1) = 14000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.RowHeight(0) = 600
    
End Sub

Public Sub ReadFiles(file_name As String)
Dim lngFreefile As Long
Call TableTxn_BenposDatewise
If RsTxn_BenposDateWise.EOF = False Then
    RsTxn_BenposDateWise.MoveFirst
    RsTxn_BenposDateWise.Find "ImportDate = " & DTPImportDate.Value & " "
    If RsTxn_BenposDateWise.EOF = False Then
        MsgBox "Benpos Details available for this date "
        Exit Sub
    End If
End If
MsgBox "Wait till next message"
Call TableTxn_CreateBenposTemp   'CreateBenposTemp
Call TableTxn_BenposTemp         'BenposTempTable

Set objFS = New Scripting.FileSystemObject
Set objFile = objFS.GetFile(file_name)
Set objTS = objFile.OpenAsTextStream(ForReading, TristateUseDefault)
sFileLineNo = 1
lblStatus.Refresh
lblStatus.Caption = "Reading File These will take few mins..Please Wait"
sFileLine = objTS.ReadLine
If CheckFile = False Then
   MsgBox "Please Select proper file!!!"
   Exit Sub
End If


tmp = "delete from  Txn_BenposDatewise"
Call TmpTable
RsTxn_BenposDateWise.Requery

'Dim tmp As Variant
Do While Not objTS.AtEndOfStream
    sFileLine = objTS.ReadLine
    sFileLineNo = sFileLineNo + 1
    If Left(sFileLine, 3) = "INC" Then
        ps = Split(sFileLine, "~")
        If UBound(ps) = 7 Then
            For I = 0 To 7
                sArray(I) = Trim(ps(I))
            Next I
            RsTxn_Benpos.AddNew
            RsTxn_Benpos!ISINID = sArray(0)
            RsTxn_Benpos!ISINDescription = sArray(1)
            RsTxn_Benpos!DPID = sArray(2)
            RsTxn_Benpos!ClientID = sArray(3)
            RsTxn_Benpos!ClientName = sArray(4)
            RsTxn_Benpos!Depository = sArray(5)
            RsTxn_Benpos!ElectronicPosition = sArray(6)
            RsTxn_Benpos!unit = sArray(7)
            RsTxn_Benpos!ImportDate = DTPImportDate.Value
        End If
        lblLines.Caption = sFileLineNo
    End If
Loop
RsTxn_Benpos.MoveLast
lblStatus.Refresh
lblStatus.Caption = "Benpos File Imported Sucessfully...."
CmdGetFile.Enabled = False
CmdUpdate.Enabled = False
CmdClose.Enabled = True
CmdCancel.Enabled = True
Call TableTxn_InsertBenposDateWise(DTPImportDate.Value)  'InsertBenposDateWise(DTPImportDate.Value)
MsgBox "Import Successfully"
End Sub

Private Function CheckFile() As Boolean
Dim Retval As Boolean
Retval = True
ps = Split(sFileLine, "~")
If UBound(ps) = 7 Then
   If Trim(ps(0)) <> "ISIN" Then
      Retval = False
   ElseIf Trim(ps(1)) <> "ISINDescription" Then
      Retval = False
   ElseIf Trim(ps(2)) <> "DPID" Then
      Retval = False
   ElseIf Trim(ps(3)) <> "ClientID" Then
      Retval = False
   ElseIf Trim(ps(4)) <> "ClientName" Then
      Retval = False
   ElseIf Trim(ps(5)) <> "Depository" Then
      Retval = False
   ElseIf Trim(ps(6)) <> "ElectronicPosition" Then
      Retval = False
   ElseIf Trim(ps(7)) <> "Unit" Then
      Retval = False
   End If
'ISIN~ISINDescription~DPID~ClientID~ClientName~Depository~ElectronicPosition~Unit
Else
   Retval = False
End If
'lblLines.Caption = sFileLineNo
CheckFile = Retval
End Function

Private Sub MSHFlexGrid1_Click()
With MSHFlexGrid1
    If .Col = 0 Then  '.Row = 1 And
      If .TextMatrix(.row, .Col) = "" Then Exit Sub
      .Col = 0
         If .TextMatrix(.row, .Col) = strUnChecked Then
            .TextMatrix(.row, .Col) = strChecked
         Else
         .TextMatrix(.row, .Col) = strUnChecked
      End If
   End If
End With
End Sub
