VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_StateWiseGSTNo 
   Caption         =   "State Wise GST No Report"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   9480
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbExchangeType 
         Height          =   315
         ItemData        =   "FrmRpt_StateWiseGSTNo.frx":0000
         Left            =   7560
         List            =   "FrmRpt_StateWiseGSTNo.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   263
         Width           =   1635
      End
      Begin VB.OptionButton OptPending 
         Caption         =   "Pending Data"
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
         Left            =   4800
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton OptUploaded 
         Caption         =   "Uploaded Data"
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
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
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
         Left            =   13845
         TabIndex        =   6
         Top             =   225
         Width           =   1095
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
         Left            =   11535
         TabIndex        =   4
         Top             =   225
         Width           =   1215
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
         Left            =   12750
         TabIndex        =   5
         Top             =   225
         Width           =   1095
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
         ItemData        =   "FrmRpt_StateWiseGSTNo.frx":0004
         Left            =   795
         List            =   "FrmRpt_StateWiseGSTNo.frx":0017
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2160
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Exchange"
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
         Left            =   6405
         TabIndex        =   13
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Total Record :"
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
         Left            =   9240
         TabIndex        =   11
         Top             =   285
         Width           =   1335
      End
      Begin VB.Label LblTotalRecord 
         AutoSize        =   -1  'True
         Caption         =   "-"
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
         Left            =   10560
         TabIndex        =   10
         Top             =   285
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Type"
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
         Left            =   75
         TabIndex        =   9
         Top             =   285
         Width           =   600
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8295
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   14631
      _Version        =   393216
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmRpt_StateWiseGSTNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mR, mC As Double
Dim mExchangeTypeID As Variant

Private Sub CmbExchangeType_GotFocus()
tmp = CmbExchangeType.Text
mTmp = CmbExchangeType.Text
Call TableMst_ExchangeType(" Where Flag = 'A' And ExchangeTypeID In (1,2) ")
CmbExchangeType.Clear
If RsMst_ExchangeType.RecordCount = 0 Then
   MsgBox "No Exchange Type found"
   CmbExchangeType.SetFocus
   Exit Sub
End If
For I = 1 To RsMst_ExchangeType.RecordCount
    CmbExchangeType.AddItem RsMst_ExchangeType!ExchangeType
    RsMst_ExchangeType.MoveNext
Next
CmbExchangeType.Text = tmp
End Sub
Private Sub CmbExchangeType_LostFocus()
If CmbExchangeType.Text = "" Then
   CmbClient.Text = ""
   Exit Sub
End If
RsMst_ExchangeType.MoveFirst
RsMst_ExchangeType.Find "ExchangeType = '" & CmbExchangeType.Text & "'"
If RsMst_ExchangeType.EOF Then
   MsgBox "Invalid Exchange Type selection "
   CmbExchangeType.SetFocus
   Exit Sub
End If
mExchangeTypeID = RsMst_ExchangeType.Fields(0)
End Sub

Private Sub CmbType_LostFocus()
If CmbType.Text = "" Then Exit Sub
If LCase(CmbType.Text) <> LCase("Client") And LCase(CmbType.Text) <> LCase("Depository") And LCase(CmbType.Text) <> LCase("Lessor") And LCase(CmbType.Text) <> LCase("Security") And LCase(CmbType.Text) <> LCase("Vendor") Then
   MsgBox "Invalid selection!!!"
   CmbType.SetFocus
   Exit Sub
End If
CmbExchangeType.Text = ""
LblTotalRecord.Caption = ""

If LCase(CmbType.Text) = LCase("Client") Then
    Label2.Visible = True
    CmbExchangeType.Visible = True
    Call GridHeadClient
ElseIf LCase(CmbType.Text) = LCase("Depository") Then
    Label2.Visible = False
    CmbExchangeType.Visible = False
    Call GridHeadDepository
ElseIf LCase(CmbType.Text) = LCase("Lessor") Then
    Label2.Visible = False
    CmbExchangeType.Visible = False
    Call GridHeadLessor
ElseIf LCase(CmbType.Text) = LCase("Security") Then
    Label2.Visible = False
    CmbExchangeType.Visible = False
    Call GridHeadSecurity
ElseIf LCase(CmbType.Text) = LCase("Vendor") Then
    Label2.Visible = False
    CmbExchangeType.Visible = False
    Call GridHeadVendor
End If

End Sub

Private Sub CmdClear_Click()
Call ClearData
End Sub

Private Sub CmdGetReport_Click()
Dim I As Double

If CmbType.Text = "" Then
   MsgBox "Select Type."
   CmbType.SetFocus
   Exit Sub
End If
If LCase(CmbType.Text) = LCase("Client") Then
    If CmbExchangeType.Text = "" Then
        MsgBox "Select Exchange."
        CmbExchangeType.SetFocus
        Exit Sub
    End If
End If

CmdExcel.Enabled = False

If OptUploaded.Value = True Then
    If LCase(CmbType.Text) = LCase("Client") Then
        Call GridHeadClient
        ''-- Client GST Data
        tmp = " SELECT MC.ClientName,MC.ClientID,MET.ExchangeType,MC.NAVCode,MC.CRFID,MSI.StateIndiaName,MCSWG.ClientStateOfficeAddress,MCSWG.GSTNo,MCSWG.GSTNavCode,MCSWG.RGSTNavCode,MCSWG.ClientIDRow,MCSWG.StateIndiaID,MCSWG.ClientStateWiseGSTID"
        tmp = tmp & " FROM Mst_ClientStateWiseGST MCSWG"
        tmp = tmp & " Inner Join Mst_Client MC On MCSWG.ClientIDRow = MC.ClientIDRow"
        tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MCSWG.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where MET.ExchangeTypeID = " & mExchangeTypeID & ""
        tmp = tmp & " Order By MCSWG.ClientStateWiseGSTID"
    ElseIf LCase(CmbType.Text) = LCase("Depository") Then
        Call GridHeadDepository
        ''-- Depository GST Data
        tmp = " Select MDSWG.DepositoryID,MD.DepositoryName,MET.ExchangeType,MD.NavCode,MSI.StateIndiaName,MDSWG.DepositoryStateOfficeAddress,MDSWG.GSTNo, MDSWG.GSTNavCode,MDSWG.StateIndiaID,MDSWG.DepositoryStateWiseGSTID"
        tmp = tmp & " From Mst_DepositoryStateWiseGST MDSWG"
        tmp = tmp & " Inner Join Mst_Depository MD On MDSWG.DepositoryID = MD.DepositoryID"
        tmp = tmp & " Inner Join Mst_ExchangeType MET On MDSWG.ExchangeTypeID = MET.ExchangeTypeID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MDSWG.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Order By MDSWG.DepositoryStateWiseGSTID"
    ElseIf LCase(CmbType.Text) = LCase("Lessor") Then
        Call GridHeadLessor
        ''-- Lessor GST Data
        tmp = " Select ML.LessorName,ML.NAVCode,MSI.StateIndiaName,MLSWG.LessorStateOfficeAddress,MLSWG.GSTNo,MLSWG.GSTNavCode,MLSWG.LessorID,MLSWG.StateIndiaID,MLSWG.LessorStateWiseGSTID"
        tmp = tmp & " From Mst_LessorStateWiseGST MLSWG"
        tmp = tmp & " Inner Join Mst_Lessor ML On MLSWG.LessorID = ML.LessorID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MLSWG.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Order By MLSWG.LessorStateWiseGSTID"
    ElseIf LCase(CmbType.Text) = LCase("Security") Then
        Call GridHeadSecurity
        ''-- Security GST Data
        tmp = " Select MSA.AgencyName,MSA.NAVCode,MSI.StateIndiaName,MSASWG.SecurityAgencyStateOfficeAddress,MSASWG.GSTNo,MSASWG.GSTNavCode,MSASWG.SecurityAgencyID,MSASWG.StateIndiaID,MSASWG.SecurityAgencyStateWiseGSTID"
        tmp = tmp & " FROM Mst_SecurityAgencyStateWiseGST MSASWG"
        tmp = tmp & " Inner Join Mst_SecurityAgency MSA On MSASWG.SecurityAgencyID = MSA.SecurityAgencyID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MSASWG.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Order By MSASWG.SecurityAgencyStateWiseGSTID"
    ElseIf LCase(CmbType.Text) = LCase("Vendor") Then
        Call GridHeadVendor
        ''-- Vendor GST Data
        tmp = " Select MV.VendorName,MV.NAVCode,MSI.StateIndiaName,MVSWG.VendorStateOfficeAddress,MVSWG.GSTNo,MVSWG.GSTNavCode,MVSWG.VendorID,MVSWG.StateIndiaID,MVSWG.VendorStateWiseGSTID"
        tmp = tmp & " FROM Mst_VendorStateWiseGST MVSWG"
        tmp = tmp & " Inner Join Mst_Vendor MV On MVSWG.VendorID = MV.VendorID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MVSWG.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Order By MVSWG.VendorStateWiseGSTID"
    End If
ElseIf OptPending.Value = True Then
    If LCase(CmbType.Text) = LCase("Client") Then
        Call GridHeadClient
        ''-- Client GST Data
        tmp = " SELECT Distinct MC.ClientName,MC.ClientID,MET.ExchangeType,MC.NAVCode,MC.CRFID,MSI.StateIndiaName,''[StateWiseClientAddress],''[GSTNo],''[GSTNAVCode],''[RGSTNAVCode],MC.ClientIDRow,MSI.StateIndiaID"
        tmp = tmp & " FROM Mst_GSL GSL"
        tmp = tmp & " Inner Join Mst_Location ML ON GSL.SM_Prefix = ML.SM_Prefix"
        tmp = tmp & " Inner Join Mst_Client MC On GSL.ClientIDRow = MC.ClientIDRow"
        tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where MC.Status = 'A' And GSL.BalanceBags > 0"
        tmp = tmp & " And MET.ExchangeTypeID = " & mExchangeTypeID & ""
        tmp = tmp & " And CONVERT(Varchar(10),MC.ClientIDRow)+''+CONVERT(Varchar(10),MSI.StateIndiaID) NOT IN (SELECT CONVERT(Varchar(10),ClientIDRow)+''+CONVERT(Varchar(10),StateIndiaID) FROM Mst_ClientStateWiseGST )"
        tmp = tmp & " Order By MC.ClientName,MSI.StateIndiaName"
    ElseIf LCase(CmbType.Text) = LCase("Depository") Then
        Call GridHeadDepository
        ''-- Depository GST Data
        tmp = " SELECT Distinct MD.DepositoryID,MD.DepositoryName,MET.ExchangeType,MD.NavCode,MSI.StateIndiaName,''[StateWiseDepositoryAddress],''[GSTNo],''[GSTNAVCode],MSI.StateIndiaID"
        tmp = tmp & " FROM Mst_GSL GSL"
        tmp = tmp & " Inner Join Mst_Client MC On GSL.ClientIDRow = MC.ClientIDRow"
        tmp = tmp & " Inner Join Mst_Location ML ON GSL.SM_Prefix = ML.SM_Prefix"
        tmp = tmp & " Inner Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Inner Join Mst_Depository MD On MC.DepositoryID = MD.DepositoryID"
        tmp = tmp & " Where (MD.Status = 'A' Or MD.Status Is Null) And GSL.BalanceBags > 0 " ''- --And GSL.ExchangeTypeID = 2
        tmp = tmp & " And CONVERT(Varchar(10),MD.DepositoryID)+''+CONVERT(Varchar(10),MSI.StateIndiaID) NOT IN (SELECT CONVERT(Varchar(10),DepositoryID)+''+CONVERT(Varchar(10),StateIndiaID) FROM Mst_DepositoryStateWiseGST)"
        tmp = tmp & " Order By MD.DepositoryName,MSI.StateIndiaName"
    ElseIf LCase(CmbType.Text) = LCase("Lessor") Then
        Call GridHeadLessor
        ''-- Lessor GST Data
        tmp = " SELECT Distinct ML.LessorName,ML.NAVCode,MSI.StateIndiaName,''[StateWiseLessorAddress],''[GSTNo],''[GSTNAVCode],TGLAD.LessorID,MS.StateIndiaID"
        tmp = tmp & " FROM Txn_GodownLessorAgreeDetail TGLAD"
        tmp = tmp & " Inner Join Mst_Lessor ML On TGLAD.LessorID = ML.LessorID"
        tmp = tmp & " Inner Join Mst_Godown MG On TGLAD.GodownID = MG.GodownID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location Location On CMP.LocationID = Location.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On Location.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where (ML.Status = 'A' or ML.Status is Null ) And MG.CloseDate is Null"
        tmp = tmp & " And CONVERT(Varchar(10),ML.LessorID)+''+CONVERT(Varchar(10),MSI.StateIndiaID) NOT IN (SELECT CONVERT(Varchar(10),LessorID)+''+CONVERT(Varchar(10),StateIndiaID) FROM Mst_LessorStateWiseGST )"
        tmp = tmp & " Order By ML.LessorName,MSI.StateIndiaName"
    ElseIf LCase(CmbType.Text) = LCase("Security") Then
        Call GridHeadSecurity
        ''-- Security GST Data
        tmp = " SELECT Distinct MSA.AgencyName,MSA.NAVCode,MSI.StateIndiaName,''[StateWiseSecurityAgencyAddress],''[GSTNo],''[GSTNAVCode],TSA.SecurityAgencyID,MSI.StateIndiaID"
        tmp = tmp & " FROM Txn_SecurityAgreement TSA"
        tmp = tmp & " Inner Join Mst_SecurityAgency MSA On TSA.SecurityAgencyID = MSA.SecurityAgencyID"
        tmp = tmp & " Inner Join Mst_State MS On TSA.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where MSA.Flag = 'Y' And CONVERT(VarChar(10),TSA.AgreeExpiryDate,112) > Convert(Varchar(10),getdate(),112)"
        tmp = tmp & " And CONVERT(Varchar(10),TSA.SecurityAgencyID)+''+CONVERT(Varchar(10),MSI.StateIndiaID) NOT IN (SELECT CONVERT(Varchar(10),SecurityAgencyID)+''+CONVERT(Varchar(10),StateIndiaID) FROM Mst_SecurityAgencyStateWiseGST )"
        tmp = tmp & " Order By MSA.AgencyName,MSI.StateIndiaName"
    ElseIf LCase(CmbType.Text) = LCase("Vendor") Then
        Call GridHeadVendor
        ''-- Vendor GST Data
        tmp = " Select A.VendorName,A.NavCode,A.StateIndiaName,A.Address,A.GSTNo[GSTNo],A.GSTNAVCode[GSTNAVCode],A.VendorID,A.StateIndiaID,A.Type From"
        tmp = tmp & " ("
        tmp = tmp & " SELECT Distinct MV.VendorName,Isnull(MV.NavCode,'')[NavCode],MSI.StateIndiaName,''[Address],''[GSTNo],''[GSTNAVCode],TF.VendorID,MSI.StateIndiaID,'Fumigation'[Type]"
        tmp = tmp & " FROM Txn_Fumigation TF"
        tmp = tmp & " Inner Join Mst_Vendor MV On TF.VendorID = MV.VendorID"
        tmp = tmp & " Inner Join Mst_Godown MG On TF.GodownID = MG.GodownID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location Location On CMP.LocationID = Location.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On Location.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where MSI.Flag = 'A' And MG.CloseDate is Null"
        tmp = tmp & " Union"
        tmp = tmp & " SELECT Distinct MV.VendorName,Isnull(MV.NavCode,'')[NavCode],MSI.StateIndiaName,''[Address],''[GSTNo],''[GSTNAVCode],TF.VendorID,MSI.StateIndiaID,'Franchisee'[Type]"
        tmp = tmp & " FROM Txn_FranchiseeRent TF"
        tmp = tmp & " Inner Join Mst_Vendor MV On TF.VendorID = MV.VendorID"
        tmp = tmp & " Inner Join Mst_Location ML On TF.LocationID = ML.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where MSI.Flag = 'A'"
        tmp = tmp & " Union"
        tmp = tmp & " SELECT Distinct MV.VendorName,Isnull(MV.NavCode,'')[NavCode],MSI.StateIndiaName,''[Address],''[GSTNo],''[GSTNAVCode],TS.VendorID,MSI.StateIndiaID,'Spraying'[Type]"
        tmp = tmp & " FROM Txn_Spraying TS"
        tmp = tmp & " Inner Join Mst_Vendor MV On TS.VendorID = MV.VendorID"
        tmp = tmp & " Inner Join Mst_Godown MG On TS.GodownID = MG.GodownID"
        tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
        tmp = tmp & " Inner Join Mst_Location Location On CMP.LocationID = Location.LocationID"
        tmp = tmp & " Inner Join Mst_State MS On Location.StateID = MS.StateID"
        tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
        tmp = tmp & " Where MSI.Flag = 'A'"
        tmp = tmp & " ) A"
        tmp = tmp & " Where CONVERT(Varchar(10),A.VendorID)+''+CONVERT(Varchar(10),A.StateIndiaID) NOT IN (SELECT CONVERT(Varchar(10),VendorID)+''+CONVERT(Varchar(10),StateIndiaID) FROM Mst_VendorStateWiseGST )"
        tmp = tmp & " Order By A.VendorName,A.StateIndiaName"
    End If
End If
Call TmpTable

    With MSHFlexGrid1
    If TmpRs.RecordCount > 0 Then
        ProgressBar1.Max = TmpRs.RecordCount
        For I = 1 To TmpRs.RecordCount
        
         If LCase(CmbType.Text) = LCase("Client") Then
            
            If OptUploaded.Value = True Then
                .TextMatrix(I, 0) = TmpRs!ClientName
                .TextMatrix(I, 1) = "'" & TmpRs!ClientID
                .TextMatrix(I, 2) = TmpRs!ExchangeType
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!CRFID), "", TmpRs!CRFID)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!ClientStateOfficeAddress), "", TmpRs!ClientStateOfficeAddress)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 9) = IIf(IsNull(TmpRs!RGSTNavCode), "", TmpRs!RGSTNavCode)
                .TextMatrix(I, 10) = IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
                .TextMatrix(I, 11) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 12) = IIf(IsNull(TmpRs!ClientStateWiseGSTID), "", TmpRs!ClientStateWiseGSTID)
            ElseIf OptPending.Value = True Then
                .TextMatrix(I, 0) = TmpRs!ClientName
                .TextMatrix(I, 1) = "'" & TmpRs!ClientID
                .TextMatrix(I, 2) = TmpRs!ExchangeType
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!CRFID), "", TmpRs!CRFID)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!StateWiseClientAddress), "", TmpRs!StateWiseClientAddress)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 9) = IIf(IsNull(TmpRs!RGSTNavCode), "", TmpRs!RGSTNavCode)
                .TextMatrix(I, 10) = IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow)
                .TextMatrix(I, 11) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 12) = ""
            End If
             
         ElseIf LCase(CmbType.Text) = LCase("Depository") Then
            If OptUploaded.Value = True Then
                .TextMatrix(I, 0) = TmpRs!DepositoryID
                .TextMatrix(I, 1) = TmpRs!DepositoryName
                .TextMatrix(I, 2) = TmpRs!ExchangeType
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!DepositoryStateOfficeAddress), "", TmpRs!DepositoryStateOfficeAddress)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 9) = IIf(IsNull(TmpRs!DepositoryStateWiseGSTID), "", TmpRs!DepositoryStateWiseGSTID)
            ElseIf OptPending.Value = True Then
                .TextMatrix(I, 0) = TmpRs!DepositoryID
                .TextMatrix(I, 1) = TmpRs!DepositoryName
                .TextMatrix(I, 2) = TmpRs!ExchangeType
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!StateWiseDepositoryAddress), "", TmpRs!StateWiseDepositoryAddress)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 9) = ""
            End If
         ElseIf LCase(CmbType.Text) = LCase("Lessor") Then
            If OptUploaded.Value = True Then
                .TextMatrix(I, 0) = TmpRs!LessorName
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!LessorStateOfficeAddress), "", TmpRs!LessorStateOfficeAddress)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!LessorID), "", TmpRs!LessorID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!LessorStateWiseGSTID), "", TmpRs!LessorStateWiseGSTID)
            ElseIf OptPending.Value = True Then
                .TextMatrix(I, 0) = TmpRs!LessorName
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!StateWiseLessorAddress), "", TmpRs!StateWiseLessorAddress)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!LessorID), "", TmpRs!LessorID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 8) = ""
            End If
         ElseIf LCase(CmbType.Text) = LCase("Security") Then
            If OptUploaded.Value = True Then
                .TextMatrix(I, 0) = TmpRs!AgencyName
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!SecurityAgencyStateOfficeAddress), "", TmpRs!SecurityAgencyStateOfficeAddress)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!SecurityAgencyID), "", TmpRs!SecurityAgencyID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!SecurityAgencyStateWiseGSTID), "", TmpRs!SecurityAgencyStateWiseGSTID)
            ElseIf OptPending.Value = True Then
                .TextMatrix(I, 0) = TmpRs!AgencyName
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!StateWiseSecurityAgencyAddress), "", TmpRs!StateWiseSecurityAgencyAddress)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!SecurityAgencyID), "", TmpRs!SecurityAgencyID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 8) = ""
            End If
         ElseIf LCase(CmbType.Text) = LCase("Vendor") Then
            If OptUploaded.Value = True Then
                .TextMatrix(I, 0) = TmpRs!VendorName
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!VendorStateOfficeAddress), "", TmpRs!VendorStateOfficeAddress)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!VendorID), "", TmpRs!VendorID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 8) = IIf(IsNull(TmpRs!VendorStateWiseGSTID), "", TmpRs!VendorStateWiseGSTID)
            ElseIf OptPending.Value = True Then
                .TextMatrix(I, 0) = TmpRs!VendorName
                .TextMatrix(I, 1) = IIf(IsNull(TmpRs!NavCode), "", TmpRs!NavCode)
                .TextMatrix(I, 2) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
                .TextMatrix(I, 3) = IIf(IsNull(TmpRs!Address), "", TmpRs!Address)
                .TextMatrix(I, 4) = IIf(IsNull(TmpRs!GSTNo), "", TmpRs!GSTNo)
                .TextMatrix(I, 5) = IIf(IsNull(TmpRs!GSTNavCode), "", TmpRs!GSTNavCode)
                .TextMatrix(I, 6) = IIf(IsNull(TmpRs!VendorID), "", TmpRs!VendorID)
                .TextMatrix(I, 7) = IIf(IsNull(TmpRs!StateIndiaID), "", TmpRs!StateIndiaID)
                .TextMatrix(I, 8) = ""
            End If
         End If
        
        .Rows = .Rows + 1
        TmpRs.MoveNext
        ProgressBar1.Value = I
        
        Next
       
    End If
    End With

LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox "Done!!!"
CmdExcel.Enabled = True

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
Call CmdClear_Click
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_" & Trim(CmbType.Text) & "_StateWiseGSTNoReport.xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till Next Message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_" & Trim(CmbType.Text) & "_StateWiseGSTNoReport.xls")
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

Private Sub ClearData()
Call GridHeadClient
LblTotalRecord.Caption = ""
CmbType.Text = ""
OptUploaded.Value = True
Label2.Visible = False
CmbExchangeType.Visible = False
End Sub

Private Sub GridHeadClient()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 13
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Client Name":                  .ColWidth(0) = 2500
    .TextMatrix(0, 1) = "Client ID":                    .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "Exchange Type":                .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "NAV Code":                     .ColWidth(3) = 2000
    .TextMatrix(0, 4) = "CRF ID":                       .ColWidth(4) = 1500
    .TextMatrix(0, 5) = "State India Name":             .ColWidth(5) = 2500
    .TextMatrix(0, 6) = "State Wise Office Address":    .ColWidth(6) = 2500
    .TextMatrix(0, 7) = "GST No":                       .ColWidth(7) = 2000
    .TextMatrix(0, 8) = "GST Nav Code":                 .ColWidth(8) = 2000
    .TextMatrix(0, 9) = "GST RNav Code":                .ColWidth(9) = 2000
    .TextMatrix(0, 10) = "ClientIDRow":                 .ColWidth(10) = 0
    .TextMatrix(0, 11) = "StateIndiaID":                .ColWidth(11) = 0
    .TextMatrix(0, 12) = "ID":                          .ColWidth(12) = 0
    
    .RowHeight(0) = 600
    .Font.Size = 10
End With
End Sub
Private Sub GridHeadDepository()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 10
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "Depository ID":                  .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "Depository Name":                .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "Exchange Type":                  .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "Nav Code":                       .ColWidth(3) = 2000
    .TextMatrix(0, 4) = "State India Name":               .ColWidth(4) = 1500
    .TextMatrix(0, 5) = "State Wise Office Address":      .ColWidth(5) = 2500
    .TextMatrix(0, 6) = "GST No":                         .ColWidth(6) = 2000
    .TextMatrix(0, 7) = "GST Nav Code":                   .ColWidth(7) = 2000
    .TextMatrix(0, 8) = "State India ID":                 .ColWidth(8) = 0
    .TextMatrix(0, 9) = "ID":                             .ColWidth(9) = 0
    
    .RowHeight(0) = 600
    .Font.Size = 10
End With
End Sub
Private Sub GridHeadLessor()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 9
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "Lessor Name":                  .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "NAV Code":                     .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "State India Name":             .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "State Wise Office Address":    .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "GST No":                       .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "GST Nav Code":                 .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "Lessor ID":                    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "State India ID":               .ColWidth(7) = 0
    .TextMatrix(0, 8) = "ID":                           .ColWidth(8) = 0
    
    .RowHeight(0) = 600
    .Font.Size = 10
End With
End Sub
Private Sub GridHeadSecurity()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 9
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "Agency Name":                  .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "NAV Code":                     .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "State India Name":             .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "State Wise Office Address":    .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "GST No":                       .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "GST Nav Code":                 .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "Agency ID":                    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "State India ID":               .ColWidth(7) = 0
    .TextMatrix(0, 8) = "ID":                           .ColWidth(8) = 0
    
    .RowHeight(0) = 600
    .Font.Size = 10
End With
End Sub
Private Sub GridHeadVendor()
With MSHFlexGrid1

    .Clear
    .Rows = 2
    .FixedRows = 1
    .Cols = 9
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(0, 0) = "Vendor Name":                  .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "NAV Code":                     .ColWidth(1) = 2000
    .TextMatrix(0, 2) = "State India Name":             .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "State Wise Office Address":    .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "GST No":                       .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "GST Nav Code":                 .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "Vendor ID":                    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "State India ID":               .ColWidth(7) = 0
    .TextMatrix(0, 8) = "ID":                           .ColWidth(8) = 0
    
    .RowHeight(0) = 600
    .Font.Size = 10
End With
End Sub


