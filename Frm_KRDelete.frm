VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm_KRDelete 
   Caption         =   "Deletion of KR"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   120
      TabIndex        =   14
      Top             =   8520
      Width           =   15000
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
         Height          =   405
         Left            =   9720
         TabIndex        =   3
         Top             =   600
         Width           =   1725
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   13170
         TabIndex        =   5
         Top             =   600
         Width           =   1725
      End
      Begin VB.CommandButton CmdDeleteKR 
         Caption         =   "Delete KR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11445
         TabIndex        =   4
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   7560
         TabIndex        =   24
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label8 
         Caption         =   "GSL having withdrawal quantity"
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
         Left            =   7935
         TabIndex        =   23
         Top             =   255
         Width           =   2955
      End
      Begin VB.Label Label5 
         Caption         =   "Warning GSL Lot is Rejected Lot"
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
         Left            =   4455
         TabIndex        =   20
         Top             =   255
         Width           =   2955
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4080
         TabIndex        =   19
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label3 
         Caption         =   "Warning ISIN already Mapped to GSL"
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
         Left            =   495
         TabIndex        =   18
         Top             =   255
         Width           =   3435
      End
      Begin VB.Label Label1 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8520
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.Frame Frame3 
         Caption         =   "GSL Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6720
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6375
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   11245
            _Version        =   393216
            WordWrap        =   -1  'True
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
      End
      Begin VB.TextBox TxtRemark 
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
         Height          =   615
         Left            =   1200
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1035
         Width           =   13680
      End
      Begin VB.TextBox TxtKRNo 
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
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   1
         Top             =   240
         Width           =   3720
      End
      Begin VB.TextBox TxtCMPID 
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
         Height          =   360
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   645
         Width           =   3885
      End
      Begin VB.TextBox TxtLocationID 
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
         Height          =   360
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   645
         Width           =   3720
      End
      Begin VB.TextBox TxtCommodity 
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
         Height          =   360
         Left            =   10905
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   645
         Width           =   3720
      End
      Begin VB.Label Label7 
         Caption         =   "Total GSL :-"
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
         Left            =   8040
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   9600
         TabIndex        =   21
         Top             =   240
         Width           =   90
      End
      Begin VB.Label Label2 
         Caption         =   "Remarks"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1222
         Width           =   1800
      End
      Begin VB.Label LblRRNNo 
         Caption         =   "KR No"
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
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblCommodity 
         Caption         =   "Commodity"
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
         Left            =   9840
         TabIndex        =   11
         Top             =   675
         Width           =   1140
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP"
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
         Left            =   5160
         TabIndex        =   10
         Top             =   675
         Width           =   1080
      End
      Begin VB.Label LblLocationID 
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
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   675
         Width           =   1200
      End
   End
End
Attribute VB_Name = "Frm_KRDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Double
Private Sub CmdDeleteKR_Click()

If TxtKRNo = "" Then
   MsgBox "Blank KR No not allowed!!!"
   TxtKRNo.SetFocus
   Exit Sub
End If

If Trim(TxtRemark) = "" Then
   MsgBox "Blank Remarks not allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If MSHFlexGrid2.Rows < 3 Then
   MsgBox "No GSL details in Grid!!!"
   TxtKRNo.SetFocus
   Exit Sub
End If

If ChkGSL = False Then
   MsgBox "Some/All GSL are having withdrawal quantity!!!"
   CmdDeleteKR.Enabled = False
   Exit Sub
End If

Dim ans As Variant

ans = MsgBox("Do you really want to DELETE this KR??? ", vbYesNo)

If ans = vbYes Then
   Call TableMst_GSL("Where ExchangeTypeID = " & mExchangeTypeID & " And KRNo = " & Val(TxtKRNo) & " And (Flag = 'O' Or Flag = 'J') ")
   For I = 1 To MSHFlexGrid2.Rows - 2
       RsLog_KRDelete.AddNew
       RsLog_KRDelete!GSLID = Val(MSHFlexGrid2.TextMatrix(I, 0))
       RsLog_KRDelete!SM_Prefix = MSHFlexGrid2.TextMatrix(I, 1)
       RsLog_KRDelete!KRNo = Val(TxtKRNo)
       RsLog_KRDelete!ISINID = MSHFlexGrid2.TextMatrix(I, 16)
       RsLog_KRDelete!Flag = MSHFlexGrid2.TextMatrix(I, 15)
       RsLog_KRDelete!Remarks = TxtRemark
       RsLog_KRDelete!DeletedBy = Gen_UserLoginID
       RsLog_KRDelete!DeletedDate = Now 'MSHFlexGrid2.TextMatrix(i, 1)
       RsLog_KRDelete!G_UID = GetGUID
       
       RsLog_KRDelete.Update
              
       RsMst_GSL.MoveFirst
       RsMst_GSL.Filter = "GSLID = " & MSHFlexGrid2.TextMatrix(I, 0) & " And SM_Prefix = '" & MSHFlexGrid2.TextMatrix(I, 1) & "'"
       
       If RsMst_GSL.RecordCount > 0 Then
          RsMst_GSL!KRNo = Null
          RsMst_GSL!ISINID = Null
          RsMst_GSL!KRDate = Null
          If MSHFlexGrid2.TextMatrix(I, 15) = "J" Then
             RsMst_GSL!Flag = "O"
          End If
          RsMst_GSL!UpdatedDate = Now
          RsMst_GSL!UpdatedBy = Gen_UserLoginID
          RsMst_GSL.Update
       End If
       RsMst_GSL.Filter = ""
   Next
Else
   TxtKRNo.SetFocus
End If

MsgBox "All KR Deleted from GSL!!!"
Call ClearContents
Call Grid_Head
TxtKRNo.SetFocus

End Sub
Private Function ChkGSL() As Boolean
Dim Retval As Boolean

Retval = True

For I = 1 To MSHFlexGrid2.Rows - 2
    If Val(MSHFlexGrid2.TextMatrix(I, 8)) <> 0 And Val(MSHFlexGrid2.TextMatrix(I, 11)) <> 0 Then
       Retval = False
       MSHFlexGrid2.row = I
       MSHFlexGrid2.Col = 0
       MSHFlexGrid2.CellBackColor = vbRed
    End If
Next

ChkGSL = Retval
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Grid_Head

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   Frame1.Enabled = False
   Frame2.Enabled = False
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
mExchangeTypeID = GetExchangeID("NCDEX")
Call ClearContents

Call TableLog_KRDelete

End Sub

Private Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 36

MSHFlexGrid2.TextMatrix(0, 0) = "GSLID"
MSHFlexGrid2.TextMatrix(0, 1) = "Location ID"
MSHFlexGrid2.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid2.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid2.TextMatrix(0, 4) = "Godown No"
MSHFlexGrid2.TextMatrix(0, 5) = "Stack No"
MSHFlexGrid2.TextMatrix(0, 6) = "Lot No"
MSHFlexGrid2.TextMatrix(0, 7) = "Deposited Bags"
MSHFlexGrid2.TextMatrix(0, 8) = "Withdrawn Bags"
MSHFlexGrid2.TextMatrix(0, 9) = "Balance Bags"
MSHFlexGrid2.TextMatrix(0, 10) = "Deposited Weight"
MSHFlexGrid2.TextMatrix(0, 11) = "Withdrawn Weight"
MSHFlexGrid2.TextMatrix(0, 12) = "Balance Weight"
MSHFlexGrid2.TextMatrix(0, 13) = "Exchange Type"
MSHFlexGrid2.TextMatrix(0, 14) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 15) = "Flag"
MSHFlexGrid2.TextMatrix(0, 16) = "ISIN ID / CDC No"
MSHFlexGrid2.TextMatrix(0, 17) = "CDTF No"
MSHFlexGrid2.TextMatrix(0, 18) = "CDTF Date"
MSHFlexGrid2.TextMatrix(0, 19) = "Pledge No"
MSHFlexGrid2.TextMatrix(0, 20) = "Pledge Bags"
MSHFlexGrid2.TextMatrix(0, 21) = "Pledge Weight"
MSHFlexGrid2.TextMatrix(0, 22) = "Spillage Flag"
MSHFlexGrid2.TextMatrix(0, 23) = "Deposit Adj Wt"
MSHFlexGrid2.TextMatrix(0, 24) = "Withdrawal Adj Wt"
MSHFlexGrid2.TextMatrix(0, 25) = "G_uid"

MSHFlexGrid2.TextMatrix(0, 26) = "KR No"
MSHFlexGrid2.TextMatrix(0, 27) = "Gross Qty"
MSHFlexGrid2.TextMatrix(0, 28) = "From GUID"
MSHFlexGrid2.TextMatrix(0, 29) = "To GUID"
MSHFlexGrid2.TextMatrix(0, 30) = "WHR Bags"
MSHFlexGrid2.TextMatrix(0, 31) = "WHR Qty"

MSHFlexGrid2.TextMatrix(0, 32) = "Contract "
MSHFlexGrid2.TextMatrix(0, 33) = "Bag Size"
MSHFlexGrid2.TextMatrix(0, 34) = "Bag Equivalent"
MSHFlexGrid2.TextMatrix(0, 35) = "Weighted Qty"

MSHFlexGrid2.RowHeight(0) = 500
MSHFlexGrid2.ColWidth(25) = 0
End Sub
Private Sub TxtKRNo_LostFocus()

If Trim(TxtKRNo) = "" Then
   Call Grid_Head
   Call ClearContents
   Exit Sub
End If

If IsNumeric(TxtKRNo) = False Then
   MsgBox "Invalid KR No format!!!"
   TxtKRNo.SetFocus
   Exit Sub
End If


Call Grid_Head

tmp = "SELECT Mst_GSL.GSLID, Mst_GSL.sm_prefix , Mst_CMP.CMPName, Mst_Commodity.Commodity, Mst_Godown.GodownNo,"
tmp = tmp & " Mst_GSL.StackNo, Mst_GSL.LotNo , Mst_GSL.DepositeBags, Mst_GSL.CDTFBags, Mst_GSL.BalanceBags, Mst_GSL.DepositeWeight,"
tmp = tmp & " Mst_GSL.CDTFWeight , Mst_GSL.BalanceWeight, Mst_ExchangeType.ExchangeType, Mst_Client.ClientName, Mst_GSL.Flag ,"
tmp = tmp & " Mst_GSL.ISINID, Mst_GSL.CDTFNo, Mst_GSL.CDTFDate, Mst_GSL.PledgeNo, Mst_GSL.PledgeBags, Mst_GSL.PledgeWeight ,"
tmp = tmp & " Mst_GSL.SpillageFlag , Mst_GSL.DepositAdjWt,Mst_GSL.WithdrawalAdjWt,mst_gsl.g_uid,  Mst_GSL.KRNo, Mst_GSL.GrossWt,"
tmp = tmp & " Mst_GSL.BGSLID , Mst_GSL.FGSLID, mst_gsl.WHRBags,mst_gsl.WHRWeight,Mst_StorageContract.ContractName, "
tmp = tmp & " Mst_GSL.BagSize,Mst_GSL.BagEquivalent,Mst_GSL.WtQty   FROM "
tmp = tmp & " Mst_GSL"
tmp = tmp & " left JOIN Mst_CMP on Mst_CMP.CMPID = Mst_GSL.CMPID"
tmp = tmp & " left JOIN Mst_Godown on Mst_Godown.GodownID = Mst_GSL.GodownID"
tmp = tmp & " left JOIN Mst_Commodity on Mst_Commodity.CommodityID = Mst_GSL.CommodityID"
tmp = tmp & " LEFT JOIN Mst_ExchangeType ON Mst_GSL.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID"
tmp = tmp & " left JOIN Mst_Client on Mst_Client.ClientIDRow = Mst_GSL.ClientIDRow"
tmp = tmp & " LEFT JOIN Mst_ISIN ON Mst_GSL.ISINID = Mst_ISIN.ISINID"
tmp = tmp & " LEFT JOIN Mst_StorageContract ON Mst_GSL.ContractID = Mst_StorageContract.ContractID"

wc = " Where Mst_GSL.KRNo = " & Val(TxtKRNo) & " And Mst_GSL.ExchangeTypeID = " & mExchangeTypeID & " And (Mst_GSL.Flag = 'O' Or Mst_GSL.Flag = 'J') "

tmp = tmp & " " & wc
tmp = tmp & " ORDER BY mst_gsl.sm_prefix,Mst_CMP.CMPName,Mst_GSL.GSLID;"

Call TmpTable

If TmpRs.RecordCount > 0 Then
   
   If TxtKRNo < 1000 Then
      TxtCMPID = "Opening Balance" 'TmpRs!CMPName
      TxtCommodity = "Opening Balance" 'TmpRs!Commodity
      TxtLocationID = "Opening Balance" 'TmpRs!SM_Prefix
   Else
      TxtCMPID = TmpRs!CMPName
      TxtCommodity = TmpRs!Commodity
      TxtLocationID = GetLocationFromCMPName(TxtCMPID) ' TmpRs!SM_Prefix
   End If
   
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
   For I = 1 To TmpRs.RecordCount
       For C = 0 To TmpRs.Fields.Count - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
       Next
        
       MSHFlexGrid2.TextMatrix(I, 7) = Format(MSHFlexGrid2.TextMatrix(I, 7), "########0")
       MSHFlexGrid2.TextMatrix(I, 8) = Format(MSHFlexGrid2.TextMatrix(I, 8), "########0")
       MSHFlexGrid2.TextMatrix(I, 9) = Format(MSHFlexGrid2.TextMatrix(I, 9), "########0")
       MSHFlexGrid2.TextMatrix(I, 10) = Format(MSHFlexGrid2.TextMatrix(I, 10), "########0.000")
       MSHFlexGrid2.TextMatrix(I, 11) = Format(MSHFlexGrid2.TextMatrix(I, 11), "########0.000")
       MSHFlexGrid2.TextMatrix(I, 12) = Format(MSHFlexGrid2.TextMatrix(I, 12), "########0.000")
       MSHFlexGrid2.TextMatrix(I, 20) = Format(MSHFlexGrid2.TextMatrix(I, 20), "########0")
       MSHFlexGrid2.TextMatrix(I, 21) = Format(MSHFlexGrid2.TextMatrix(I, 21), "########0.000")
       
       MSHFlexGrid2.TextMatrix(I, 18) = Format(MSHFlexGrid2.TextMatrix(I, 18), "dd-mmm-yyyy")
       
       If MSHFlexGrid2.TextMatrix(I, 15) = "J" Then
          MSHFlexGrid2.Col = 0
          MSHFlexGrid2.row = I
          MSHFlexGrid2.CellBackColor = vbBlue
       End If
       If MSHFlexGrid2.TextMatrix(I, 15) <> "J" Then
          If MSHFlexGrid2.TextMatrix(I, 16) <> "" Then
             MSHFlexGrid2.Col = 0
             MSHFlexGrid2.row = I
             MSHFlexGrid2.CellBackColor = vbYellow
          End If
       End If
       TmpRs.MoveNext
   Next
   Label6.Caption = TmpRs.RecordCount
   CmdExcel.Enabled = True
   CmdDeleteKR.Enabled = True
Else
   MsgBox "No GSL Record found for given KR!!!"
   Call ClearContents
   TxtKRNo.SetFocus
   Exit Sub
End If


End Sub
Private Sub ClearContents()

TxtCMPID = "" 'TmpRs!CMPName
TxtCommodity = "" 'TmpRs!Commodity
TxtLocationID = "" 'TmpRs!SM_Prefix
TxtRemark = ""
CmdExcel.Enabled = False
CmdDeleteKR.Enabled = False
Label6.Caption = ""

End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_DeleteKR.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_DeleteKR.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, C), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid2.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
           End If
        Else
           oWS.Cells(mRow, C + 1) = MSHFlexGrid2.TextMatrix(I, C)
        End If
        MSHFlexGrid2.row = I
        MSHFlexGrid2.Col = C
        
        If MSHFlexGrid2.CellBackColor = vbBlue Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbBlue
        ElseIf MSHFlexGrid2.CellBackColor = vbYellow Then
           oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
