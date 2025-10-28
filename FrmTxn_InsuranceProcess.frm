VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_InsuranceProcess 
   Caption         =   "Insurance Process"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12990
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   12915
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8760
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   16260
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   105
         Left            =   120
         TabIndex        =   2
         Top             =   855
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   185
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   130
         Width           =   15975
         Begin VB.CommandButton CmdInsuranceRiskWiseReport 
            Caption         =   "Insurance Risk Wise Report"
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
            Left            =   9405
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   2745
         End
         Begin VB.CommandButton CmdClose 
            Caption         =   "Close"
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
            Left            =   14760
            TabIndex        =   6
            Top             =   240
            Width           =   1140
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
            Left            =   12450
            TabIndex        =   4
            Top             =   240
            Width           =   1140
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
            Left            =   13600
            TabIndex        =   5
            Top             =   240
            Width           =   1140
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   975
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   6588
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   3855
         Left            =   120
         TabIndex        =   7
         Top             =   4800
         Width           =   15975
         _ExtentX        =   28178
         _ExtentY        =   6800
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "FrmTxn_InsuranceProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExportReportList() As Variant, mFilePath As Variant
Dim mInsuranceRiskCoveredId As Variant, mInsuranceRiskCovered As Variant
Dim mAutoProcess As Boolean

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_InsuranceProcess" & " " & mInsuranceRiskCovered & ".xls")
Call Xls_Detail_Rpt
End Sub

Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Integer, mCol As Integer
Dim mColor_ As ColorConstants

If mAutoProcess = False Then
    MsgBox ("Wait till Next Message")
End If
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_InsuranceProcess" & " " & mInsuranceRiskCovered & ".xls")
mFilePath = Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_InsuranceProcess" & " " & mInsuranceRiskCovered & ".xls"
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1

ProgressBar1.Max = MSHFlexGrid1.Rows + MSHFlexGrid2.Rows

With MSHFlexGrid1
For I = 0 To .Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To .Cols - 2
        If InStr(.TextMatrix(I, C), "/") > 0 Then
           If IsDate(.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        Else
           If C = 1 Or C = 2 Then
            oWS.Cells(mRow, C + 1) = "'" & .TextMatrix(I, C)
           Else
            oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        End If
    Next
    ProgressBar1.Value = I
Next
End With
oWB.Save

Set oWS = oWB.Worksheets("Sheet2")
oWS.ClearArrows
mRow = 1

With MSHFlexGrid2
For I = 0 To .Rows - 2
    mRow = I + 1 ' + 6
    For C = 0 To .Cols - 1
        If InStr(.TextMatrix(I, C), "/") > 0 Then
           If IsDate(.TextMatrix(I, C)) Then
              oWS.Cells(mRow, C + 1) = Format(.TextMatrix(I, C), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        Else
           If C = 3 Or C = 17 Or C = 18 Then
            oWS.Cells(mRow, C + 1) = "'" & .TextMatrix(I, C)
           Else
            oWS.Cells(mRow, C + 1) = .TextMatrix(I, C)
           End If
        End If
        
        If I > 0 Then
            .row = mRow - 1
            .Col = C
''            mColor_ = .CellBackColor
''            If mColor_ <> vbWhite Then
''                oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = mColor_
''            End If
        
            If .CellBackColor = vbRed Then
                oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbRed
            ElseIf .CellBackColor = vbCyan Then
                oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbCyan
            ElseIf .CellBackColor = vbYellow Then
                oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbYellow
            ElseIf .CellBackColor = vbMagenta Then
                oWS.Range(oWS.Cells(mRow, C + 1), oWS.Cells(mRow, C + 1)).Interior.Color = vbMagenta
            End If
        End If
    
    Next
    ProgressBar1.Value = MSHFlexGrid1.Rows + I
    
Next
End With
oWB.Save

If mAutoProcess = False Then
    MsgBox ("Excel File Created !!! ")
    oXL.Visible = True
Else
    'Call Gen_Email
End If

End Sub

Private Sub CmdGo_Click()
''--- Geather GSL Stock WIth Current AUM  Client/Godown/Commodity /AUM Wise
Dim mR As Double, mRow As Double
Dim mInsurancePolicyID As Double, mInsurancePolicyAmt As Double, mInsurancePolicyUsedAmt As Double, mStructureWiseLimitAmt As Double
Dim mCommodityInclude As Double, mCommodityExclude As Double, mAUM As Double
Dim mInsuranceByClient As Boolean
Dim mInsuranceRemainAmt As Double, mRemainAMUAmt As Double
Dim mPreviousInsuranceCovered As Double
Dim mYearMonth As Variant
Dim mCMPIDList As Variant

Dim mPolicyDayPremium As Double

If mAutoProcess = False Then
    MsgBox ("Wait till Next Message")
End If

Call Grid_Head
Call Grid_Head1
mYearMonth = Format(Date, "yyyy-mm")

''-- Add two More Columns for PolicyDayPremium and Process flag 'M' / 'S'
tmp = " Insert into Txn_ArchiveInsurancePolicyData Select ClientID,CommodityID,CMPID,GodownID,LocationID,BalanceBags,BalanceWeight,Rate,YearMonth,AUM,InsuranceCovered,LinkByPolicyID,CoveredByPolicyID,ClientNCMLFlagID,StructureWiseCMPLimitOver,ClientNCMLStockFlag,Flag,G_UID,Concurrency,CreatedBy,CreatedDate,UpdatedBy,UpdatedDate,PolicyDayPremium,ProcessFlag from Txn_InsurancePolicyData  "
Call TmpTable

'tmp = " Truncate Table Txn_InsurancePolicyData "
'Call TmpTable

tmp = " Delete from Txn_InsurancePolicyData "
Call TmpTable


''-- Insert All AUM Data In Txn_InsurancePolicyData Table
'-- All Client Stock and Spillage Bags too
''-- LinkByPolicyID,CoveredByPolicyID,ClientNCMLFlagID,StructureWiseCMPLimitOver,Flag,LinkByPolicyID,CoveredByPolicyID,ClientNCMLFlagID,StructureWiseCMPLimitOver,Flag,

tmp = " Insert into Txn_InsurancePolicyData(ClientID,CommodityID,CMPID,GodownID,LocationID,BalanceBags,BalanceWeight,Rate,YearMonth,AUM,InsuranceCovered,G_UID,CreatedBy,CreatedDate,ProcessFlag)"
tmp = tmp & " Select GSL.ClientID,GSL.CommodityID,GSL.CMPID,GSL.GodownID,ML.LocationID,Sum(GSL.BalanceBags)[BalanceBags],Sum(GSL.BalanceWeight)[BalanceWeight],Isnull(CRate.Rate,0)Rate,CRate.YearMonth,Sum(GSL.BalanceWeight)*Isnull(CRate.Rate,0)[AUM],0,NEWID(),'" & Gen_UserLoginID & "',GETDATE(),'M'"
tmp = tmp & " From Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Location ML On GSL.SM_Prefix = ML.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
tmp = tmp & " Inner Join Mst_Commodity MC On GSL.CommodityID = MC.CommodityID"
tmp = tmp & " left Join"
tmp = tmp & " (   Select MCP.StateID,MCP.CommodityID,MCP.YearMonth,"
''-- Commented on 26 June 2018 By Ullhas Due to Date wise Commodity Price Upload
'tmp = tmp & "     Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
'tmp = tmp & "     Then"
'tmp = tmp & "       Case DatePart(Day, GETDATE())"
'tmp = tmp & "       When 1 Then MCP.Day1 When 2 Then MCP.Day2 When 3 Then MCP.Day3 When 4 Then MCP.Day4 When 5 Then MCP.Day5 When 6 Then MCP.Day6 When 7 Then MCP.Day7 When 8 Then MCP.Day8 When 9 Then MCP.Day9 When 10 Then MCP.Day10"
'tmp = tmp & "       When 11 Then MCP.Day11 When 12 Then MCP.Day12 When 13 Then MCP.Day13 When 14 Then MCP.Day14 When 15 Then MCP.Day15 When 16 Then MCP.Day16 When 17 Then MCP.Day17 When 18 Then MCP.Day18 When 19 Then MCP.Day19 When 20 Then MCP.Day20"
'tmp = tmp & "       When 21 Then MCP.Day21 When 22 Then MCP.Day22 When 23 Then MCP.Day23 When 24 Then MCP.Day24 When 25 Then MCP.Day25 When 26 Then MCP.Day26 When 27 Then MCP.Day27 When 28 Then MCP.Day28 When 29 Then MCP.Day29 When 30 Then MCP.Day30"
'tmp = tmp & "       When 31 Then MCP.Day31 Else 0 End"
'tmp = tmp & "     Else"
'tmp = tmp & "       Case DATEPART(Day,DATEADD(s,-1,DATEADD(mm, DATEDIFF(m,0,Convert(datetime,MCP.YearMonth+'-01'))+1,0)))"
'tmp = tmp & "       When 1 Then MCP.Day1 When 2 Then MCP.Day2 When 3 Then MCP.Day3 When 4 Then MCP.Day4 When 5 Then MCP.Day5 When 6 Then MCP.Day6 When 7 Then MCP.Day7 When 8 Then MCP.Day8 When 9 Then MCP.Day9 When 10 Then MCP.Day10"
'tmp = tmp & "       When 11 Then MCP.Day11 When 12 Then MCP.Day12 When 13 Then MCP.Day13 When 14 Then MCP.Day14 When 15 Then MCP.Day15 When 16 Then MCP.Day16 When 17 Then MCP.Day17 When 18 Then MCP.Day18 When 19 Then MCP.Day19 When 20 Then MCP.Day20"
'tmp = tmp & "       When 21 Then MCP.Day21 When 22 Then MCP.Day22 When 23 Then MCP.Day23 When 24 Then MCP.Day24 When 25 Then MCP.Day25 When 26 Then MCP.Day26 When 27 Then MCP.Day27 When 28 Then MCP.Day28 When 29 Then MCP.Day29 When 30 Then MCP.Day30"
'tmp = tmp & "       When 31 Then MCP.Day31 Else 0 End"
'tmp = tmp & "     End [Rate]"

tmp = tmp & "     Case When Convert(date,MCP.YearMonth+'-01') = Convert(date,Convert(Varchar,DATEPART(Year,GETDATE()))+'-'+Convert(Varchar,DATEPART(Month,GETDATE()))+'-01')"
tmp = tmp & "     Then"
tmp = tmp & "       COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & "     Else"
tmp = tmp & "       COALESCE(MCP.Day31,MCP.Day30,MCP.Day29,MCP.Day28,MCP.Day27,MCP.Day26,MCP.Day25,MCP.Day24,MCP.Day23,MCP.Day22,MCP.Day21,MCP.Day20,MCP.Day19,MCP.Day18,MCP.Day17,MCP.Day16,MCP.Day15,MCP.Day14,MCP.Day13,MCP.Day12,MCP.Day11,MCP.Day10,MCP.Day9,MCP.Day8,MCP.Day7,MCP.Day6,MCP.Day5,MCP.Day4,MCP.Day3,MCP.Day2,MCP.Day1)"
tmp = tmp & "     End [Rate]"

tmp = tmp & "     from Mst_CommodityPrice MCP"
tmp = tmp & "     Where Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,YearMonth) In  (Select Convert(Varchar,StateID)+'~'+Convert(Varchar,CommodityID)+'~'+Convert(Varchar,MAX(YearMonth)) From Mst_CommodityPrice Group By StateID,CommodityID ) "
tmp = tmp & " ) As CRate On MS.StateID = CRate.StateID And GSL.CommodityID = CRate.CommodityID"
'Commented on 18 May 2018 by Ullhas
'tmp = tmp & " Where GSL.Flag In ('O','Z','J','R') And GSL.BalanceWeight > 0 "
'tmp = tmp & " Where GSL.Flag Not In ('B','Z','E') And GSL.BalanceWeight >= 0 "
tmp = tmp & " Where GSL.Status not In ('E','Z','B') And GSL.BalanceWeight > 0 "
tmp = tmp & " Group By GSL.ClientID,GSL.CommodityID,GSL.CMPID,GSL.GodownID,ML.LocationID,CRate.Rate,CRate.YearMonth"
'Commented on 18 May 2018 by Ullhas
'tmp = tmp & " Having Sum(GSL.BalanceBags) > 0 And Sum(GSL.BalanceWeight) > 0"
tmp = tmp & " Order By GSL.ClientID Desc,[AUM] Desc,GSL.CommodityID,ML.LocationID,GSL.CMPID,GSL.GodownID"
Call TmpTable

''-- List All the Valid Insurance Policy Order By CLient TO NCML (Client/Commodity)
tmp = " Select MID.InsurancePolicyID,MID.PolicyNo,MID.CoverNoteNo,MICM.InsuranceCoName,MID.Amount,MID.StructureWiseLimit,MID.ValidFrom,MID.ValidTo,MIB.InsuranceByName,MIRC.InsuranceRiskCovered,MID.InsuranceByID,MICD.CommodityID,MICD.EntryType,MICDD.ClientID"
tmp = tmp & " From Mst_InsuranceDetails MID"
tmp = tmp & " Inner Join Mst_InsuranceBy MIB On MID.InsuranceByID =MIB.InsuranceByID"
tmp = tmp & " Inner Join Mst_InsuranceCoMaster MICM On MID.InsuranceCoID = MICM.InsuranceCoID"
tmp = tmp & " Left Join Mst_InsuranceRiskCovered MIRC On MID.InsuranceRiskCoveredID = MIRC.InsuranceRiskCoveredID"
tmp = tmp & " Left Join Mst_InsuranceCommodityDetail MICD On MID.InsurancePolicyID = MICD.InsurancePolicyID"
tmp = tmp & " Left Join Mst_InsuranceClientDetail MICDD On MID.InsurancePolicyID = MICDD.InsurancePolicyID"
tmp = tmp & " Where MID.Flag = 'A' "
''-- Commented on 17 July 2018
''tmp = tmp & " And '" & Format(Date, "yyyy-mm-dd") & "' Between MID.ValidFrom and MID.ValidTo "
''tmp = tmp & " And (MID.InsuranceByID Is Not Null And MID.InsuranceRiskCoveredID Is Not Null)"
tmp = tmp & " " & IIf(Val(mInsuranceRiskCoveredId) = 0, "", " And MID.InsuranceRiskCoveredID = " & mInsuranceRiskCoveredId & "  ")
tmp = tmp & " Order By MID.InsuranceByID DESC,MICD.CommodityID Desc,MICDD.ClientID"
Call TmpTable

ProgressBar1.Value = 0
If TmpRs.RecordCount > 0 Then
    With MSHFlexGrid1
    ProgressBar1.Max = TmpRs.RecordCount
    For mR = 1 To TmpRs.RecordCount
        
        If PolicyIDAvailable(IIf(IsNull(TmpRs!InsurancePolicyID), 0, TmpRs!InsurancePolicyID)) = False Then
            mRow = .Rows - 1
            .TextMatrix(mRow, 0) = IIf(IsNull(TmpRs!InsurancePolicyID), 0, TmpRs!InsurancePolicyID)
            .TextMatrix(mRow, 1) = IIf(IsNull(TmpRs!PolicyNo), "", TmpRs!PolicyNo)
            .TextMatrix(mRow, 2) = IIf(IsNull(TmpRs!CoverNoteNo), "", TmpRs!CoverNoteNo)
            .TextMatrix(mRow, 3) = IIf(IsNull(TmpRs!InsuranceCoName), "", TmpRs!InsuranceCoName)
            .TextMatrix(mRow, 4) = IIf(IsNull(TmpRs!Amount), 0, Format(TmpRs!Amount, "#0.00"))
            .TextMatrix(mRow, 5) = IIf(IsNull(TmpRs!StructureWiseLimit), 0, TmpRs!StructureWiseLimit)
            .TextMatrix(mRow, 6) = IIf(IsNull(TmpRs!ValidFrom), "", Format(TmpRs!ValidFrom, "DD-MMM-YYYY"))
            .TextMatrix(mRow, 7) = IIf(IsNull(TmpRs!ValidTo), "", Format(TmpRs!ValidTo, "DD-MMM-YYYY"))
            .TextMatrix(mRow, 8) = IIf(IsNull(TmpRs!InsuranceByName), "", TmpRs!InsuranceByName)
            .TextMatrix(mRow, 9) = IIf(IsNull(TmpRs!InsuranceRiskCovered), "", TmpRs!InsuranceRiskCovered)
            .TextMatrix(mRow, 10) = MappingDataforInsurance(Val(.TextMatrix(mRow, 0)), "ClientName") ''-- ClientName
            .TextMatrix(mRow, 11) = MappingDataforInsurance(Val(.TextMatrix(mRow, 0)), "Commodity") ''-- CommodityName
            .TextMatrix(mRow, 12) = MappingDataforInsurance(Val(.TextMatrix(mRow, 0)), "GodownType")  ''-- GodownType
            .TextMatrix(mRow, 13) = 0#
            .TextMatrix(mRow, 14) = 0#
            .TextMatrix(mRow, 15) = ""  ''-- Storage Policy Amt
            .TextMatrix(mRow, 16) = IIf(IsNull(TmpRs!InsuranceByID), 0, TmpRs!InsuranceByID)
            .Rows = .Rows + 1
        End If
        
        ProgressBar1.Value = mR
        TmpRs.MoveNext
    Next mR
    End With
End If

mPolicyDayPremium = 0#  ''-- Ullhas Patil 17 July 2018

If MSHFlexGrid1.Rows > 2 Then
    
    ''------------------------------------------------------------------------------------------------
    ''-- Link the Policy to Stock
    For mR = 1 To MSHFlexGrid1.Rows - 2
        ProgressBar1.Max = MSHFlexGrid1.Rows - 2
        With MSHFlexGrid1
        
            mInsurancePolicyID = Val(.TextMatrix(mR, 0))
            mInsuranceByClient = IIf(Val(.TextMatrix(mR, 16)) = 2, True, False)
            mInsurancePolicyAmt = Val(.TextMatrix(mR, 4))
            mInsurancePolicyUsedAmt = Val(.TextMatrix(mR, 13))
            mStructureWiseLimitAmt = Val(.TextMatrix(mR, 5))
            
            mCommodityInclude = 0
            mCommodityExclude = 0
            mPreviousInsuranceCovered = 0
            
            tmp = " Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I' "
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
                mCommodityInclude = TmpRs1.RecordCount
            End If
            
            tmp = " Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E' "
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
                mCommodityExclude = TmpRs1.RecordCount
            End If
            
            tmp = " Select * from Txn_InsurancePolicyData Where (AUM - InsuranceCovered) > 0 "
            Call TmpTable
        
            ''-- Find Stock As per Mapping and Link the Policy No
            ''-- FInd The Strture Wise Limit And Mark the Over Limit of Stock Lines (Note -: Avoid The Client Which Having there own Policy and Avoid the Commodity if it is there)
            If mInsuranceByClient = True Then
                
                ''------------------------------------------------------------
                ''--- Link The Policy No
                tmp = " Update Txn_InsurancePolicyData SET ClientNCMLStockFlag = 'C',LinkByPolicyID = Isnull(LinkByPolicyID,'') + '" & mInsurancePolicyID & ",' Where InsuranceCovered = 0 "
                tmp = tmp & " And ClientID In (Select ClientID  from Mst_InsuranceClientDetail Where InsurancePolicyID = " & mInsurancePolicyID & " )"
                If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                    tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                    tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                End If
                tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                tmp = tmp & " And IsNull(ClientNCMLStockFlag,'') In ('','C') "
                Call Tmp1Table
                ''------------------------------------------------------------
                
                ''------------------------------------------------------------
                ''-- Update The StructureWiseLimitOver = Y if CMP Wise AUM is More then Limit
                If mStructureWiseLimitAmt > 0 Then
                    mCMPID = "0"
                    tmp = "SELECT CMPID,SUM(AUM)[AUMTotal] FROM Txn_InsurancePolicyData Where InsuranceCovered = 0 "
                    tmp = tmp & " And ClientID In (Select ClientID  from Mst_InsuranceClientDetail Where InsurancePolicyID = " & mInsurancePolicyID & " )"
                    If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                        tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                    ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                        tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                    End If
                    tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                    tmp = tmp & " Group By CMPID Having SUM(AUM) > " & mStructureWiseLimitAmt & " Order BY CMPID "
                    Call Tmp2Table
                    If TmpRs2.RecordCount > 0 Then
                        For mRow = 1 To TmpRs2.RecordCount
                            mCMPID = mCMPID & "," & IIf(IsNull(TmpRs2!CMPID), 0, TmpRs2!CMPID)
                            TmpRs2.MoveNext
                        Next mRow
                    End If
                    
                    If mCMPID <> "" Then
                        tmp = " UPDATE Txn_InsurancePolicyData Set StructureWiseCMPLimitOver = 'Y' Where InsuranceCovered = 0 "
                        tmp = tmp & " And ClientID In (Select ClientID  from Mst_InsuranceClientDetail Where InsurancePolicyID = " & mInsurancePolicyID & " )"
                        If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                            tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                        ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                            tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                        End If
                        tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                        tmp = tmp & " And CMPID In (" & mCMPID & ") "
                        Call Tmp2Table
                    End If
                End If
                ''------------------------------------------------------------
            
            ElseIf mInsuranceByClient = False Then
                
                ''------------------------------------------------------------
                ''--- Link The Policy No
                tmp = " Update Txn_InsurancePolicyData SET ClientNCMLStockFlag = 'N',LinkByPolicyID = Isnull(LinkByPolicyID,'') + '" & mInsurancePolicyID & ",' Where InsuranceCovered = 0 "
                If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                    tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                    tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                End If
                tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                tmp = tmp & " And IsNull(ClientNCMLStockFlag,'') In ('','N') "
                Call Tmp1Table
                ''------------------------------------------------------------
                
                ''------------------------------------------------------------
                ''-- Update The StructureWiseLimitOver = Y if CMP Wise AUM is More then Limit
                If mStructureWiseLimitAmt > 0 Then
                    mCMPID = "0"
                    tmp = "SELECT CMPID,SUM(AUM)[AUMTotal] FROM Txn_InsurancePolicyData Where InsuranceCovered = 0 "
                    If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                        tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                    ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                        tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                    End If
                    tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                    tmp = tmp & " And ClientID Not In (Select ClientID  from Mst_InsuranceClientDetail ) "
                    tmp = tmp & " Group By CMPID Having SUM(AUM) > " & mStructureWiseLimitAmt & " Order BY CMPID "
                    Call Tmp2Table
                    If TmpRs2.RecordCount > 0 Then
                        For mRow = 1 To TmpRs2.RecordCount
                            mCMPID = mCMPID & "," & IIf(IsNull(TmpRs2!CMPID), 0, TmpRs2!CMPID)
                            TmpRs2.MoveNext
                        Next mRow
                    End If
                    
                    If mCMPID <> "" Then
                        tmp = " UPDATE Txn_InsurancePolicyData Set StructureWiseCMPLimitOver = 'Y' Where InsuranceCovered = 0 "
                        If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                            tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                        ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                            tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                        End If
                        tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                        tmp = tmp & " And ClientID Not In (Select ClientID  from Mst_InsuranceClientDetail ) "
                        tmp = tmp & " And CMPID In (" & mCMPID & ") "
                        Call Tmp2Table
                    End If
                End If
                ''--------------------------------------------------
                
            End If
        
        End With
    ProgressBar1.Value = mR
    ProgressBar1.Refresh
    Next mR
    ''------------------------------------------------------------------------------------------------

    For mR = 1 To MSHFlexGrid1.Rows - 2
        ProgressBar1.Max = MSHFlexGrid1.Rows - 2
        With MSHFlexGrid1
            
            mInsurancePolicyID = Val(.TextMatrix(mR, 0))
            mInsuranceByClient = IIf(Val(.TextMatrix(mR, 16)) = 2, True, False)
            mInsurancePolicyAmt = Val(.TextMatrix(mR, 4))
            mInsurancePolicyUsedAmt = Val(.TextMatrix(mR, 13))
            mStructureWiseLimitAmt = Val(.TextMatrix(mR, 5))
                                
            ''-- Commented on 25 July 2018
            ''-- Ullhas Patil 17 July 2018
            ''mPolicyDayPremium = PolicyDayPremium(CDbl(mInsurancePolicyID))
                    
            mCommodityInclude = 0
            mCommodityExclude = 0
            mPreviousInsuranceCovered = 0
            
            tmp = " Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I' "
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
                mCommodityInclude = TmpRs1.RecordCount
            End If
            
            tmp = " Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E' "
            Call Tmp1Table
            If TmpRs1.RecordCount > 0 Then
                mCommodityExclude = TmpRs1.RecordCount
            End If
            
            tmp = " Select * from Txn_InsurancePolicyData Where (AUM - InsuranceCovered) > 0 "
            Call TmpTable
            
            
            ''-- Find Stock Client Wise / Commodity / Godown and Apply Insurance
            ''--->> If Insurance By Client is True then
            If mInsuranceByClient = True Then
            
                tmp = "SELECT * FROM Txn_InsurancePolicyData Where (AUM - InsuranceCovered) > 0 And ClientNCMLStockFlag = 'C' "
                tmp = tmp & " And ClientID In (Select ClientID  from Mst_InsuranceClientDetail Where InsurancePolicyID = " & mInsurancePolicyID & " )"
                If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                    tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                    tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                End If
                tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                tmp = tmp & " Order BY ClientID,AUM Desc,CommodityID,GodownID"
                Call Tmp1Table
                If TmpRs1.RecordCount > 0 Then
                    For mRow = 1 To TmpRs1.RecordCount
                        mInsuranceRemainAmt = Format(mInsurancePolicyAmt - mInsurancePolicyUsedAmt, "#0.00")
                        TmpRs.Filter = ""
                        If TmpRs.RecordCount > 0 Then
                           
                               TmpRs.Filter = " ClientID = " & Val(TmpRs1!ClientID) & " And LocationID = " & Val(TmpRs1!LocationID) & " And GodownID = " & Val(TmpRs1!GodownID) & " And CommodityID = " & Val(TmpRs1!CommodityID) & " "
                               If TmpRs.EOF = False Then
                                    If mInsuranceRemainAmt > 0 Then
                                        mPreviousInsuranceCovered = Format(TmpRs!InsuranceCovered, "#0.00")
                                        mRemainAMUAmt = Format(TmpRs!AUM - TmpRs!InsuranceCovered, "#0.00")
                                        TmpRs!InsuranceCovered = Format(IIf(IsNull(TmpRs!InsuranceCovered), 0, TmpRs!InsuranceCovered) + Val(IIf(mInsuranceRemainAmt > mRemainAMUAmt, mRemainAMUAmt, mInsuranceRemainAmt)), "#0.00")
                                        TmpRs!CoveredByPolicyID = IIf(IsNull(TmpRs!CoveredByPolicyID), "", TmpRs!CoveredByPolicyID) & mInsurancePolicyID & ","
                                        mInsurancePolicyUsedAmt = mInsurancePolicyUsedAmt + TmpRs!InsuranceCovered
                                        .TextMatrix(mR, 13) = Format(mInsurancePolicyUsedAmt - mPreviousInsuranceCovered, "#0.00")
                                        TmpRs!Flag = IIf((TmpRs!AUM - TmpRs!InsuranceCovered) = 0, 0, 1)
                                        
                                        ''-- Commented on 25 July 2018
                                        ''-- Ullhas Patil 17 July 2018
                                        ''TmpRs!PolicyDayPremium = IIf(IsNull(TmpRs!PolicyDayPremium), 0, TmpRs!PolicyDayPremium) + mPolicyDayPremium
                                    Else
                                        TmpRs!Flag = 1
                                    End If
                                    TmpRs.Update
                                    TmpRs.Requery
                               End If

                        End If
                        TmpRs1.MoveNext
                    Next mRow
                End If
                
            ''--->> If Insurance By Client is False then
            ElseIf mInsuranceByClient = False Then
            
                tmp = "SELECT * FROM Txn_InsurancePolicyData Where (AUM - InsuranceCovered) > 0 And Isnull(Flag,'0') = 0 And IsNUll(ClientNCMLStockFlag,'') In ('','N') "
                If mCommodityInclude > 0 And mCommodityExclude = 0 Then
                    tmp = tmp & " And CommodityID In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'I')"
                ElseIf mCommodityInclude = 0 And mCommodityExclude > 0 Then
                    tmp = tmp & " And CommodityID Not In (Select CommodityID  from Mst_InsuranceCommodityDetail Where InsurancePolicyID = " & mInsurancePolicyID & " And EntryType = 'E')"
                End If
                tmp = tmp & " And GodownID IN ( Select GodownID  from Mst_InsuranceGodownMapping Where InsurancePolicyID = " & mInsurancePolicyID & " And Connected = 'Y' )"
                tmp = tmp & " Order BY ClientID Desc,AUM Desc,CommodityID,GodownID"
                Call Tmp1Table
                If TmpRs1.RecordCount > 0 Then
                    For mRow = 1 To TmpRs1.RecordCount
                        
                        mInsuranceRemainAmt = Format(mInsurancePolicyAmt - mInsurancePolicyUsedAmt, "#0.00")
                        TmpRs.Filter = ""
                        If TmpRs.RecordCount > 0 Then
                               If IsNull(TmpRs1!ClientID) = True Then
                                TmpRs.Filter = " ClientID = Null And LocationID = " & Val(TmpRs1!LocationID) & " And GodownID = " & Val(TmpRs1!GodownID) & " And CommodityID = " & Val(TmpRs1!CommodityID) & " "
                               Else
                                TmpRs.Filter = " ClientID = " & Val(TmpRs1!ClientID) & " And LocationID = " & Val(TmpRs1!LocationID) & " And GodownID = " & Val(TmpRs1!GodownID) & " And CommodityID = " & Val(TmpRs1!CommodityID) & " "
                               End If
                               If TmpRs.EOF = False Then
                                    If mInsuranceRemainAmt > 0 Then
                                        mRemainAMUAmt = Format(TmpRs!AUM - TmpRs!InsuranceCovered, "#0.00")
                                        TmpRs!InsuranceCovered = Format(IIf(IsNull(TmpRs!InsuranceCovered), 0, TmpRs!InsuranceCovered) + Val(IIf(mInsuranceRemainAmt > mRemainAMUAmt, mRemainAMUAmt, mInsuranceRemainAmt)), "#0.00")
                                        TmpRs!CoveredByPolicyID = IIf(IsNull(TmpRs!CoveredByPolicyID), "", TmpRs!CoveredByPolicyID) & mInsurancePolicyID & ","
                                        mInsurancePolicyUsedAmt = mInsurancePolicyUsedAmt + TmpRs!InsuranceCovered
                                        .TextMatrix(mR, 13) = Format(mInsurancePolicyUsedAmt, "#0.00")
                                        
                                        ''-- Commented on 25 July 2018
                                        ''-- Ullhas Patil 17 July 2018
                                        ''TmpRs!PolicyDayPremium = IIf(IsNull(TmpRs!PolicyDayPremium), 0, TmpRs!PolicyDayPremium) + mPolicyDayPremium
                                    Else
                                        Exit For
                                    End If
                                    TmpRs.Update
                                    TmpRs.Requery
                               End If
                        End If
                        TmpRs1.MoveNext
                    
                    Next mRow
                End If
                
            End If
            .TextMatrix(mR, 14) = Format(.TextMatrix(mR, 4) - .TextMatrix(mR, 13), "#0.00")
        
        End With
    ProgressBar1.Value = mR
    ProgressBar1.Refresh
    Next mR
End If

''-- Storage Policy Amount
For mR = 1 To MSHFlexGrid1.Rows - 2
    ProgressBar1.Max = MSHFlexGrid1.Rows - 2
    With MSHFlexGrid1
        If Val(.TextMatrix(mR, 14)) = 0 Then
        
            mInsurancePolicyID = Val(.TextMatrix(mR, 0))
            mInsuranceByClient = IIf(Val(.TextMatrix(mR, 16)) = 2, True, False)
            mInsurancePolicyAmt = Val(.TextMatrix(mR, 4))
            mInsurancePolicyUsedAmt = Val(.TextMatrix(mR, 13))
            mAUM = 0#
            
            If mInsuranceByClient = True Then
                tmp = " Select Isnull(Sum(Aum-InsuranceCovered),0)[Aum] from Txn_InsurancePolicyData Where ClientNCMLStockFlag = 'C' And LinkByPolicyID Like '" & mInsurancePolicyID & ",%' And (AUM-InsuranceCovered) > 0 "
            Else
                tmp = " Select Isnull(Sum(AUM-InsuranceCovered),0)[Aum] from Txn_InsurancePolicyData Where ClientNCMLStockFlag = 'N' And LinkByPolicyID Like '" & mInsurancePolicyID & ",%' "
            End If
            Call TmpTable
            If TmpRs!AUM > 0 Then
              mAUM = IIf(IsNull(TmpRs!AUM), 0#, TmpRs!AUM)
              .TextMatrix(mR, 15) = Format(mAUM, "#0.00")
            Else
              .TextMatrix(mR, 15) = ""
            End If
        End If
    End With
Next mR
''------------------------

''-- Ullhas Patil 25 July 2018
Dim mListofPolicy() As Variant
Dim mSplit As Variant
Dim mPolicy As Variant
Dim mTotalPolicyCoveredAmount As Variant, mCountCoveredByPolicy As Variant

tmp = " Select Distinct CoveredByPolicyID from Txn_InsurancePolicyData Where ClientNCMLStockFlag = 'N' ANd CoveredByPolicyID Is Not Null "
Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   ReDim mListofPolicy(RecordCount)
   ReDim mListofPolicy(TmpRs1.RecordCount)
   For mR = 0 To TmpRs1.RecordCount - 1
    mListofPolicy(mR) = IIf(IsNull(TmpRs1!CoveredByPolicyID), "", TmpRs1!CoveredByPolicyID)
    TmpRs1.MoveNext
   Next mR
End If

If UBound(mListofPolicy) > 0 Then
   For mR = 0 To UBound(mListofPolicy) - 1
    mPolicy = Split(mListofPolicy(mR), ",")
        mPolicyDayPremium = 0
        For mRow = 0 To UBound(mPolicy) - 1
            If mPolicy(mRow) <> "" And IsNumeric(mPolicy(mRow)) = True Then
                mPolicyDayPremium = Val(mPolicyDayPremium) + PolicyDayPremium(CDbl(mPolicy(mRow)))
            End If
        Next mRow
        
        tmp = " Select SUM(InsuranceCovered)[TotalPolicyCoveredAmount],Count(CoveredByPolicyID)[CountCoveredByPolicy] from Txn_InsurancePolicyData Where ClientNCMLStockFlag = 'N' And CoveredByPolicyID = '" & mListofPolicy(mR) & "' "
        Call Tmp1Table
        If TmpRs1.RecordCount > 0 Then
            mTotalPolicyCoveredAmount = IIf(IsNull(TmpRs1!TotalPolicyCoveredAmount), 0#, TmpRs1!TotalPolicyCoveredAmount)
            mCountCoveredByPolicy = IIf(IsNull(TmpRs1!CountCoveredByPolicy), 0#, TmpRs1!CountCoveredByPolicy)
        Else
            mTotalPolicyCoveredAmount = 0#
            mCountCoveredByPolicy = 0#
        End If
        
        ''-- Apply PolicyDayPremium to the Stock Line with are connected to Listed policy
        ''-- Formula
        ''-- PolicyDayPremium = InsuranceCovered /Sum(InsuranceCovered) * OneDayPolicyPremium
        tmp = " Update Txn_InsurancePolicyData Set PolicyDayPremium = ( InsuranceCovered / " & mTotalPolicyCoveredAmount & ") * " & mPolicyDayPremium & " Where ClientNCMLStockFlag = 'N' ANd CoveredByPolicyID = '" & mListofPolicy(mR) & "' "
        Call Tmp1Table
        ''---
   Next mR
End If
''-----------------------

''-- Ullhas Patil 23 July 2018
''-- >>Insert the data Godown Wise In Txn_GodownWiseInsurancePremiumDetail table for PnL
''-- >>If data in already present in same date then delete and insert (i.e. Unique data for same day)
tmp = " Select Distinct ProcessDate from Txn_GodownWiseInsurancePremiumDetail Where ProcessDate = '" & Format(Date, "yyyy-MM-dd") & "' "
Call TmpTable
If TmpRs.RecordCount > 0 Then
    tmp = " Delete From Txn_GodownWiseInsurancePremiumDetail Where ProcessDate = '" & Format(Date, "yyyy-MM-dd") & "' "
    Call Tmp1Table
End If

''tmp = " INSERT INTO Txn_GodownWiseInsurancePremiumDetail (ProcessDate,GodownID,Amount,ProcessFlag,G_UID,CreatedDate,CreatedBy) Select CONVERT(DATE,getdate()),GodownID,Sum(Amount) 'Amount','M',NEWID(),GETDATE(),'" & Gen_UserLoginID & "' From (Select GodownID,CoveredByPolicyID,Sum(Distinct PolicyDayPremium) 'Amount' from Txn_InsurancePolicyData Group by GodownID,CoveredByPolicyID Having Sum(Distinct PolicyDayPremium) > 0)  A Group by GodownID "
tmp = " INSERT INTO Txn_GodownWiseInsurancePremiumDetail (ProcessDate,GodownID,Amount,ProcessFlag,G_UID,CreatedDate,CreatedBy) Select CONVERT(DATE,getdate()),GodownID,SUM(PolicyDayPremium) As 'Amount','M',NEWID(),GETDATE(),'" & Gen_UserLoginID & "' From Txn_InsurancePolicyData Where Isnull(PolicyDayPremium,0) > 0 Group By  GodownID "
Call TmpTable



''-- Call Insurance Detail Record Set
Call TableMst_PolicyDetail(" Where Flag = 'A' Order By InsurancePolicyID ")
Call TableMst_InsuranceBy
Call Grid_Head1

''-- Call Stock Report
With MSHFlexGrid2
    tmp = " SELECT A.ClientID,MC.ClientName,MC.ClientIDRow,MET.ExchangeType,Comm.Commodity,MG.GodownNo,MUT.UnitType,CMP.CMPName,ML.LocationName,ML.SM_Prefix,"
    tmp = tmp & " MS.StateName,MSI.StateIndiaName,MR.Region,A.BalanceBags,A.BalanceWeight,A.Rate,A.AUM,A.Flag,A.InsuranceCovered,A.LinkByPolicyID,A.CoveredByPolicyID,A.StructureWiseCMPLimitOver,A.ClientNCMLStockFlag,A.YearMonth,MG.Area"
    tmp = tmp & " FROM Txn_InsurancePolicyData AS A"
    tmp = tmp & " Left Join Mst_Client MC On A.ClientID = MC.ClientID"
    tmp = tmp & " Left Join Mst_ExchangeType MET On MC.ExchangeTypeID = MET.ExchangeTypeID"
    tmp = tmp & " Inner Join Mst_Commodity Comm On A.CommodityID = Comm.CommodityID"
    tmp = tmp & " Inner Join Mst_Godown MG On A.GodownID = MG.GodownID"
    tmp = tmp & " Inner Join Mst_CMP CMP On MG.CMPID = CMP.CMPID"
    tmp = tmp & " Inner Join Mst_Location ML On A.LocationID = ML.LocationID"
    tmp = tmp & " Inner Join Mst_State MS On ML.StateID = MS.StateID"
    tmp = tmp & " Inner Join Mst_StateIndia MSI On MS.StateIndiaID = MSI.StateIndiaID"
    tmp = tmp & " Inner Join Mst_Region MR On MS.RegionID = MR.RegionID"
    tmp = tmp & " Inner Join Mst_UnitType MUT On MG.UnitTypeID = MUT.UnitTypeID"
    tmp = tmp & " Order By MR.Region,MSI.StateIndiaName,MS.StateName,ML.LocationName,Comm.Commodity,MC.ClientName,A.AUM Desc"
    Call TmpTable
    If TmpRs.RecordCount > 0 Then
        ProgressBar1.Max = TmpRs.RecordCount
        For mR = 1 To TmpRs.RecordCount
                        
            .TextMatrix(mR, 0) = IIf(IsNull(TmpRs!ClientID), "", IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName) & "#" & IIf(IsNull(TmpRs!ClientIDRow), "", TmpRs!ClientIDRow) & "#" & IIf(IsNull(TmpRs!ClientID), "", TmpRs!ClientID))
            .TextMatrix(mR, 1) = IIf(IsNull(TmpRs!ExchangeType), "", TmpRs!ExchangeType)
            .TextMatrix(mR, 2) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
            .TextMatrix(mR, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
            .TextMatrix(mR, 4) = IIf(IsNull(TmpRs!UnitType), "", TmpRs!UnitType)
            .TextMatrix(mR, 5) = IIf(IsNull(TmpRs!CMPName), "", TmpRs!CMPName)
            .TextMatrix(mR, 6) = IIf(IsNull(TmpRs!Locationname), "", TmpRs!Locationname)
            .TextMatrix(mR, 7) = IIf(IsNull(TmpRs!SM_Prefix), "", TmpRs!SM_Prefix)
            .TextMatrix(mR, 8) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
            .TextMatrix(mR, 9) = IIf(IsNull(TmpRs!StateIndiaName), "", TmpRs!StateIndiaName)
            .TextMatrix(mR, 10) = IIf(IsNull(TmpRs!Region), "", TmpRs!Region)
            .TextMatrix(mR, 11) = IIf(IsNull(TmpRs!BalanceBags), 0, TmpRs!BalanceBags)
            .TextMatrix(mR, 12) = IIf(IsNull(TmpRs!BalanceWeight), 0, TmpRs!BalanceWeight)
            .TextMatrix(mR, 13) = IIf(IsNull(TmpRs!Rate), 0, TmpRs!Rate)
            .TextMatrix(mR, 14) = IIf(IsNull(TmpRs!AUM), "", TmpRs!AUM)
            .TextMatrix(mR, 15) = IIf(IsNull(TmpRs!InsuranceCovered), 0, TmpRs!InsuranceCovered)
            .TextMatrix(mR, 16) = .TextMatrix(mR, 14) - .TextMatrix(mR, 15)
            .TextMatrix(mR, 17) = ListOfInsurance(IIf(IsNull(TmpRs!CoveredByPolicyID), "", TmpRs!CoveredByPolicyID), "PolicyNo")
            .TextMatrix(mR, 18) = ListOfInsurance(IIf(IsNull(TmpRs!CoveredByPolicyID), "", TmpRs!CoveredByPolicyID), "CoverNoteNo")
            .TextMatrix(mR, 19) = ListOfInsurance(IIf(IsNull(TmpRs!LinkByPolicyID), "", TmpRs!LinkByPolicyID), "PolicyNo")
            .TextMatrix(mR, 20) = IIf(IsNull(TmpRs!ClientNCMLStockFlag), "", IIf(LCase(TmpRs!ClientNCMLStockFlag) = "c", "Client", "NCML")) ''ListOfInsurance(IIf(IsNull(TmpRs!CoveredByPolicyID), "", TmpRs!CoveredByPolicyID), "InsuranceBy")
            .TextMatrix(mR, 21) = IIf(IsNull(TmpRs!YearMonth), 0, TmpRs!YearMonth)
            .TextMatrix(mR, 22) = IIf(IsNull(TmpRs!Area), "", TmpRs!Area)
            
            ''-- Red if Rate is not vailable
            If Val(.TextMatrix(mR, 13)) = 0 Then
                Call ColorCell(MSHFlexGrid2, mR, vbRed, "Row", 0) ''-- Red
            End If
            
            ''-- Yellow if Storage Amount
            ''-- Color Both (Covered Amount and Shortage Amount) Column if AUM is partially covered
            ''-- Color Shortage Amount Column if Shortage Amt is remain
            If Val(.TextMatrix(mR, 15)) > 0 And Val(.TextMatrix(mR, 16)) > 0 Then   ''-- Yellow
                Call ColorCell(MSHFlexGrid2, mR, vbYellow, "Column", 15)
            End If
            If Val(.TextMatrix(mR, 16)) > 0 Then                                    ''-- Yellow
                Call ColorCell(MSHFlexGrid2, mR, vbYellow, "Column", 16)
            End If
            ''-------------
            
            ''-- Magenta if CMP Limit is Over then Limit
            If LCase(IIf(IsNull(TmpRs!StructureWiseCMPLimitOver), "", TmpRs!StructureWiseCMPLimitOver)) = "y" Then
                Call ColorCell(MSHFlexGrid2, mR, vbMagenta, "Row", 0)   ''-- Light Brown
            End If
            
            ''-- Cyan if Rate vailable and is old Year-Month
            If Val(.TextMatrix(mR, 13)) > 0 And .TextMatrix(mR, 21) <> mYearMonth Then
                Call ColorCell(MSHFlexGrid2, mR, vbCyan, "Column", 21) ''-- Sky Blue
            End If
            
            
            TmpRs.MoveNext
            .Rows = .Rows + 1
            ProgressBar1.Value = mR
            ProgressBar1.Refresh
        Next mR
    End If
End With

If mAutoProcess = False Then
    MsgBox "Done"
End If
''---
End Sub

Private Sub CmdInsuranceRiskWiseReport_Click()
Dim mCount As Double
Call TableMst_InsuranceRiskCovered
mAutoProcess = True
If RsMst_InsuranceRiskCovered.RecordCount > 0 Then
    ReDim mExportReportList(RsMst_InsuranceRiskCovered.RecordCount)
    For mCount = 1 To RsMst_InsuranceRiskCovered.RecordCount
        mInsuranceRiskCoveredId = IIf(IsNull(RsMst_InsuranceRiskCovered!InsuranceRiskCoveredID), 0, RsMst_InsuranceRiskCovered!InsuranceRiskCoveredID)
        mInsuranceRiskCovered = IIf(IsNull(RsMst_InsuranceRiskCovered!InsuranceRiskCovered), "", RsMst_InsuranceRiskCovered!InsuranceRiskCovered)
        Call CmdGo_Click
        Call CmdExcel_Click
        mExportReportList(mCount - 1) = mFilePath
        RsMst_InsuranceRiskCovered.MoveNext
    Next mCount
End If
'MsgBox "Done."
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Call Grid_Head
Call Grid_Head1
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 2
    .Cols = 17
    .FixedCols = 0
    .WordWrap = True
    .Font.Size = 10
    .TextMatrix(0, 0) = "Insurance Policy ID":                                  .ColWidth(0) = 1000
    .TextMatrix(0, 1) = "Policy No":                                            .ColWidth(1) = 2500
    .TextMatrix(0, 2) = "Covert Note No":                                       .ColWidth(2) = 2500
    .TextMatrix(0, 3) = "Insurance Company Name":                               .ColWidth(3) = 2500
    .TextMatrix(0, 4) = "Insurance Policy Covered Amount":                      .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "Structure Wise Insurance Amount Limit":                .ColWidth(5) = 2000
    .TextMatrix(0, 6) = "Valid From":                                           .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "Valid To":                                             .ColWidth(7) = 1500
    .TextMatrix(0, 8) = "Insurance By Name":                                    .ColWidth(8) = 1200
    .TextMatrix(0, 9) = "Insurance Risk Covered":                               .ColWidth(9) = 2500
    .TextMatrix(0, 10) = "Client Name":                                         .ColWidth(10) = 3500
    .TextMatrix(0, 11) = "Commodity":                                           .ColWidth(11) = 2500
    .TextMatrix(0, 12) = "Godown Type":                                         .ColWidth(12) = 2500
    .TextMatrix(0, 13) = "Value of Commodity Covered":                          .ColWidth(13) = 2000
    .TextMatrix(0, 14) = "Remaining Insurance Amount":                          .ColWidth(14) = 2000
    .TextMatrix(0, 15) = "Storage Insurance Amount":                            .ColWidth(15) = 2000
    .TextMatrix(0, 16) = "InsuranceByID":                                       .ColWidth(16) = 0
    .RowHeight(0) = 600
End With
End Sub

Function MappingDataforInsurance(InsurancePolicyID_ As Double, Type_ As Variant) As Variant
Dim mResult As Variant
Dim mR As Double
If InsurancePolicyID_ > 0 Then
    mResult = ""
    If LCase(Type_) = LCase("ClientName") Then
        tmp = " Select MC.ClientIDRow,MC.ClientName,MC.ClientID from Mst_InsuranceClientDetail MICD Inner Join Mst_Client MC On MICD.ClientID = MC.ClientID Where InsurancePolicyID = " & InsurancePolicyID_ & " "
        Call Tmp1Table
        If TmpRs1.RecordCount > 0 Then
            For mR = 1 To TmpRs1.RecordCount
              mResult = IIf(mResult = "", "", mResult & ", ") & IIf(IsNull(TmpRs1!ClientName), "", TmpRs1!ClientName) & "#" & IIf(IsNull(TmpRs1!ClientIDRow), "", TmpRs1!ClientIDRow) & "#" & IIf(IsNull(TmpRs1!ClientID), "", TmpRs1!ClientID)
              TmpRs1.MoveNext
            Next mR
        End If
    ElseIf LCase(Type_) = LCase("Commodity") Then
        tmp = " Select MC.Commodity,MICD.EntryType from Mst_InsuranceCommodityDetail MICD Inner Join Mst_Commodity MC On MICD.CommodityID = MC.CommodityID Where InsurancePolicyID = " & InsurancePolicyID_ & " "
        Call Tmp1Table
        If TmpRs1.RecordCount > 0 Then
            If LCase(TmpRs1!EntryType) = LCase("I") Then
                For mR = 1 To TmpRs1.RecordCount
                  mResult = IIf(mResult = "", "", mResult & ", ") & IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
                  TmpRs1.MoveNext
                Next mR
                mResult = "(Include) " & mResult
            ElseIf LCase(TmpRs1!EntryType) = LCase("E") Then
                For mR = 1 To TmpRs1.RecordCount
                  mResult = IIf(mResult = "", "", mResult & ", ") & IIf(IsNull(TmpRs1!Commodity), "", TmpRs1!Commodity)
                  TmpRs1.MoveNext
                Next mR
                mResult = "(Exclude) " & mResult
            End If
        Else
            mResult = "(Include all Commodity)"
        End If
    ElseIf LCase(Type_) = LCase("GodownType") Then
        tmp = " Select MTY.UnitType From Mst_InsuranceForUnitTypeDetail MIUTD Inner Join Mst_UnitType MTY On MIUTD.UnitTypeID = MTY.UnitTypeID Where MIUTD.InsurancePolicyID = " & InsurancePolicyID_ & ""
        Call Tmp1Table
        If TmpRs1.RecordCount > 0 Then
            For mR = 1 To TmpRs1.RecordCount
              mResult = IIf(mResult = "", "", mResult & ", ") & IIf(IsNull(TmpRs1!UnitType), "", TmpRs1!UnitType)
              TmpRs1.MoveNext
            Next mR
        End If
    End If
End If
MappingDataforInsurance = mResult
End Function

Private Sub Grid_Head1()
With MSHFlexGrid2
    
    .Clear
    .Rows = 2
    .Cols = 23
    .FixedCols = 0
    .WordWrap = True
    .Font.Size = 10
    .TextMatrix(0, 0) = "Client Name":                          .ColWidth(0) = 1500
    .TextMatrix(0, 1) = "Exchange":                             .ColWidth(1) = 1500
    .TextMatrix(0, 2) = "Commodity":                            .ColWidth(2) = 1500
    .TextMatrix(0, 3) = "Godown No":                            .ColWidth(3) = 1200
    .TextMatrix(0, 4) = "Godown Type":                          .ColWidth(4) = 1200
    .TextMatrix(0, 5) = "CMP Name":                             .ColWidth(5) = 1500
    .TextMatrix(0, 6) = "Location Name":                        .ColWidth(6) = 1500
    .TextMatrix(0, 7) = "SM-Prefix":                            .ColWidth(7) = 1000
    .TextMatrix(0, 8) = "State Name":                           .ColWidth(8) = 1800
    .TextMatrix(0, 9) = "State India Name":                     .ColWidth(9) = 1800
    .TextMatrix(0, 10) = "Region":                              .ColWidth(10) = 1800
    .TextMatrix(0, 11) = "Balance Bags":                        .ColWidth(11) = 1500
    .TextMatrix(0, 12) = "Balance Weight":                      .ColWidth(12) = 1500
    .TextMatrix(0, 13) = "Rate":                                .ColWidth(13) = 1500
    .TextMatrix(0, 14) = "AUM":                                 .ColWidth(14) = 1500
    .TextMatrix(0, 15) = "Covered Amount":                      .ColWidth(15) = 1500
    .TextMatrix(0, 16) = "Shortage Amount":                     .ColWidth(16) = 1500
    .TextMatrix(0, 17) = "Covered Under Policy Nos":            .ColWidth(17) = 3000
    .TextMatrix(0, 18) = "Covered Under Covert Note Nos":       .ColWidth(18) = 3000
    .TextMatrix(0, 19) = "Godown Endorse to Policy Nos":        .ColWidth(19) = 3000
    .TextMatrix(0, 20) = "Insurance By Name":                   .ColWidth(20) = 1500
    .TextMatrix(0, 21) = "Rate Year-Month":                     .ColWidth(21) = 1500
    .TextMatrix(0, 22) = "Godown Area (Square Feet)":           .ColWidth(22) = 1500
    .RowHeight(0) = 600
    
End With
End Sub

Function PolicyIDAvailable(Id_ As Double) As Boolean
Dim mR As Double
Dim mResult As Boolean
mResult = False
With MSHFlexGrid1
    For mR = 1 To .Rows - 2
        If Val(.TextMatrix(mR, 0)) = Id_ Then
            mResult = True
            Exit For
        End If
    Next mR
End With
PolicyIDAvailable = mResult
End Function

Function ListOfInsurance(Id_ As Variant, Action_ As Variant) As Variant
Dim mSplit As Variant
Dim mResult As Variant
Dim mR As Double
mSplit = Split(Id_, ",")
mResult = ""
If UBound(mSplit) > 0 Then
    For mR = 0 To UBound(mSplit)
       If Trim(mSplit(mR)) <> "" Then
            RsMst_PolicyDetail.Filter = ""
            If RsMst_PolicyDetail.RecordCount > 0 Then
                RsMst_PolicyDetail.Filter = "InsurancePolicyID = " & mSplit(mR)
                If RsMst_PolicyDetail.EOF = False Then
                    If LCase(Action_) = LCase("PolicyNo") Then
                        'mResult = mResult & IIf(IsNull(RsMst_PolicyDetail!PolicyNo), "", RsMst_PolicyDetail!PolicyNo) & "#" & IIf(IsNull(RsMst_PolicyDetail!CoverNoteNo), "", RsMst_PolicyDetail!CoverNoteNo) & ","
                        mResult = mResult & IIf(IsNull(RsMst_PolicyDetail!PolicyNo), "", RsMst_PolicyDetail!PolicyNo) & ","
                    ElseIf LCase(Action_) = LCase("CoverNoteNo") Then
                        mResult = mResult & IIf(IsNull(RsMst_PolicyDetail!CoverNoteNo), "", RsMst_PolicyDetail!CoverNoteNo) & ","
                    ElseIf LCase(Action_) = LCase("InsuranceBy") Then
                        RsMst_InsuranceBy.Filter = ""
                        If RsMst_InsuranceBy.RecordCount > 0 Then
                            RsMst_InsuranceBy.Filter = "InsuranceByID = " & RsMst_PolicyDetail!InsuranceByID
                            If RsMst_InsuranceBy.EOF = False Then
                                If mResult = "" Then
                                    mResult = mResult & IIf(IsNull(RsMst_InsuranceBy!InsuranceByName), "", RsMst_InsuranceBy!InsuranceByName) & ","
                                Else
                                    If InStr(IIf(IsNull(RsMst_InsuranceBy!InsuranceByName), "", RsMst_InsuranceBy!InsuranceByName) & ",", mResult) = 0 Then
                                        mResult = mResult & IIf(IsNull(RsMst_InsuranceBy!InsuranceByName), "", RsMst_InsuranceBy!InsuranceByName) & ","
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
       End If
    Next mR
Else
    mResult = ""
End If
If mResult <> "" Then
    mResult = Left(mResult, Len(mResult) - 1)
End If
ListOfInsurance = mResult
End Function

Function ColorCell(Grid As MSHFlexGrid, Index_ As Double, Color_ As ColorConstants, Type_ As Variant, ColIndex_ As Double)
Dim mC As Double
With Grid
    If LCase(Type_) = "row" Then
    For mC = 0 To .Cols - 1
        .Col = mC: .row = Index_: .CellBackColor = Color_
    Next mC
    ElseIf LCase(Type_) = "column" Then
        .Col = ColIndex_: .row = Index_: .CellBackColor = Color_
    End If
End With
End Function

Private Sub Gen_Email()
Dim TOADD As Variant
Dim TOCC, mMsg As Variant
Dim mR As Variant

mMsg = ""
TOADD = ""
TOCC = ""
mCMPMEmailList = ""

mMsg = "Please find the attachment for Insurance MIS :-" & Format(Date, "dd-MMM-yyyy") & "."

''mMsg = "<table border=1 width='80%' align='center'>"
''mMsg = mMsg & "<tr><td width='80%'><b>Agreement ID</b></td><td width='60%'>" & TxtAgreementID.Text & "</td></tr>"
''mMsg = mMsg & "</table>"

''tmp = " Select EmployeeNo,EmployeeName,EmployeeEmailID from Mst_Employee Where EmployeeNo In (" & mListOfEmailEmployeeNo & ") ANd Isnull(Flag,'Y') = 'Y' "
''Call TmpTable
''If TmpRs.RecordCount > 0 Then
''   For mR = 1 To TmpRs.RecordCount
''    TOADD = TOADD & "," & TmpRs!EmployeeEmailID
''    TmpRs.MoveNext
''   Next mR
''   TOADD = Right(TOADD, Len(TOADD) - 1)
''End If
''
''If mForwardID <> Gen_UserID Then
''End If

TOADD = "cmg@ncml.com;accounts@ncml.com"
TOCC = Gen_UserEmailID & ",manu.ag@ncml.com,ullhas.p@ncml.com,parag.r@ncml.com"
''""  ''"ullhas.p@ncml.com" ''
mMsgSubject = "Insurance MIS"
Call SendMail(TOADD, TOCC, mMsgSubject, mMsg, "N")
End Sub

Private Function SendMail(ToADD_ As Variant, TOCC_ As Variant, MsgSubject_ As Variant, Message1_ As Variant, mMailTypeFlg As Variant)
Dim htmlStringhead As String
Dim htmlStringtail As String
Set oSMTP = New OSSMTP.SMTPSession
servertxt = "webmail.ncml.com"
htmlStringhead = "<html>"
htmlStringhead = htmlStringhead & " <head>"
htmlStringhead = htmlStringhead & "<style> body{ color: #000000;margin: 5px;font-family: Candara;} table, th, td { border: 1px solid black; border-collapse: collapse;} th { padding: 5px; text-align: center;} td { padding: 5px; text-align: left;} </style>"
htmlStringhead = htmlStringhead & "</head><body>"
htmlStringtail = "</body></html>"
   
''With oSMTP
''  'connection
''  .Server = servertxt
''  If LCase(mMailTypeFlg) = "n" Then
''     .MailFrom = "info@ncmsl.com"
''  Else
''     .MailFrom = Gen_UserEmailID
''     '.MailFrom = "ullhas.p@ncml.com"
''  End If
'''    .SendTo = "ullhas.p@ncml.com"
'''    .CC = Gen_UserEmailID
''  .SendTo = ToADD_
''  If ToADD_ <> TOCC_ Then
''     .CC = TOCC_
''  End If
''  If LCase(mMailTypeFlg) = "n" Then
''     '.Attachments.Add (mPDFFileName)
''  End If
''  .MessageSubject = MsgSubject_ & " (from Store-Man)" '"Confidential Mail"
''  .MessageHTML = htmlStringhead & Message1_ & htmlStringtail
''  .TimeStamp = Now()
''  If LCase(Gen_EmailConfigFlag) = "y" Then
''     .SendEmail
''  End If
''End With

Dim objSMTP
Set objSMTP = CreateObject("MailBee.SMTP")
objSMTP.LicenseKey = "MBC900-5F470FAD67-D8E202A2E22927335F24F5A667BFD196"
objSMTP.ServerName = Gen_ReadMailServerIPINI("MailIP") ' = "mail.ncml.com" '"mail.ncml.com"
'If LCase(mMailTypeFlg) = "n" Then
'   objSMTP.FromAddr = "info@ncmsl.com"
'Else
   objSMTP.FromAddr = Gen_UserEmailID
'End If
If IsNull(Trim(ToADD_)) = True Or Trim(ToADD_) = "" Then
   ToADD_ = "manu.ag@ncml.com"
End If
objSMTP.ToAddr = ToADD_
If ToADD_ <> TOCC_ Then
   objSMTP.CCAddr = TOCC_
End If
If LCase(mMailTypeFlg) = "n" Then
   objSMTP.AddAttachment mFilePath
End If
objSMTP.Subject = MsgSubject_ & " (from Store-Man)" 'MsgSubject_ '"Message in HTML format"
objSMTP.BodyFormat = 1
objSMTP.BodyText = htmlStringhead & Message1_ & htmlStringtail 'htmlStringhead & MsgSubject_ & " :- " & Message1_ & htmlStringtail '"<html>This word is <b>bold</b></html>"
If LCase(Gen_EmailConfigFlag) = "y" Then
   objSMTP.Send
End If
End Function


''-- Ullhas Patil 17 July 2018
Function PolicyDayPremium(mPolicyID As Double) As Double
Dim mNoofDays As Double
Dim mPremium As Double
tmp = " Select B.InsurancePolicyID,A.PremiumMonth,A.TotalMonthlyPremium,C.ValidFrom,C.ValidTo from Txn_InsurancePremiumDetail A"
tmp = tmp & " Inner Join Txn_InsurancePremium B On A.InsurancePremiumID = B.InsurancePremiumID"
tmp = tmp & " Inner Join Mst_InsuranceDetails C On B.InsurancePolicyID = C.InsurancePolicyID"
tmp = tmp & " Where B.InsurancePolicyID = " & mPolicyID & " And A.PremiumMonth = '" & Format(Date, "yyyy-mm-01") & "' "
Call TmpTable
mPremium = 0
If TmpRs.RecordCount = 1 Then
    mNoofDays = Format(DateDiff("d", Format(Date, "yyyy-mm-01"), Format(DateAdd("m", 1, Date), "yyyy-mm-01")), "#0")
    mPremium = Format(TmpRs!TotalMonthlyPremium / mNoofDays, "#0.00")
End If
PolicyDayPremium = mPremium
End Function
