VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmRpt_WHRMonitor 
   Caption         =   "WHR Monitor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "Location Wise Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   11
      Top             =   5355
      Width           =   15000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4215
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   7435
         _Version        =   393216
         WordWrap        =   -1  'True
         MergeCells      =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CMP Wise Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   15000
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   3855
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   6800
         _Version        =   393216
         WordWrap        =   -1  'True
         MergeCells      =   1
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbState 
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   3120
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
         Left            =   5325
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   210
         Width           =   3240
      End
      Begin VB.ComboBox CmbCmpName 
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
         Left            =   1080
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   7485
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
         Height          =   350
         Left            =   9720
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdSearch 
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
         Height          =   350
         Left            =   8625
         TabIndex        =   4
         Top             =   240
         Width           =   1095
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
         Height          =   350
         Left            =   10695
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "State"
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
         Left            =   120
         TabIndex        =   15
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label8 
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
         Left            =   4440
         TabIndex        =   14
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   660
         Width           =   435
      End
      Begin VB.Label LblTotalRecord 
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
         Left            =   14160
         TabIndex        =   8
         Top             =   270
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label Label4 
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
         Left            =   11850
         TabIndex        =   7
         Top             =   270
         Visible         =   0   'False
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmRpt_WHRMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCMPID As Variant
Dim mStateID, mLocationID As Variant
Private Sub CmbCMPName_GotFocus()
If CmbLocationID.Text = "" Then
   CmbCMPName.Clear
   MsgBox "Please select location!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

tmp = CmbCMPName.Text
wc = ""
If Gen_WcCMP <> "" Then
   wc = "And " & Gen_WcCMP
End If

Call TableMst_CMP(" where LocationID=" & mLocationID & " " & tmp1)

CmbCMPName.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPName.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPName.Text = tmp
End Sub
'----
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
'TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub
Private Sub CmbLocationID_GotFocus()
If CmbState.Text = "" Then
   MsgBox "Select State!!!"
   CmbState.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
CmbCMPName.Text = ""
tmp1 = ""
If Gen_WcLocation <> "" Then
   tmp1 = "And " & Gen_WcLocation
End If
Call TableMst_Location(" Where StateID = " & mStateID & " And LocationID <> 0 " & tmp1)
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record Found!"
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
   CmbCMPName.Text = ""
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If RsMst_Location!LocationName <> CmbLocationID.Text Then
   CmbCMPName.Text = ""
End If

mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmbState_GotFocus()
tmp = CmbState.Text
CmbLocationID.Text = ""
CmbCMPName.Text = ""
tmp1 = ""
If Gen_WcState <> "" Then
   tmp1 = "And " & Gen_WcState
End If
Call TableMst_State(" Where StateID <> 0 " & tmp1)
CmbState.Clear
If RsMst_State.RecordCount = 0 Then
   MsgBox "No Record Found!"
   Exit Sub
End If
For I = 1 To RsMst_State.RecordCount
    CmbState.AddItem RsMst_State!StateName
    RsMst_State.MoveNext
Next
CmbState.Text = tmp

End Sub

Private Sub CmbState_LostFocus()
If CmbState.Text = "" Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbState.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbState.SetFocus
   Exit Sub
End If
If RsMst_State!StateName <> CmbState.Text Then
   CmbLocationID.Text = ""
   CmbCMPName.Text = ""
End If

mStateID = RsMst_State!StateID
End Sub

Private Sub CmdClear_Click()
Call Grid_Head
Call Grid_Head1
End Sub

Private Sub CmdSearch_Click()
Dim wc As String
Call Grid_Head
Call Grid_Head1
LblTotalRecord.Caption = ""
MsgBox ("Wait till Next Message")

'If CmbSCmp.Text <> "" Then
'   wc = GenWc(wc, mCMPName, "Mst_CMP.Name", "S")
'End If

Call GetCMPWiseRpt

MsgBox ("CMP Wise Report completed. Processing for Location wise report, please wait!!!")

Call GetLocationWiseRpt

'LblTotalRecord.Caption = TmpRs.RecordCount
MsgBox ("Done")



End Sub
Private Sub GetCMPWiseRpt()
Dim wc1 As Variant
tmp = ""
If CmbLocationID.Text <> "" Then
   wc1 = " And mst_location.LocationName='" & CmbLocationID.Text & "'"
ElseIf Gen_WcLocation <> "" Then
   wc1 = " And mst_location." & Gen_WcLocation
End If

If CmbState.Text <> "" Then
   wc1 = " And mst_state.statename='" & CmbState.Text & "'"
ElseIf Gen_WcState <> "" Then
   wc1 = " And mst_state." & Gen_WcState
End If

If CmbCMPName.Text <> "" Then
   wc1 = wc1 & " And Mst_CMP.CMPName='" & CmbCMPName.Text & "' "
Else
   If Gen_WcCMP <> "" Then
      wc1 = wc1 & " And Mst_CMP." & Gen_WcCMP
   End If
End If
tmp = "Select mst_state.statename, mst_location.LocationName,Mst_CMP.CMPName, Mst_Commodity.Commodity , Txn_CxTf_Gsl.ClientIDRow,Mst_Client.ClientName, " & _
      " Sum(Txn_CxTf_Gsl.NoofBags)  'NoOfBags',sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight'  " & _
      " From Txn_CxTf_Details Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And  " & _
      " Txn_CxTf_Details.CxTFNo = Txn_CxTf_Gsl.CxTFNo And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType  " & _
      " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID  " & _
      " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID  " & _
      " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID  " & _
      " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID  " & _
      " left Join Mst_Client On Txn_CxTf_Gsl.ClientIDRow = Mst_Client.ClientIDRow  " & _
      " left Join mst_location On Txn_CxTf_Details.SM_Prefix = mst_location.SM_Prefix  " & _
      " left Join mst_state On mst_location.stateid = mst_state.stateid " & _
      " Where Txn_CxTf_Details.TxnType = 'D'  And Mst_ExchangeType.ExchangeTypeID= 1  " & _
      wc1 & _
      " group by Mst_CMP.CMPName, Mst_Commodity.Commodity,Txn_CxTf_Gsl.ClientIDRow,Mst_Client.ClientName,mst_location.LocationName, mst_state.statename  " & _
      " ORDER BY mst_state.statename, mst_location.locationname, Mst_Client.ClientName "

Call TmpTable

MSHFlexGrid1.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid1.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    TmpRs.MoveNext
Next


Call DispCMPWHRDetails

End Sub
Private Sub DispCMPWHRDetails()
Dim wc1 As Variant
tmp = ""

If CmbState.Text <> "" Then
   wc1 = " Where mst_state.stateName='" & CmbState.Text & "'"
Else
   If Gen_WcState <> "" Then
      wc1 = " Where mst_state." & Gen_WcState
   End If
End If

If CmbLocationID.Text <> "" Then
   If wc1 = "" Then
      wc1 = " Where mst_location.LocationName='" & CmbLocationID.Text & "'"
   Else
      wc1 = wc1 & " and mst_location.LocationName='" & CmbLocationID.Text & "'"
   End If
Else
   If Gen_WcLocation <> "" Then
      If wc1 = "" Then
         wc1 = " Where mst_location." & Gen_WcLocation
      Else
         wc1 = wc & " and mst_location." & Gen_WcLocation
      End If
   End If
End If
If CmbCMPName.Text <> "" Then
   wc1 = wc1 & " And CMP.CMPName='" & CmbCMPName.Text & "' "
Else
   
   If Gen_WcCMP <> "" Then
      If wc1 <> "" Then
         wc1 = wc1 & " And CMP." & Gen_WcCMP
      Else
         wc1 = " Where CMP." & Gen_WcCMP
      End If
   End If
End If

tmp = "SELECT  mst_state.statename,mst_location.LocationName,CMP.CMPName,ComM.Commodity,TWD.ClientIDRow,Mst_Client.ClientName,sum(TWD.WHRNoofBags) 'No of Bags',Sum(TWD.WHRQuantity) 'WHR Qty', " & _
      " Sum(TWD.DeliveredNoOfBags)  'Del Bags',sum(DeliveredQuantity) 'Del Qty'  " & _
      " FROM    Txn_WHR_Detail TWD " & _
      " Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID  " & _
      " Inner Join Mst_Client CM ON TWD.ClientIDRow=CM.ClientIDRow  " & _
      " Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID  " & _
      " Left Join Mst_Client On TWD.ClientIDRow = Mst_Client.ClientIDRow  " & _
      " inner Join mst_location On TWD.SM_Prefix = mst_location.SM_Prefix  " & _
      " inner Join mst_state On mst_state.stateid = mst_location.stateid " & _
      wc1 & _
      " Group by mst_state.statename,mst_location.LocationName,TWD.ClientIDRow,CMP.CMPName,ComM.Commodity,Mst_Client.ClientName " & _
      " Order By mst_state.statename,mst_location.LocationName, Mst_Client.ClientName "
Call TmpTable
Dim flg As Boolean

For I = 1 To TmpRs.RecordCount
    flg = False
    iu = 0
    For x = 1 To MSHFlexGrid1.Rows - 1
        If MSHFlexGrid1.TextMatrix(x, 0) = TmpRs.Fields(0) And MSHFlexGrid1.TextMatrix(x, 1) = TmpRs.Fields(1) And MSHFlexGrid1.TextMatrix(x, 2) = TmpRs.Fields(2) And MSHFlexGrid1.TextMatrix(x, 3) = TmpRs.Fields(3) And MSHFlexGrid1.TextMatrix(x, 4) = TmpRs.Fields(4) And MSHFlexGrid1.TextMatrix(x, 5) = TmpRs.Fields(5) Then '&
           For c = 8 To 9
               MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c - 2)), 0, TmpRs.Fields(c - 2))
           Next
           flg = True
           Exit For
        End If
        MSHFlexGrid1.TextMatrix(x, 10) = Val(MSHFlexGrid1.TextMatrix(x, 6)) - Val(MSHFlexGrid1.TextMatrix(x, 8))
        MSHFlexGrid1.TextMatrix(x, 11) = Val(MSHFlexGrid1.TextMatrix(x, 7)) - Val(MSHFlexGrid1.TextMatrix(x, 9))
        If Val(MSHFlexGrid1.TextMatrix(x, 6)) = Val(MSHFlexGrid1.TextMatrix(x, 10)) Then
           MSHFlexGrid1.TextMatrix(x, 8) = "0"
           MSHFlexGrid1.TextMatrix(x, 9) = "0"
        End If
        iu = x
    Next
        
    If flg = False Then
       MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
       For c = 0 To 5
           MSHFlexGrid1.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid1.TextMatrix(iu, 8) = IIf(IsNull(TmpRs.Fields(6)), 0, TmpRs.Fields(6))
       MSHFlexGrid1.TextMatrix(iu, 9) = IIf(IsNull(TmpRs.Fields(7)), 0, TmpRs.Fields(7))
       MSHFlexGrid1.TextMatrix(iu, 10) = Val(MSHFlexGrid1.TextMatrix(iu, 6)) - Val(MSHFlexGrid1.TextMatrix(iu, 8))
       MSHFlexGrid1.TextMatrix(iu, 11) = Val(MSHFlexGrid1.TextMatrix(iu, 7)) - Val(MSHFlexGrid1.TextMatrix(iu, 9))
    End If
    TmpRs.MoveNext
Next


End Sub




Private Sub GetLocationWiseRpt()

tmp = ""
Dim wc1 As Variant
tmp = ""
If CmbState.Text <> "" Then
   wc1 = " And mst_state.stateName='" & CmbState.Text & "'"
Else
   If Gen_WcState <> "" Then
      wc1 = " And mst_state." & Gen_WcState
   End If
End If

If CmbLocationID.Text <> "" Then
   wc1 = " And mst_location.LocationName='" & CmbLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      wc1 = " And mst_location." & Gen_WcLocation
   End If
End If

tmp = "Select mst_state.statename,mst_location.LocationName, Mst_Commodity.Commodity , Txn_CxTf_Gsl.ClientIDRow,Mst_Client.ClientName, " & _
      " Sum(Txn_CxTf_Gsl.NoofBags)  'NoOfBags',sum(Txn_CxTf_Gsl.GrossWeight) 'GrossWeight'  " & _
      " From Txn_CxTf_Details Inner Join Txn_CxTf_Gsl  on Txn_CxTf_Details.SM_Prefix = Txn_CxTf_Gsl.SM_Prefix And  " & _
      " Txn_CxTf_Details.CxTFNo = Txn_CxTf_Gsl.CxTFNo And Txn_CxTf_Details.TxnType = Txn_CxTf_Gsl.TxnType  " & _
      " Inner Join Mst_CMP On Txn_CxTf_Details.CMPID = Mst_CMP.CMPID  " & _
      " Inner Join Mst_Commodity On Txn_CxTf_Details.CommodityID = Mst_Commodity.CommodityID  " & _
      " Inner Join Mst_ExchangeType ON Txn_CxTf_Details.ExchangeTypeID = Mst_ExchangeType.ExchangeTypeID  " & _
      " Left Join Mst_Agent ON Txn_CxTf_Details.AgentID = Mst_Agent.AgentID  " & _
      " left Join Mst_Client On Txn_CxTf_Gsl.ClientIDRow = Mst_Client.ClientIDRow  " & _
      " left Join mst_location On Txn_CxTf_Details.SM_Prefix = mst_location.SM_Prefix  " & _
      " left Join mst_state On mst_state.stateid = mst_location.stateid" & _
      " Where Txn_CxTf_Details.TxnType = 'D'  And Mst_ExchangeType.ExchangeTypeID= 1  " & _
      wc1 & _
      " group by mst_state.statename,Mst_Commodity.Commodity,Txn_CxTf_Gsl.ClientIDRow,Mst_Client.ClientName,mst_location.LocationName  " & _
      " ORDER BY mst_state.statename,mst_location.locationname, Mst_Client.ClientName "

Call TmpTable

MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
For I = 1 To TmpRs.RecordCount
    For c = 0 To TmpRs.Fields.Count - 1
        MSHFlexGrid2.TextMatrix(I, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
    Next
    TmpRs.MoveNext
Next


Call DispLocationWHRDetails

End Sub
Private Sub DispLocationWHRDetails()
Dim wc1 As Variant
tmp = ""
If CmbState.Text <> "" Then
   wc1 = " Where mst_state.stateName='" & CmbState.Text & "'"
Else
   If Gen_WcState <> "" Then
      wc1 = " Where mst_state." & Gen_WcState
   End If
End If
If CmbLocationID.Text <> "" Then
   If wc1 = "" Then
      wc1 = " Where mst_location.LocationName='" & CmbLocationID.Text & "'"
   Else
      wc1 = wc1 & " and mst_location.LocationName='" & CmbLocationID.Text & "'"
   End If
Else
   If Gen_WcLocation <> "" Then
      If wc1 = "" Then
         wc1 = " Where mst_location." & Gen_WcLocation
      Else
         wc1 = wc1 & " and mst_location." & Gen_WcLocation
      End If
   End If
End If
tmp = "SELECT  mst_state.statename,mst_location.LocationName,ComM.Commodity,TWD.ClientIDRow,Mst_Client.ClientName,sum(TWD.WHRNoofBags) 'No of Bags',Sum(TWD.WHRQuantity) 'WHR Qty', " & _
      " Sum(TWD.DeliveredNoOfBags)  'Del Bags',sum(DeliveredQuantity) 'Del Qty'  " & _
      " FROM    Txn_WHR_Detail TWD " & _
      " Inner Join Mst_CMP CMP ON TWD.CMPID=CMP.CMPID Inner Join Mst_Godown GM ON TWD.GodownID=GM.GodownID  " & _
      " Inner Join Mst_Client CM ON TWD.ClientIDRow=CM.ClientIDRow  " & _
      " Inner Join Mst_Commodity ComM ON TWD.CommodityID=ComM.CommodityID  " & _
      " Left Join Mst_Client On TWD.ClientIDRow = Mst_Client.ClientIDRow  " & _
      " inner Join mst_location On TWD.SM_Prefix = mst_location.SM_Prefix  " & _
      " inner Join mst_state On mst_state.stateid = mst_location.stateid " & _
      wc1 & _
      " Group by mst_state.statename,mst_location.LocationName,TWD.ClientIDRow,ComM.Commodity,Mst_Client.ClientName " & _
      " Order By mst_state.statename,mst_location.LocationName,Mst_Client.ClientName "
Call TmpTable
Dim flg As Boolean

For I = 1 To TmpRs.RecordCount
    flg = False
    iu = 0
    For x = 1 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(x, 0) = TmpRs.Fields(0) And MSHFlexGrid2.TextMatrix(x, 1) = TmpRs.Fields(1) And MSHFlexGrid2.TextMatrix(x, 2) = TmpRs.Fields(2) And MSHFlexGrid2.TextMatrix(x, 3) = TmpRs.Fields(3) And MSHFlexGrid2.TextMatrix(x, 4) = TmpRs.Fields(4) Then '&
           For c = 7 To 8
               MSHFlexGrid2.TextMatrix(x, c) = IIf(IsNull(TmpRs.Fields(c - 2)), 0, TmpRs.Fields(c - 2))
           Next
           flg = True
           Exit For
        End If
        MSHFlexGrid2.TextMatrix(x, 9) = Val(MSHFlexGrid2.TextMatrix(x, 5)) - Val(MSHFlexGrid2.TextMatrix(x, 7))
        MSHFlexGrid2.TextMatrix(x, 10) = Val(MSHFlexGrid2.TextMatrix(x, 6)) - Val(MSHFlexGrid2.TextMatrix(x, 8))
        If Val(MSHFlexGrid2.TextMatrix(x, 5)) = Val(MSHFlexGrid2.TextMatrix(x, 9)) Then
           MSHFlexGrid2.TextMatrix(x, 7) = "0"
           MSHFlexGrid2.TextMatrix(x, 8) = "0"
        End If
        iu = x
    Next
        
    If flg = False Then
       MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
       For c = 0 To 5
           MSHFlexGrid2.TextMatrix(iu, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid2.TextMatrix(iu, 8) = IIf(IsNull(TmpRs.Fields(6)), 0, TmpRs.Fields(6))
       MSHFlexGrid2.TextMatrix(iu, 9) = IIf(IsNull(TmpRs.Fields(7)), 0, TmpRs.Fields(7))
       MSHFlexGrid2.TextMatrix(iu, 10) = Val(MSHFlexGrid2.TextMatrix(iu, 6)) - Val(MSHFlexGrid2.TextMatrix(iu, 8))
       MSHFlexGrid2.TextMatrix(iu, 11) = Val(MSHFlexGrid2.TextMatrix(iu, 7)) - Val(MSHFlexGrid2.TextMatrix(iu, 9))
    End If
    TmpRs.MoveNext
Next

End Sub




Function GenWc(wc_ As String, ID As Variant, Field_ As String, Type_ As String) As String
Dim Retval
Retval = wc_
If Retval = "" Then
   If Type_ = "S" Then
      Retval = "And " & Field_ & "= '" & ID & "'"
   Else
      Retval = "And " & Field_ & "= " & ID
   End If
Else
   If Type_ = "S" Then
      Retval = Retval & " And " & Field_ & "= '" & ID & "'"
   Else
      Retval = Retval & " And " & Field_ & "= " & ID
   End If
End If
GenWc = Retval
End Function

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Frame2.Enabled = False
   Frame3.Enabled = False
   Call CmdClear_Click
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If
Frame1.Enabled = True
Frame2.Enabled = True
Frame3.Enabled = True

'Call ClearFrame
Call CmdClear_Click
'Call Grid_Head
'Call Grid_Head1
End Sub

Private Sub Grid_Head()

MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 12


MSHFlexGrid1.TextMatrix(0, 0) = "State Name" '"SM_Prefix"
MSHFlexGrid1.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP Name"
MSHFlexGrid1.TextMatrix(0, 3) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 4) = "ClientIDRow"
MSHFlexGrid1.TextMatrix(0, 5) = "Client Name"
MSHFlexGrid1.TextMatrix(0, 6) = "No Of Bags"
MSHFlexGrid1.TextMatrix(0, 7) = "Gross Weight"
MSHFlexGrid1.TextMatrix(0, 8) = "WHR No Of Bags"
MSHFlexGrid1.TextMatrix(0, 9) = "WHR Qty"
MSHFlexGrid1.TextMatrix(0, 10) = "Balance No Of Bags"
MSHFlexGrid1.TextMatrix(0, 11) = "Balance Qty"



MSHFlexGrid1.ColWidth(0) = 1900 '1000
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 2000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 0
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1600
MSHFlexGrid1.ColWidth(8) = 1600
MSHFlexGrid1.ColWidth(9) = 1600
MSHFlexGrid1.ColWidth(10) = 1800
MSHFlexGrid1.ColWidth(11) = 1800


MSHFlexGrid1.RowHeight(0) = 400


End Sub

Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.Cols = 11


MSHFlexGrid2.TextMatrix(0, 0) = "State Name"
MSHFlexGrid2.TextMatrix(0, 1) = "Location Name"
MSHFlexGrid2.TextMatrix(0, 2) = "Commodity"
MSHFlexGrid2.TextMatrix(0, 3) = "ClientIDRow"
MSHFlexGrid2.TextMatrix(0, 4) = "Client Name"
MSHFlexGrid2.TextMatrix(0, 5) = "No Of Bags"
MSHFlexGrid2.TextMatrix(0, 6) = "Gross Weight"
MSHFlexGrid2.TextMatrix(0, 7) = "WHR No Of Bags"
MSHFlexGrid2.TextMatrix(0, 8) = "WHR Qty"
MSHFlexGrid2.TextMatrix(0, 9) = "Balance No Of Bags"
MSHFlexGrid2.TextMatrix(0, 10) = "Balance Qty"



MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 0
MSHFlexGrid2.ColWidth(4) = 2000
MSHFlexGrid2.ColWidth(5) = 1600
MSHFlexGrid2.ColWidth(6) = 1600
MSHFlexGrid2.ColWidth(7) = 1600
MSHFlexGrid2.ColWidth(8) = 1600
MSHFlexGrid2.ColWidth(9) = 1600
MSHFlexGrid2.ColWidth(10) = 1800



MSHFlexGrid2.RowHeight(0) = 400


End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_WHRMonitor.xls")
Call Xls_Detail_Rpt
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_WHRMonitor.xls")
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


Set oWS = oWB.Worksheets("Sheet2")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = I + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        If InStr(MSHFlexGrid2.TextMatrix(I, c), "/") > 0 Then
           If IsDate(MSHFlexGrid2.TextMatrix(I, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid2.TextMatrix(I, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
        End If
    Next
Next


oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

