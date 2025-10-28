VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmRpt_WHRvsActualStock 
   Caption         =   "WHR vs Actual Stock"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1395
      Left            =   105
      TabIndex        =   7
      Top             =   0
      Width           =   15000
      Begin VB.ComboBox CmbSCommodity 
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
         Left            =   7755
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   345
         Width           =   4950
      End
      Begin VB.ComboBox CmbSClientName 
         Height          =   315
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   1020
         Width           =   7335
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
         Left            =   3675
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   4080
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
         Left            =   75
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   3555
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
         Left            =   12675
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CmdGetReport 
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
         Left            =   11580
         TabIndex        =   4
         Top             =   960
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
         Height          =   375
         Left            =   13770
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label30 
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
         Left            =   9668
         TabIndex        =   22
         Top             =   120
         Width           =   1125
      End
      Begin VB.Label Label28 
         Caption         =   "Client Name"
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
         Left            =   3165
         TabIndex        =   21
         Top             =   765
         Width           =   1155
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
         Left            =   5333
         TabIndex        =   11
         Top             =   105
         Width           =   765
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
         Left            =   1320
         TabIndex        =   10
         Top             =   105
         Width           =   1065
      End
      Begin VB.Label lblRecordCount 
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
         Left            =   9840
         TabIndex        =   9
         Top             =   1050
         Width           =   105
      End
      Begin VB.Label lblRec 
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
         Left            =   7440
         TabIndex        =   8
         Top             =   1050
         Width           =   2325
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Comparison of WHR vs Actual Stock "
      TabPicture(0)   =   "FrmRpt_WHRvsActualStock.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "WHR Stock"
      TabPicture(1)   =   "FrmRpt_WHRvsActualStock.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Actual Stock"
      TabPicture(2)   =   "FrmRpt_WHRvsActualStock.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "WHR Having Excess delivery"
      TabPicture(3)   =   "FrmRpt_WHRvsActualStock.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   7425
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid4 
            Height          =   7095
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12515
            _Version        =   393216
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
      End
      Begin VB.Frame Frame4 
         Height          =   7395
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
            Height          =   7095
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12515
            _Version        =   393216
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
      End
      Begin VB.Frame Frame2 
         Height          =   7395
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
            Height          =   6975
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12303
            _Version        =   393216
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
      End
      Begin VB.Frame Frame3 
         Height          =   7395
         Left            =   -74880
         TabIndex        =   13
         Top             =   360
         Width           =   14775
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
            Height          =   6975
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   14535
            _ExtentX        =   25638
            _ExtentY        =   12303
            _Version        =   393216
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
      End
   End
End
Attribute VB_Name = "FrmRpt_WHRvsActualStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mStateID, mLocationID As Variant

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
 CmbSCommodity.AddItem RsMst_Commodity!Commodity
 RsMst_Commodity.MoveNext
Next
CmbSCommodity.Text = tmp
End Sub

Private Sub CmbSCommodity_LostFocus()
If CmbSCommodity.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbSCommodity.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbSCommodity.SetFocus
   Exit Sub
End If
mSCommodityID = RsMst_Commodity!CommodityID
End Sub


Private Sub CmbStateID_GotFocus()
tmp = CmbStateID.Text
If Gen_WcState <> "" Then
   Call TableMst_State("Where" & Gen_WcState)
Else
   Call TableMst_State
End If
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
   Exit Sub
End If
RsMst_State.MoveFirst
RsMst_State.Find "StateName = '" & CmbStateID.Text & "'"
If RsMst_State.EOF Then
   MsgBox "Invalid selection "
   CmbStateID.SetFocus
   Exit Sub
End If

If tmp <> RsMst_State!StateName Then
   CmbLocationID.Text = ""
End If

mStateID = RsMst_State!StateID
End Sub
Private Sub CmbLocationID_GotFocus()
If CmbStateID.Text = "" Then
   MsgBox "Please Select State Name !!"
   CmbStateID.SetFocus
   Exit Sub
End If

tmp = CmbLocationID.Text
CmbLocationID.Clear

If Gen_WcLocation = "" Then
   Call TableMst_Location("Where StateID=" & mStateID & "")
Else
   Call TableMst_Location("Where StateID=" & mStateID & " And " & Gen_WcLocation)
End If

If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Location found"
   CmbLocationID.SetFocus
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
   mLocationID = 0
   'CmbCMPName.Text = ""
   'CmbSGodown.Text = ""
   Exit Sub
End If

RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid Location selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
'mSM_Prefix = RsMst_Location!SM_Prefix
End Sub

Private Sub ClearData()
Call Grid_Head
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
CmbStateID.Text = ""
CmbLocationID.Text = ""
CmbSCommodity.Text = ""
CmbSClientName.Text = ""
lblRecordCount.Caption = ""
End Sub

Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 3
    .FixedRows = 2
    
    .Cols = 9
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "State Name"
    .TextMatrix(1, 1) = "Location Name"
    .TextMatrix(1, 2) = "Client Name"
    .TextMatrix(1, 3) = "Actual No of Bags"
    .TextMatrix(1, 4) = "WHR No of Bags"
    .TextMatrix(1, 5) = "Difference in Bags"
    .TextMatrix(1, 6) = "Actual Weight"
    .TextMatrix(1, 7) = "WHR Weight"
    .TextMatrix(1, 8) = "Difference in Weight"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    
    .RowHeight(1) = 600
    .Font.Size = 10
End With
End Sub

Private Sub Grid_Head2()
With MSHFlexGrid2
    .Clear
    .Rows = 3
    .FixedRows = 2
    
    .Cols = 7
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "State Name"
    .TextMatrix(1, 1) = "Location Name"
    .TextMatrix(1, 2) = "Client Name"
    .TextMatrix(1, 3) = "WHR No "
    .TextMatrix(1, 4) = "Commodity"
    .TextMatrix(1, 5) = "WHR No of Bags"
    .TextMatrix(1, 6) = "WHR Weight"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1800
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    
    .RowHeight(1) = 600
    .Font.Size = 10
End With
End Sub

Private Sub Grid_Head3()
With MSHFlexGrid3
    .Clear
    .Rows = 3
    .FixedRows = 2
    
    .Cols = 9
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "State Name"
    .TextMatrix(1, 1) = "Location Name"
    .TextMatrix(1, 2) = "Client Name"
    .TextMatrix(1, 3) = "Godown No"
    .TextMatrix(1, 4) = "Stack No"
    .TextMatrix(1, 5) = "Lot No"
    .TextMatrix(1, 6) = "Commodity"
    .TextMatrix(1, 7) = "Balance No of Bags"
    .TextMatrix(1, 8) = "Balance Weight"
    
    .ColWidth(0) = 2500
    .ColWidth(1) = 2500
    .ColWidth(2) = 3500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1800
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    
    .RowHeight(1) = 600
    .Font.Size = 10
End With
End Sub

Private Sub Grid_Head4()
With MSHFlexGrid4
    .Clear
    .Rows = 3
    .FixedRows = 2
    
    .Cols = 7
    .WordWrap = True
    .AllowUserResizing = flexResizeColumns
    
    .TextMatrix(1, 0) = "State Name"
    .TextMatrix(1, 1) = "Location Name"
    .TextMatrix(1, 2) = "Client Name"
    .TextMatrix(1, 3) = "WHR No "
    .TextMatrix(1, 4) = "Commodity"
    .TextMatrix(1, 5) = "WHR No of Bags"
    .TextMatrix(1, 6) = "WHR Weight"
    
    .ColWidth(0) = 2000
    .ColWidth(1) = 2000
    .ColWidth(2) = 3500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1800
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    
    .RowHeight(1) = 600
    .Font.Size = 10
End With
End Sub


Private Sub CmdClear_Click()
Call ClearData
End Sub

Private Sub CmdGetReport_Click()

Dim wc As String

lblRecordCount.Caption = ""

Call Grid_Head
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4

MsgBox "Wait till next message!!!"

If CmbStateID.Text <> "" Then
   wc = GenWc(wc, CmbStateID.Text, "MS.StateName", "S", "N")
Else
   If Gen_WcState <> "" Then
      tmp2 = Replace(Gen_WcState, "StateID", "MS.StateID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbLocationID.Text <> "" Then
   wc = GenWc(wc, CmbLocationID.Text, "ML.LocationName", "S", "N")
Else
   If Gen_WcLocation <> "" Then
      tmp2 = Replace(Gen_WcLocation, "LocationID", "ML.LocationID")
      wc = GenWc(wc, tmp2, "", "S", "I")
   End If
End If

If CmbSClientName.Text <> "" Then
   wc = GenWc(wc, CmbSClientName.Text, "MCL.ClientName", "S", "N")
End If

If CmbSCommodity.Text <> "" Then
   wc = GenWc(wc, CmbSCommodity.Text, "MCom.Commodity", "S", "N")
End If


tmp = " Select MS.StateName, ML.LocationName, MCL.ClientName,"
tmp = tmp & " convert(varchar,isnull(TWD.CMPID,0))+'-'+convert(varchar, (TWD.WHRNo-(isnull(TWD.CMPID,0)*1000000))) as 'WHRNo',"
tmp = tmp & " MCom.Commodity, sum(TWD.WHRNoOfBags-TWD.DeliveredNoOfBags) As BalanceBags,  sum(TWD.WHRQuantity - TWD.DeliveredQuantity) As BalanceWeight"
tmp = tmp & " From Txn_WHR_Detail TWD"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TWD.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = TWD.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity MCom on MCom.CommodityID = TWD.CommodityID"
If wc <> "" Then
    tmp = tmp & "  " & wc & " And TWD.WHRNoofBags - TWD.DeliveredNoOfBags > 0"
Else
    tmp = tmp & " Where TWD.WHRNoofBags - TWD.DeliveredNoOfBags > 0"
End If
tmp = tmp & " Group by MS.StateName, ML.LocationName, MCL.ClientName, TWD.WHRNo, TWD.CMPID, MCom.Commodity"
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCL.ClientName, TWD.WHRNo, TWD.CMPID, MCom.Commodity"

Call TmpTable
   
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid2.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid2.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid2.TextMatrix(I, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
       MSHFlexGrid2.TextMatrix(I, 3) = IIf(IsNull(TmpRs!WHRNo), "", TmpRs!WHRNo)
       MSHFlexGrid2.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid2.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
       MSHFlexGrid2.TextMatrix(0, 5) = Val(MSHFlexGrid2.TextMatrix(0, 5)) + Val(MSHFlexGrid2.TextMatrix(I, 5))
       MSHFlexGrid2.TextMatrix(I, 6) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
       MSHFlexGrid2.TextMatrix(0, 6) = Val(MSHFlexGrid2.TextMatrix(0, 6)) + Val(MSHFlexGrid2.TextMatrix(I, 6))
       TmpRs.MoveNext
   Next
End If

tmp = " select MS.StateName, ML.LocationName, MCL.ClientName, MG.GodownNo, GSL.StackNo, GSL.LotNo, MCom.Commodity,  sum(GSL.BalanceBags) as BalanceBags, sum(GSL.BalanceWeight) as BalanceWeight"
tmp = tmp & "  From Mst_GSL GSL"
tmp = tmp & "  Inner Join Mst_Location ML on ML.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & "  Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & "  Inner Join Mst_Client MCL on MCL.ClientIDRow = GSL.ClientIDRow"
tmp = tmp & "  Inner Join Mst_Godown MG on MG.GodownID = GSL.GodownID"
tmp = tmp & "  Inner Join Mst_Commodity MCom on MCom.CommodityID = GSL.CommodityID"
If wc <> "" Then
    tmp = tmp & "  " & wc & " And GSL.Flag = 'O'  and GSL.ExchangeTypeID = 1"
Else
    tmp = tmp & " Where GSL.Flag = 'O'  and GSL.ExchangeTypeID = 1"
End If
tmp = tmp & " Group By MS.StateName, ML.LocationName, MCL.ClientName, MG.GodownNo, GSL.StackNo, GSL.LotNo, MCom.Commodity"
tmp = tmp & " Order By MS.StateName, ML.LocationName, MCL.ClientName, MG.GodownNo, GSL.StackNo, GSL.LotNo, MCom.Commodity"

Call TmpTable
   
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = MSHFlexGrid3.Rows + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid3.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid3.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid3.TextMatrix(I, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
       MSHFlexGrid3.TextMatrix(I, 3) = IIf(IsNull(TmpRs!GodownNo), "", TmpRs!GodownNo)
       MSHFlexGrid3.TextMatrix(I, 4) = IIf(IsNull(TmpRs!StackNo), "", TmpRs!StackNo)
       MSHFlexGrid3.TextMatrix(I, 5) = IIf(IsNull(TmpRs!LotNo), "", TmpRs!LotNo)
       MSHFlexGrid3.TextMatrix(I, 6) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid3.TextMatrix(I, 7) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
       MSHFlexGrid3.TextMatrix(0, 7) = Val(MSHFlexGrid3.TextMatrix(0, 7)) + Val(MSHFlexGrid3.TextMatrix(I, 7))
       MSHFlexGrid3.TextMatrix(I, 8) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
       MSHFlexGrid3.TextMatrix(0, 8) = Val(MSHFlexGrid3.TextMatrix(0, 8)) + Val(MSHFlexGrid3.TextMatrix(I, 8))
       TmpRs.MoveNext
   Next
End If

tmp = " Select MS.StateName, ML.LocationName, MCL.ClientName,"
tmp = tmp & " convert(varchar,isnull(TWD.CMPID,0))+'-'+convert(varchar, (TWD.WHRNo-(isnull(TWD.CMPID,0)*1000000))) as 'WHRNo',"
tmp = tmp & " MCom.Commodity, sum(TWD.WHRNoOfBags-TWD.DeliveredNoOfBags) As BalanceBags,  sum(TWD.WHRQuantity - TWD.DeliveredQuantity) As BalanceWeight"
tmp = tmp & " From Txn_WHR_Detail TWD"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TWD.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = TWD.ClientIDRow"
tmp = tmp & " Inner Join Mst_Commodity MCom on MCom.CommodityID = TWD.CommodityID"
If wc <> "" Then
    tmp = tmp & "  " & wc & " And TWD.WHRNoofBags - TWD.DeliveredNoOfBags < 0"
Else
    tmp = tmp & " Where TWD.WHRNoofBags - TWD.DeliveredNoOfBags < 0"
End If
tmp = tmp & " Group by MS.StateName, ML.LocationName, MCL.ClientName, TWD.WHRNo, TWD.CMPID, MCom.Commodity"
tmp = tmp & " Order by MS.StateName, ML.LocationName, MCL.ClientName, TWD.WHRNo, TWD.CMPID, MCom.Commodity"

Call TmpTable
   
If TmpRs.RecordCount > 0 Then
   MSHFlexGrid4.Rows = MSHFlexGrid4.Rows + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid4.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid4.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid4.TextMatrix(I, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
       MSHFlexGrid4.TextMatrix(I, 3) = IIf(IsNull(TmpRs!WHRNo), "", TmpRs!WHRNo)
       MSHFlexGrid4.TextMatrix(I, 4) = IIf(IsNull(TmpRs!Commodity), "", TmpRs!Commodity)
       MSHFlexGrid4.TextMatrix(I, 5) = IIf(IsNull(TmpRs!BalanceBags), "", TmpRs!BalanceBags)
       MSHFlexGrid4.TextMatrix(0, 5) = Val(MSHFlexGrid4.TextMatrix(0, 5)) + Val(MSHFlexGrid4.TextMatrix(I, 5))
       MSHFlexGrid4.TextMatrix(I, 6) = IIf(IsNull(TmpRs!BalanceWeight), "", TmpRs!BalanceWeight)
       MSHFlexGrid4.TextMatrix(0, 6) = Val(MSHFlexGrid4.TextMatrix(0, 6)) + Val(MSHFlexGrid4.TextMatrix(I, 6))
       TmpRs.MoveNext
   Next
End If


tmp = " select StateName,LocationName, ClientName, StateID, LocationID, ClientIDRow,"
tmp = tmp & " sum(ActualBalanceBags) ActualBalanceBags, sum(ActualBalanceWeight) ActualBalanceWeight, sum(WHRBalanceBags) WHRBalanceBags, sum(WHRBalanceWeight) WHRBalanceWeight"
tmp = tmp & " from("
tmp = tmp & " select MS.StateName, ML.LocationName, MCL.ClientName, MS.StateID, ML.LocationID, MCL.ClientIDRow,"
tmp = tmp & " sum(GSL.BalanceBags) as ActualBalanceBags, sum(GSL.BalanceWeight) as ActualBalanceWeight, '0' WHRBalanceBags, '0' WHRBalanceWeight"
tmp = tmp & " From Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = GSL.ClientIDRow"
If wc <> "" Then
    tmp = tmp & "  " & wc & " And GSL.Flag = 'O'  and GSL.ExchangeTypeID = 1"
Else
    tmp = tmp & " Where GSL.Flag = 'O'  and GSL.ExchangeTypeID = 1"
End If
tmp = tmp & " Group By MS.StateName, ML.LocationName, MCL.ClientName, MS.StateID, ML.LocationID, MCL.ClientIDRow"
tmp = tmp & " Union All"
tmp = tmp & " Select MS.StateName, ML.LocationName, MCL.ClientName, MS.StateID, ML.LocationID, MCL.ClientIDRow,"
tmp = tmp & " '0' ActualBalanceBags, '0' ActualBalanceWeight, sum(TWD.WHRNoOfBags-TWD.DeliveredNoOfBags) As WHRBalanceBags, sum(TWD.WHRQuantity - TWD.DeliveredQuantity) As WHRBalanceWeight"
tmp = tmp & " From Txn_WHR_Detail  TWD"
tmp = tmp & " Inner Join Mst_Location ML on ML.SM_Prefix = TWD.SM_Prefix"
tmp = tmp & " Inner Join Mst_State MS on MS.StateID = ML.StateID"
tmp = tmp & " Inner Join Mst_Client MCL on MCL.ClientIDRow = TWD.ClientIDRow"
If wc <> "" Then
    tmp = tmp & "  " & wc & " And TWD.WHRNoofBags - TWD.DeliveredNoOfBags > 0"
Else
    tmp = tmp & " Where TWD.WHRNoofBags - TWD.DeliveredNoOfBags > 0"
End If
tmp = tmp & " Group By MS.StateName, ML.LocationName, MCL.ClientName, MS.StateID, ML.LocationID, MCL.ClientIDRow"
tmp = tmp & " )A"
tmp = tmp & " Group By"
tmp = tmp & " StateName , LocationName, ClientName, StateID, LocationID, ClientIDRow"
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + TmpRs.RecordCount
   For I = 2 To TmpRs.RecordCount + 1
       MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs!StateName), "", TmpRs!StateName)
       MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs!LocationName), "", TmpRs!LocationName)
       MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs!ClientName), "", TmpRs!ClientName)
       MSHFlexGrid1.TextMatrix(I, 3) = IIf(IsNull(TmpRs!ActualBalanceBags), "", TmpRs!ActualBalanceBags)
       MSHFlexGrid1.TextMatrix(I, 4) = IIf(IsNull(TmpRs!WHRBalanceBags), "", TmpRs!WHRBalanceBags)
       MSHFlexGrid1.TextMatrix(I, 5) = Val(MSHFlexGrid1.TextMatrix(I, 3)) - Val(MSHFlexGrid1.TextMatrix(I, 4))
       MSHFlexGrid1.TextMatrix(I, 6) = IIf(IsNull(TmpRs!ActualBalanceWeight), "", TmpRs!ActualBalanceWeight)
       MSHFlexGrid1.TextMatrix(I, 7) = IIf(IsNull(TmpRs!WHRBalanceWeight), "", TmpRs!WHRBalanceWeight)
       MSHFlexGrid1.TextMatrix(I, 8) = Val(MSHFlexGrid1.TextMatrix(I, 6)) - Val(MSHFlexGrid1.TextMatrix(I, 7))
       
       MSHFlexGrid1.TextMatrix(0, 3) = Val(MSHFlexGrid1.TextMatrix(0, 3)) + Val(MSHFlexGrid1.TextMatrix(I, 3))
       MSHFlexGrid1.TextMatrix(0, 4) = Val(MSHFlexGrid1.TextMatrix(0, 4)) + Val(MSHFlexGrid1.TextMatrix(I, 4))
       MSHFlexGrid1.TextMatrix(0, 5) = Val(MSHFlexGrid1.TextMatrix(0, 5)) + Val(MSHFlexGrid1.TextMatrix(I, 5))
       MSHFlexGrid1.TextMatrix(0, 6) = Val(MSHFlexGrid1.TextMatrix(0, 6)) + Val(MSHFlexGrid1.TextMatrix(I, 6))
       MSHFlexGrid1.TextMatrix(0, 7) = Val(MSHFlexGrid1.TextMatrix(0, 7)) + Val(MSHFlexGrid1.TextMatrix(I, 7))
       MSHFlexGrid1.TextMatrix(0, 8) = Val(MSHFlexGrid1.TextMatrix(0, 8)) + Val(MSHFlexGrid1.TextMatrix(I, 8))
       TmpRs.MoveNext
   Next
End If

lblRecordCount.Caption = TmpRs.RecordCount
MsgBox "Done!!!"

End Sub

Private Sub Form_Load()

Call Grid_Head
Call Grid_Head2
Call Grid_Head3
Call Grid_Head4
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
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Rpt_WHRvsActualStock.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Rpt_WHRvsActualStock.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.Name = "WHR vs Actual Stock"
oWS.ClearArrows
mRow = 1
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 15
oWS.Range("c:c").ColumnWidth = 25
oWS.Range("d:d").ColumnWidth = 11
oWS.Range("e:e").ColumnWidth = 11
oWS.Range("f:f").ColumnWidth = 11
oWS.Range("g:g").ColumnWidth = 11
oWS.Range("h:h").ColumnWidth = 11
oWS.Range("i:i").ColumnWidth = 11

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "WHR vs Actual Stock"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(I, c)
        If I = 1 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, c + 1).WrapText = True
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next

Set oWS = oWB.Worksheets("Sheet2")
oWS.Name = "WHR Stock"
oWS.ClearArrows
mRow = 1
c = 0
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 15
oWS.Range("c:c").ColumnWidth = 25
oWS.Range("d:d").ColumnWidth = 11
oWS.Range("e:e").ColumnWidth = 25
oWS.Range("f:f").ColumnWidth = 11
oWS.Range("g:g").ColumnWidth = 11

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "WHR Stock"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

For I = 0 To MSHFlexGrid2.Rows - 1
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid2.Cols - 1
        
        If c = 3 Then
            oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid2.TextMatrix(I, c)
        Else
            oWS.Cells(mRow, c + 1) = MSHFlexGrid2.TextMatrix(I, c)
        End If
        If I = 1 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, c + 1).WrapText = True
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next


Set oWS = oWB.Worksheets("Sheet3")
oWS.Name = "Actual Stock"
oWS.ClearArrows
mRow = 1
c = 0
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 15
oWS.Range("c:c").ColumnWidth = 25
oWS.Range("d:d").ColumnWidth = 9
oWS.Range("e:e").ColumnWidth = 9
oWS.Range("f:f").ColumnWidth = 9
oWS.Range("g:g").ColumnWidth = 25
oWS.Range("h:h").ColumnWidth = 11
oWS.Range("i:i").ColumnWidth = 11

tmp = "a" & mRow & ":" & "i" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "Actual Stock"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

For I = 0 To MSHFlexGrid3.Rows - 1
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid3.Cols - 1
        oWS.Cells(mRow, c + 1) = MSHFlexGrid3.TextMatrix(I, c)
        
        If I = 1 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, c + 1).WrapText = True
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next

Set oWS = oWB.Worksheets("Sheet4")
oWS.Name = "WHR having excess delivery"
oWS.ClearArrows
mRow = 1
c = 0
oWS.Range("a:a").ColumnWidth = 15
oWS.Range("b:b").ColumnWidth = 15
oWS.Range("c:c").ColumnWidth = 25
oWS.Range("d:d").ColumnWidth = 11
oWS.Range("e:e").ColumnWidth = 25
oWS.Range("f:f").ColumnWidth = 11
oWS.Range("g:g").ColumnWidth = 11

tmp = "a" & mRow & ":" & "g" & mRow
oWS.Range(tmp).Merge
oWS.Range(tmp).Borders.Color = RGB(0, 0, 0)
oWS.Cells(mRow, c + 1) = "WHR Stock"
oWS.Cells(mRow, c + 1).HorizontalAlignment = -4108
oWS.Cells(mRow, c + 1).Font.Bold = True
oWS.Cells(mRow, c + 1).Font.Size = 14
oWS.Cells(mRow, c + 1).Interior.Color = &HFF9F9F
oWS.Cells(mRow, c + 1).WrapText = True

For I = 0 To MSHFlexGrid4.Rows - 1
    mRow = mRow + 1 ' + 6
    For c = 0 To MSHFlexGrid4.Cols - 1
        
        If c = 3 Then
            oWS.Cells(mRow, c + 1) = "'" & MSHFlexGrid4.TextMatrix(I, c)
        Else
            oWS.Cells(mRow, c + 1) = MSHFlexGrid4.TextMatrix(I, c)
        End If
        If I = 1 Then
           oWS.Cells(mRow, c + 1).Interior.Color = &H55FFFF
           oWS.Cells(mRow, c + 1).Font.Bold = True
           oWS.Cells(mRow, c + 1).RowHeight = 30
        End If
        oWS.Cells(mRow, c + 1).WrapText = True
        oWS.Cells(mRow, c + 1).Borders.Color = RGB(0, 0, 0)
    Next
Next


oWB.Save
MsgBox ("Excel File Created !!! ")
oXL.Visible = True
End Sub

Private Sub CmbSClientName_GotFocus()
tmp = CmbSClientName.Text

Call TableMst_Client("Where ExchangeTypeId = 1 ") '" & mExchangeTypeID)
CmbSClientName.Clear
If RsMst_Client.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Client.RecordCount
    CmbSClientName.AddItem RsMst_Client!ClientName
    RsMst_Client.MoveNext
Next
CmbSClientName.Text = tmp

End Sub

Private Sub CmbSClientName_LostFocus()
If CmbSClientName.Text = "" Then
   Exit Sub
End If
RsMst_Client.MoveFirst
RsMst_Client.Find "ClientName = '" & CmbSClientName.Text & "'"
If RsMst_Client.EOF Then
   MsgBox "Invalid selection "
   CmbSClientName.SetFocus
   Exit Sub
End If
mSClientID = RsMst_Client!ClientIDRow

End Sub
