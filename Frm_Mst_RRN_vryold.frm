VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_RRN 
   Caption         =   "RRN Master"
   ClientHeight    =   7095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10890
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2205
      Left            =   8520
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   4080
      Begin VB.TextBox TxtSearchISIN 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         TabIndex        =   60
         Top             =   585
         Width           =   2205
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   10590
         TabIndex        =   59
         Top             =   585
         Width           =   2565
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   6630
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   585
         Width           =   3930
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   13230
         TabIndex        =   57
         Top             =   180
         Width           =   1000
      End
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   3675
         Sorted          =   -1  'True
         TabIndex        =   56
         Top             =   585
         Width           =   2910
      End
      Begin VB.TextBox TxtSearchRRN 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   585
         Width           =   1215
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13230
         TabIndex        =   15
         Top             =   600
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1455
         Left            =   75
         TabIndex        =   6
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label LblISINNo 
         Caption         =   "ISIN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2190
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   11550
         TabIndex        =   63
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "CMP"
         Height          =   300
         Left            =   7590
         TabIndex        =   62
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label4 
         Caption         =   "Location"
         Height          =   300
         Left            =   4110
         TabIndex        =   61
         Top             =   300
         Width           =   1320
      End
      Begin VB.Label LblSearchRRN 
         Caption         =   "RRN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   195
         TabIndex        =   36
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "CITF Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   74
      Top             =   5790
      Width           =   14805
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid3 
         Height          =   1920
         Left            =   75
         TabIndex        =   75
         Top             =   270
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   3387
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "CA Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   135
      TabIndex        =   72
      Top             =   3480
      Width           =   14805
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1920
         Left            =   75
         TabIndex        =   73
         Top             =   270
         Width           =   14640
         _ExtentX        =   25823
         _ExtentY        =   3387
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   135
      TabIndex        =   29
      Top             =   8160
      Width           =   14805
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11775
         TabIndex        =   30
         Top             =   645
         Width           =   2835
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11775
         TabIndex        =   31
         Top             =   240
         Width           =   2835
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8940
         TabIndex        =   8
         Top             =   645
         Width           =   2835
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8940
         TabIndex        =   9
         Top             =   240
         Width           =   2835
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   5985
         TabIndex        =   10
         Top             =   645
         Width           =   2955
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   5985
         TabIndex        =   11
         Top             =   240
         Width           =   2955
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3030
         TabIndex        =   12
         Top             =   645
         Width           =   2955
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3030
         TabIndex        =   7
         Top             =   240
         Width           =   2955
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   75
         TabIndex        =   13
         Top             =   645
         Width           =   2955
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Width           =   2955
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3315
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   14820
      Begin VB.TextBox TxtUpdated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   8565
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   615
         Width           =   6120
      End
      Begin VB.TextBox TxtCreated 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   8565
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   150
         Width           =   6120
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   1
         TabIndex        =   69
         Top             =   3990
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.TextBox TxtRRNQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   68
         Top             =   2880
         Width           =   1560
      End
      Begin VB.TextBox TxtCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1785
         Width           =   3120
      End
      Begin VB.ComboBox CmbClientDP 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1410
         Visible         =   0   'False
         Width           =   8445
      End
      Begin VB.TextBox TxtOldDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.TextBox TxtISINDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   1800
         Width           =   5565
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2145
         Width           =   3120
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2505
         Width           =   8445
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   2160
         Width           =   8445
      End
      Begin VB.TextBox TxtGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   2505
         Width           =   3120
      End
      Begin VB.TextBox TxtDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1425
         Width           =   2850
      End
      Begin VB.TextBox TxtDPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1425
         Width           =   4125
      End
      Begin VB.TextBox TxtISINID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1800
         TabIndex        =   5
         Top             =   1800
         Width           =   2850
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1035
         Width           =   4125
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         TabIndex        =   3
         Top             =   1005
         Width           =   2880
      End
      Begin VB.TextBox TxtBalanceQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Height          =   360
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   2880
         Width           =   1605
      End
      Begin VB.CommandButton CmdClientID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4710
         TabIndex        =   34
         Top             =   1035
         Width           =   375
      End
      Begin VB.CommandButton CmdISINID 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10305
         TabIndex        =   33
         Top             =   1395
         Width           =   375
      End
      Begin VB.TextBox TxtRRNID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   2880
      End
      Begin VB.TextBox TxtRRNNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   1
         Top             =   615
         Width           =   2880
      End
      Begin MSComCtl2.DTPicker RRNDate 
         Height          =   360
         Left            =   6120
         TabIndex        =   2
         Top             =   630
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Format          =   59179009
         CurrentDate     =   39864
      End
      Begin VB.TextBox TxtDeliveredQty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   2865
         Width           =   1845
      End
      Begin MSComCtl2.DTPicker EntryDate 
         Height          =   360
         Left            =   6135
         TabIndex        =   27
         Top             =   270
         Visible         =   0   'False
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   59179009
         CurrentDate     =   39864
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   7680
         TabIndex        =   66
         Top             =   645
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Created  :"
         Height          =   255
         Left            =   7680
         TabIndex        =   65
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label lblNameDid 
         Caption         =   "ClientDPID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   255
         TabIndex        =   37
         Top             =   1425
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Exchange Type"
         Height          =   300
         Left            =   10305
         TabIndex        =   53
         Top             =   2205
         Width           =   1440
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         Height          =   300
         Left            =   240
         TabIndex        =   52
         Top             =   2520
         Width           =   1200
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP"
         Height          =   300
         Left            =   240
         TabIndex        =   51
         Top             =   2175
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "Grade"
         Height          =   180
         Left            =   10320
         TabIndex        =   50
         Top             =   2565
         Width           =   900
      End
      Begin VB.Label LblCommodity 
         Caption         =   "Commodity"
         Height          =   300
         Left            =   10320
         TabIndex        =   49
         Top             =   1875
         Width           =   1140
      End
      Begin VB.Label Label8 
         Caption         =   "Client Name "
         Height          =   255
         Left            =   5160
         TabIndex        =   40
         Top             =   1035
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "DP ID"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1455
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Balance Qty in MT"
         Height          =   300
         Left            =   6960
         TabIndex        =   32
         Top             =   2895
         Width           =   1440
      End
      Begin VB.Label LblRRNID 
         Caption         =   "RRN ID"
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label LblRRNNo 
         Caption         =   "RRN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   19
         Top             =   615
         Width           =   1800
      End
      Begin VB.Label LblRRNDate 
         Caption         =   "RRN Date"
         Height          =   300
         Left            =   5160
         TabIndex        =   20
         Top             =   630
         Width           =   1800
      End
      Begin VB.Label LblISINID 
         Caption         =   "ISIN No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   21
         Top             =   1875
         Width           =   1800
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   22
         Top             =   1035
         Width           =   1800
      End
      Begin VB.Label LblRRNQty 
         Caption         =   "RRN Qty in MT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   23
         Top             =   2970
         Width           =   1560
      End
      Begin VB.Label LblDeliveredQty 
         Caption         =   "Delivered Qty in MT"
         Height          =   300
         Left            =   3480
         TabIndex        =   24
         Top             =   2925
         Width           =   1800
      End
      Begin VB.Label LblEntryDate 
         Caption         =   "Entry Date"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5175
         TabIndex        =   26
         Top             =   270
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   300
         Left            =   255
         TabIndex        =   28
         Top             =   4020
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "DP Name "
         Height          =   255
         Left            =   5160
         TabIndex        =   38
         Top             =   1440
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmMst_RRN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg As Variant
Dim mISINID, mClientID As String
Dim mClientIDRow As Double
Dim mCMPName, mCMPID, mLocationID, mCommodityID, mExchangeTypeID As Double
Dim x As Integer


Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
Call TableMst_CMP
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_CMP.RecordCount
    CmbSCmp.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbSCmp.Text = tmp
End Sub
Private Sub CmbSCmp_LostFocus()
If CmbSCmp.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbSCmp.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbSCmp.SetFocus
   Exit Sub
End If
mCMPName = RsMst_CMP!CMPID
End Sub

Private Sub CmbSCommodity_GotFocus()
tmp = CmbSCommodity.Text
Call TableMst_Commodity
CmbSCommodity.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Commodity.RecordCount
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

mCommodityID = RsMst_Commodity!CommodityID

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbSLocation.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbSLocation.Text = tmp
End Sub
Private Sub CmbSLocation_LostFocus()
If CmbSLocation.Text = "" Then
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocation.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocation.SetFocus
   Exit Sub
End If
mLocationID = RsMst_Location!LocationID
End Sub

Private Sub CmdSearch_Click()
Dim x As Integer
Dim Wc As String
Wc = ""
RsMst_ISIN.MoveFirst
Call Grid_Head
If TxtSearchRRN <> "" Then
   'tmp = "Select * from Mst_ISIN where ISINID like '%" & Trim(TxtSearchISIN) & "%'"
   Wc = " where rrn.RRNNo like '%" & Trim(TxtSearchRRN) & "%' "
Else
    If CmbSCmp.Text <> "" Then
       Wc = GenWc(Wc, mCMPName, "cmp.CMPID", "N", "N")
    End If
    If CmbSLocation.Text <> "" Then
       Wc = GenWc(Wc, mLocationID, "l.LocationID", "N", "N")
    End If
    If CmbSCommodity.Text <> "" Then
       Wc = GenWc(Wc, mCommodityID, "com.CommodityID", "N", "N")
    End If
    If TxtSearchISIN.Text <> "" Then
        Wc = GenWc(Wc, TxtSearchISIN, "rrn.ISINID", "S", "L")
    End If
End If

tmp = "select rrn.RRNID, rrn.RRNNo, rrn.RRNDate, rrn.ISINID, l.LocationName, cmp.CMPName, com.Commodity, "
tmp = tmp & " c.ClientID + ':'+ c.ClientName As Client , dp.DepositoryID + ':'+ dp.DepositoryName As [DP Name], "
tmp = tmp & " rrn.RRNQty,rrn.DeliveredQty,rrn.EntryDate,rrn.Flag,rrn.G_UID "
tmp = tmp & " from mst_rrn rrn "
tmp = tmp & " inner join Mst_ISIN isin on rrn.ISINID = isin.ISINID "
tmp = tmp & " inner join Mst_CMP cmp on isin.CMPID = cmp.CMPID "
tmp = tmp & " inner join mst_location l on cmp.LocationID = l.LocationID "
tmp = tmp & " inner join Mst_Commodity com on isin.CommodityID = com.CommodityID "
tmp = tmp & " inner join Mst_Client c on  rrn.ClientIDRow = c.ClientIDRow  "
tmp = tmp & " inner join Mst_Depository dp on c.DepositoryID = dp.DepositoryID and dp.ExchangeTypeID = 2 "
tmp = tmp & Wc
tmp = tmp & " order by rrn.RRNNo,cmp.CMPName,com.Commodity "
Call Tmp4Table

If TmpRs4.RecordCount <= 0 Then
   MsgBox "No Record found!!!"
   TxtSearchRRN.SetFocus
   Exit Sub
End If
Call ShowGrid
End Sub
Private Sub ShowGrid()
MSHFlexGrid1.Rows = TmpRs4.RecordCount + 2
For x = 1 To TmpRs4.RecordCount
    For c = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(x, c) = IIf(IsNull(TmpRs4.Fields(c)), "", TmpRs4.Fields(c))
    Next
    TmpRs4.MoveNext
Next
TmpRs4.Close
End Sub

Private Sub Form_Load()
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call TableMst_RRN
If RsMst_RRN.RecordCount = 0 Then
   If UserType = "Auditor" Then
      Call GButtonLock(Me, False)
      SaveCmd.Enabled = False
      SearchCmd.Enabled = False
      Exit Sub
   End If
   Call AddCmd_Click
   Exit Sub
End If
Call GButtonLock(Me, True)
Call InData
Call Grid_Head1
Call Grid_Head2
End Sub
Private Sub MSHFlexGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyC And Shift = vbCtrlMask Then
   With MSHFlexGrid3
        .Redraw = False              'Find range to copy
        min_row = .Row:   max_row = .RowSel
        If min_row > max_row Then tmp = max_row: max_row = min_row: min_row = tmp
        min_col = .Col:   max_col = .ColSel
        If min_col > max_col Then tmp = max_col: max_col = min_col: min_col = tmp

        my_text = ""                   'Get the text
        For i = min_row To max_row
            For j = min_col To max_col
                my_text = my_text & .TextMatrix(i, j)
                If j < max_col Then my_text = my_text & vbTab
            Next j
            If i < max_row Then my_text = my_text & vbCrLf
        Next i
        .Redraw = True
   End With
   If my_text <> "" Then
      Clipboard.Clear
      Clipboard.SetText (my_text)
   Else
      Beep
      MsgBox "Please select some text.", vbExclamation, "Nothing selected!"
   End If
End If
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_RRN.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_RRN.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For i = 0 To MSHFlexGrid1.Rows - 1
    mRow = i + 1 ' + 6
    For c = 0 To MSHFlexGrid1.Cols - 2
        If InStr(MSHFlexGrid1.TextMatrix(i, c), "/") > 0 Then
           If IsDate(MSHFlexGrid1.TextMatrix(i, c)) Then
              oWS.Cells(mRow, c + 1) = Format(MSHFlexGrid1.TextMatrix(i, c), "MM/dd/yyyy")
           Else
              oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
           End If
        Else
           oWS.Cells(mRow, c + 1) = MSHFlexGrid1.TextMatrix(i, c)
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub



Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_RRN.MoveNext
If RsMst_RRN.EOF Then
   RsMst_RRN.MoveLast
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    NextCmd.Enabled = False
    LastCmd.Enabled = False
End If
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
End Sub
Private Sub PreviousCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_RRN.MovePrevious
If RsMst_RRN.BOF Then
   RsMst_RRN.MoveFirst
   LF_Flag = "Y"
End If
Call InData
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_RRN.MoveFirst
Call InData
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_RRN.MoveLast
Call InData
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsMst_RRN.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   Call FirstCmd_Click
   Exit Sub
End If
Frame0.Height = Me.Height - 900
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Clear
MSHFlexGrid1.Cols = 3
MSHFlexGrid1.Width = Frame0.Width - 300
MSHFlexGrid1.Height = Frame0.Height - 1600
RsMst_RRN.MoveFirst
Call Grid_Head
TxtSearchRRN = ""



End Sub
Public Sub InData()

TxtRRNID.Locked = True
TxtRRNNo.Locked = True
RRNDate.Enabled = False
TxtISINID.Locked = True
'CmbISINID.Visible = False
CmdISINID.Visible = False
TxtISINID.Visible = True
TxtClientID.Locked = True
'CmbClientId.Visible = False
CmdClientId.Visible = False
CmbClientDP.Visible = False
lblNameDid.Visible = False
TxtClientID.Visible = True
TxtRRNQty.Locked = True
TxtDeliveredQty.Locked = True
TxtBalanceQty.Locked = True
EntryDate.Enabled = False
TxtFlag.Locked = True

TxtRRNID = IIf(IsNull(RsMst_RRN!RRNID), "", RsMst_RRN!RRNID)
TxtRRNNo = IIf(IsNull(RsMst_RRN!RRNNo), "", RsMst_RRN!RRNNo)
RRNDate = IIf(IsNull(RsMst_RRN!RRNDate), "", RsMst_RRN!RRNDate)
mISINID = IIf(IsNull(RsMst_RRN!ISINID), "", RsMst_RRN!ISINID)
'TxtISINID = GetISINName(RsMst_RRN!ISINID)
Call TableMst_ISIN(" Where ISINID = '" & RsMst_RRN!ISINID & "'")
TxtISINID = RsMst_RRN!ISINID
TxtISINDesc = RsMst_ISIN!ISINDescription
mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)
TxtCMPID = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
TxtLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
TxtCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
mExchangeTypeID = IIf(IsNull(RsMst_ISIN!exchangetypeid), "", RsMst_ISIN!exchangetypeid)
TxtExchangeTypeID = GetExchangeName(RsMst_ISIN!exchangetypeid) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
TxtGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
'TxtISINID = IIf(IsNull(RsMst_RRN!ISINDescription), "", RsMst_RRN!ISINDescription)

mClientIDRow = IIf(IsNull(RsMst_RRN!ClientIDRow), "", RsMst_RRN!ClientIDRow)


Dim RetGetVal As Variant

RetGetVal = GetClientNameByRow((RsMst_RRN!ClientIDRow))   '(IsNull(RsMst_RRN!ClientIDRow), "", RsMst_RRN!ClientIDRow)

TxtClientID = RetGetVal(0) 'IIf(IsNull(RsMst_RRN!ClientID), "", RsMst_RRN!ClientID)
TxtClientName = RetGetVal(1)
TxtDPID = RetGetVal(2) 'GetDPFromClient(TxtClientID)
TxtOldDPID = RetGetVal(2) 'GetDPFromClient(TxtClientID)
TxtDPName = RetGetVal(3) 'GetDepositoryName(TxtDPID)

TxtRRNQty = IIf(IsNull(RsMst_RRN!RRNQty), "", RsMst_RRN!RRNQty)
TxtDeliveredQty = IIf(IsNull(RsMst_RRN!DeliveredQty), "", RsMst_RRN!DeliveredQty)
TxtBalanceQty = Format(Val(TxtRRNQty) - Val(TxtDeliveredQty), "########0.000")
If IsNull(RsMst_RRN!EntryDate) = False Then
   'EntryDate = IIf(IsNull(RsMst_RRN!EntryDate), "", RsMst_RRN!EntryDate)
   EntryDate = RsMst_RRN!EntryDate
End If
TxtFlag = IIf(IsNull(RsMst_RRN!Flag), "", RsMst_RRN!Flag)


TxtCreated = IIf(IsNull(RsMst_RRN!CreatedBy), "", RsMst_RRN!CreatedBy) & " :- " & IIf(IsNull(RsMst_RRN!CreatedDate), "", RsMst_RRN!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_RRN!UpdatedBy), "", RsMst_RRN!UpdatedBy) & " :- " & IIf(IsNull(RsMst_RRN!UpdatedDate), "", RsMst_RRN!UpdatedDate)

Call Indata_CA
Call Indata_CITF

If RsMst_RRN.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub
Private Sub Indata_CA()
Call Grid_Head1

tmp = "Select CANo,CADate,Quantity,PartialAcceptance from Mst_CA Where RRNID = " & TxtRRNID & ""

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs.RecordCount + 2
   mRow = 2
   For x = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(mRow, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid2.TextMatrix(mRow, 1) = Format(MSHFlexGrid2.TextMatrix(mRow, 1), "dd-mmm-yyyy")
       MSHFlexGrid2.TextMatrix(0, 3) = Val(MSHFlexGrid2.TextMatrix(0, 3)) + Val(MSHFlexGrid2.TextMatrix(mRow, 3))
       mRow = mRow + 1
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub Indata_CITF()
Call Grid_Head2

'tmp = "Select CANo,CADate,Quantity,PartialAcceptance from Mst_CA Where RRNID = " & TxtRRNID & ""



tmp = "select a.SM_Prefix,a.CxTFNo,b.CxTFDate,a.AdjRRNQty,a.AdjRRNNoofBags From Txn_CxTF_ClientDetail a " & _
"inner join Txn_CxTF_Details b on a.SM_Prefix = b.SM_Prefix And a.cxtfno = b.cxtfno  And a.txntype = b.txntype " & _
"where b.exchangetypeid= 2 and b.txntype = 'W' and a.rrnid = " & Val(TxtRRNID)

Call TmpTable

If TmpRs.RecordCount > 0 Then
   MSHFlexGrid3.Rows = TmpRs.RecordCount + 2
   mRow = 2
   For x = 1 To TmpRs.RecordCount
       For c = 0 To MSHFlexGrid3.Cols - 1
           MSHFlexGrid3.TextMatrix(mRow, c) = IIf(IsNull(TmpRs.Fields(c)), "", TmpRs.Fields(c))
       Next
       MSHFlexGrid3.TextMatrix(mRow, 2) = Format(MSHFlexGrid3.TextMatrix(mRow, 2), "dd-mmm-yyyy")
       MSHFlexGrid3.TextMatrix(0, 3) = Val(MSHFlexGrid3.TextMatrix(0, 3)) + Val(MSHFlexGrid3.TextMatrix(mRow, 3))
       MSHFlexGrid3.TextMatrix(0, 4) = Val(MSHFlexGrid3.TextMatrix(0, 4)) + Val(MSHFlexGrid3.TextMatrix(mRow, 4))
       mRow = mRow + 1
       TmpRs.MoveNext
   Next
End If

End Sub

Private Sub TxtISINID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtISINID.Text = "" Then
    mCommodityID = 0
    mCMPID = 0
    'mExchangeTypeID = 0
    TxtCMPID = ""
    TxtLocationID = ""
    TxtCommodity = ""
    TxtExchangeTypeID = ""
    TxtGrade = ""
    Call Grid_Head
    Exit Sub
End If
Call TableMst_ISIN("where ISINID = '" & TxtISINID.Text & "'")

If RsMst_ISIN.RecordCount = 0 Then
   MsgBox "No Such ISIN Exist !!!"
   TxtISINID.SetFocus
   Exit Sub
End If
mISINID = RsMst_ISIN!ISINID
mCMPID = RsMst_ISIN!CMPID
TxtCMPID = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
TxtISINDesc = RsMst_ISIN!ISINDescription
mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
TxtLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
TxtCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
mExchangeTypeID = IIf(IsNull(RsMst_ISIN!exchangetypeid), "", RsMst_ISIN!exchangetypeid)
TxtExchangeTypeID = GetExchangeName(RsMst_ISIN!exchangetypeid) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
TxtGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
End Sub
Private Sub TxtRRNID_LostFocus()
If TxtRRNID = "" Then Exit Sub
If IsNumeric(TxtRRNID) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtRRNID.SetFocus
   Exit Sub
End If
TxtRRNID = Val(TxtRRNID)
End Sub
Private Sub TxtRRNNo_LostFocus()
If TxtRRNNo = "" Then Exit Sub
If RsMst_RRN.RecordCount = 0 Then Exit Sub
TxtRRNNo = Trim(TxtRRNNo)
'RsMst_RRN.MoveFirst
'
'RsMst_RRN.Find "RRNNo = '" & TxtRRNNo & "'"
'If Not RsMst_RRN.EOF Then
'   Call Indata
'   Exit Sub
'End If
End Sub

Private Sub TxtRRNQty_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
If TxtRRNQty = "" Then Exit Sub
If IsNumeric(TxtRRNQty) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtRRNQty.SetFocus
   Exit Sub
End If
TxtRRNQty = Val(TxtRRNQty)
TxtBalanceQty = Val(TxtRRNQty) - Val(TxtDeliveredQty)
End Sub
Private Sub TxtDeliveredQty_LostFocus()
If TxtDeliveredQty = "" Then Exit Sub
If IsNumeric(TxtDeliveredQty) = False Then
   MsgBox "Invalid numeric format !!!!"
   TxtDeliveredQty.SetFocus
   Exit Sub
End If
TxtDeliveredQty = Val(TxtDeliveredQty)
'TxtBalanceQty = Val(TxtRRNQty) - Val(TxtDeliveredQty)
End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)

TxtRRNID.Locked = True
TxtRRNNo.Locked = False
RRNDate.Enabled = True
TxtISINID.Locked = False
'CmbISINID.Visible = True
CmdISINID.Visible = True
'TxtISINID.Visible = False
TxtClientID.Locked = False
'CmbClientId.Visible = True
CmdClientId.Visible = True
'TxtClientID.Visible = False
TxtRRNQty.Locked = False
TxtDeliveredQty.Locked = True
CmbClientDP.Visible = False
lblNameDid.Visible = False
EntryDate.Enabled = False
TxtFlag.Locked = False
TxtRRNID = ""
TxtRRNNo = ""
RRNDate = Date
TxtISINID = ""
TxtISINDesc = ""
TxtCMPID = ""
TxtLocationID = ""
TxtCommodity = ""
TxtGrade = ""
TxtExchangeTypeID = ""
CmbClientDP.Text = ""
TxtClientID = ""
TxtClientName = ""
TxtRRNQty = ""
TxtDPID.Text = ""
TxtDPName.Text = ""
TxtDeliveredQty = ""
TxtBalanceQty = ""
EntryDate = Date
TxtFlag = ""

TxtCreated = ""
TxtUpdated = ""

If RsMst_RRN.RecordCount > 0 Then
   TxtRRNNo.SetFocus
End If


End Sub
Private Sub EditCmd_Click()
If Val(TxtDeliveredQty) <> 0 Then
   MsgBox "you Can't Edit This RRN since, CITF is generated !!! "
   Exit Sub
End If
Edit_Flg = "E"
Call GButtonLock(Me, False)

TxtRRNID.Locked = True
TxtRRNNo.Locked = False
RRNDate.Enabled = True
TxtISINID.Locked = False
TxtISINDesc.Locked = True
'CmbISINID.Visible = True
CmdISINID.Visible = True
'TxtISINID.Visible = False
TxtClientID.Locked = False
'CmbClientId.Visible = True
CmbClientDP.Visible = False
lblNameDid.Visible = False
CmdClientId.Visible = True
'TxtClientID.Visible = False
TxtRRNQty.Locked = False
'TxtDeliveredQty.Locked = False
EntryDate.Enabled = False
TxtFlag.Locked = False
End Sub
Private Sub DeleteCmd_Click()

If Val(TxtDeliveredQty) <> 0 Then
   MsgBox "you Can't Delete This RRN since, CITF is generated !!! "
   Exit Sub
End If

tmp = "Select * from Mst_CA Where RRNID = " & TxtRRNID & ""
Call TmpTable

If TmpRs.RecordCount > 0 Then
   MsgBox "You Can't Delete This RRN !!! "
   Exit Sub
End If


Dim Ans As Variant
On Error GoTo msgError
Ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If Ans = vbYes Then
   RsMst_RRN.Delete
   RsMst_RRN.Update
   If RsMst_RRN.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_RRN.MoveNext
   If RsMst_RRN.EOF() Then
      RsMst_RRN.MoveLast
   End If
   Call InData
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_RRN.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtRRNNo.Text = "" Then
   MsgBox "Blank RRNNo Not Allowed !!! "
   TxtRRNNo.SetFocus
   Exit Sub
End If

If TxtISINID.Text = "" Then
    MsgBox "Blank ISIN ID not Allowed!!!"
    CmbISINID.SetFocus
    Exit Sub
End If

If TxtClientID.Text = "" Then
   MsgBox "Blank Client Id Not Allowed!!!"
   CmbClientID.SetFocus
   Exit Sub
End If

If TxtRRNQty.Text = "" Then
   MsgBox "Blank RRNQty Not Allowed !!! "
   TxtRRNQty.SetFocus
   Exit Sub
End If


If Edit_Flg = "A" Then
   If RsMst_RRN.RecordCount > 0 Then
      RsMst_RRN.MoveFirst
      tmp = "DepositoryID = '" & TxtDPID & "'"
      RsMst_RRN.Filter = tmp
      
      RsMst_RRN.Find "RRNNO= '" & TxtRRNNo.Text & "'"
      If Not RsMst_RRN.EOF Then
        MsgBox "Duplicate RRN No Not Allowed !!! "
         TxtRRNNo.SetFocus
         Exit Sub
      End If
   End If
      RsMst_RRN.AddNew
      RsMst_RRN!RRNID = GetRRNID
      RsMst_RRN!G_UID = GetGUID
      TxtRRNID.Text = RsMst_RRN!RRNID
Else
    RsMst_RRN.MoveFirst
    tmp = "DepositoryID = '" & TxtDPID & "'"
    RsMst_RRN.Filter = tmp
    RsMst_RRN.Find "RRNNo= '" & TxtRRNNo.Text & "'"
    If Not RsMst_RRN.EOF Then
       If RsMst_RRN!RRNID <> TxtRRNID.Text Then
          MsgBox "Duplicate RRN No Not Allowed !!! "
          TxtRRNNo.SetFocus
          Exit Sub
       End If
    End If
    If RsMst_RRN.RecordCount > 0 Then
       RsMst_RRN.MoveFirst
    End If
    tmp = "DepositoryID = '" & TxtOldDPID & "'"
    RsMst_RRN.Filter = tmp
    RsMst_RRN.Find " RRNID= " & TxtRRNID.Text
End If

RsMst_RRN!RRNID = TxtRRNID.Text
RsMst_RRN!RRNNo = TxtRRNNo.Text
RsMst_RRN!RRNDate = RRNDate.Value
RsMst_RRN!ISINID = mISINID
RsMst_RRN!ClientIDRow = mClientIDRow
RsMst_RRN!RRNQty = Val(TxtRRNQty.Text)
RsMst_RRN!DeliveredQty = Val(TxtDeliveredQty.Text)
RsMst_RRN!EntryDate = EntryDate.Value
RsMst_RRN!Flag = TxtFlag.Text
RsMst_RRN!DepositoryID = TxtDPID

If IsNull(RsMst_RRN!CreatedBy) = True Or RsMst_RRN!CreatedBy = "" Then
   RsMst_RRN!CreatedBy = Gen_UserLoginID
   RsMst_RRN!CreatedDate = Now
Else
   RsMst_RRN!UpdatedBy = Gen_UserLoginID
   RsMst_RRN!UpdatedDate = Now
End If

RsMst_RRN.Update
RsMst_RRN.Filter = ""
RsMst_RRN.Find "RRNID = " & TxtRRNID
Call InData
End Sub
Private Function GetRRNID() As Double
Dim retVal As Double
tmp = "Select max(RRNID) from Mst_RRN"
Call TmpTable
retVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
retVal = retVal + 1
GetRRNID = retVal
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 14
MSHFlexGrid1.TextMatrix(0, 0) = "RRN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "RRN No"
MSHFlexGrid1.TextMatrix(0, 2) = "RRN Date"
MSHFlexGrid1.TextMatrix(0, 3) = "ISIN No"
MSHFlexGrid1.TextMatrix(0, 4) = "Location"
MSHFlexGrid1.TextMatrix(0, 5) = "CMP"
MSHFlexGrid1.TextMatrix(0, 6) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = "Client"
MSHFlexGrid1.TextMatrix(0, 8) = "DP"
MSHFlexGrid1.TextMatrix(0, 9) = "RRN Qty in MT"
MSHFlexGrid1.TextMatrix(0, 10) = "Delivered Qty in MT"
MSHFlexGrid1.TextMatrix(0, 11) = "Entry Date"
MSHFlexGrid1.TextMatrix(0, 12) = "Flag"
MSHFlexGrid1.TextMatrix(0, 13) = "G-UID"

MSHFlexGrid1.ColWidth(0) = 1200
MSHFlexGrid1.ColWidth(1) = 2000
MSHFlexGrid1.ColWidth(2) = 1000
MSHFlexGrid1.ColWidth(3) = 2000
MSHFlexGrid1.ColWidth(4) = 2000
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1600
MSHFlexGrid1.ColWidth(7) = 1400
MSHFlexGrid1.ColWidth(8) = 1200

End Sub
Private Sub Grid_Head2()
MSHFlexGrid3.Clear
MSHFlexGrid3.Rows = 3
MSHFlexGrid3.FixedCols = 0
MSHFlexGrid3.Cols = 5

MSHFlexGrid3.TextMatrix(1, 0) = "Location ID"
MSHFlexGrid3.TextMatrix(1, 1) = "CITF No"
MSHFlexGrid3.TextMatrix(1, 2) = "CITF Date"
MSHFlexGrid3.TextMatrix(1, 3) = "CITF Qty"
MSHFlexGrid3.TextMatrix(1, 4) = "CITF No of Bags"


MSHFlexGrid3.ColWidth(0) = 2500
MSHFlexGrid3.ColWidth(1) = 2500
MSHFlexGrid3.ColWidth(2) = 2500
MSHFlexGrid3.ColWidth(3) = 2500
MSHFlexGrid3.ColWidth(4) = 2500


End Sub
Private Sub Grid_Head1()
MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.Cols = 4

MSHFlexGrid2.TextMatrix(1, 0) = "CA N0"
MSHFlexGrid2.TextMatrix(1, 1) = "CA Date"
MSHFlexGrid2.TextMatrix(1, 2) = "RRN Qty "
MSHFlexGrid2.TextMatrix(1, 3) = "CA Qty"


MSHFlexGrid2.ColWidth(0) = 2000
MSHFlexGrid2.ColWidth(1) = 2000
MSHFlexGrid2.ColWidth(2) = 2000
MSHFlexGrid2.ColWidth(3) = 2000


End Sub

Private Sub MSHFlexGrid1_Click()
Dim i As Variant
i = MSHFlexGrid1.RowSel
If i <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(i, 0) = "" Then Exit Sub
RsMst_RRN.MoveFirst
RsMst_RRN.Find "G_UID = '" & MSHFlexGrid1.TextMatrix(i, 13) & "'"
If Not RsMst_RRN.EOF Then
   Call InData
   Frame0.Visible = False
End If
End Sub
Private Sub CmdISINID_Click()
    FrmMst_ISIN.Show
End Sub

'----
Private Sub CmdClientID_Click()
    FrmMst_Client.Show
End Sub
Private Sub CmbClientDP_LostFocus()

Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst


TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow

End Sub


Private Sub ShowGridData(RSType As Variant, k As Integer)
For c = 0 To MSHFlexGrid1.Cols - 1
   MSHFlexGrid1.TextMatrix(k, c) = IIf(IsNull(RSType.Fields(c)), "", RSType.Fields(c))
Next
End Sub
Private Sub TxtClientID_LostFocus()
If SaveCmd.Enabled = False Then Exit Sub
CmbClientDP.Visible = False
If TxtClientID = "" Then Exit Sub
'If CmbExchangeTypeID.Text = "" Then Exit Sub

Dim Ans As Variant
mExchangeTypeID = GetExchangeID("NCDEX")
Ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)

If Ans(0, 0) = False Then
   MsgBox "Not Found !!!"
   TxtClientID.SetFocus
   Exit Sub
End If
If Ans(3, 0) > 1 Then
   CmbClientDP.Visible = True
End If
CmbClientDP.Text = ""
CmbClientDP.Clear
tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
      "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
      "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
      "WHERE Mst_Depository.ExchangeTypeID = " & mExchangeTypeID & _
      "  And  (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
Call Tmp1Table

For op = 1 To TmpRs1.RecordCount
    If CmbClientDP.Text = "" Then
       CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    End If
    CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
    TmpRs1.MoveNext
Next
If Ans(3, 0) <= 1 Then
   CmbClientDP_LostFocus
Else
   If CmbClientDP.Visible = True Then
      CmbClientDP.SetFocus
   End If
End If
'Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)

 
End Sub

