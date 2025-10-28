VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMst_ISIN 
   Caption         =   "ISIN Master"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      Height          =   2925
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   15105
      Begin VB.ComboBox CmbSLocation 
         Height          =   315
         Left            =   7485
         Sorted          =   -1  'True
         TabIndex        =   41
         Top             =   525
         Width           =   3030
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
         Height          =   350
         Left            =   12960
         TabIndex        =   49
         Top             =   960
         Width           =   1000
      End
      Begin VB.ComboBox CmbSCmp 
         Height          =   315
         Left            =   10560
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   525
         Width           =   4410
      End
      Begin VB.ComboBox CmbSCommodity 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   43
         Top             =   960
         Width           =   3045
      End
      Begin VB.TextBox TxtSearchISIN 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3000
         TabIndex        =   40
         Top             =   525
         Width           =   4425
      End
      Begin VB.CommandButton CmdExcel 
         Caption         =   "Excel"
         Height          =   350
         Left            =   13965
         TabIndex        =   19
         Top             =   960
         Width           =   1000
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   5775
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   12855
         _ExtentX        =   22675
         _ExtentY        =   10186
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   300
         Left            =   1575
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   525
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   63766529
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   300
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   525
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
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
         Format          =   63766529
         CurrentDate     =   36494
      End
      Begin VB.Label Label45 
         Caption         =   "To"
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
         Left            =   2092
         TabIndex        =   59
         Top             =   240
         Width           =   300
      End
      Begin VB.Label Label19 
         Caption         =   "FED From Date"
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
         Left            =   307
         TabIndex        =   58
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label4 
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
         Left            =   8663
         TabIndex        =   50
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "CMP"
         Height          =   300
         Left            =   12570
         TabIndex        =   48
         Top             =   180
         Width           =   405
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   990
         Width           =   1170
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
         Left            =   4837
         TabIndex        =   44
         Top             =   180
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5025
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   7245
      Begin VB.CommandButton CmdNCommodityID 
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
         Height          =   285
         Left            =   6675
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3810
         Width           =   375
      End
      Begin VB.ComboBox CmbNCommodityID 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   3810
         Width           =   4590
      End
      Begin VB.TextBox TxtUpdated 
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
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   4560
         Width           =   4575
      End
      Begin VB.TextBox TxtCreated 
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
         Height          =   375
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4155
         Width           =   4575
      End
      Begin VB.TextBox TxtISINGrade 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         TabIndex        =   7
         Top             =   2985
         Width           =   4560
      End
      Begin VB.CommandButton CmdCMPID 
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
         Height          =   285
         Left            =   6720
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   990
         Width           =   375
      End
      Begin VB.CommandButton CmdCommodityID 
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
         Height          =   285
         Left            =   6720
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2190
         Width           =   375
      End
      Begin VB.ComboBox CmbCommodityID 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   2190
         Width           =   4590
      End
      Begin VB.ComboBox CmbCMPID 
         Height          =   315
         Left            =   2040
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   990
         Width           =   4575
      End
      Begin VB.TextBox TxtISINID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Height          =   360
         Left            =   2040
         TabIndex        =   1
         Top             =   240
         Width           =   4560
      End
      Begin VB.TextBox TxtISINDescription 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         MaxLength       =   255
         TabIndex        =   2
         Top             =   615
         Width           =   4560
      End
      Begin VB.TextBox TxtCMPID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   990
         Width           =   4560
      End
      Begin MSComCtl2.DTPicker NE_Date 
         Height          =   360
         Left            =   2040
         TabIndex        =   4
         Top             =   1785
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   63766529
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker FED 
         Height          =   360
         Left            =   5295
         TabIndex        =   5
         Top             =   1800
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Format          =   63766529
         CurrentDate     =   39864
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1380
         Width           =   4560
      End
      Begin VB.TextBox TxtCommodityID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2040
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   2190
         Width           =   4560
      End
      Begin VB.TextBox TxtExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2580
         Width           =   4560
      End
      Begin VB.TextBox TxtFlag 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   360
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   46
         Top             =   3390
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.TextBox TxtNCommodityID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2025
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3810
         Width           =   4560
      End
      Begin VB.Label Label5 
         Caption         =   "New Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   55
         Top             =   3855
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Updated :"
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   4620
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "Created  :"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   4215
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Grade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   9
         Top             =   3060
         Width           =   1800
      End
      Begin VB.Label LblISINID 
         Caption         =   "ISIN ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   21
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label LblISINDescription 
         Caption         =   "ISIN Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   22
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label LblCMPID 
         Caption         =   "CMP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   23
         Top             =   1042
         Width           =   1800
      End
      Begin VB.Label LblNE_Date 
         Caption         =   "NED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   25
         Top             =   1853
         Width           =   1650
      End
      Begin VB.Label LblFED 
         Caption         =   "FED"
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
         Left            =   4665
         TabIndex        =   26
         Top             =   1830
         Width           =   360
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
         Height          =   225
         Left            =   240
         TabIndex        =   27
         Top             =   1448
         Width           =   1800
      End
      Begin VB.Label LblCommodityID 
         Caption         =   "Commodity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   29
         Top             =   2227
         Width           =   1800
      End
      Begin VB.Label LblExchangeTypeID 
         Caption         =   "Exchange Type"
         Height          =   240
         Left            =   240
         TabIndex        =   31
         Top             =   2640
         Width           =   1800
      End
      Begin VB.Label LblFlag 
         Caption         =   "Flag"
         Enabled         =   0   'False
         Height          =   240
         Left            =   255
         TabIndex        =   33
         Top             =   3450
         Visible         =   0   'False
         Width           =   1800
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   34
      Top             =   5040
      Width           =   7245
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   5685
         TabIndex        =   11
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   5685
         TabIndex        =   35
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   4290
         TabIndex        =   13
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   4290
         TabIndex        =   14
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   2895
         TabIndex        =   15
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   2895
         TabIndex        =   16
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   1500
         TabIndex        =   17
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   1500
         TabIndex        =   10
         Top             =   240
         Width           =   1400
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   18
         Top             =   660
         Width           =   1400
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   100
         TabIndex        =   45
         Top             =   240
         Width           =   1400
      End
   End
End
Attribute VB_Name = "FrmMst_ISIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Edit_Flg, mISINDescription, mSplitDescription As Variant
Dim mCMPName, mCMPID, mLocationID, mCommodityId, mNCommodityID, mExchangeTypeID As Double
Private Sub CmbSCmp_GotFocus()
tmp = CmbSCmp.Text
Call TableMst_CMP
CmbSCmp.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
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

mCommodityId = RsMst_Commodity!CommodityID

End Sub

Private Sub CmbSLocation_GotFocus()
tmp = CmbSLocation.Text
Call TableMst_Location
CmbSLocation.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
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

Private Sub CmdNCommodityID_Click()
    FrmMst_Commodity.Show
End Sub

Private Sub CmdSearch_Click()
Dim x As Integer
Dim Wc As String

Wc = ""
RsMst_ISIN.MoveFirst
Call Grid_Head

Wc = " Where i.FED between '" & SFrom.Value & "' And '" & STo.Value & "'"

If TxtSearchISIN <> "" Then
   'tmp = "Select * from Mst_ISIN where ISINID like '%" & Trim(TxtSearchISIN) & "%'"
   Wc = Wc & " And ISINID like '%" & Trim(TxtSearchISIN) & "%' "
Else
    If CmbSCmp.Text <> "" Then
       Wc = GenWc(Wc, mCMPName, "c.CMPID", "N", "N")
    End If
    If CmbSLocation.Text <> "" Then
       Wc = GenWc(Wc, mLocationID, "l.LocationID", "N", "N")
    End If
    If CmbSCommodity.Text <> "" Then
       Wc = GenWc(Wc, mCommodityId, "com.CommodityID", "N", "N")
    End If
End If
tmp = "select i.ISINID,i.ISINDescription,c.CMPName,i.NE_Date,i.FED,l.LocationName,com.Commodity,i.Grade "
tmp = tmp & "  from Mst_ISIN I inner join Mst_CMP c on  i.CMPID = c.CMPID "
tmp = tmp & " inner join Mst_location l on  c.LocationID = l.LocationID "
tmp = tmp & " inner join Mst_Commodity com on i.CommodityID = com.CommodityID "
tmp = tmp & " inner join Mst_ExchangeType et on i.ExchangeTypeID = et.ExchangeTypeID "
tmp = tmp & Wc
tmp = tmp & " order by i.ISINID,c.CMPName,com.Commodity "
Call Tmp4Table

If TmpRs4.RecordCount <= 0 Then
   MsgBox "No Such ISIN Exist in DataBase!!!"
   TxtSearchISIN.SetFocus
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
Call Grid_Head
SFrom.Value = Date - 30
STo.Value = Date
TxtSearchISIN = ""
CmbSCmp.Text = ""
CmbSCommodity.Text = ""
mCMPName = 0
mCommodityId = 0

Call TableMst_ISIN
If RsMst_ISIN.RecordCount = 0 Then
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
Call Indata
End Sub
Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_ISIN.xls")
Call Xls_Detail_Rpt
End Sub
Public Sub Xls_Detail_Rpt()
Dim oXL As Excel.Application
Dim oWB As Excel.Workbook
Dim mRow As Double
Dim mCol As Double
MsgBox ("Wait till next message")
Set oXL = New Excel.Application
Dim Pat As String
Pat = App.Path
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_ISIN.xls")
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
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsMst_ISIN.MoveNext
If RsMst_ISIN.EOF Then
   RsMst_ISIN.MoveLast
   LF_Flag = "Y"
End If
Call Indata
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
RsMst_ISIN.MovePrevious
If RsMst_ISIN.BOF Then
   RsMst_ISIN.MoveFirst
   LF_Flag = "Y"
End If
Call Indata
If LF_Flag = "Y" Then
    PreviousCmd.Enabled = False
    FirstCmd.Enabled = False
End If
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub FirstCmd_Click()
RsMst_ISIN.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub
Private Sub LastCmd_Click()
RsMst_ISIN.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub ExitCmd_Click()
Unload Me
End Sub
Private Sub SearchCmd_Click()
Dim x As Integer
If SearchCmd.Caption = "Cancel" Then
  If RsMst_ISIN.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   'Call FirstCmd_Click
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 850
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top - 100
Frame0.Width = Me.Width - 150
Frame0.Visible = True
MSHFlexGrid1.Width = Frame0.Width - 200
MSHFlexGrid1.Height = Frame0.Height - 1600
RsMst_ISIN.MoveFirst

''===========comment on 7 july 2012 for not clearing search grid=============
'MSHFlexGrid1.Clear
'MSHFlexGrid1.Cols = 3
''Call Grid_Head
''TxtSearchISIN = ""
''CmbSCmp.Text = ""
''CmbSCommodity.Text = ""
''mCMPName = 0
''mCommodityID = 0

'MSHFlexGrid1.Rows = RsMst_ISIN.RecordCount + 1
'For x = 1 To RsMst_ISIN.RecordCount
'     Call ShowGrid
'     RsMst_ISIN.MoveNext
'Next
'TxtSearchISIN.SetFocus
End Sub
Public Sub Indata()

TxtISINID.Locked = True
TxtISINDescription.Locked = True
TxtCMPID.Locked = True
CmbCMPID.Visible = False
CmdCMPID.Visible = False
TxtCMPID.Visible = True
NE_Date.Enabled = False
FED.Enabled = False
TxtLocationID.Locked = True

TxtNCommodityID.Locked = True
CmbNCommodityID.Visible = False
CmdNCommodityID.Visible = False
TxtNCommodityID.Visible = True

TxtLocationID.Visible = True
TxtCommodityID.Locked = True
CmbCommodityID.Visible = False
CmdCommodityID.Visible = False
TxtCommodityID.Visible = True
TxtExchangeTypeID.Locked = True

TxtExchangeTypeID.Visible = True
TxtISINGrade.Locked = True
TxtFlag.Locked = True

TxtISINID = IIf(IsNull(RsMst_ISIN!ISINID), "", RsMst_ISIN!ISINID)
TxtISINDescription = IIf(IsNull(RsMst_ISIN!ISINDescription), "", RsMst_ISIN!ISINDescription)

mCMPID = IIf(IsNull(RsMst_ISIN!CMPID), "", RsMst_ISIN!CMPID)

TxtCMPID = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)

NE_Date = IIf(IsNull(RsMst_ISIN!NE_Date), "", RsMst_ISIN!NE_Date)
FED = IIf(IsNull(RsMst_ISIN!FED), "", RsMst_ISIN!FED)

mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)

TxtLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)

mCommodityId = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)

TxtCommodityID = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeID), "", RsMst_ISIN!ExchangeTypeID)

TxtExchangeTypeID = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)

TxtISINGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
TxtFlag = IIf(IsNull(RsMst_ISIN!Flag), "", RsMst_ISIN!Flag)


TxtCreated = IIf(IsNull(RsMst_ISIN!CreatedBy), "", RsMst_ISIN!CreatedBy) & " :- " & IIf(IsNull(RsMst_ISIN!CreatedDate), "", RsMst_ISIN!CreatedDate)
TxtUpdated = IIf(IsNull(RsMst_ISIN!UpdatedBy), "", RsMst_ISIN!UpdatedBy) & " :- " & IIf(IsNull(RsMst_ISIN!UpdatedDate), "", RsMst_ISIN!UpdatedDate)

If RsMst_ISIN.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If

EditCmd.Enabled = False
If LCase(Gen_UserRole) = "power" Or LCase(Gen_UserRole) = "superadmin" Then
   If mCommodityId = 31 Or mCommodityId = 114 Then
      EditCmd.Enabled = True
   End If
End If
End Sub

Private Sub TxtISINDescription_LostFocus()
TxtISINDescription = Trim(TxtISINDescription)
End Sub

Private Sub TxtISINGrade_LostFocus()
TxtISINGrade = Trim(TxtISINGrade)
End Sub

Private Sub TxtISINID_LostFocus()
If TxtISINID = "" Then Exit Sub
TxtISINID = Trim(TxtISINID)
If RsMst_ISIN.RecordCount = 0 Then Exit Sub
RsMst_ISIN.MoveFirst
RsMst_ISIN.Find "ISINID = '" & TxtISINID & "'"
If Not RsMst_ISIN.EOF Then
   Call Indata
   Exit Sub
End If


End Sub

Private Sub AddCmd_Click()
Edit_Flg = "A"
Call GButtonLock(Me, False)
TxtISINID.Enabled = True
TxtISINID.Locked = False
TxtISINDescription.Locked = False
TxtCMPID.Locked = False
CmbCMPID.Visible = True
CmdCMPID.Visible = True
TxtCMPID.Visible = False
NE_Date.Enabled = True
FED.Enabled = True
TxtLocationID.Locked = False
TxtLocationID.Visible = True
TxtCommodityID.Locked = False
CmbCommodityID.Visible = True
CmdCommodityID.Visible = True
TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = True
TxtExchangeTypeID.Visible = True
TxtISINGrade.Locked = False
TxtFlag.Locked = False
TxtISINID = ""
TxtISINDescription = ""
TxtCMPID = ""
CmbCMPID = ""
NE_Date = Date
FED = Date
TxtCreated = ""
TxtUpdated = ""

TxtLocationID = ""
TxtCommodityID = ""
CmbCommodityID = ""
TxtExchangeTypeID = "NCDEX"
TxtISINGrade = ""
TxtFlag = ""

TxtNCommodityID.Locked = True
CmbNCommodityID.Visible = False
CmdNCommodityID.Visible = False
TxtNCommodityID.Visible = True

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Call GButtonLock(Me, False)
TxtISINID.Enabled = False
TxtISINID.Locked = True
TxtISINDescription.Locked = False
'TxtCMPID.Locked = False
'CmbCMPID.Visible = True
'CmdCMPID.Visible = True
'TxtCMPID.Visible = False
NE_Date.Enabled = True
FED.Enabled = True
TxtLocationID.Locked = True
TxtLocationID.Visible = True
'TxtCommodityID.Locked = False
'CmbCommodityID.Visible = True
'CmdCommodityID.Visible = True
'TxtCommodityID.Visible = False
TxtExchangeTypeID.Locked = True
TxtExchangeTypeID.Visible = True
TxtISINGrade.Locked = False
TxtFlag.Locked = False

If mCommodityId = 31 Then
   CmbNCommodityID.Visible = True
   CmdNCommodityID.Visible = True
   CmbNCommodityID.Text = "Chilli - 25 KG bag"
   TxtNCommodityID.Visible = False
   CmbNCommodityID.Enabled = False
   mNCommodityID = 114
ElseIf mCommodityId = 114 Then
    CmbNCommodityID.Visible = True
    CmdNCommodityID.Visible = True
    CmbNCommodityID.Text = "Chilli"
    TxtNCommodityID.Visible = False
    CmbNCommodityID.Enabled = False
    mNCommodityID = 31
End If

TxtISINDescription.SetFocus
'CmbCMPID.Text = TxtCMPID
'CmbCommodityID.Text = TxtCommodityID

End Sub
Private Function CheckForUseage() As Boolean
Dim Retval As Boolean

Retval = True

tmp = "select 'GSL' As  'Type', Mst_GSL.ISINID  from Mst_GSL where Mst_GSL.ISINID = '" & TxtISINID & "'"
tmp = tmp & " union "
tmp = tmp & " select 'GSL-First' As  'Type', Mst_GSL.FirstISINID  from Mst_GSL where Mst_GSL.FirstISINID = '" & TxtISINID & "' "
tmp = tmp & " union"
tmp = tmp & " select 'Mst_RRN' As  'Type', Mst_RRN.ISINID  from Mst_RRN where Mst_RRN.ISINID = '" & TxtISINID & "' "
tmp = tmp & " union"
tmp = tmp & " select 'Txn_PledgeCreation' As  'Type', Txn_PledgeCreation.ISINID  from Txn_PledgeCreation where Txn_PledgeCreation.ISINID = '" & TxtISINID & "' "
If Gen_DBType = "MDB" Then
   Call TmpTable
   If TmpRs.RecordCount > 0 Then
      Retval = False
   End If
Else
   Call Tmp4Table
   If TmpRs4.RecordCount > 0 Then
      Retval = False
   End If
End If
CheckForUseage = Retval
End Function

Private Sub DeleteCmd_Click()
Dim ans As Variant

If CheckForUseage() = False Then
   MsgBox "You can't delete this record since it is used in Transaction !!!  "
   Exit Sub
End If

On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then
   RsMst_ISIN.Delete
   RsMst_ISIN.Update
   If RsMst_ISIN.RecordCount = 0 Then
      Call AddCmd_Click
      Exit Sub
   End If
   RsMst_ISIN.MoveNext
   If RsMst_ISIN.EOF() Then
      RsMst_ISIN.MoveLast
   End If
   Call Indata
End If
Exit Sub
msgError:
MsgBox "Child record Present "
RsMst_ISIN.CancelUpdate
End Sub
Private Sub SaveCmd_Click()
If TxtISINID.Text = "" Then
   MsgBox "Blank ISINID not Allowed !!!"
   TxtISINID.SetFocus
   Exit Sub
End If

If TxtISINDescription.Text = "" Then
   MsgBox "Blank ISIN Description Not Allowed!!!"
   TxtISINDescription.SetFocus
   Exit Sub
End If

If CmbCMPID.Visible = True And CmbCMPID.Text = "" Then
  MsgBox "Blank CMP Not Allowed!!!"
  CmbCMPID.SetFocus
  Exit Sub
End If

If CmbCommodityID.Visible = True And CmbCommodityID.Text = "" Then
  MsgBox "Blank Commodity Not Allowed!!!"
  CmbCommodityID.SetFocus
  Exit Sub
End If

If CmbNCommodityID.Visible = True And CmbNCommodityID.Text = "" Then
  MsgBox "Blank New Commodity Not Allowed!!!"
  CmbNCommodityID.SetFocus
  Exit Sub
End If

If CmbNCommodityID.Visible = True Then
   If TxtCommodityID.Text = CmbNCommodityID.Text Then
      MsgBox "New Commodity should not same as commodity!!"
      CmbNCommodityID.SetFocus
      Exit Sub
   End If
End If

If TxtISINGrade.Text = "" Then
  MsgBox "Blank Grade Not Allowed!!!"
  TxtISINGrade.SetFocus
  Exit Sub
End If

If FED.Value < NE_Date.Value Then
    MsgBox "FED Must Be Greater Than NED Always!!!"
    FED.SetFocus
    Exit Sub
End If

If Edit_Flg = "A" Then
   If RsMst_ISIN.RecordCount > 0 Then
      RsMst_ISIN.MoveFirst
      RsMst_ISIN.Find "ISINID=  '" & TxtISINID.Text & "'"
      If Not RsMst_ISIN.EOF Then
        MsgBox "Duplicate ISINID Not Allowed !!! "
         TxtISINID.SetFocus
         Exit Sub
      End If
   End If
      RsMst_ISIN.AddNew
      mISINDescription = Trim(CmbCMPID.Text) & "~" & Trim(CmbCommodityID.Text) & "~" & Trim(TxtISINGrade) & "~" & Trim(TxtISINDescription)
      RsMst_ISIN!G_UID = GetGUID
'      RsMst_ISIN!ISINID = GetISINID
'      TxtISINID.Text = RsMst_ISIN!ISINID
Else
    RsMst_ISIN.MoveFirst
    RsMst_ISIN.Find "ISINID = '" & TxtISINID.Text & "'"
    If Not RsMst_ISIN.EOF Then
       If RsMst_ISIN!ISINID <> TxtISINID.Text Then
          MsgBox "Duplicate ISIN ID  Not Allowed !!! "
          TxtISINID.SetFocus
          Exit Sub
       End If
    End If
    RsMst_ISIN.MoveFirst
    RsMst_ISIN.Find " ISINID= '" & TxtISINID.Text & "'"
    mSplitDescription = Split(TxtISINDescription, "~")
    'mISINDescription = Trim(CmbCMPID.Text) & "~" & Trim(CmbCommodityID.Text) & "~" & Trim(TxtISINGrade) & "~" & mSplitDescription(3)
    'mISINDescription = Trim(TxtCMPID.Text) & "~" & Trim(TxtCommodityID.Text) & "~" & Trim(TxtISINGrade) & "~" & mSplitDescription(3)
    mISINDescription = TxtISINDescription
End If

RsMst_ISIN!ISINID = UCase(TxtISINID.Text)
RsMst_ISIN!ISINDescription = mISINDescription 'Trim(CmbCMPID.Text) & "~" & Trim(CmbCommodityID.Text) & "~" & Trim(TxtISINGrade) & "~" & mISINDescription(3)
RsMst_ISIN!CMPID = mCMPID
RsMst_ISIN!NE_Date = NE_Date.Value
RsMst_ISIN!FED = FED.Value
RsMst_ISIN!LocationID = mLocationID

If CmbNCommodityID.Visible = True And CmbNCommodityID.Text <> "" Then
   RsMst_ISIN!CommodityID = mNCommodityID
Else
   RsMst_ISIN!CommodityID = mCommodityId
End If

RsMst_ISIN!ExchangeTypeID = GetExchangeID(TxtExchangeTypeID)
RsMst_ISIN!Flag = "0" 'TxtFlag.Text
RsMst_ISIN!Grade = TxtISINGrade
If IsNull(RsMst_ISIN!CreatedBy) = True Or RsMst_ISIN!CreatedBy = "" Then
   RsMst_ISIN!CreatedBy = Gen_UserLoginID
   RsMst_ISIN!CreatedDate = Now
Else
   RsMst_ISIN!UpdatedBy = Gen_UserLoginID
   RsMst_ISIN!UpdatedDate = Now
End If

RsMst_ISIN.Update
If CmbNCommodityID.Visible = True And CmbNCommodityID.Text <> "" Then
   Call AddLog_ISINCommodity
End If

Call Indata
End Sub
Private Function GetISINID() As Double
Dim Retval As Double
tmp = "Select max(ISINID) from Mst_ISIN"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetISINID = Retval
End Function
Private Sub Grid_Head()
MSHFlexGrid1.Clear
MSHFlexGrid1.Rows = 2
MSHFlexGrid1.Cols = 8
MSHFlexGrid1.TextMatrix(0, 0) = "ISIN ID"
MSHFlexGrid1.TextMatrix(0, 1) = "ISIN Description"
MSHFlexGrid1.TextMatrix(0, 2) = "CMP"
MSHFlexGrid1.TextMatrix(0, 3) = "NED"
MSHFlexGrid1.TextMatrix(0, 4) = "FED"
MSHFlexGrid1.TextMatrix(0, 5) = "Location"
MSHFlexGrid1.TextMatrix(0, 6) = "Commodity"
MSHFlexGrid1.TextMatrix(0, 7) = "Grade"


MSHFlexGrid1.ColWidth(0) = 1600
MSHFlexGrid1.ColWidth(1) = 4000
MSHFlexGrid1.ColWidth(2) = 2600
MSHFlexGrid1.ColWidth(3) = 1000
MSHFlexGrid1.ColWidth(4) = 1000
MSHFlexGrid1.ColWidth(5) = 1600
MSHFlexGrid1.ColWidth(6) = 1400
MSHFlexGrid1.ColWidth(7) = 1400

End Sub
Private Sub MSHFlexGrid1_Click()
Dim I As Variant
I = MSHFlexGrid1.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid1.TextMatrix(I, 0) = "" Then Exit Sub
RsMst_ISIN.MoveFirst
RsMst_ISIN.Find "ISINID = '" & MSHFlexGrid1.TextMatrix(I, 0) & "'"
If Not RsMst_ISIN.EOF Then
   Call Indata
   Frame0.Visible = False
End If
End Sub
Private Sub CmdCMPID_Click()
    FrmMst_CMP.Show
End Sub
'----
Private Sub CmbCMPID_GotFocus()
tmp = CmbCMPID.Text
Call TableMst_CMP
CmbCMPID.Clear
If RsMst_CMP.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_CMP.RecordCount
    CmbCMPID.AddItem RsMst_CMP!CMPName
    RsMst_CMP.MoveNext
Next
CmbCMPID.Text = tmp
End Sub
'----
Private Sub CmbCMPID_LostFocus()
If CmbCMPID.Text = "" Then
   Exit Sub
End If
RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
mCMPID = RsMst_CMP!CMPID
mLocationID = RsMst_CMP!LocationID
TxtLocationID = GetLocationName(RsMst_CMP!LocationID)
End Sub
'----
Private Sub CmdCommodityID_Click()
    FrmMst_Commodity.Show
End Sub
'----
Private Sub CmbCommodityID_GotFocus()
tmp = CmbCommodityID.Text
Call TableMst_Commodity
CmbCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbCommodityID.Text = tmp
End Sub
'----
Private Sub CmbCommodityID_LostFocus()
If CmbCommodityID.Text = "" Then
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbCommodityID.SetFocus
   Exit Sub
End If
mCommodityId = RsMst_Commodity!CommodityID
End Sub
'----

Private Sub CmbNCommodityID_GotFocus()
tmp = CmbNCommodityID.Text
Call TableMst_Commodity(" Where CommodityID in (114,31)")
CmbNCommodityID.Clear
If RsMst_Commodity.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Commodity.RecordCount
    CmbNCommodityID.AddItem RsMst_Commodity!Commodity
    RsMst_Commodity.MoveNext
Next
CmbNCommodityID.Text = tmp
End Sub
'----
Private Sub CmbNCommodityID_LostFocus()
If CmbNCommodityID.Text = "" Then
   mNCommodityID = 0
   Exit Sub
End If
RsMst_Commodity.MoveFirst
RsMst_Commodity.Find "Commodity = '" & CmbNCommodityID.Text & "'"
If RsMst_Commodity.EOF Then
   MsgBox "Invalid selection "
   CmbNCommodityID.SetFocus
   Exit Sub
End If
mNCommodityID = RsMst_Commodity!CommodityID
End Sub

Private Sub AddLog_ISINCommodity()
Call TableLog_ISINCommodity(" Where ISINID='" & TxtISINID.Text & "'")
RsLog_ISINCommodity.AddNew
RsLog_ISINCommodity!ISINID = TxtISINID
RsLog_ISINCommodity!OldCommodityID = mCommodityId
RsLog_ISINCommodity!NewCommodityID = mNCommodityID
RsLog_ISINCommodity!UpdatedDate = Date
RsLog_ISINCommodity!UpdatedBy = Gen_UserLoginID
RsLog_ISINCommodity.Update
RsLog_ISINCommodity.Requery
End Sub



