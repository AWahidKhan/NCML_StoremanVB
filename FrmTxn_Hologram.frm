VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTxn_Hologram 
   Caption         =   "Hologram"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8910
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame0 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   14955
      Begin VB.TextBox TxtHologramNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6360
         MaxLength       =   12
         TabIndex        =   11
         Top             =   600
         Width           =   1755
      End
      Begin VB.CommandButton CmdSearch 
         Caption         =   "Search"
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
         Left            =   8400
         TabIndex        =   13
         Top             =   600
         Width           =   1000
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
         Left            =   9420
         TabIndex        =   15
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox CmbSLocationID 
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
         ItemData        =   "FrmTxn_Hologram.frx":0000
         Left            =   3210
         List            =   "FrmTxn_Hologram.frx":0002
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   3000
      End
      Begin MSComCtl2.DTPicker STo 
         Height          =   375
         Left            =   1695
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   9895937
         CurrentDate     =   36494
      End
      Begin MSComCtl2.DTPicker SFrom 
         Height          =   375
         Left            =   135
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   555
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   9895937
         CurrentDate     =   36494
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   1455
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2566
         _Version        =   393216
         WordWrap        =   -1  'True
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
      Begin VB.Label Label1 
         Caption         =   "Hologram No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6720
         TabIndex        =   44
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label45 
         Caption         =   "To Date"
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
         Left            =   1995
         TabIndex        =   17
         Top             =   255
         Width           =   855
      End
      Begin VB.Label Label44 
         Caption         =   "From Date"
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
         Left            =   360
         TabIndex        =   16
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label28 
         Caption         =   "Total Records Found :"
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
         Left            =   10740
         TabIndex        =   14
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label22 
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
         Left            =   13215
         TabIndex        =   12
         Top             =   615
         Width           =   90
      End
      Begin VB.Label Label23 
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
         Left            =   4320
         TabIndex        =   9
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4170
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   7935
      Begin VB.TextBox TxtTotalQty 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1785
         Width           =   1755
      End
      Begin VB.TextBox TxtHologramID 
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   285
         Width           =   1755
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
         ItemData        =   "FrmTxn_Hologram.frx":0004
         Left            =   2040
         List            =   "FrmTxn_Hologram.frx":0006
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   785
         Width           =   5535
      End
      Begin VB.TextBox TxtRemark 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   2040
         MaxLength       =   1000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2285
         Width           =   5535
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
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3585
         Width           =   5535
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
         Height          =   360
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3085
         Width           =   5535
      End
      Begin VB.TextBox TxtFromNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         MaxLength       =   12
         TabIndex        =   3
         Top             =   1285
         Width           =   1755
      End
      Begin VB.TextBox TxtToNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   6000
         MaxLength       =   12
         TabIndex        =   4
         Top             =   1285
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker TxtHologramDate 
         Height          =   360
         Left            =   6060
         TabIndex        =   1
         Top             =   285
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   635
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
         Format          =   9895937
         CurrentDate     =   39884
      End
      Begin VB.TextBox TxtLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   785
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
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
         TabIndex        =   43
         Top             =   1838
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Hologram Date"
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
         Left            =   4080
         TabIndex        =   42
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Hologram ID"
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
         Left            =   240
         TabIndex        =   41
         Top             =   315
         Width           =   1335
      End
      Begin VB.Label LblLocationID 
         Caption         =   "Location"
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
         Left            =   240
         TabIndex        =   40
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label5 
         Caption         =   "Remark"
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
         Left            =   240
         TabIndex        =   39
         Top             =   2488
         Width           =   960
      End
      Begin VB.Label Label25 
         Caption         =   "Updated :"
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
         TabIndex        =   38
         Top             =   3638
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "Created  :"
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
         Left            =   240
         TabIndex        =   36
         Top             =   3145
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "From Hologram No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   240
         TabIndex        =   34
         Top             =   1245
         Width           =   1800
      End
      Begin VB.Label Label11 
         Caption         =   "To Hologram No."
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
         TabIndex        =   32
         Top             =   1315
         Width           =   1860
      End
   End
   Begin VB.Frame ButtonFrm 
      Height          =   1155
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   7920
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New "
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
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
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
         Left            =   1695
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
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
         Left            =   3150
         TabIndex        =   23
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
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
         Left            =   4725
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "&Search"
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
         Left            =   6180
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
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
         Left            =   240
         TabIndex        =   29
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
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
         Left            =   1695
         TabIndex        =   31
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
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
         Left            =   3150
         TabIndex        =   33
         Top             =   615
         Width           =   1575
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
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
         Left            =   4725
         TabIndex        =   35
         Top             =   615
         Width           =   1455
      End
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
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
         Left            =   6180
         TabIndex        =   37
         Top             =   615
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmTxn_Hologram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mVendorID, mLocationID, mCommodityID, mSVendorID, mSLocationID, mSCommodityID As Variant
Dim mCMPID As Variant
Dim Edit_Flg As Variant

Private Sub AddCmd_Click()

Edit_Flg = "A"

Frame1.Enabled = True
TxtHologramID.Text = ""

CmbLocationID.Text = ""
TxtLocationID.Text = ""

TxtLocationID.Visible = False
CmbLocationID.Visible = True

TxtFromNo.Text = ""
TxtToNo.Text = ""
TxtTotalQty.Text = ""

TxtRemark.Text = ""
TxtCreated = ""
TxtUpdated = ""

TxtHologramDate.Value = Date

Call GButtonLock(Me, False)
If RsTxn_Hologram.RecordCount > 0 Then
   TxtHologramDate.SetFocus
End If
End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location '(" Where LocationID <> 0 ")
End If
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp
End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbLocationID.SetFocus
   Exit Sub
End If

mLocationID = RsMst_Location!LocationID
End Sub


Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Txn_Hologram.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Txn_Hologram.xls")
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
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub



Private Sub CmdSearch_Click()
Dim Wc1 As Variant
Call Grid_Head1

If SFrom.Value > STo.Value Then
   MsgBox "'From Date' Must be less than 'To Date'!!!"
   SFrom.SetFocus
   Exit Sub
End If

tmp = "Select HologramID,HologramDate,"
tmp = tmp & " LocationName , FromHologramNo, ToHologramNo, TotalQty, Remark"
tmp = tmp & " From Txn_Hologram TF"
tmp = tmp & " Inner Join Mst_Location ML on TF.LocationID=ML.LocationID"

If Gen_DBType = "MDB" Then
   Wc1 = " Where HologramDate Between #" & Format(SFrom.Value, Gen_DatFormat) & "# And #" & Format(STo.Value, Gen_DatFormat) & "# "
Else
   Wc1 = " Where HologramDate Between '" & Format(SFrom.Value, Gen_DatFormat) & "' And '" & Format(STo.Value, Gen_DatFormat) & "' "
End If
If CmbSLocationID.Text <> "" Then
   Wc1 = Wc1 & " And LocationName='" & CmbSLocationID.Text & "'"
Else
   If Gen_WcLocation <> "" Then
      Wc1 = Wc1 & " And ML." & Gen_WcLocation
   End If
End If

If TxtHologramNo.Text <> "" Then
   Wc1 = Wc1 & " And FromHologramNo<=" & Val(TxtHologramNo.Text) & " And ToHologramNo>=" & Val(TxtHologramNo.Text) & ""
End If
tmp = tmp & Wc1 & " Order By FromHologramNo"

Call Tmp1Table

If TmpRs1.RecordCount > 0 Then
   MSHFlexGrid2.Rows = TmpRs1.RecordCount + 2
   For I = 1 To TmpRs1.RecordCount
       For C = 0 To MSHFlexGrid2.Cols - 1
           MSHFlexGrid2.TextMatrix(I, C) = IIf(IsNull(TmpRs1.Fields(C)), "", TmpRs1.Fields(C))
       Next
       MSHFlexGrid2.TextMatrix(I, 1) = Format(MSHFlexGrid2.TextMatrix(I, 1), "dd-MMM-yyyy")
       TmpRs1.MoveNext
   Next
   
End If

Label22.Caption = TmpRs1.RecordCount
Label22.Refresh


End Sub

Private Sub DeleteCmd_Click()

'Dim ans As Variant
'On Error GoTo msgError
'
'ans = MsgBox("Do you really want to DELETE this record??? ", vbYesNo)
'If ans = vbYes Then
'
'   RsTxn_Hologram.Delete
'   RsTxn_Hologram.Update
'   If RsTxn_Hologram.RecordCount = 0 Then
'      Call AddCmd_Click
'      Exit Sub
'   End If
'   RsTxn_Hologram.MoveNext
'   If RsTxn_Hologram.EOF() Then
'      RsTxn_Hologram.MoveLast
'   End If
'   Call InData
'End If
'
'Exit Sub
'msgError:
'MsgBox "Child record Present "
'RsTxn_Hologram.CancelUpdate


End Sub

Private Sub EditCmd_Click()

Edit_Flg = "E"

Call GButtonLock(Me, False)

Frame1.Enabled = True
'TxtHologramDate.Enabled = False

CmbLocationID.Text = TxtLocationID.Text
CmbLocationID.Visible = True
TxtLocationID.Visible = False
End Sub

Private Sub ExitCmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim wc As Variant

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Frame1.Enabled = False
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

Call Grid_Head1
SFrom.Value = Date - 365
STo.Value = Date
Label22.Caption = ""
CmbSLocationID.Text = ""
TxtHologramNo.Text = ""

Frame1.Enabled = True
If Gen_WcLocation <> "" Then
   Call TableTxn_Hologram(" Where " & Gen_WcLocation & " Order by HologramID")
Else
   Call TableTxn_Hologram(" Order by HologramID")
End If

If RsTxn_Hologram.RecordCount = 0 Then
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


Private Sub MSHFlexGrid2_Click()
Dim I As Variant
I = MSHFlexGrid2.RowSel
If I <= 0 Then Exit Sub
If MSHFlexGrid2.TextMatrix(I, 0) = "" Then Exit Sub
RsTxn_Hologram.MoveFirst
RsTxn_Hologram.Find "HologramID= " & Val(MSHFlexGrid2.TextMatrix(I, 0)) & ""
If Not RsTxn_Hologram.EOF Then
   Call Indata
   Frame0.Visible = False
End If

End Sub

Private Sub SaveCmd_Click()

If TxtHologramDate.Value > Date Then
   MsgBox "Future Date not allowed!!"
   TxtHologramDate.SetFocus
   Exit Sub
End If

If CmbLocationID.Text = "" Then
   MsgBox "Select Location Name!!"
   CmbLocationID.SetFocus
   Exit Sub
End If

If TxtFromNo.Text = "" Then
   MsgBox "Blank From Hologram No. not Allowed!!"
   TxtFromNo.SetFocus
   Exit Sub
End If

If TxtToNo.Text = "" Then
   MsgBox "Blank To Hologram No. not Allowed!!"
   TxtToNo.SetFocus
   Exit Sub
End If

TxtTotalQty.Text = Val(TxtToNo.Text) - Val(TxtFromNo.Text) + 1
If Val(TxtTotalQty.Text) <= 0 Then
   MsgBox "'To Hologram No' must be greater than 'From Hologram No'!!!"
   TxtToNo.SetFocus
   Exit Sub
End If

tmp = "Select HologramID,Hologramdate,LocationID,FromHologramNo,ToHologramNo From Txn_Hologram"
''tmp = tmp & " Where LocationID = " & mLocationID

Call Tmp1Table
If TmpRs1.RecordCount > 0 Then
   For I = 0 To TmpRs1.RecordCount - 1
       If Edit_Flg = "E" Then
          If Val(TxtHologramID.Text) <> Val(TmpRs1!HologramID) Then
             If Val(TxtFromNo.Text) >= Val(TmpRs1!FromHologramNo) And Val(TxtFromNo.Text) <= Val(TmpRs1!ToHologramNo) Then
                MsgBox "From Hologram No. Already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromHologramNo & " To " & TmpRs1!ToHologramNo
                TxtFromNo.SetFocus
                Exit Sub
             End If
             If Val(TxtFromNo.Text) <= Val(TmpRs1!FromHologramNo) And Val(TxtToNo.Text) >= Val(TmpRs1!ToHologramNo) Then
                MsgBox "Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromHologramNo & " To " & TmpRs1!ToHologramNo
                TxtFromNo.SetFocus
                Exit Sub
             End If
             If Val(TxtToNo.Text) >= Val(TmpRs1!FromHologramNo) And Val(TxtToNo.Text) <= Val(TmpRs1!ToHologramNo) Then
                MsgBox "To Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromHologramNo & " To " & TmpRs1!ToHologramNo
                TxtToNo.SetFocus
                Exit Sub
             End If
          End If
       Else
          If Val(TxtFromNo.Text) >= Val(TmpRs1!FromHologramNo) And Val(TxtFromNo.Text) <= Val(TmpRs1!ToHologramNo) Then
             MsgBox "From Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromHologramNo & " To " & TmpRs1!ToHologramNo
             TxtFromNo.SetFocus
             Exit Sub
          End If
          If Val(TxtFromNo.Text) <= Val(TmpRs1!FromHologramNo) And Val(TxtToNo.Text) >= Val(TmpRs1!ToHologramNo) Then
                MsgBox "Hologram No. already  in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromHologramNo & " To " & TmpRs1!ToHologramNo
                TxtFromNo.SetFocus
                Exit Sub
          End If
          If Val(TxtToNo.Text) >= Val(TmpRs1!FromHologramNo) And Val(TxtToNo.Text) <= Val(TmpRs1!ToHologramNo) Then
             MsgBox "To Hologram No. already in " & GetLocationName(TmpRs1!LocationID) & ", Range " & TmpRs1!FromHologramNo & " To " & TmpRs1!ToHologramNo
             TxtToNo.SetFocus
             Exit Sub
          End If
 
       End If
       TmpRs1.MoveNext
   Next
End If

If TxtRemark.Text = "" Then
   MsgBox "Blank Remark not Allowed!!"
   TxtRemark.SetFocus
   Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_Hologram.AddNew
   RsTxn_Hologram!HologramID = getMaxHologramID
   RsTxn_Hologram!G_UID = GetGUID
   TxtHologramID.Text = RsTxn_Hologram!HologramID
Else
   RsTxn_Hologram.MoveFirst
   RsTxn_Hologram.Find " HologramID= '" & Val(TxtHologramID.Text) & "'"
End If

'RsTxn_Hologram!HologramID = Val(TxtHologramID.Text)
RsTxn_Hologram!Hologramdate = Format(TxtHologramDate.Value, Gen_DatFormat)
RsTxn_Hologram!LocationID = mLocationID
RsTxn_Hologram!Remark = TxtRemark.Text
RsTxn_Hologram!FromHologramNo = Val(TxtFromNo.Text)
RsTxn_Hologram!ToHologramNo = Val(TxtToNo.Text)
RsTxn_Hologram!TotalQty = Val(TxtTotalQty.Text)

If IsNull(RsTxn_Hologram!CreatedBy) = True Or RsTxn_Hologram!CreatedBy = "" Then
   RsTxn_Hologram!CreatedBy = Gen_UserLoginID
   RsTxn_Hologram!CreatedDate = Now
Else
   RsTxn_Hologram!UpdatedBy = Gen_UserLoginID
   RsTxn_Hologram!UpdatedDate = Now
End If

RsTxn_Hologram.Update

SaveCmd.Enabled = False
Call Indata

End Sub
Private Function getMaxHologramID() As Double
Dim Retval As Double
tmp = "Select max(HologramID) from Txn_Hologram"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
getMaxHologramID = Retval
End Function
Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_Hologram.MoveNext
If RsTxn_Hologram.EOF Then
   RsTxn_Hologram.MoveLast
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
RsTxn_Hologram.MovePrevious
If RsTxn_Hologram.BOF Then
   RsTxn_Hologram.MoveFirst
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


Private Sub LastCmd_Click()
RsTxn_Hologram.MoveLast
Call Indata
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub
Private Sub FirstCmd_Click()
RsTxn_Hologram.MoveFirst
Call Indata
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub SearchCmd_Click()
If SearchCmd.Caption = "Cancel" Then
   If RsTxn_Hologram.RecordCount = 0 Then Exit Sub
   Call GButtonLock(Me, True)
   RsTxn_Hologram.MoveFirst
   RsTxn_Hologram.Find "HologramID = " & Val(TxtHologramID) & ""
   If RsTxn_Hologram.EOF Then
      RsTxn_Hologram.MoveLast
   End If
   Call Indata
   Exit Sub
End If
Frame0.Height = Me.Height - 500
Frame0.Left = Frame1.Left - 100
Frame0.Top = Frame1.Top
Frame0.Width = Me.Width - 200
Frame0.Visible = True

MSHFlexGrid2.Width = Frame0.Width - 200
MSHFlexGrid2.Height = Frame0.Height - 2500

RsTxn_Hologram.MoveFirst
RsTxn_Hologram.Filter = ""

''===========comment on 7 july 2012 for not clearing search grid=============
'Call Grid_Head1
'SFrom.Value = Date - 365
'STo.Value = Date
'Label22.Caption = ""
'CmbSLocationID.Text = ""
'TxtHologramNo.Text = ""
''===========comment on 7 july 2012 for not clearing search grid=============

End Sub

Private Sub TxtFromNo_LostFocus()
If TxtFromNo = "" Then Exit Sub
If Not IsNumeric(TxtFromNo) Or Val(TxtFromNo) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtFromNo.SetFocus
End If
If TxtToNo = "" Then Exit Sub
TxtTotalQty.Text = Val(TxtToNo.Text) - Val(TxtFromNo.Text) + 1

End Sub

Private Sub TxtToNo_LostFocus()
If TxtToNo = "" Then Exit Sub
If Not IsNumeric(TxtToNo) Or Val(TxtToNo) < 0 Then
   MsgBox "Invalid numeric format !!!! "
   TxtToNo.SetFocus
End If
TxtTotalQty.Text = Val(TxtToNo.Text) - Val(TxtFromNo.Text) + 1

End Sub

Public Sub Indata()

Frame1.Enabled = False

TxtHologramDate.Enabled = True
CmbLocationID.Visible = False
TxtLocationID.Visible = True


TxtHologramID = RsTxn_Hologram!HologramID
TxtHologramDate.Value = RsTxn_Hologram!Hologramdate
mLocationID = RsTxn_Hologram!LocationID
TxtLocationID.Text = GetLocationName(RsTxn_Hologram!LocationID)
TxtRemark.Text = RsTxn_Hologram!Remark
TxtFromNo.Text = RsTxn_Hologram!FromHologramNo
TxtToNo.Text = RsTxn_Hologram!ToHologramNo
TxtTotalQty.Text = RsTxn_Hologram!TotalQty

TxtCreated = IIf(IsNull(RsTxn_Hologram!CreatedBy), "", RsTxn_Hologram!CreatedBy) & " :- " & IIf(IsNull(RsTxn_Hologram!CreatedDate), "", RsTxn_Hologram!CreatedDate)
TxtUpdated = IIf(IsNull(RsTxn_Hologram!UpdatedBy), "", RsTxn_Hologram!UpdatedBy) & " :- " & IIf(IsNull(RsTxn_Hologram!UpdatedDate), "", RsTxn_Hologram!UpdatedDate)


If RsTxn_Hologram.RecordCount = 1 Then
   Call GButtonLock_1(Me, True)
Else
   Call GButtonLock(Me, True)
End If
End Sub

Private Sub Grid_Head1()

MSHFlexGrid2.Clear
MSHFlexGrid2.Rows = 2
MSHFlexGrid2.FixedRows = 1
MSHFlexGrid2.Cols = 7
MSHFlexGrid2.WordWrap = True
MSHFlexGrid2.TextMatrix(0, 0) = "Hologram ID"

MSHFlexGrid2.TextMatrix(0, 1) = "Hologram Date"
MSHFlexGrid2.TextMatrix(0, 2) = "Location"
MSHFlexGrid2.TextMatrix(0, 3) = "From Hologram No."
MSHFlexGrid2.TextMatrix(0, 4) = "To Hologram No."
MSHFlexGrid2.TextMatrix(0, 5) = "Total Qty"
MSHFlexGrid2.TextMatrix(0, 6) = "Remark"

MSHFlexGrid2.ColWidth(0) = 1000
MSHFlexGrid2.ColWidth(1) = 1400
MSHFlexGrid2.ColWidth(2) = 2400
MSHFlexGrid2.ColWidth(3) = 2000
MSHFlexGrid2.ColWidth(4) = 2000
MSHFlexGrid2.ColWidth(5) = 1400
MSHFlexGrid2.ColWidth(6) = 2800

MSHFlexGrid2.RowHeight(0) = 800

End Sub


Private Sub CmbSLocationID_GotFocus()
tmp = CmbSLocationID.Text
If Gen_WcLocation <> "" Then
   Call TableMst_Location(" Where " & Gen_WcLocation)
Else
   Call TableMst_Location '(" Where LocationID <> 0 ")
End If
CmbSLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For I = 1 To RsMst_Location.RecordCount
    CmbSLocationID.AddItem RsMst_Location!Locationname
    RsMst_Location.MoveNext
Next
CmbSLocationID.Text = tmp
End Sub

Private Sub CmbSLocationID_LostFocus()
If CmbSLocationID.Text = "" Then
   mSLocationID = 0
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbSLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid selection "
   CmbSLocationID.SetFocus
   Exit Sub
End If


mSLocationID = RsMst_Location!LocationID

End Sub

