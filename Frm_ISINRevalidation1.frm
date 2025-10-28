VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ISINRevalidation 
   Caption         =   "ISIN Revalidation"
   ClientHeight    =   10635
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10635
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   52
      Top             =   8895
      Width           =   14925
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11625
         TabIndex        =   61
         Top             =   600
         Width           =   2580
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11625
         TabIndex        =   60
         Top             =   195
         Width           =   2580
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   8925
         TabIndex        =   59
         Top             =   600
         Width           =   2700
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   8925
         TabIndex        =   58
         Top             =   195
         Width           =   2700
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6225
         TabIndex        =   57
         Top             =   600
         Width           =   2700
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6225
         TabIndex        =   56
         Top             =   195
         Width           =   2700
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3405
         TabIndex        =   55
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3405
         TabIndex        =   7
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   585
         TabIndex        =   54
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   585
         TabIndex        =   53
         Top             =   195
         Width           =   2820
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "New ISIN"
      Height          =   2655
      Left            =   7680
      TabIndex        =   30
      Top             =   1560
      Width           =   7260
      Begin VB.TextBox TxtNewISINID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1800
         TabIndex        =   6
         Top             =   165
         Width           =   2640
      End
      Begin VB.TextBox TxtNewCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1725
         Width           =   5325
      End
      Begin VB.TextBox TxtNewGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   5130
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2000
      End
      Begin VB.TextBox TxtNewCMP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   540
         Width           =   5325
      End
      Begin VB.TextBox TxtNewLocationId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   915
         Width           =   5325
      End
      Begin VB.TextBox TxtNewExchangeTypeId 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2000
      End
      Begin MSComCtl2.DTPicker NewNE_Date 
         Height          =   360
         Left            =   1800
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57147393
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker NewFED 
         Height          =   360
         Left            =   5490
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1305
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57147393
         CurrentDate     =   39864
      End
      Begin VB.Label LblNewExchangeType 
         Caption         =   "Exchange Type"
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
         Left            =   90
         TabIndex        =   45
         Top             =   2160
         Width           =   1920
      End
      Begin VB.Label LblNewLocation 
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
         Left            =   90
         TabIndex        =   44
         Top             =   990
         Width           =   1560
      End
      Begin VB.Label LblNewFED 
         Caption         =   "FED"
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
         Left            =   4560
         TabIndex        =   43
         Top             =   1335
         Width           =   600
      End
      Begin VB.Label LblNewNED 
         Caption         =   "NED"
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
         Left            =   90
         TabIndex        =   42
         Top             =   1380
         Width           =   1050
      End
      Begin VB.Label LblNewCMP 
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
         Left            =   90
         TabIndex        =   41
         Top             =   615
         Width           =   1080
      End
      Begin VB.Label LblNewGrade 
         Caption         =   "Grade"
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
         Left            =   4200
         TabIndex        =   40
         Top             =   2115
         Width           =   780
      End
      Begin VB.Label LblNewISINID 
         Caption         =   "ISIN ID"
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
         TabIndex        =   39
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label LblNewCommodity 
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
         Left            =   90
         TabIndex        =   38
         Top             =   1800
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Old ISIN"
      Height          =   2655
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Width           =   7335
      Begin VB.TextBox TxtOldExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2000
      End
      Begin VB.TextBox txtOldCMP 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   585
         Width           =   5325
      End
      Begin VB.TextBox TxtOldLocationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   5325
      End
      Begin VB.TextBox TxtOldGrade 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   5205
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2145
         Width           =   2000
      End
      Begin VB.TextBox TxtOldISINID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1890
         TabIndex        =   5
         Top             =   195
         Width           =   2640
      End
      Begin VB.TextBox TxtOldCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1755
         Width           =   5325
      End
      Begin MSComCtl2.DTPicker OldNE_Date 
         Height          =   360
         Left            =   1890
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57147393
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker OldFED 
         Height          =   360
         Left            =   5925
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   57147393
         CurrentDate     =   39864
      End
      Begin VB.Label LblOldExchangeTypeID 
         Caption         =   "Exchange Type"
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
         Left            =   135
         TabIndex        =   29
         Top             =   2130
         Width           =   1920
      End
      Begin VB.Label LbOldlLocationID 
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
         Left            =   135
         TabIndex        =   28
         Top             =   945
         Width           =   1200
      End
      Begin VB.Label LblOldFED 
         Caption         =   "FED"
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
         Left            =   4650
         TabIndex        =   27
         Top             =   1335
         Width           =   840
      End
      Begin VB.Label LblOldNE_Date 
         Caption         =   "NED"
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
         Left            =   135
         TabIndex        =   26
         Top             =   1335
         Width           =   810
      End
      Begin VB.Label LblOldCMPID 
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
         Left            =   120
         TabIndex        =   25
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label lblOldGrade 
         Caption         =   "Grade"
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
         Left            =   4320
         TabIndex        =   24
         Top             =   2145
         Width           =   780
      End
      Begin VB.Label LblOldISINID 
         Caption         =   "ISIN ID"
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
         Left            =   135
         TabIndex        =   23
         Top             =   195
         Width           =   1500
      End
      Begin VB.Label LblOldCommodity 
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
         Left            =   135
         TabIndex        =   22
         Top             =   1755
         Width           =   1500
      End
   End
   Begin MSComCtl2.DTPicker RevalidationDate 
      Height          =   360
      Left            =   9480
      TabIndex        =   1
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   635
      _Version        =   393216
      Format          =   57147393
      CurrentDate     =   39889
   End
   Begin VB.Frame Frame1 
      Height          =   8775
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   14895
      Begin VB.TextBox TxtRevalidationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
      End
      Begin VB.ComboBox CmbClientDP 
         Height          =   315
         Left            =   1995
         Sorted          =   -1  'True
         TabIndex        =   4
         Top             =   1035
         Visible         =   0   'False
         Width           =   12705
      End
      Begin VB.TextBox TxtDPID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1035
         Width           =   2640
      End
      Begin VB.TextBox TxtDPName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1020
         Width           =   5325
      End
      Begin VB.TextBox TxtClientName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   9360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   645
         Width           =   5325
      End
      Begin VB.TextBox TxtClientID 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1995
         TabIndex        =   3
         Top             =   645
         Width           =   2640
      End
      Begin VB.TextBox TxtRevalidationReqNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   5385
         TabIndex        =   0
         Top             =   240
         Width           =   1950
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4455
         Left            =   120
         TabIndex        =   20
         Top             =   4200
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   7858
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "Revalidation Id"
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
         Height          =   315
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label LblDPName 
         Caption         =   "DP ID"
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
         Left            =   120
         TabIndex        =   50
         Top             =   1035
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "DP Name"
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
         Left            =   7635
         TabIndex        =   49
         Top             =   1020
         Width           =   1260
      End
      Begin VB.Label lblNameDid 
         Caption         =   "ClientDPID"
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
         Left            =   840
         TabIndex        =   48
         Top             =   1035
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label LblClientID 
         Caption         =   "Client ID"
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
         TabIndex        =   47
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label LblClientName 
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
         Left            =   7635
         TabIndex        =   46
         Top             =   660
         Width           =   1260
      End
      Begin VB.Label LblRevalidationDate 
         Caption         =   "Revalidation Date"
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
         Left            =   7635
         TabIndex        =   18
         Top             =   255
         Width           =   1740
      End
      Begin VB.Label LblRevalidationID 
         Caption         =   "Revalidation ReqNo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3240
         TabIndex        =   17
         Top             =   240
         Width           =   2145
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   10260
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Frm_ISINRevalidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mExchangeTypeID As Double
Dim mCommodityID  As Double
Dim mCMPID As Double
Dim mClientIDRow As Double
Dim Edit_Flg As Variant

Private Sub FirstCmd_Click()
'RsMst_Assayer.MoveFirst
Call InData
'TmpRs.MoveFirst
PreviousCmd.Enabled = False
FirstCmd.Enabled = False
NextCmd.Enabled = True
LastCmd.Enabled = True
End Sub

Private Sub Form_Load()
    Call Grid_Head
    Call TableTxn_ISINRevalidation
    If RsTxn_ISINRevalidation.RecordCount = 0 Then
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
End Sub

Private Sub AddCmd_Click()
    Edit_Flg = "A"
    TxtRevalidationReqNo.Locked = False
    TxtClientID.Locked = False
    TxtOldISINID.Locked = False
    TxtNewISINID.Locked = False
    RevalidationDate.Value = Date
    
    TxtRevalidationID.Text = ""
    TxtRevalidationReqNo.Text = ""
    TxtClientID.Text = ""
    TxtClientName.Text = ""
    TxtDPID.Text = ""
    TxtDPName.Text = ""
    
    TxtOldISINID.Text = ""
    txtOldCMP.Text = ""
    TxtOldLocationID.Text = ""
    TxtOldCommodity.Text = ""
    TxtOldExchangeTypeID.Text = ""
    TxtOldGrade.Text = ""
    
    TxtNewISINID.Text = ""
    TxtNewCMP.Text = ""
    TxtNewLocationId.Text = ""
    TxtNewCommodity.Text = ""
    TxtNewExchangeTypeId.Text = ""
    TxtNewGrade.Text = ""
    
    Call GButtonLock(Me, False)
    
End Sub

Private Sub SaveCmd_Click()
    If TxtRevalidationReqNo.Text = "" Then
        MsgBox "Blank Revalidation Req No Not Allowed !!! "
        TxtRevalidationReqNo.SetFocus
        Exit Sub
    End If

    If TxtOldISINID.Text = "" Then
        MsgBox "Blank Old ISIN ID Not Allowed !!! "
        TxtOldISINID.SetFocus
        Exit Sub
    End If

    If TxtNewISINID.Text = "" Then
        MsgBox "Blank New ISIN ID Not Allowed !!! "
        TxtNewISINID.SetFocus
        Exit Sub
    End If
     
    If TxtOldISINID.Text = TxtNewISINID.Text Then
        MsgBox "Same Old And New ISIN Not Allowed!!!"
        TxtNewISINID.SetFocus
        Exit Sub
    End If
    
    If TxtOldCommodity.Text <> TxtNewCommodity.Text Then
        MsgBox "Old Commodity and New Commodity Name must be same !!! "
        TxtNewISINID.SetFocus
        Exit Sub
    End If

    If txtOldCMP.Text <> TxtNewCMP.Text Then
        MsgBox "Old CMP and New CMP Name must be same !!! "
        TxtNewISINID.SetFocus
        Exit Sub
    End If

    If TxtOldLocationID.Text <> TxtNewLocationId.Text Then
        MsgBox "Old Location and New Location Name must be same !!! "
        TxtNewISINID.SetFocus
        Exit Sub
    End If
    
    If MSHFlexGrid2.Rows < 4 Then
        MsgBox "Grid Should Contain Atleast One Record!!!"
        TxtOldISINID.SetFocus
        Exit Sub
    End If
    
    If Edit_Flg = "A" Then
        If RsTxn_ISINRevalidation.RecordCount > 0 Then
            RsTxn_ISINRevalidation.MoveFirst
            RsTxn_ISINRevalidation.Find "ISINRevalReqNo= '" & TxtRevalidationReqNo.Text & "'"
            If Not RsTxn_ISINRevalidation.EOF Then
                MsgBox "Duplicate Revalidation Req No Not Allowed !!! "
                TxtRevalidationReqNo.SetFocus
            Exit Sub
        End If
    End If
    RsTxn_ISINRevalidation.AddNew
    RsTxn_ISINRevalidation!ISINRevalID = GetISINRevalID
    TxtRevalidationID.Text = RsTxn_ISINRevalidation!ISINRevalID
    Else
        RsTxn_ISINRevalidation.MoveFirst
        RsTxn_ISINRevalidation.Find "ISINRevalReqNo= '" & TxtRevalidationReqNo.Text & "'"
        If Not RsTxn_ISINRevalidation.EOF Then
            If RsTxn_ISINRevalidation!ISINRevalReqNo <> TxtISINRevalReqNo.Text Then
                MsgBox "Duplicate ISINReval Req No Not Allowed !!! "
                TxtISINRevalReqNo.SetFocus
            Exit Sub
            End If
        End If
    RsTxn_ISINRevalidation.MoveFirst
    RsTxn_ISINRevalidation.Find " ISINRevalID= " & TxtISINRevalReqNo.Text
    End If
    RsTxn_ISINRevalidation!ISINRevalID = TxtRevalidationID.Text
    RsTxn_ISINRevalidation!ISINRevalReqNo = TxtRevalidationReqNo.Text
    RsTxn_ISINRevalidation!ISINRevalDate = RevalidationDate.Value
    RsTxn_ISINRevalidation!ClientRowID = mClientIDRow
    RsTxn_ISINRevalidation!OldISIN = TxtOldISINID.Text
    RsTxn_ISINRevalidation!NewISIN = TxtNewISINID.Text
    RsTxn_ISINRevalidation.Update
    'Call InData
    SaveGridData
End Sub

Private Sub ExitCmd_Click()
    Unload Me
End Sub
Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub
    If MSHFlexGrid2.Rows > 3 Then
        If MSHFlexGrid2.RowSel <> "0" Then
            If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub
                If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "" Then
                    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
                Else
                    MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
                End If
            End If
        End If
End Sub
Private Sub SaveGridData()
    If MSHFlexGrid2.Rows > 3 Then
        Call TableTxn_ISINGSLRevalidation
        For i = 2 To MSHFlexGrid2.Rows - 1
        If MSHFlexGrid2.TextMatrix(i, 1) = "" Then Exit For
            If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
                If Edit_Flg = "A" Then
                    RsTxn_ISINGSLRevalidation.AddNew
                    'RsTxn_ISINGSLRevalidation.Find ("GSLID = " & MSHFlexGrid2.TextMatrix(i, 6))
                    RsTxn_ISINGSLRevalidation!ISINGSLRevalId = GetISINGSLRevalId
                Else
                    RsTxn_ISINGSLRevalidation.Find ("ISINGSLRevalId = " & MSHFlexGrid2.TextMatrix(i, 7))
                End If
                RsTxn_ISINGSLRevalidation!ISINRevalID = Trim(TxtRevalidationID.Text)
                RsTxn_ISINGSLRevalidation!GslID = MSHFlexGrid2.TextMatrix(i, 6)
                RsTxn_ISINGSLRevalidation.Update
          
            End If
        Next
      
  End If
  RsTxn_ISINGSLRevalidation.Requery
  Call InData
End Sub
Private Sub InData()
  Call TableTxn_ISINRevalidation
  TxtRevalidationID = RsTxn_ISINRevalidation!ISINRevalID
  TxtRevalidationReqNo = RsTxn_ISINRevalidation!ISINRevalReqNo
  RevalidationDate = RsTxn_ISINRevalidation!ISINRevalDate
  mClientIDRow = RsTxn_ISINRevalidation!ClientRowID
  ans = GetClientNameByRow(mClientIDRow)
  TxtClientID = ans(0)
  TxtOldISINID = RsTxn_ISINRevalidation!OldISIN
  TxtNewISINID = RsTxn_ISINRevalidation!NewISIN
  
  Call TxtClientID_LostFocus
  Call TxtOldISINID_LostFocus
  Call TxtNewISINID_LostFocus
  
  Call GridInData
End Sub

Private Sub GridInData()
    Call Grid_Head
    Call TableTxn_ISINGSLRevalidation("Where ISINRevalId=" & TxtRevalidationID.Text)
    If RsTxn_ISINGSLRevalidation.RecordCount > 0 Then
    
    RsTxn_ISINGSLRevalidation.MoveFirst
    k = 1
    For i = 0 To RsTxn_ISINGSLRevalidation.RecordCount - 1
    k = k + 1
    MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
    'MSHFlexGrid2.TextMatrix(i, 6) = RsTxn_ISINGSLRevalidation!GSLId
        tmp = "SELECT Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, Mst_GSL.DepositeWeight,Mst_GSL.ISINID, "
        tmp = tmp & " Mst_GSL.GSLID FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
        tmp = tmp & " WHERE (Mst_GSL.GSLId=" & RsTxn_ISINGSLRevalidation!GslID & ")"
        Call TmpTable
        MSHFlexGrid2.TextMatrix(k, 0) = "*"
        MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!GodownNo
        MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!StackNo
        MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!LotNo
        MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!DepositeBags
        MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!DepositeWeight
        MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!GslID
        MSHFlexGrid2.TextMatrix(k, 7) = RsTxn_ISINGSLRevalidation!ISINGSLRevalId
        RsTxn_ISINGSLRevalidation.MoveNext
    Next
    End If
End Sub

Private Sub SearchCmd_Click()
  If SearchCmd.Caption = "Cancel" Then
    If RsTxn_ISINRevalidation.RecordCount = 0 Then Exit Sub
    Call GButtonLock(Me, True)
    Call FirstCmd_Click
    Exit Sub
  End If
End Sub

Private Sub TxtClientID_LostFocus()
    CmbClientDP.Visible = False
    If TxtClientID = "" Then Exit Sub
    'If CmbExchangeTypeID.Text = "" Then Exit Sub
    Dim ans As Variant
    mExchangeTypeID = GetExchangeID("NCDEX")
    ans = GetClientNameDtl(TxtClientID, mExchangeTypeID)
    If ans(0, 0) = False Then
       MsgBox "Not Found !!!"
       TxtClientID.SetFocus
       Exit Sub
    End If
    If ans(3, 0) > 1 Then
       CmbClientDP.Visible = True
    End If
    CmbClientDP.Text = ""
    CmbClientDP.Clear
    tmp = "SELECT Mst_Depository.DepositoryID, Mst_Depository.DepositoryName, Mst_Client.ClientID, " & _
          "Mst_Client.ClientName,  Mst_Client.ClientIDRow " & _
          "FROM Mst_Depository INNER JOIN Mst_Client ON Mst_Depository.DepositoryID = Mst_Client.DepositoryID " & _
          "WHERE (((Mst_Client.ExchangeTypeID)=" & mExchangeTypeID & ") AND ((Mst_Client.ClientID)='" & TxtClientID & "'));"
    Call Tmp1Table
    
    For op = 1 To TmpRs1.RecordCount
        If CmbClientDP.Text = "" Then
           CmbClientDP.Text = TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
        End If
        CmbClientDP.AddItem TmpRs1!DepositoryID & " ~ " & TmpRs1!DepositoryName
        TmpRs1.MoveNext
    Next
    If ans(3, 0) <= 1 Then
       CmbClientDP_LostFocus
    End If
    
    Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
    
End Sub
Private Sub CmbClientDP_LostFocus()
Dim Str_() As String
Str_ = Split(CmbClientDP.Text, "~")
If UBound(Str_) <> 1 Then
   MsgBox "Invalid Selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If
TmpRs1.MoveFirst

TmpRs1.MoveFirst
TmpRs1.Find "DepositoryID  = '" & Str_(0) & "'"
If TmpRs1.EOF Then
   MsgBox "Invalid Selection !!!!! "
   CmbClientDP.SetFocus
   Exit Sub
End If

TxtClientName = TmpRs1!ClientName
TxtDPID = TmpRs1!DepositoryID
TxtDPName = TmpRs1!DepositoryName
CmbClientDP.Visible = False
mClientIDRow = TmpRs1!ClientIDRow
'CmdClientAdd2List.SetFocus
End Sub


Private Sub TxtNewISINID_LostFocus()
    If TxtNewISINID.Text = "" Then
'     mCommodityID = 0
'     mCMPID = ""
'     mExchangeTypeID = 0
     TxtNewCMP = ""
     TxtNewLocationId = ""
     TxtNewCommodity = ""
     TxtNewExchangeTypeId = ""
     TxtNewGrade = ""
     'Call Grid_Head
     Exit Sub
    End If
    Call TableMst_ISIN("where ISINID = '" & TxtNewISINID.Text & "'")
    
    If RsMst_ISIN.RecordCount = 0 Then
       MsgBox "No Such ISIN Exist !!!"
       TxtNewISINID.SetFocus
       Exit Sub
    End If
    'mCMPID = RsMst_ISIN!CMPID
    TxtNewCMP = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
    NewNE_Date = IIf(IsNull(RsMst_ISIN!NE_Date), "", RsMst_ISIN!NE_Date)
    NewFED = IIf(IsNull(RsMst_ISIN!FED), "", RsMst_ISIN!FED)
    'mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
    TxtNewLocationId = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
    'mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
    TxtNewCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
    'mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeID), "", RsMst_ISIN!ExchangeTypeID)
    TxtNewExchangeTypeId = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
    TxtNewGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
End Sub

Private Sub TxtOldISINID_LostFocus()
    If TxtOldISINID.Text = "" Then
     mCommodityID = 0
     mCMPID = 0
     mExchangeTypeID = 0
     txtOldCMP = ""
     TxtOldLocationID = ""
     TxtOldCommodity = ""
     TxtOldExchangeTypeID = ""
     TxtOldGrade = ""
     'Call Grid_Head
     Exit Sub
    End If
    Call TableMst_ISIN("where ISINID = '" & TxtOldISINID.Text & "'")
    
    If RsMst_ISIN.RecordCount = 0 Then
       MsgBox "No Such ISIN Exist !!!"
       TxtOldISINID.SetFocus
       Exit Sub
    End If
    mCMPID = RsMst_ISIN!CMPID
    txtOldCMP = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
    OldNE_Date = IIf(IsNull(RsMst_ISIN!NE_Date), "", RsMst_ISIN!NE_Date)
    OldFED = IIf(IsNull(RsMst_ISIN!FED), "", RsMst_ISIN!FED)
    mLocationID = IIf(IsNull(RsMst_ISIN!LocationID), "", RsMst_ISIN!LocationID)
    TxtOldLocationID = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
    mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
    TxtOldCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
    mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeID), "", RsMst_ISIN!ExchangeTypeID)
    TxtOldExchangeTypeID = GetExchangeName(RsMst_ISIN!ExchangeTypeID) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
    TxtOldGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
    
    Call GetGridData(mClientIDRow, mExchangeTypeID, mCMPID, mCommodityID)
    

End Sub

Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 8
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "   "
MSHFlexGrid2.TextMatrix(1, 1) = "Godown  "
MSHFlexGrid2.TextMatrix(1, 2) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 3) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 4) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 5) = "Gross Weight"
MSHFlexGrid2.TextMatrix(1, 6) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 7) = "ISINGSLRevalId"

MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 2600
MSHFlexGrid2.ColWidth(3) = 2200
MSHFlexGrid2.ColWidth(4) = 2200
MSHFlexGrid2.ColWidth(5) = 2600
MSHFlexGrid2.ColWidth(6) = 0
MSHFlexGrid2.ColWidth(7) = 0
'MSHFlexGrid2.RowHeight(1) = 400
End Sub


Private Sub GetGridData(mClientRowId_, mExchangeID_, mCMPID_, mCommodityID_)
Call Grid_Head
If SaveCmd.Enabled = False Then Exit Sub
If IsNull(mClientRowId_) Then
   Exit Sub
End If
If IsNull(mExchangeID_) Then
   Exit Sub
End If
If IsNull(mCMPID_) Or mCMPID_ = 0 Then
   Exit Sub
End If
If IsNull(mCommodityID_) Or mCommodityID_ = 0 Then
   Exit Sub
End If

tmp = "SELECT Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, Mst_GSL.DepositeWeight,Mst_GSL.ISINID, "
tmp = tmp & " Mst_GSL.GSLID FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
tmp = tmp & " WHERE (Mst_GSL.ClientIDRow= " & mClientRowId_ & " AND Mst_GSL.CommodityID=" & mCommodityID_ & "  AND "
tmp = tmp & " Mst_GSL.ExchangeTypeID=" & mExchangeID_ & " AND Mst_Godown.CMPID= " & mCMPID_ & ") AND  "

If SaveCmd.Enabled = False Then
   tmp = tmp & " (Mst_GSL.ISINID= '" & TxtOldISINID & "')"
Else
   'tmp = tmp & " (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID= '' OR Mst_GSL.ISINID= '" & TxtOldISINID & "')"
   tmp = tmp & " (Mst_GSL.ISINID= '" & TxtOldISINID & "')"
End If
  
Call TmpTable

If TmpRs.RecordCount <= 0 Then
 MsgBox "No Record Found For Grid!!!"
 'TxtClientID.SetFocus
 Exit Sub
End If
k = 1
For i = 0 To TmpRs.RecordCount - 1
 k = k + 1
  MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
  MSHFlexGrid2.TextMatrix(k, 0) = ""
  MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!GodownNo
  MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!StackNo
  MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!LotNo
  MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!DepositeBags
  MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!DepositeWeight
  MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!GslID
  TmpRs.MoveNext
Next

End Sub

Private Function GetISINRevalID() As Double
Dim RetVal As Double
tmp = "Select max(ISINRevalID) from Txn_ISINRevalidation"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetISINRevalID = RetVal
End Function

Private Function GetISINGSLRevalId() As Double
Dim RetVal As Double
tmp = "Select max(ISINGSLRevalId) from Txn_ISINGSLRevalidation"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
GetISINGSLRevalId = RetVal
End Function

