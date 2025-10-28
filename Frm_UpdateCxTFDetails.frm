VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_UpdateCxtfDetails 
   Caption         =   "Update CxTF Details"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1575
      Left            =   120
      TabIndex        =   49
      Top             =   8160
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   2778
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   7800
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   6495
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   12015
      Begin VB.CheckBox Check3 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1680
         TabIndex        =   9
         Top             =   5160
         Width           =   225
      End
      Begin VB.CheckBox Check2 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1680
         TabIndex        =   7
         Top             =   4688
         Width           =   225
      End
      Begin VB.TextBox TxtNewManualNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   2055
         MaxLength       =   15
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4695
         Width           =   2370
      End
      Begin VB.TextBox TxtPrevManualNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1395
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
         Height          =   720
         Left            =   1680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   5640
         Width           =   10155
      End
      Begin VB.CheckBox Check1 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1680
         TabIndex        =   5
         Top             =   4237
         Width           =   225
      End
      Begin VB.TextBox TxtAgent 
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
         Left            =   7440
         Locked          =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3480
         Width           =   4410
      End
      Begin VB.ComboBox CmbAgent 
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
         Left            =   2055
         Sorted          =   -1  'True
         TabIndex        =   10
         Top             =   5130
         Width           =   4440
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   10125
      End
      Begin VB.TextBox TxtPreFix 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   718
         Width           =   4395
      End
      Begin VB.ComboBox CmbCMPID 
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
         Left            =   1680
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   1196
         Width           =   4440
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2640
         Width           =   4410
      End
      Begin VB.TextBox TxtCMPAddress 
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
         Height          =   1305
         Left            =   6240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Width           =   5565
      End
      Begin VB.TextBox TxtEHoPCMPID 
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1674
         Width           =   4425
      End
      Begin VB.ComboBox CmbModuleName 
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
         ItemData        =   "Frm_UpdateCxTFDetails.frx":0000
         Left            =   1680
         List            =   "Frm_UpdateCxTFDetails.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   2630
         Width           =   4440
      End
      Begin VB.TextBox TxtModuleNo 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         Top             =   3108
         Width           =   1395
      End
      Begin VB.TextBox TxtExchangeType 
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
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3060
         Width           =   4410
      End
      Begin VB.TextBox TxtStartDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1395
      End
      Begin VB.TextBox TxtCurrentDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1395
      End
      Begin VB.TextBox TxtModuleDate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   4680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3108
         Width           =   1395
      End
      Begin VB.TextBox TxtHologramNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   360
         Left            =   10440
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   3960
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComCtl2.DTPicker dtpNewCxtfDate 
         Height          =   375
         Left            =   2055
         TabIndex        =   6
         Top             =   4200
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
         Format          =   3932161
         CurrentDate     =   39884
      End
      Begin VB.Label Label4 
         Caption         =   "New Manual CxTF No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         TabIndex        =   44
         Top             =   4635
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Manual CxTF No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         TabIndex        =   43
         Top             =   3600
         Width           =   1260
      End
      Begin VB.Label Label14 
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
         Left            =   195
         TabIndex        =   41
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Agent"
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
         Left            =   6240
         TabIndex        =   40
         Top             =   3600
         Width           =   525
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "New CxTF Date"
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
         Left            =   150
         TabIndex        =   39
         Top             =   4230
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "New Agent"
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
         Left            =   150
         TabIndex        =   38
         Top             =   5190
         Width           =   975
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
         Height          =   300
         Left            =   195
         TabIndex        =   36
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label31 
         Caption         =   "SM_Prefix"
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
         Left            =   195
         TabIndex        =   35
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "CMP Address"
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
         Left            =   7710
         TabIndex        =   34
         Top             =   885
         Width           =   1575
      End
      Begin VB.Label Label12 
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
         Left            =   6240
         TabIndex        =   33
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "CMP"
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
         Left            =   195
         TabIndex        =   32
         Top             =   1253
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "EHoP CMP ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   195
         TabIndex        =   31
         Top             =   1770
         Width           =   1320
      End
      Begin VB.Label LblStartDate 
         Caption         =   "CMP Start Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   195
         TabIndex        =   30
         Top             =   2168
         Width           =   1440
      End
      Begin VB.Label Label8 
         Caption         =   "CMP Current Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   3240
         TabIndex        =   29
         Top             =   2160
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Module Name"
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
         Left            =   195
         TabIndex        =   28
         Top             =   2683
         Width           =   1335
      End
      Begin VB.Label LblModuleNo 
         Caption         =   "No"
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
         Left            =   195
         TabIndex        =   27
         Top             =   3161
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   3195
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Exchange"
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
         Left            =   6255
         TabIndex        =   25
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Hologram"
         Height          =   180
         Left            =   9600
         TabIndex        =   24
         Top             =   4050
         Visible         =   0   'False
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   12000
      Begin VB.CommandButton CmdPendingGSL 
         Caption         =   "Pending GSLs"
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
         Left            =   10275
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton CmdUpdateBillingRate 
         Caption         =   "Update Billing Rate on GSL"
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
         Left            =   8340
         TabIndex        =   48
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdUpdateProcess 
         Caption         =   "Update Data"
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
         Left            =   6765
         TabIndex        =   46
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Cmd_GSLFlg 
         Caption         =   "Set 'E' Flag in GSL Master where Bags is zero and Pledge Release Bags update"
         Height          =   615
         Left            =   3270
         TabIndex        =   45
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton CmdUpdate 
         Caption         =   "Update"
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
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   1455
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
         Height          =   615
         Left            =   1695
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Frm_UpdateCxtfDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mLocationID, mCMPID, mWHRNo, mCommodityID, mExchangeTypeID As Variant
Dim mCurrentDate, mModuleDate, mAgentID, mType, mPrevAgentID As Variant
Private Sub Check1_Click()
If Check1.Value = 1 Then
   dtpNewCxtfDate.Enabled = True
   dtpNewCxtfDate.Value = Date
Else
   dtpNewCxtfDate.Enabled = False
   dtpNewCxtfDate.Value = Date
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   TxtNewManualNo.Locked = False
   TxtNewManualNo = ""
Else
   TxtNewManualNo.Locked = True
   TxtNewManualNo = ""
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   CmbAgent.Enabled = True
   CmbAgent.Text = ""
Else
   CmbAgent.Enabled = False
   CmbAgent.Text = ""
End If
End Sub

Private Sub CmbAgent_GotFocus()
tmp = CmbAgent.Text
Call TableMst_Agent("")
CmbAgent.Clear
If RsMst_Agent.RecordCount = 0 Then
   MsgBox "No record found!!! "
   Exit Sub
End If
For I = 1 To RsMst_Agent.RecordCount
    CmbAgent.AddItem RsMst_Agent!AgentName
    RsMst_Agent.MoveNext
Next
CmbAgent.Text = tmp
End Sub

Private Sub CmbAgent_LostFocus()
If CmbAgent.Text = "" Then
   Exit Sub
End If
RsMst_Agent.MoveFirst
RsMst_Agent.Find "AgentName = '" & CmbAgent.Text & "'"
If RsMst_Agent.EOF Then
   MsgBox "Invalid selection!!! "
   CmbAgent.SetFocus
   Exit Sub
End If
mAgentID = RsMst_Agent!AgentID
End Sub

Private Sub CmbModuleName_LostFocus()
If CmbModuleName.Text = "" Then
   LblModuleNo.Caption = "No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   TxtModuleDate = ""
   TxtHologramNo = ""
   Exit Sub
End If

If CmbModuleName.Text = "CDTF" Then
   mType = "D"
   LblModuleNo.Caption = "CDTF No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   CmbAgent.Text = ""
   TxtAgent = ""
   Check1.Value = 0
   Check2.Value = 0
   Check3.Value = 0
   mAgentID = Null
   mPrevAgentID = Null
ElseIf CmbModuleName.Text = "CITF" Then
   mType = "W"
   LblModuleNo.Caption = "CITF No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   CmbAgent.Text = ""
   TxtAgent = ""
   Check1.Value = 0
   Check2.Value = 0
   Check3.Value = 0
   mAgentID = Null
   mPrevAgentID = Null
Else
   MsgBox "Invalid Selection!!!"
   LblModuleNo.Caption = "No"
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   CmbAgent.Text = ""
   TxtAgent = ""
   Check1.Value = 0
   Check2.Value = 0
   Check3.Value = 0
   mAgentID = Null
   mPrevAgentID = Null
   CmbModuleName.SetFocus
   Exit Sub
End If

End Sub

Private Sub Cmd_GSLFlg_Click()


tmp = "update mst_gsl set flag = 'E' where BalanceBags <= 0 and flag = 'J'"
Call TmpTable

tmp = "update mst_gsl set flag = 'E' where BalanceBags <= 0 and flag = 'O'"
Call TmpTable

tmp = "update mst_gsl set flag = 'O' where BalanceBags > 0 and flag = 'E'"
Call TmpTable

'Call TableTxn_WHRDetail


'tmp = "Select a.sm_prefix,a.WHRNo,a.pledgeid , a.WHRNoOfBags , a.DeliveredNoOfBags from Txn_WHR_Detail a "
'tmp = tmp & " inner join Txn_PledgeRelease b on b.pledgeid = a.pledgeid and b.sm_prefix = a.sm_prefix"
'tmp = tmp & " where  (a.WHRNoOfBags -a.DeliveredNoOfBags ) = 0 "
'
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   For I = 1 To TmpRs.RecordCount
'       RsTxn_WHRDetail.MoveFirst
'       RsTxn_WHRDetail.Filter = "SM_Prefix = '" & TmpRs!SM_Prefix & "' And WHRNo = " & TmpRs!WHRNo & ""
'       If RsTxn_WHRDetail.RecordCount > 0 Then
'          RsTxn_WHRDetail!PledgeID = 0
'          RsTxn_WHRDetail.Update
'       End If
'       RsTxn_WHRDetail.Filter = ""
'       TmpRs.MoveNext
'   Next
'End If


'tmp = " Select b.SM_Prefix,b.PledgeID,b.GSLID,Sum(b.ReleasedNoOfBags) 'ReleasedNoOfBags',Sum (b.ReleasedQuantity)'ReleasedQuantity',g.Flag,g.FGSLID,g.FSM_Prefix,g.PledgeBags,g.PledgeWeight  ,g.BalanceBags,g.BalanceWeight  ,g.DepositeBags,g.DepositeWeight"
'tmp = tmp & " From Txn_PledgeRelease a"
'tmp = tmp & " Inner Join Txn_PledgeRelease_GSL b on a.PledgeReleaseID = b.PledgeReleaseID and a.SM_Prefix = b.SM_Prefix"
'tmp = tmp & " Inner Join Mst_GSL g on b.GSLID = g.GSLID And b.SM_Prefix = g.SM_Prefix"
'tmp = tmp & " Where a.CreatedDate > '" & Date - 200 & "' And g.FGSLID is not null " '--and g.GSLID = 19 and g.SM_Prefix = 'KNGMP'"
'tmp = tmp & " Group By b.SM_Prefix,b.PledgeID,b.GSLID,g.Flag,g.FGSLID,g.FSM_Prefix,g.PledgeBags,g.PledgeWeight  ,g.BalanceBags,g.BalanceWeight  ,g.DepositeBags,g.DepositeWeight"
'tmp = tmp & " Order by b.SM_Prefix"
'
'Call TmpTable
'
'For I = 1 To TmpRs.RecordCount
'
'    mFGSLID = TmpRs!FGSLID
'    mFSM_Prefix = TmpRs!FSM_Prefix
'
'query:    tmp = "Select * from Mst_GSL Where GSLID = " & mFGSLID & " And SM_Prefix = '" & mFSM_Prefix & "' "
'    Call Tmp1Table
'
'    If TmpRs1.RecordCount > 0 Then
'       If IsNull(TmpRs1!FGSLID) = True Then
'          If LCase(TmpRs1!Flag) <> "e" Then
'             mPledgeRelBags = IIf(IsNull(TmpRs!ReleasedNoOfBags), 0, TmpRs!ReleasedNoOfBags)
'             mPledgeRelQty = IIf(IsNull(TmpRs!ReleasedQuantity), 0, TmpRs!ReleasedQuantity)
'
'             tmp = "Select * from Txn_PledgeCreation_GSL Where PledgeID = " & TmpRs!PledgeID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "'  And GSLID = " & TmpRs!GSLID & "            "
'             Call Tmp2Table
'
'             If TmpRs2.RecordCount > 0 Then
'                'mOrigPledgeBags = IIf(IsNull(TmpRs2!PledgeNoOfBags), 0, TmpRs2!PledgeNoOfBags)
'                'If Val(mOrigPledgeBags) < TmpRs1!PledgeBags - Val(mPledgeRelBags) Then
'                   'If TmpRs1!PledgeBags - Val(mPledgeRelBags) >= 0 Then
'                      TmpRs1!PledgeBags = Val(mPledgeRelBags)
'                      TmpRs1!PledgeWeight = Val(mPledgeRelQty)
'                      TmpRs1.Update
'                   'End If
'                'End If
'             End If
'
'          End If
'       Else
'          mFGSLID = TmpRs1!FGSLID
'          mFSM_Prefix = TmpRs1!FSM_Prefix
'          GoTo query
'       End If
'    End If
'    TmpRs.MoveNext
'Next

MsgBox "Done!!!"
'Cmd_GSLFlg.Enabled = False
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub ClearData()
CmbLocationID.Text = ""
CmbCMPID.Text = ""
CmbModuleName = ""
TxtRemark.Text = ""
Call CmbLocationID_LostFocus
Call CmbCMPID_LostFocus
Call CmbModuleName_LostFocus
End Sub
Private Sub CmdPendingGSL_Click()

If LCase(Gen_UserRole) = "power" Then
   Call CmdUpdateBillingRate_Click
End If
'tmp = " Select StateName,LocationName,SM_Prefix,CMPName,GSLID,GodownNo,Commodity,Client,BalanceBags,BalanceWeight,BagSize,ContractID,NContractID,G_UID,CDTFDate,CDTFNo From ("
'tmp = tmp & " Select s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.GSLID,MG.GodownNo,mcom.Commodity,"
'tmp = tmp & " Mclient.ClientName+'#'+Mclient.ClientID+'#'+Convert(Varchar,Mclient.ClientIDRow) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight,"
'tmp = tmp & " (GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,GSL.CDTFDate,GSL.CDTFNo "
'tmp = tmp & " From  Mst_GSL GSL"
'tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
'tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
'tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
'tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
'tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientIDRow = Mclient.ClientIDRow"
'tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate > '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0' And GSL.Flag Not in ('Z') and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1')"
'tmp = tmp & " And Convert(Varchar,GSL.GSLID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where b.HologramNo is not null And b.HologramNo <> '' )"
'tmp = tmp & " Union All"
'tmp = tmp & " Select s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.GSLID,MG.GodownNo,mcom.Commodity,Mclient.ClientName+'#'+Mclient.ClientID+'#'+Convert(Varchar,Mclient.ClientIDRow) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,GSL.CDTFDate,GSL.CDTFNo "
'tmp = tmp & " From  Mst_GSL GSL"
'tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
'tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
'tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
'tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
'tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
'tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientIDRow = Mclient.ClientIDRow"
'tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate > '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Flag Not in ('Z') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1') "
'tmp = tmp & " And Convert(Varchar,GSL.GSLID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.GSLID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where b.CxTFDate >= '2017-12-01')) A"
'tmp = tmp & " Order By StateName,LocationName,SM_Prefix,CMPName,GSLID"


tmp = " Select StateName,LocationName,SM_Prefix,CMPName,Mst_GSL_ID,GodownNo,Commodity,Client,BalanceBags,BalanceWeight,BagSize,ContractID,NContractID,G_UID,CDTFDate,CxTF_DetailsID,Madeupflag From ("
tmp = tmp & " Select Distinct s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.Mst_GSL_ID,MG.GodownNo,mcom.Commodity,Mclient.ClientName+'#'+Mclient.ClientIDRow+'#'+Convert(Varchar,"
tmp = tmp & " Mclient.ClientID) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight, (GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,"
tmp = tmp & " GSL.CDTFDate , GSL.CxTF_DetailsID, GSL.Madeupflag"
tmp = tmp & " From  Mst_GSL GSL"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on MG.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on GSL.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0' and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And  GSL.CDTFDate <= '2018-02-01' And GSL.Flag not in ('Z') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1')"
tmp = tmp & " Union All"
tmp = tmp & " Select s.StateName,l.LocationName,l.SM_Prefix,Mc.CMPName,GSL.Mst_GSL_ID,MG.GodownNo,mcom.Commodity,Mclient.ClientName+'#'+Mclient.ClientIDRow+'#'+Convert(Varchar,Mclient.ClientID) 'Client'  ,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,GSL.G_UID,GSL.CDTFDate,GSL.CxTF_DetailsID, GSL.Madeupflag"
tmp = tmp & " From  Txn_CxTF_GSL  a"
tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
tmp = tmp & " Inner Join Mst_GSL GSL On a.Mst_GSL_ID = GSL.Mst_GSL_ID and a.SM_Prefix = GSL.SM_Prefix"
tmp = tmp & " Inner Join Mst_Godown MG On GSL.GodownID = MG.GodownID"
tmp = tmp & " Inner Join Mst_CMP mC on b.CMPID = mc.CMPID"
tmp = tmp & " Inner Join Mst_Location l on mc.LocationID = l.LocationID"
tmp = tmp & " Inner Join Mst_State s on l.StateID = s.StateID"
tmp = tmp & " Inner Join Mst_Commodity mcom on b.CommodityID = mcom.CommodityID"
tmp = tmp & " Inner Join Mst_Client Mclient On GSL.ClientID = Mclient.ClientID"
tmp = tmp & " Where a.TxnType = 'W' And b.CxTFDate > '2017-12-01' and GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate > '2017-01-01'   And b.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And isnull(GSL.ContractID,0)=0 And GSL.Flag in ('E') And MG.YearlyBillingFlg+'-'+Mcom.YearlyBillingFlg not in ('1-1')) A"
tmp = tmp & " Order By StateName,LocationName,SM_Prefix,CMPName,Mst_GSL_ID"


Call TmpTable

Call Grid_Head

For iu = 1 To TmpRs.RecordCount
    MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
    For C = 0 To MSHFlexGrid1.Cols - 1
        MSHFlexGrid1.TextMatrix(iu, C) = IIf(IsNull(TmpRs.Fields(C)), "", TmpRs.Fields(C))
    Next
    TmpRs.MoveNext
Next

Call CmdExcel_Click

End Sub

Private Sub CmdUpdate_Click()

If CmbLocationID.Text = "" Then
   MsgBox "Blank Location not allowed!!!"
   CmbLocationID.SetFocus
   Exit Sub
End If
If CmbCMPID.Text = "" Then
   MsgBox "Blank CMP not allowed!!!"
   CmbCMPID.SetFocus
   Exit Sub
End If
If CmbModuleName.Text = "" Then
   MsgBox "Blank Module Name not allowed!!!"
   CmbModuleName.SetFocus
   Exit Sub
End If
If TxtModuleNo = "" Then
   MsgBox "Blank " & LblModuleNo.Caption & " not allowed!!!"
   TxtModuleNo.SetFocus
   Exit Sub
End If

If Check1.Value = 1 Then
   If dtpNewCxtfDate.Value > CDate(TxtCurrentDate) Then
      MsgBox "New date must be less than CMP Current Date!!"
      dtpNewCxtfDate.SetFocus
      Exit Sub
   End If
End If

If Check3.Value = 1 Then
   If CmbAgent.Text = "" Then
      MsgBox "Blank Agent Name not allowed!!!"
      CmbAgent.SetFocus
      Exit Sub
   End If
End If

If TxtRemark = "" Then
   MsgBox "Blank Remark not allowed!!!"
   TxtRemark.SetFocus
   Exit Sub
End If

'tmp = "Select * from Txn_CxTF_Details Where SM_Prefix = '" & TxtPreFix & "' And CxTF_DetailsID = " & TxtModuleNo & " And TxnType = '" & mType & "'"
'Call TmpTable
If Check1.Value = 1 Then
   tmp = "Select Max(YearMonth) as 'YearMonth' From Txn_InvMonGSLDetail Where SM_Prefix = '" & TxtPreFix & "'"
   Call TmpTable
   
   If TmpRs.RecordCount > 0 Then
      If IsNull(TmpRs!YearMonth) = False Then
         'If Format(dtpNewCxtfDate, "mmyyyy") <= Format(TmpRs!YearMonth, "mmyyyy") Then
         'If Month(TmpRs!YearMOnth) = 1 Then
         '   mDate = TmpRs!YearMOnth + 31
         'Else
         '   mDate = TmpRs!YearMOnth + 31
         'End If
         
         mTotalDays = Day(DateSerial(Year(TmpRs!YearMonth), Month(TmpRs!YearMonth) + 1, 0)) ' + 1
         mToDate = Format(Year(TmpRs!YearMonth) & "-" & Month(TmpRs!YearMonth) & "-" & mTotalDays, "yyyy-mm-dd")

         
         If dtpNewCxtfDate.Value <= CDate(mToDate) Then 'TmpRs!YearMOnth Then
            MsgBox "Billing Process executed for selected location for new date. Can't edit!!!"
            dtpNewCxtfDate.SetFocus
            Exit Sub
         End If
      End If
   End If
   
   RsTxn_CxTF_Details!CxTFDate = dtpNewCxtfDate.Value
End If
If Check2.Value = 1 Then
   RsTxn_CxTF_Details!ManualCDTFNo = TxtNewManualNo
End If
If Check3.Value = 1 Then
   RsTxn_CxTF_Details!AgentID = mAgentID 'dtpNewCxtfDate.Value
   RsTxn_CxTF_Details!Agent = CmbAgent.Text
End If
RsTxn_CxTF_Details.Update

If CmbModuleName.Text = "CDTF" Then
   If Check1.Value = 1 Then
      tmp = "Update Mst_GSL Set CDTFDate='" & dtpNewCxtfDate.Value & "' Where CxTF_DetailsID=" & Val(TxtModuleNo.Text) & " And SM_Prefix='" & TxtPreFix.Text & "' And CMPID=" & mCMPID & ""
      Call TmpTable
   End If
End If
Call SaveLogHistory
RsTxn_CxTF_Details.Filter = ""
RsTxn_CxTF_Details.Requery
MsgBox " Done!!"
Call ClearData
End Sub
Private Sub SaveLogHistory()
Call TableLog_UpdateCxTFDetails
RsLog_UpdateCxTFDetails.AddNew
RsLog_UpdateCxTFDetails!CxTFUpdateID = GetMaxID
RsLog_UpdateCxTFDetails!ModuleType = CmbModuleName.Text
RsLog_UpdateCxTFDetails!ModuleNo = Val(TxtModuleNo.Text)
RsLog_UpdateCxTFDetails!PrevCXTFDate = TxtModuleDate.Text
If Check1.Value = 1 Then
   RsLog_UpdateCxTFDetails!NewCXTFDate = dtpNewCxtfDate.Value
End If
RsLog_UpdateCxTFDetails!PrevAgentID = mPrevAgentID
RsLog_UpdateCxTFDetails!NewAgentID = mAgentID
RsLog_UpdateCxTFDetails!PrevManualCDTFNo = TxtPrevManualNo
RsLog_UpdateCxTFDetails!NewManualCDTFNo = TxtNewManualNo

RsLog_UpdateCxTFDetails!Remark = TxtRemark.Text
RsLog_UpdateCxTFDetails!UpdatedBy = Gen_UserLoginID
RsLog_UpdateCxTFDetails.Update
End Sub
Private Function GetMaxID() As Double
Dim Retval As Double
tmp = "Select max(CxTFUpdateID) from Log_UpdateCxTFDetails"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxID = Retval
End Function

Private Sub CmdUpdateBillingRate_Click()

'Commented by Parag on 02 Jan 2018 for implementing Billing Details to GSL in transaction and master from  Deposit Contract
'''''''''''mFromDate = Year(Date) & "-" & Month(Date) & "-01"
'''''''''''mSMonth = Month(Date)
'''''''''''mSYear = Year(Date)
'''''''''''mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
'''''''''''
'''''''''''
'''''''''''tmp = "Select Distinct SM_Prefix from Mst_GSL Where BillingCycleID is null And Flag not in ('E','Z','B') And SpillageFlag = '0' And MadeupFlag = '0' " ' ANd SM_Prefix = 'NVARM'"
'''''''''''Call Tmp1Table
'''''''''''ProgressBar1.Value = 0
'''''''''''ProgressBar1.Max = TmpRs1.RecordCount + 5
'''''''''''
'''''''''''
'''''''''''If TmpRs1.RecordCount > 0 Then
'''''''''''   For I = 1 To TmpRs1.RecordCount
'''''''''''
'''''''''''       tmp = " Update Mst_GSL Set BagSize = (DepositeWeight/DepositeBags)*1000 where flag = 'O' and SpillageFlag = 0 And DepositeBags <> 0 And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''       Call TmpTable
'''''''''''
'''''''''''       tmp = " Update Mst_GSL Set BagSize = Round(Bagsize,0) Where SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''       Call TmpTable
'''''''''''
'''''''''''
'''''''''''       tmp = " Update Mst_GSL Set Mst_GSL.BillingCycleID ="
'''''''''''       tmp = tmp & "  (Select a.BillingCycleID"
'''''''''''       tmp = tmp & "  from  Mst_GSL b"
'''''''''''       tmp = tmp & "  inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''       tmp = tmp & "  inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''       tmp = tmp & "  left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''       tmp = tmp & "  And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''       tmp = tmp & "  where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''       tmp = tmp & "  And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''       tmp = tmp & "  Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''       tmp = tmp & "  a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''''''''''       tmp = tmp & "  a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''       tmp = tmp & "  Where  Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''       Call TmpTable
'''''''''''
'''''''''''
'''''''''''       tmp = "  Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''       tmp = tmp & "  (Select a.DepoBillingCycleID"
'''''''''''       tmp = tmp & "  from  Txn_CxTF_GSL b"
'''''''''''       tmp = tmp & "  inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''       tmp = tmp & "  inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''       tmp = tmp & "  Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''       tmp = tmp & "  inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''       tmp = tmp & "  left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''       tmp = tmp & "  And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''       tmp = tmp & "  Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''       tmp = tmp & "  And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''       tmp = tmp & "  Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''       tmp = tmp & "  a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''       tmp = tmp & "  a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''       tmp = tmp & "  Where Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''       Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select a.BillingCycleID"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
''''''''''' '       ---------------------------------------------------------------------------------------------------
''''''''''' '       ---------------------------------------------------------------------------------------------------
'''''''''''
''''''''''' '       --Billing Unit
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select a.BillingUnit"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where  Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select a.BillingUnit"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and b.BillingCycleID is not null"
'''''''''''        tmp = tmp & "  Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select a.BillingUnit"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' And b.BillingCycleID is not null"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
''''''''''''        ---------------------------------------------------------------------------------------------------
''''''''''''        ---------------------------------------------------------------------------------------------------
'''''''''''
''''''''''''        ---Rate
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select a.Amount"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.BillingUnit = 'Q'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select a.Amount"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select a.Amount"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.RatePerBag"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''        tmp = tmp & " And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCardDetail  ad on a.SPRCID = ad.SPRCID"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.BillingUnit = 'B'"
'''''''''''        tmp = tmp & " And b.BagSize >= ad.FromBagSize And b.BagSize <= ad.ToBagSize"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.BillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount,ad.RatePerBag)"
'''''''''''        tmp = tmp & " Where  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.RatePerBag"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCardDepoDetail  ad on a.SPRCID = ad.SPRCID"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
'''''''''''        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount,ad.RatePerBag)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.RatePerBag"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.ClientIDRow = b.ClientIDRow And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " left join Mst_SpecialRateCardDetail  ad on a.SPRCID = ad.SPRCID"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SPRCID is not null And a.FLag = 'A' And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
'''''''''''        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SPRCID, a.DepoBillingCycleID,a.BillingUnit,"
'''''''''''        tmp = tmp & " a.RateOnQty,a.BaseRate,a.AbsoluteDiscount,a.Amount,ad.RatePerBag)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateCardType = 'S' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Flag = 'O' And RateCardType is null"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'S' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'D' And RateCardType is null"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'S' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'W' And RateCardType is null"
'''''''''''        Call TmpTable
''''''''''''        ---------------------------------------------------------------------------------------------------
''''''''''''        ---------------------------------------------------------------------------------------------------
'''''''''''
''''''''''''        ---Base Rate Card
'''''''''''
''''''''''''        ---Billing Cycle ID
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select a.BillingCycleID"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SRCID is not null And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
'''''''''''        tmp = tmp & " a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Mst_GSL.RateCardType is null And  Mst_GSL.BillingCycleID is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select a.DepoBillingCycleID"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.flag,a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select a.BillingCycleID"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.SRCID, a.BillingCycleID,a.flag,a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
''''''''''''        ---------------------------------------------------------------------------------------------------
''''''''''''        ---------------------------------------------------------------------------------------------------
'''''''''''
''''''''''''        --Billing Unit
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select a.flag"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SRCID is not null And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
'''''''''''        tmp = tmp & " a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Mst_GSL.RateCardType is null And  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select a.flag"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " And  b.BillingCycleID is not null"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.flag,a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select a.flag"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " And  b.BillingCycleID is not null"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.SRCID, a.BillingCycleID,a.flag,a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And  Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
''''''''''''        ---------------------------------------------------------------------------------------------------
''''''''''' '       ---------------------------------------------------------------------------------------------------
'''''''''''
''''''''''''        --Billing Rate
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select a.RateOnQty"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid"
'''''''''''        tmp = tmp & " And  a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.SRCID is not null  And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.Flag = 'Q' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
'''''''''''        tmp = tmp & " a.RateOnQty)"
'''''''''''        tmp = tmp & " Where  Mst_GSL.RateCardType is null And Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select a.RateOnQty"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select a.RateOnQty"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid"
'''''''''''        tmp = tmp & " And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q' and a.DefaultRate = '1'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.SRCID, a.DepoBillingCycleID,a.Flag,a.RateOnQty)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select distinct ad.RatePerBag"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =b.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCardDetail  ad on a.SRCID = ad.SRCID"
'''''''''''        tmp = tmp & " where (b.ContractID = 0 Or b.ContractID is Null) And a.DefaultRate = 1 And a.SRCID is not null  And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And a.Flag = 'B'"
'''''''''''        tmp = tmp & " And b.BagSize >= ad.FromBagSize And b.BagSize <= ad.ToBagSize"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
'''''''''''        tmp = tmp & " a.RateOnQty,ad.RatePerBag)"
'''''''''''        tmp = tmp & " Where  Mst_GSL.RateCardType is null And Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.RatePerBag"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCardDepoDetail  ad on a.SRCID = ad.SRCID"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.DefaultRate = 1 And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
'''''''''''        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SRCID, a.DepoBillingCycleID,a.Flag,"
'''''''''''        tmp = tmp & " a.RateOnQty,ad.RatePerBag)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.RatePerBag"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCard  a on a.locationid = l.locationid And a.exchangetypeid =d.exchangetypeid And a.commoditygroupid = c.commoditygroupid"
'''''''''''        tmp = tmp & " left join Mst_StorageRateCardDetail  ad on a.SRCID = ad.SRCID"
'''''''''''        tmp = tmp & " Where (g.ContractID = 0 Or g.ContractID is Null) And a.DefaultRate = 1 And a.SRCID is not null  And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 1 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'B'"
'''''''''''        tmp = tmp & " And g.BagSize >= ad.FromBagSize And g.BagSize <= ad.ToBagSize"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.SRCID, a.BillingCycleID,a.Flag,"
'''''''''''        tmp = tmp & " a.RateOnQty,ad.RatePerBag)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.RateCardType is null And Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateCardType = 'N' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Flag = 'O' And RateCardType is null"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'N' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'D' And RateCardType is null"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateCardType = 'N' Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And BillingCycleID is not null And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And TxnType = 'W' And RateCardType is null"
'''''''''''        Call TmpTable
'''''''''''
''''''''''''        ------------------------------------------------------------------------------------------------------------
''''''''''''        ------------------------------------------------------------------------------------------------------------
'''''''''''
''''''''''''        --NCDEX Rate
'''''''''''
''''''''''''        --Billing Cycle and Billing Unit
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set BillingCycleID =  1 , BillingUnit = 'Q' Where ExchangeTypeID = 2 And Flag = 'O' And BillingCycleID is null And SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select Distinct 1 'BillingCycleID'"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Where d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "')"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null and Txn_CxTF_GSL.TxnType in ('D','W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select Distinct 'Q' 'BillingUnit'"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Where d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "')"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null and Txn_CxTF_GSL.TxnType in ('D','W') And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType in ('D','W') And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.Rate"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_NCDEXStorageRateCard  a on a.locationid = l.locationid And a.CommodityGroupID =c.CommodityGroupID"
'''''''''''        tmp = tmp & " left Join Mst_NCDEXStorageRateCardDetail ad On a.NSRCID = ad.NSRCID And b.GodownID = ad.GodownID "
'''''''''''        tmp = tmp & " Where (b.ContractID = 0 Or b.ContractID Is Null) And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix and b.BillingCycleID is not null And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " Group by b.GodownID,ad.Rate)"
'''''''''''        tmp = tmp & " Where  Mst_GSL.BillingCycleID is not null And Mst_GSL.ExchangeTypeID = 2 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O' And Mst_GSL.BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.Rate"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_NCDEXStorageRateCard  a on a.locationid = l.locationid And a.CommodityGroupID =c.CommodityGroupID"
'''''''''''        tmp = tmp & " left Join Mst_NCDEXStorageRateCardDetail ad On a.NSRCID = ad.NSRCID And g.GodownID = ad.GodownID "
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,ad.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('D') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select ad.Rate"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_NCDEXStorageRateCard  a on a.locationid = l.locationid And a.CommodityGroupID =c.CommodityGroupID"
'''''''''''        tmp = tmp & " left Join Mst_NCDEXStorageRateCardDetail ad On a.NSRCID = ad.NSRCID And g.GodownID = ad.GodownID "
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix And d.ExchangeTypeID = 2 And d.CxTFDate >= '" & CDate(mFromDate) & "'"
'''''''''''        tmp = tmp & " and b.BillingCycleID is not null And b.BillingUnit = 'Q'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,ad.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is not null And Txn_CxTF_GSL.TxnType in ('W') And Txn_CxTF_GSL.BillingUnit = 'Q' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 2)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select d.BillingCycleID"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
'''''''''''        tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is null And Mst_GSL.RateCardType is null  And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select 'B'"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'No Of Bags'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
'''''''''''        tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is not null And  Mst_GSL.BillingUnit is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select 'W'"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'Weight In MT'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
'''''''''''        tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is not null And  Mst_GSL.BillingUnit is null And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set Mst_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select d.BillingRate"
'''''''''''        tmp = tmp & " from  Mst_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on b.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And a.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  d on a.ContractID = d.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (b.ContractID Is Not Null And b.ContractID <> 0) And b.BillingCycleID Is Not Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Mst_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Mst_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between d.FromDate and d.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,"
'''''''''''        tmp = tmp & " a.ContractID, d.BillingCycleID,a.Flag,d.BillingRate)"
'''''''''''        tmp = tmp & " Where (Mst_GSL.ContractID is not null and Mst_GSL.ContractID <>  0) And  Mst_GSL.BillingCycleID is not null And Mst_GSL.RateCardType is not null  And Mst_GSL.ExchangeTypeID = 1 And Mst_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Mst_GSL.Flag = 'O'"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select MSB.BillingCycleID"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select MSB.BillingCycleID"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,MSB.BillingRate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingCycleID ="
'''''''''''        tmp = tmp & " (Select MSB.BillingCycleID"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,MSB.BillingRate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''
''''''''''''        -----------------------------------------------------------------------
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select 'W'"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'Weight In MT'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select 'B'"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'No Of Bags'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select 'W'"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'Weight In MT'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.BillingUnit ="
'''''''''''        tmp = tmp & " (Select 'B'"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0 and a.BookedAsPer = 'No Of Bags'"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingCycleID,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select MSB.BillingRate"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingRate,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'D' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.Rate ="
'''''''''''        tmp = tmp & " (Select MSB.BillingRate"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join mst_location l on b.SM_Prefix = l.SM_Prefix"
'''''''''''        tmp = tmp & " inner join Txn_CxTF_Details d on b.CxTF_DetailsID = d.CxTF_DetailsID and b.SM_Prefix = d.SM_Prefix and b.TxnType = d.TxnType"
'''''''''''        tmp = tmp & " Inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " inner join mst_commodity c on d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_StorageContract  a on a.locationid = l.locationid And d.commodityid = c.commodityid"
'''''''''''        tmp = tmp & " left join Mst_SContractBillingDetail  MSB on a.ContractID = MSB.ContractID"
'''''''''''        tmp = tmp & " Left Join Txn_SRGodownMapping TSGM on a.ContractID  = TSGM.ContractID"
'''''''''''        tmp = tmp & " left Join Txn_SRGodownMapDetails TSGMD On TSGM.SRGodownMappingID = TSGMD.SRGodownMappingID"
'''''''''''        tmp = tmp & " Where (g.ContractID Is Not Null And g.ContractID <> 0) And b.BillingCycleID Is Null And a.ContractID Is Not Null And b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix  And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' and a.Flag = 'A' And '" & CDate(mFromDate) & "' between MSB.FromDate and MSB.ToDate and TSGMD.ReservationQty = 0"
'''''''''''        tmp = tmp & " Group by b.Mst_GSL_ID,b.SM_Prefix,a.ContractID, MSB.BillingRate,a.Flag,b.Rate)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID is null And Txn_CxTF_GSL.RateCardType is null  And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "'"
'''''''''''        tmp = tmp & " And Txn_CxTF_GSL.CxTF_DetailsID in (Select Distinct CxTF_DetailsID from Txn_CxTF_Details Where TxnType = 'W' And SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTFDate >= '" & CDate(mFromDate) & "' and ExchangeTypeID = 1)"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''
''''''''''''-------------------------------------------------
'''''''''''        tmp = "Update Mst_GSL Set RateForDailyPnL = (Rate) Where BillingCycleID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = "Update Mst_GSL Set RateForDailyPnL = Round((((Rate/BagSize)*1000)),0) Where BillingCycleID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = "Update Mst_GSL Set RateForDailyPnL = (((Rate)*4) / " & mTotalDays & ") Where BillingCycleID = 2 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateForDailyPnL = Round(((((Rate/BagSize)*1000)*4)/ " & mTotalDays & "),0) Where BillingCycleID = 2 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateForDailyPnL = (((Rate)*2)/ " & mTotalDays & ") Where BillingCycleID = 3 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateForDailyPnL = Round(((((Rate/BagSize)*1000)*2)/ " & mTotalDays & "),0) Where BillingCycleID = 3 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = "Update Mst_GSL Set RateForDailyPnL = ((Rate)/ " & mTotalDays & ") Where BillingCycleID = 4 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = "Update Mst_GSL Set RateForDailyPnL = ((Rate)/ " & mTotalDays & ") Where BillingCycleID = 5 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateForDailyPnL = Round((((Rate/BagSize)*1000)/ " & mTotalDays & "),0) Where BillingCycleID = 4 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Mst_GSL Set RateForDailyPnL = Round((((Rate/BagSize)*1000)/ " & mTotalDays & "),0) Where BillingCycleID = 5 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'B'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = Rate Where RateForDailyPnL is null and BillingCycleID = 1 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (((Rate)*4)/ " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 2 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (((Rate)*2)/ " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 3 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (Rate / " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 4 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set RateForDailyPnL = (Rate / " & mTotalDays & ") Where RateForDailyPnL is null and BillingCycleID = 5 And SM_Prefix = '" & TmpRs1!SM_Prefix & "' ANd BillingUnit = 'Q'"
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
'''''''''''        tmp = tmp & " (Select Distinct Round((((b.Rate/g.BagSize)*1000)),0)"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 1 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 1 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
'''''''''''        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000)*4)/" & mTotalDays & "),0)"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 2 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 2 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
'''''''''''        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000)*2)/" & mTotalDays & "),0)"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 3 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 3 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
'''''''''''        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000))/" & mTotalDays & "),0)"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 4 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 4 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''
'''''''''''        tmp = " Update Txn_CxTF_GSL Set Txn_CxTF_GSL.RateForDailyPnL ="
'''''''''''        tmp = tmp & " (Select Distinct Round(((((b.Rate/g.BagSize)*1000))/" & mTotalDays & "),0)"
'''''''''''        tmp = tmp & " from  Txn_CxTF_GSL b"
'''''''''''        tmp = tmp & " inner join Mst_GSL g On b.Mst_GSL_ID = g.Mst_GSL_ID and b.SM_Prefix = g.SM_Prefix"
'''''''''''        tmp = tmp & " Where b.Mst_GSL_ID = Txn_CxTF_GSL.Mst_GSL_ID"
'''''''''''        tmp = tmp & " And b.SM_Prefix = Txn_CxTF_GSL.SM_Prefix and b.BillingCycleID = 5 And b.BillingUnit = 'B' And b.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And b.RateForDailyPnL is null)"
'''''''''''        tmp = tmp & " Where Txn_CxTF_GSL.BillingCycleID = 5 And Txn_CxTF_GSL.BillingUnit = 'B' And Txn_CxTF_GSL.SM_Prefix = '" & TmpRs1!SM_Prefix & "' And Txn_CxTF_GSL.RateForDailyPnL is null"
'''''''''''
'''''''''''        Call TmpTable
'''''''''''
'''''''''''        ProgressBar1.Value = ProgressBar1.Value + 1
'''''''''''
'''''''''''        TmpRs1.MoveNext
'''''''''''   Next
'''''''''''End If


'tmp = " Update Txn_CxTF_GSL Set RateCardType = Null,BillingCycleID = Null,BillingUnit = Null,Rate = Null,RateForDailyPnL = Null,DepConDID =  Null,NoofDays = Null Where TxnType='D' And DepConDID is not null  "
'Call TmpTable
'
'tmp = " Update Txn_CxTF_GSL Set RateCardType = Null,BillingCycleID = Null,BillingUnit = Null,Rate = Null,RateForDailyPnL = Null,DepConDID =  Null,NoofDays = Null Where TxnType='W' And DepConDID is Not null "
'Call TmpTable
'
'tmp = "Update Mst_GSL Set RateCardType = Null,BillingCycleID = Null,BillingUnit = Null,Rate = Null,RateForDailyPnL = Null,DepConDID =  Null,NoofDays = Null Where ExchangeTypeID = 1 And DepConDID is not null "
'Call Tmp1Table


mCutoffDate = "2017-12-01"

'tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientIDRow,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,DepBillingCycleID,BagSizeFlag ,FinalPrice,DepositFinalPrice,G_UID,DepConDID"
'tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientIDRow,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
'tmp = tmp & " From  Txn_CxTF_GSL TCG"
'tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
'tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
'tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((TCG.DematWeight/TCG.NoOfBags)*1000) Between  MinBagSize And MaxBagSize"
'tmp = tmp & " Left Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'N'"
'tmp = tmp & " Where TC.CxTFDate >= '" & mCutoffDate & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
'tmp = tmp & " Union All"
'tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientIDRow,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
'tmp = tmp & " From  Txn_CxTF_GSL TCG"
'tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
'tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
'tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
'tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'N'"
'tmp = tmp & " Where TC.CxTFDate >= '" & mCutoffDate & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
'tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"
'
'Call TmpTable


''''''''''''''tmp = " Select SM_Prefix,ClientIDRow,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,BalanceBags,BalanceWeight,BagSize,ContractID,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID,CDTFDate,DepConID,GSLDCDID"
''''''''''''''tmp = tmp & " From (Select GSL.SM_Prefix ,GSL.ClientIDRow,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID'"
''''''''''''''tmp = tmp & " From  Mst_GSL GSL"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'N'"
''''''''''''''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
''''''''''''''tmp = tmp & " Union All"
''''''''''''''tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientIDRow,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID'"
''''''''''''''tmp = tmp & " From  Mst_GSL GSL"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'N'"
''''''''''''''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags > 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N' "
''''''''''''''tmp = tmp & " Union All "
''''''''''''''tmp = tmp & " Select SM_Prefix,ClientIDRow,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,BalanceBags,BalanceWeight,BagSize,ContractID,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID,CDTFDate,DepConID,GSLDCDID"
''''''''''''''tmp = tmp & " From (Select GSL.SM_Prefix ,GSL.ClientIDRow,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID'"
''''''''''''''tmp = tmp & " From  Mst_GSL GSL"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'N'"
''''''''''''''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And GSL.Flag = 'Z' And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
''''''''''''''tmp = tmp & " Union All"
''''''''''''''tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientIDRow,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID'"
''''''''''''''tmp = tmp & " From  Mst_GSL GSL"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'N'"
''''''''''''''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0  And GSL.Flag = 'Z' And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N' "
''''''''''''''tmp = tmp & " Union All "
''''''''''''''tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientIDRow,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID'"
''''''''''''''tmp = tmp & " From  Mst_GSL GSL"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'N'"
''''''''''''''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y' And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.TxnType = 'W' And b.CxTFDate >= '2017-12-01') "
''''''''''''''tmp = tmp & " Union All"
''''''''''''''tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientIDRow,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID'"
''''''''''''''tmp = tmp & " From  Mst_GSL GSL"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientIDRow = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'N'"
''''''''''''''tmp = tmp & " Where GSL.DepositeBags > 0 And GSL.BalanceBags = 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N' And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.TxnType = 'W' And b.CxTFDate >= '2017-12-01')) A"
''''''''''''''tmp = tmp & " Order By SM_Prefix,Mst_GSL_ID"
''''''''''''''
''''''''''''''Call TmpTable
''''''''''''''
''''''''''''''
''''''''''''''
''''''''''''''
''''''''''''''
''''''''''''''For i = 1 To TmpRs.RecordCount
''''''''''''''    mDepoBillingCycleID = ""
''''''''''''''    mDepoBillingUnit = ""
''''''''''''''    mNoofdays = ""
''''''''''''''    mDepConDID = ""
''''''''''''''    mDepConID = ""
''''''''''''''    mDepoBillingCycleID = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID)
''''''''''''''    If TmpRs!BagSizeFlag = "Y" Then
''''''''''''''       mDepoBillingUnit = "B"
''''''''''''''    Else
''''''''''''''       mDepoBillingUnit = "Q"
''''''''''''''    End If
''''''''''''''    mDepoBillingRate = IIf(IsNull(TmpRs!FinalPrice), 0, TmpRs!FinalPrice)
''''''''''''''    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
''''''''''''''    mDepConID = IIf(IsNull(TmpRs!DepConID), "", TmpRs!DepConID)
''''''''''''''
''''''''''''''    If mDepoBillingCycleID = 1 Then
''''''''''''''       mNoofdays = "0"
''''''''''''''    ElseIf mDepoBillingCycleID = 2 Then
''''''''''''''       mNoofdays = "4"
''''''''''''''    ElseIf mDepoBillingCycleID = 3 Then
''''''''''''''       mNoofdays = "2"
''''''''''''''    ElseIf mDepoBillingCycleID = 4 Then
''''''''''''''       mNoofdays = "1"
''''''''''''''    End If
''''''''''''''
''''''''''''''    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofdays <> "" Then
''''''''''''''       tmp = "Update Mst_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofdays & ",UpdatedBy = '" & Gen_UserLoginID & "',UpdatedDate = getdate() Where ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And DepConDID is null "
''''''''''''''       Call Tmp1Table
''''''''''''''    End If
''''''''''''''
''''''''''''''    TmpRs.MoveNext
''''''''''''''Next
''''''''''''''
''''''''''''''
''''''''''''''tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientIDRow,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID"
''''''''''''''tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientIDRow,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
''''''''''''''tmp = tmp & " From  Txn_CxTF_GSL TCG"
''''''''''''''tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
''''''''''''''tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A')  "
''''''''''''''tmp = tmp & " Where TC.CxTFDate >= '" & mCutoffDate & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
''''''''''''''tmp = tmp & " Union All"
''''''''''''''tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientIDRow,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
''''''''''''''tmp = tmp & " From  Txn_CxTF_GSL TCG"
''''''''''''''tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
''''''''''''''tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A')  "
''''''''''''''tmp = tmp & " Where TC.CxTFDate >= '" & mCutoffDate & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
''''''''''''''tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"
''''''''''''''
''''''''''''''Call TmpTable
''''''''''''''
''''''''''''''For i = 1 To TmpRs.RecordCount
''''''''''''''    mDepoBillingCycleID = ""
''''''''''''''    mDepoBillingUnit = ""
''''''''''''''    mNoofdays = ""
''''''''''''''    mDepConDID = ""
''''''''''''''
''''''''''''''    mDepoBillingCycleID = IIf(IsNull(TmpRs!WBillingCycleID), "", TmpRs!WBillingCycleID)
''''''''''''''    If TmpRs!BagSizeFlag = "Y" Then
''''''''''''''       mDepoBillingUnit = "B"
''''''''''''''    Else
''''''''''''''       mDepoBillingUnit = "Q"
''''''''''''''    End If
''''''''''''''    mDepoBillingRate = IIf(IsNull(TmpRs!WithdrawalFinalPrice), 0, TmpRs!WithdrawalFinalPrice)
''''''''''''''    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
''''''''''''''
''''''''''''''    If mDepoBillingCycleID = 1 Then
''''''''''''''       mNoofdays = Val(Day(TmpRs!CxTFDate))
''''''''''''''    ElseIf mDepoBillingCycleID = 2 Then
''''''''''''''       If Day(TmpRs!CxTFDate) < 8 Then
''''''''''''''          mNoofdays = "1"
''''''''''''''       ElseIf Day(TmpRs!CxTFDate) < 15 Then
''''''''''''''          mNoofdays = "2"
''''''''''''''       ElseIf Day(TmpRs!CxTFDate) < 22 Then
''''''''''''''          mNoofdays = "3"
''''''''''''''       Else
''''''''''''''          mNoofdays = "4"
''''''''''''''       End If
''''''''''''''    ElseIf mDepoBillingCycleID = 3 Then
''''''''''''''       If Day(TmpRs!CxTFDate) < 16 Then
''''''''''''''          mNoofdays = "1"
''''''''''''''       Else
''''''''''''''          mNoofdays = "2"
''''''''''''''       End If
''''''''''''''    ElseIf mDepoBillingCycleID = 4 Then
''''''''''''''       mNoofdays = "1"
''''''''''''''    End If
''''''''''''''
''''''''''''''    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofdays <> "" Then
''''''''''''''       tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofdays & " Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs!CxTF_DetailsID & "' And TxnType='W' And DepConDID is null And G_UID = '" & TmpRs!G_UID & "' "
''''''''''''''       Call Tmp1Table
''''''''''''''    End If
''''''''''''''
''''''''''''''    TmpRs.MoveNext
''''''''''''''Next
''''''''''''''
''''''''''''''
''''''''''''''tmp = " Select b.CxTFDate, a.*,c.DepConDID,dc.DepBillingCycleID,c.BillingUnit,c.NoofDays,dcgd.DepositFinalPrice,dc.BagSizeFlag     from Txn_CxTF_GSL a"
''''''''''''''tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
''''''''''''''tmp = tmp & " Inner Join Mst_GSL c On a.Mst_GSL_ID = c.Mst_GSL_ID and a.SM_Prefix = c.SM_Prefix"
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContractGodownDetail dcgd on c.DepConDID = dcgd.DepConDID"
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContract dc on dcgd.DepConID = dc.DepConID"
''''''''''''''tmp = tmp & " Where a.TxnType = 'D' And b.CxTFDate >= '" & mCutoffDate & "' And b.ExchangeTypeID = 1 And a.DepConDID is null"
''''''''''''''tmp = tmp & " And c.DepConDID is not null"
''''''''''''''
''''''''''''''Call TmpTable
''''''''''''''
''''''''''''''For i = 1 To TmpRs.RecordCount
''''''''''''''    mDepoBillingCycleID = ""
''''''''''''''    mDepoBillingUnit = ""
''''''''''''''    mNoofdays = ""
''''''''''''''    mDepConDID = ""
''''''''''''''
''''''''''''''    mDepoBillingCycleID = IIf(IsNull(TmpRs!DepBillingCycleID), "", TmpRs!DepBillingCycleID)
''''''''''''''    If TmpRs!BagSizeFlag = "Y" Then
''''''''''''''       mDepoBillingUnit = "B"
''''''''''''''    Else
''''''''''''''       mDepoBillingUnit = "Q"
''''''''''''''    End If
''''''''''''''    mDepoBillingRate = IIf(IsNull(TmpRs!DepositFinalPrice), 0, TmpRs!DepositFinalPrice)
''''''''''''''    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
''''''''''''''
''''''''''''''    If mDepoBillingCycleID = 1 Then
''''''''''''''       mNoofdays = a
''''''''''''''       mFromDate = Year(TmpRs!CxTFDate) & "-" & Month(TmpRs!CxTFDate) & "-01"
''''''''''''''       mSMonth = Month(TmpRs!CxTFDate)
''''''''''''''       mSYear = Year(TmpRs!CxTFDate)
''''''''''''''       mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
''''''''''''''       mNoofdays = (Val(mTotalDays) - Val(Day(TmpRs!CxTFDate))) + 1
''''''''''''''    ElseIf mDepoBillingCycleID = 2 Then
''''''''''''''       If Day(TmpRs!CxTFDate) < 8 Then
''''''''''''''          mNoofdays = "4"
''''''''''''''       ElseIf Day(TmpRs!CxTFDate) < 15 Then
''''''''''''''          mNoofdays = "3"
''''''''''''''       ElseIf Day(TmpRs!CxTFDate) < 22 Then
''''''''''''''          mNoofdays = "2"
''''''''''''''       Else
''''''''''''''          mNoofdays = "1"
''''''''''''''       End If
''''''''''''''    ElseIf mDepoBillingCycleID = 3 Then
''''''''''''''       If Day(TmpRs!CxTFDate) < 16 Then
''''''''''''''          mNoofdays = "2"
''''''''''''''       Else
''''''''''''''          mNoofdays = "1"
''''''''''''''       End If
''''''''''''''    ElseIf mDepoBillingCycleID = 4 Then
''''''''''''''       mNoofdays = "1"
''''''''''''''    End If
''''''''''''''
''''''''''''''    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofdays <> "" Then
''''''''''''''       tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofdays & " Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs!CxTF_DetailsID & "' And TxnType='D' And DepConDID is null And G_UID = '" & TmpRs!G_UID & "' "
''''''''''''''       Call Tmp1Table
''''''''''''''    End If
''''''''''''''
''''''''''''''    TmpRs.MoveNext
''''''''''''''Next
''''''''''''''
''''''''''''''
''''''''''''''tmp = " Select b.CxTFDate, a.*,c.DepConDID,dc.DepBillingCycleID,c.BillingUnit,c.NoofDays,dcgd.DepositFinalPrice,dc.BagSizeFlag,TINV.MonG_UID  from Txn_CxTF_GSL a"
''''''''''''''tmp = tmp & " Inner Join Txn_CxTF_Details b on a.CxTF_DetailsID = b.CxTF_DetailsID and a.SM_Prefix = b.SM_Prefix and a.TxnType = b.TxnType"
''''''''''''''tmp = tmp & " Inner Join Mst_GSL c On a.Mst_GSL_ID = c.Mst_GSL_ID and a.SM_Prefix = c.SM_Prefix"
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContractGodownDetail dcgd on c.DepConDID = dcgd.DepConDID"
''''''''''''''tmp = tmp & " Inner Join Txn_DepositContract dc on dcgd.DepConID = dc.DepConID"
''''''''''''''tmp = tmp & " Inner Join Txn_InvMonGSLDetail TINV On TINV.YearMonth = '2017-12-01' And a.SM_Prefix = TINV.SM_Prefix And a.Mst_GSL_ID = TINV.Mst_GSL_ID"
''''''''''''''tmp = tmp & " Where a.TxnType = 'D' And b.CxTFDate >= '2017-12-01' And b.ExchangeTypeID = 1 And a.DepConDID is not null And c.DepConDID is not null And TINV.DRNo is null"
''''''''''''''
''''''''''''''Call TmpTable
''''''''''''''
''''''''''''''For i = 1 To TmpRs.RecordCount
''''''''''''''    mDepoBillingCycleID = ""
''''''''''''''    mDepoBillingUnit = ""
''''''''''''''    mNoofdays = ""
''''''''''''''    mDepConDID = ""
''''''''''''''
''''''''''''''    mDepoBillingCycleID = IIf(IsNull(TmpRs!DepBillingCycleID), "", TmpRs!DepBillingCycleID)
''''''''''''''    If TmpRs!BagSizeFlag = "Y" Then
''''''''''''''       mDepoBillingUnit = "B"
''''''''''''''    Else
''''''''''''''       mDepoBillingUnit = "Q"
''''''''''''''    End If
''''''''''''''    mDepoBillingRate = IIf(IsNull(TmpRs!DepositFinalPrice), 0, TmpRs!DepositFinalPrice)
''''''''''''''    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
''''''''''''''
''''''''''''''    If mDepoBillingCycleID = 1 Then
''''''''''''''       mNoofdays = a
''''''''''''''       mFromDate = Year(TmpRs!CxTFDate) & "-" & Month(TmpRs!CxTFDate) & "-01"
''''''''''''''       mSMonth = Month(TmpRs!CxTFDate)
''''''''''''''       mSYear = Year(TmpRs!CxTFDate)
''''''''''''''       mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
''''''''''''''       mNoofdays = (Val(mTotalDays) - Val(Day(TmpRs!CxTFDate))) + 1
''''''''''''''    ElseIf mDepoBillingCycleID = 2 Then
''''''''''''''       If Day(TmpRs!CxTFDate) < 8 Then
''''''''''''''          mNoofdays = "4"
''''''''''''''       ElseIf Day(TmpRs!CxTFDate) < 15 Then
''''''''''''''          mNoofdays = "3"
''''''''''''''       ElseIf Day(TmpRs!CxTFDate) < 22 Then
''''''''''''''          mNoofdays = "2"
''''''''''''''       Else
''''''''''''''          mNoofdays = "1"
''''''''''''''       End If
''''''''''''''    ElseIf mDepoBillingCycleID = 3 Then
''''''''''''''       If Day(TmpRs!CxTFDate) < 16 Then
''''''''''''''          mNoofdays = "2"
''''''''''''''       Else
''''''''''''''          mNoofdays = "1"
''''''''''''''       End If
''''''''''''''    ElseIf mDepoBillingCycleID = 4 Then
''''''''''''''       mNoofdays = "1"
''''''''''''''    End If
''''''''''''''
''''''''''''''    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofdays <> "" Then
''''''''''''''       If mDepoBillingUnit = "B" Then
''''''''''''''          tmp = " Update Txn_InvMonGSLDetail Set Flag1 = 'D',DRNo = " & mDepConDID & ",BillingCycleID = " & mDepoBillingCycleID & " ,BillingUnit ='B',InvoiceDate = '2017-12-31',InvoiceAmount = " & mDepoBillingRate & "*" & mNoofdays & "*BalanceBags,Rate = " & mDepoBillingRate & ",NoofDays = " & mNoofdays & " Where TFlag = 'D' And DRNo is null And Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "' And MonG_UID = '" & TmpRs!MonG_UID & "' And YearMonth = '2017-12-01'"
''''''''''''''          Call Tmp1Table
''''''''''''''       ElseIf mDepoBillingUnit = "Q" Then
''''''''''''''          tmp = " Update Txn_InvMonGSLDetail Set Flag1 = 'D',DRNo = " & mDepConDID & ",BillingCycleID = " & mDepoBillingCycleID & " ,BillingUnit ='Q',InvoiceDate = '2017-12-31',InvoiceAmount = " & mDepoBillingRate & "*" & mNoofdays & "*BalanceWeight,Rate = " & mDepoBillingRate & ",NoofDays = " & mNoofdays & " Where TFlag = 'D' And DRNo is null And Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And SM_Prefix = '" & TmpRs!SM_Prefix & "' And MonG_UID = '" & TmpRs!MonG_UID & "' And YearMonth = '2017-12-01'"
''''''''''''''          Call Tmp1Table
''''''''''''''       End If
''''''''''''''    End If
''''''''''''''
''''''''''''''    TmpRs.MoveNext
''''''''''''''Next
''''''''''''''
''''''''''''''tmp = "Select c.*,d.DepConDID,dc.BillingCycleID,d.BillingUnit,d.NoofDays,dcgd.FinalPrice,dc.BagSizeFlag"
''''''''''''''tmp = tmp & " from Mst_GSL c"
''''''''''''''tmp = tmp & " LEFT  Join Mst_GSL d On c.ConnectedMst_GSL_ID = d.Mst_GSL_ID and c.SM_Prefix = d.SM_Prefix"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContractGodownDetail dcgd on d.DepConDID = dcgd.DepConDID"
''''''''''''''tmp = tmp & " Left Join Txn_DepositContract dc on dcgd.DepConID = dc.DepConID"
''''''''''''''tmp = tmp & " Where c.ExchangeTypeID = 1 And d.DepConDID is not null And c.DepConDID is null And c.MadeupFlag = '1' and c.BalanceBags > 0"
''''''''''''''
''''''''''''''Call TmpTable
''''''''''''''
''''''''''''''For i = 1 To TmpRs.RecordCount
''''''''''''''    mDepoBillingCycleID = ""
''''''''''''''    mDepoBillingUnit = ""
''''''''''''''    mNoofdays = ""
''''''''''''''    mDepConDID = ""
''''''''''''''    mDepConID = ""
''''''''''''''    mDepoBillingCycleID = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID)
''''''''''''''    If TmpRs!BagSizeFlag = "Y" Then
''''''''''''''       mDepoBillingUnit = "B"
''''''''''''''    Else
''''''''''''''       mDepoBillingUnit = "Q"
''''''''''''''    End If
''''''''''''''    mDepoBillingRate = IIf(IsNull(TmpRs!FinalPrice), 0, TmpRs!FinalPrice)
''''''''''''''    mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
'''''''''''''''    mDepConID = IIf(IsNull(TmpRs!DepConID), "", TmpRs!DepConID)
''''''''''''''
''''''''''''''    If mDepoBillingCycleID = 1 Then
''''''''''''''       mNoofdays = "0"
''''''''''''''    ElseIf mDepoBillingCycleID = 2 Then
''''''''''''''       mNoofdays = "4"
''''''''''''''    ElseIf mDepoBillingCycleID = 3 Then
''''''''''''''       mNoofdays = "2"
''''''''''''''    ElseIf mDepoBillingCycleID = 4 Then
''''''''''''''       mNoofdays = "1"
''''''''''''''    End If
''''''''''''''
''''''''''''''    If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofdays <> "" Then
''''''''''''''       tmp = "Update Mst_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofdays & " Where MadeUpFlag =  '1' And ExchangeTypeID = 1 And SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And DepConDID is null "
''''''''''''''       Call Tmp1Table
''''''''''''''    End If
''''''''''''''
''''''''''''''    TmpRs.MoveNext
''''''''''''''Next




tmp = " Select Distinct SM_Prefix from Txn_DepositContract Where ContractType = 'R' And ClosedOn is null and Verified = 'Y' "
Call Tmp2Table


For x = 1 To TmpRs2.RecordCount

    tmp = " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID, TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R' And TDC.ClosedOn is null "
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And GSL.DepositeBags > 0 And GSL.BalanceBags > 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " Union All"
    tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R'  And TDC.ClosedOn is null"
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And GSL.DepositeBags > 0 And GSL.BalanceBags > 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N'"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " Union All"
    tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R'  And TDC.ClosedOn is null"
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
    tmp = tmp & " And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And a.TxnType = 'W' And b.CxTFDate >= '2017-12-01')"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " Union All"
    tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R'  And TDC.ClosedOn is null"
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And GSL.DepositeBags > 0 And GSL.BalanceBags = 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N'"
    tmp = tmp & " And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And a.TxnType = 'W' And b.CxTFDate >= '2017-12-01')"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    
    Call TmpTable
    
    For I = 1 To TmpRs.RecordCount
        mDepoBillingCycleID = ""
        mDepoBillingUnit = ""
        mNoofDays = ""
        mDepConDID = ""
        
        mDepoBillingCycleID = IIf(IsNull(TmpRs!BillingCycleID), "", TmpRs!BillingCycleID)
        If TmpRs!BagSizeFlag = "Y" Then
           mDepoBillingUnit = "B"
        Else
           mDepoBillingUnit = "Q"
        End If
        mDepoBillingRate = IIf(IsNull(TmpRs!FinalPrice), 0, TmpRs!FinalPrice)
        mDepConDID = IIf(IsNull(TmpRs!DepConDID), "", TmpRs!DepConDID)
        If mDepoBillingCycleID = 1 Then
           mNoofDays = "0"
        ElseIf mDepoBillingCycleID = 2 Then
           mNoofDays = "4"
        ElseIf mDepoBillingCycleID = 3 Then
           mNoofDays = "2"
        ElseIf mDepoBillingCycleID = 4 Then
           mNoofDays = "1"
        End If
        mLastInvDate = ""
        mNewTransDate = ""
        
        If IsNull(TmpRs!YearMonth) = False Then
           mLastInvDate = TmpRs!YearMonth
           mNewTransDate = DateAdd("M", 1, CDate(mLastInvDate))
        End If
        
        If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
           tmp = "Update Mst_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs!SM_Prefix & "' And Mst_GSL_ID = '" & TmpRs!Mst_GSL_ID & "' And DepConDID is not null And G_UID = '" & TmpRs!G_UID & "' "
           Call Tmp1Table
           
           tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID"
           tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
           tmp = tmp & " From  Txn_CxTF_GSL TCG"
           tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
           tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
           tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
           tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A')  "
           If mNewTransDate <> "" Then
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And  TC.CxTFDate >= '" & CDate(mNewTransDate) & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
           Else
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
           End If
           tmp = tmp & " Union All"
           tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
           tmp = tmp & " From  Txn_CxTF_GSL TCG"
           tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
           tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
           tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
           tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A')  "
           If mNewTransDate <> "" Then
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.CxTFDate >= '" & CDate(mNewTransDate) & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
           Else
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
           End If
           tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"
    
           Call Tmp1Table
            
           For j = 1 To TmpRs1.RecordCount
              
               mDepoBillingCycleID = ""
               mDepoBillingUnit = ""
               mNoofDays = ""
               mDepConDID = ""
                 
               mDepoBillingCycleID = IIf(IsNull(TmpRs1!WBillingCycleID), "", TmpRs1!WBillingCycleID)
               If TmpRs!BagSizeFlag = "Y" Then
                  mDepoBillingUnit = "B"
               Else
                  mDepoBillingUnit = "Q"
               End If
               mDepoBillingRate = IIf(IsNull(TmpRs1!WithdrawalFinalPrice), 0, TmpRs1!WithdrawalFinalPrice)
               mDepConDID = IIf(IsNull(TmpRs1!DepConDID), "", TmpRs1!DepConDID)
               If mDepoBillingCycleID = 1 Then
                  mNoofDays = Val(Day(TmpRs1!CxTFDate))
                ElseIf mDepoBillingCycleID = 2 Then
                  If Day(TmpRs1!CxTFDate) < 8 Then
                     mNoofDays = "1"
                  ElseIf Day(TmpRs1!CxTFDate) < 15 Then
                     mNoofDays = "2"
                  ElseIf Day(TmpRs1!CxTFDate) < 22 Then
                     mNoofDays = "3"
                  Else
                     mNoofDays = "4"
                  End If
                ElseIf mDepoBillingCycleID = 3 Then
                  If Day(TmpRs1!CxTFDate) < 16 Then
                     mNoofDays = "1"
                  Else
                     mNoofDays = "2"
                  End If
                ElseIf mDepoBillingCycleID = 4 Then
                  mNoofDays = "1"
                End If
                 
                If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
                   tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs1!CxTF_DetailsID & "' And TxnType='W' And DepConDID is null And G_UID = '" & TmpRs1!G_UID & "' "
                   Call Tmp4Table
                End If
                  
                TmpRs1.MoveNext
           Next j
           
           
           If mNewTransDate <> "" Or IsNull(TmpRs!BGSLID_Mst_GSL_ID) = True Then
              'If (CDate(mLastInvDate) <= TmpRs!CDTFDate And CDate(mNewTransDate) - 1 >= TmpRs!CDTFDate) Or (mLastInvDate = "") Then
                 tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,DepBillingCycleID,BagSizeFlag ,FinalPrice,DepositFinalPrice,G_UID,DepConDID"
                 tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
                 tmp = tmp & " From  Txn_CxTF_GSL TCG"
                 tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
                 tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
                 tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((TCG.DematWeight/TCG.NoOfBags)*1000) Between  MinBagSize And MaxBagSize"
                 tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'R'  And TDC.ClosedOn is null"
                 If IsNull(TmpRs!CxTF_DetailsID) = True Then
                    tmp = tmp & " Where TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(TC.CxTF_DetailsID,0) = 0 And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
                 Else
                    tmp = tmp & " Where TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TC.CxTF_DetailsID = " & TmpRs!CxTF_DetailsID & " And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
                 End If
                 If mLastInvDate <> "" Then
                    tmp = tmp & " And ('" & CDate(mLastInvDate) & "' <= '" & TmpRs!CDTFDate & "' And '" & CDate(mNewTransDate) - 1 & "' >= '" & TmpRs!CDTFDate & "')"
                 End If
                 tmp = tmp & " Union All"
                 tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
                 tmp = tmp & " From  Txn_CxTF_GSL TCG"
                 tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
                 tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
                 tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.GodownID = TDCD.GodownId"
                 tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And  TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'R'  And TDC.ClosedOn is null"
'                 tmp = tmp & " Where TCG.GSLID = " & TmpRs!GSLID & " And TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TC.CxTF_DetailsID = " & TmpRs!CxTF_DetailsID & " And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is Not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N'"
                 If IsNull(TmpRs!CxTF_DetailsID) = True Then
                    tmp = tmp & " Where TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And isnull(TC.CxTF_DetailsID,0) = 0 And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
                 Else
                    tmp = tmp & " Where TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TC.CxTF_DetailsID = " & TmpRs!CxTF_DetailsID & " And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
                 End If
                 If mLastInvDate <> "" Then
                    tmp = tmp & " And ('" & CDate(mLastInvDate) & "' <= '" & TmpRs!CDTFDate & "' And '" & CDate(mNewTransDate) - 1 & "' >= '" & TmpRs!CDTFDate & "')"
                 End If
                 tmp = tmp & ") A"
                 tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"
                 
                 Call Tmp1Table
                
                
                 For j = 1 To TmpRs1.RecordCount
                     mDepoBillingCycleID = ""
                     mDepoBillingUnit = ""
                     mNoofDays = ""
                     mDepConDID = ""
                     
                     
                    
                     mDepoBillingCycleID = IIf(IsNull(TmpRs1!DepBillingCycleID), "", TmpRs1!DepBillingCycleID)
                     If TmpRs!BagSizeFlag = "Y" Then
                        mDepoBillingUnit = "B"
                     Else
                        mDepoBillingUnit = "Q"
                     End If
                     mDepoBillingRate = IIf(IsNull(TmpRs1!DepositFinalPrice), 0, TmpRs1!DepositFinalPrice)
                     mDepConDID = IIf(IsNull(TmpRs1!DepConDID), "", TmpRs1!DepConDID)
                    
                     If mDepoBillingCycleID = 1 Then
                        mNoofDays = a
                        mFromDate = Year(TmpRs1!CxTFDate) & "-" & Month(TmpRs1!CxTFDate) & "-01"
                        mSMonth = Month(TmpRs1!CxTFDate)
                        mSYear = Year(TmpRs1!CxTFDate)
                        mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
                        mNoofDays = (Val(mTotalDays) - Val(Day(TmpRs1!CxTFDate))) + 1
                     ElseIf mDepoBillingCycleID = 2 Then
                        If Day(TmpRs1!CxTFDate) < 8 Then
                           mNoofDays = "4"
                        ElseIf Day(TmpRs1!CxTFDate) < 15 Then
                           mNoofDays = "3"
                        ElseIf Day(TmpRs1!CxTFDate) < 22 Then
                           mNoofDays = "2"
                        Else
                           mNoofDays = "1"
                        End If
                     ElseIf mDepoBillingCycleID = 3 Then
                        If Day(TmpRs1!CxTFDate) < 16 Then
                           mNoofDays = "2"
                        Else
                           mNoofDays = "1"
                        End If
                     ElseIf mDepoBillingCycleID = 4 Then
                        mNoofDays = "1"
                     End If
                    
                     If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
                        tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs1!CxTF_DetailsID & "' And TxnType='D' And DepConDID is not null And G_UID = '" & TmpRs1!G_UID & "' "
                        Call Tmp4Table
                    End If
                    TmpRs1.MoveNext
                Next j
              'End If
           Else
              If IsNull(TmpRs!BGSLID_Mst_GSL_ID) = False Then
                 Call UpdateBGSLDCDetails(TmpRs!BGSLID_Mst_GSL_ID, TmpRs!BSM_Prefix)
              End If
           End If
        End If
        
        tmp = " Update Txn_DepositContract Set ClosedRemark = 'Closed After process executed on " & Now & "',ClosedOn = '" & Date & "',ClosedBy = '" & Gen_UserLoginID & "'  Where  ClosedOn is null And ContractType = 'R' And DepConID = " & TmpRs!DepConID & " "
        Call Tmp5Table
        
        TmpRs.MoveNext
        

        
    Next I
    TmpRs2.MoveNext
Next x


ProgressBar1.Value = ProgressBar1.Max
MsgBox "Done " & Now & ""

End Sub
Private Sub UpdateBGSLDCDetails(mBGSLID_ As Variant, mBSM_Prefix_ As Variant)

    tmp = " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID, TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R'  And TDC.ClosedOn is null "
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.Flag = 'Z' And GSL.Mst_GSL_ID = " & mBGSLID_ & " And GSL.SM_Prefix = '" & mBSM_Prefix_ & "' And GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " Union All"
    tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R'  And TDC.ClosedOn is null "
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.Flag = 'Z' And GSL.Mst_GSL_ID = " & mBGSLID_ & " And GSL.SM_Prefix = '" & mBSM_Prefix_ & "' And  GSL.DepositeBags > 0 And GSL.BalanceBags = 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N'"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " Union All"
    tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((GSL.DepositeWeight/GSL.DepositeBags)*1000) Between  MinBagSize And MaxBagSize"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A') And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R'  And TDC.ClosedOn is null"
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.Flag = 'Z' And GSL.Mst_GSL_ID = " & mBGSLID_ & " And GSL.SM_Prefix = '" & mBSM_Prefix_ & "' And  GSL.DepositeBags > 0 And GSL.BalanceBags = 0 And  GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1  And GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
    tmp = tmp & " And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And a.TxnType = 'W' And b.CxTFDate >= '2017-12-01')"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " Union All"
    tmp = tmp & " Select GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,(GSL.DepositeWeight/GSL.DepositeBags)*1000 as 'BagSize',GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID,GSL.DepConDID 'GSLDCDID',GSL.Rate,Max(GINV.YearMonth) 'YearMonth',GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    tmp = tmp & " From  Mst_GSL GSL"
    tmp = tmp & " Left Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId"
    tmp = tmp & " Left Join Txn_DepositContract TDC On  TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And GSL.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And GSL.CDTFDate Between StartDate and EndDate  And ContractType = 'R' And TDC.ClosedOn is null"
    tmp = tmp & " Left Join Txn_InvMonGSLDetail GINV on GSL.Mst_GSL_ID = GINV.Mst_GSL_ID And GSL.SM_Prefix = GINV.SM_Prefix"
    tmp = tmp & " Where GSL.Flag = 'Z' And GSL.Mst_GSL_ID = " & mBGSLID_ & " And GSL.SM_Prefix = '" & mBSM_Prefix_ & "' And GSL.DepositeBags > 0 And GSL.BalanceBags = 0  And GSL.CDTFDate >= '2010-01-01'   And GSL.ExchangeTypeID=1 And  GSL.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 And TDC.BagSizeFlag = 'N'"
    tmp = tmp & " And Convert(Varchar,GSL.Mst_GSL_ID)+'-'+GSL.SM_Prefix in (Select Convert(Varchar,a.Mst_GSL_ID)+'-'+a.SM_Prefix From Txn_CxTF_GSL a Inner Join Txn_CxTF_Details b On a.CxTF_DetailsID = b.CxTF_DetailsID and a.TxnType = b.TxnType and a.SM_Prefix = b.SM_Prefix  Where a.SM_Prefix = '" & TmpRs2!SM_Prefix & "' And a.TxnType = 'W' And b.CxTFDate >= '2017-12-01')"
    tmp = tmp & " Group by GSL.SM_Prefix ,GSL.ClientID,GSL.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,GSL.BalanceBags,GSL.BalanceWeight,GSL.DepositeWeight,GSL.DepositeBags,GSL.ContractID,GSL.NContractID,TDC.DepBillingCycleID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.DepositFinalPrice,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,GSL.G_UID,TDCD.DepConDID,GSL.CDTFDate,TDC.DepConID ,GSL.DepConDID,GSL.Rate,GSL.CxTF_DetailsID,GSL.BGSLID_Mst_GSL_ID,GSL.BSM_Prefix"
    
    
    Call Tmp5Table
    
    
    For ab = 1 To TmpRs5.RecordCount
        mDepoBillingCycleID = ""
        mDepoBillingUnit = ""
        mNoofDays = ""
        mDepConDID = ""
        
        mDepoBillingCycleID = IIf(IsNull(TmpRs5!BillingCycleID), "", TmpRs5!BillingCycleID)
        If TmpRs!BagSizeFlag = "Y" Then
           mDepoBillingUnit = "B"
        Else
           mDepoBillingUnit = "Q"
        End If
        mDepoBillingRate = IIf(IsNull(TmpRs5!FinalPrice), 0, TmpRs5!FinalPrice)
        mDepConDID = IIf(IsNull(TmpRs5!DepConDID), "", TmpRs5!DepConDID)
        If mDepoBillingCycleID = 1 Then
           mNoofDays = "0"
        ElseIf mDepoBillingCycleID = 2 Then
           mNoofDays = "4"
        ElseIf mDepoBillingCycleID = 3 Then
           mNoofDays = "2"
        ElseIf mDepoBillingCycleID = 4 Then
           mNoofDays = "1"
        End If
        mLastInvDate = ""
        mNewTransDate = ""
        
        If IsNull(TmpRs5!YearMonth) = False Then
           mLastInvDate = TmpRs5!YearMonth
           mNewTransDate = DateAdd("M", 1, CDate(mLastInvDate))
        End If
        
        If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
           tmp = "Update Mst_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & mBSM_Prefix_ & "' And Mst_GSL_ID = '" & mBGSLID_ & "' And DepConDID is not null And G_UID = '" & TmpRs5!G_UID & "' "
           Call Tmp1Table
           
           tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,WBillingCycleID,BagSizeFlag ,FinalPrice,WithdrawalFinalPrice,G_UID,DepConDID"
           tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
           tmp = tmp & " From  Txn_CxTF_GSL TCG"
           tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
           tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
           tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
           tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A')  "
           If mNewTransDate <> "" Then
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And  TC.CxTFDate >= '" & CDate(mNewTransDate) & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
           Else
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
           End If
           tmp = tmp & " Union All"
           tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.WBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.WithdrawalFinalPrice,TCG.G_UID,TDCD.DepConDID"
           tmp = tmp & " From  Txn_CxTF_GSL TCG"
           tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
           tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
           tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.DepConDID = TDCD.DepConDID "
           tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID  And TDC.ApprovalFlag in ('S','A')  "
           If mNewTransDate <> "" Then
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.CxTFDate >= '" & CDate(mNewTransDate) & "'   And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
           Else
              tmp = tmp & " Where TC.SM_Prefix = '" & TmpRs!SM_Prefix & "' And TCG.Mst_GSL_ID = " & TmpRs!Mst_GSL_ID & " And TC.ExchangeTypeID=1 And  TC.TxnType='W'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0'  and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N') A"
           End If
           tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"

           Call Tmp1Table
        
           For j = 1 To TmpRs1.RecordCount
          
               mDepoBillingCycleID = ""
               mDepoBillingUnit = ""
               mNoofDays = ""
               mDepConDID = ""
              
               mDepoBillingCycleID = IIf(IsNull(TmpRs1!WBillingCycleID), "", TmpRs1!WBillingCycleID)
               If TmpRs!BagSizeFlag = "Y" Then
                  mDepoBillingUnit = "B"
               Else
                  mDepoBillingUnit = "Q"
               End If
               mDepoBillingRate = IIf(IsNull(TmpRs1!WithdrawalFinalPrice), 0, TmpRs1!WithdrawalFinalPrice)
               mDepConDID = IIf(IsNull(TmpRs1!DepConDID), "", TmpRs1!DepConDID)
               If mDepoBillingCycleID = 1 Then
                  mNoofDays = Val(Day(TmpRs1!CxTFDate))
               ElseIf mDepoBillingCycleID = 2 Then
                  If Day(TmpRs1!CxTFDate) < 8 Then
                     mNoofDays = "1"
                  ElseIf Day(TmpRs1!CxTFDate) < 15 Then
                     mNoofDays = "2"
                  ElseIf Day(TmpRs1!CxTFDate) < 22 Then
                     mNoofDays = "3"
                  Else
                     mNoofDays = "4"
                  End If
               ElseIf mDepoBillingCycleID = 3 Then
                  If Day(TmpRs1!CxTFDate) < 16 Then
                      mNoofDays = "1"
                  Else
                      mNoofDays = "2"
                  End If
               ElseIf mDepoBillingCycleID = 4 Then
                  mNoofDays = "1"
               End If
              
               If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
                  tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs1!CxTF_DetailsID & "' And TxnType='W' And DepConDID is null And G_UID = '" & TmpRs1!G_UID & "' "
                  Call Tmp4Table
               End If
              
               TmpRs1.MoveNext
            Next j
            If mNewTransDate <> "" Then
               'If CDate(mLastInvDate) <= TmpRs5!CDTFDate And CDate(mNewTransDate) - 1 >= TmpRs5!CDTFDate Then
                  tmp = "Select CxTF_DetailsID,SM_Prefix,CxTFDate,CDTFDate,NoOfDay,ClientID,Mst_GSL_ID,GodownID,CommodityID,SpillageFlag,NoOfBags,DematWeight,BagSize,NContractID,BillingCycleID,DepBillingCycleID,BagSizeFlag ,FinalPrice,DepositFinalPrice,G_UID,DepConDID"
                  tmp = tmp & " From (Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag ,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
                  tmp = tmp & " From  Txn_CxTF_GSL TCG"
                  tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
                  tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
                  tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On  GSL.GodownID = TDCD.GodownId And ((TCG.DematWeight/TCG.NoOfBags)*1000) Between  MinBagSize And MaxBagSize"
                  tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'R'  And TDC.ClosedOn is null "
                  tmp = tmp & " Where TCG.Mst_GSL_ID = " & mBGSLID_ & " And TC.SM_Prefix = '" & mBSM_Prefix & "' And TC.CxTF_DetailsID = " & TmpRs5!CxTF_DetailsID & " And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'Y'"
                  If mLastInvDate <> "" Then
                    tmp = tmp & " And ('" & CDate(mLastInvDate) & "' <= '" & TmpRs5!CDTFDate & "' And '" & CDate(mNewTransDate) - 1 & "' >= '" & TmpRs5!CDTFDate & "')"
                  End If
                  tmp = tmp & " Union All"
                  tmp = tmp & " Select TC.CxTF_DetailsID,TC.SM_Prefix ,TC.CxTFDate,GSL.CDTFDate,day(TC.CxTFDate) NoOfDay,GSL.ClientID,TCG.Mst_GSL_ID,GSL.GodownID,GSL.CommodityID ,GSL.SpillageFlag,TCG.NoOfBags,TCG.DematWeight,(TCG.DematWeight/TCG.NoOfBags)*1000 as 'BagSize',GSL.NContractID,TDC.BillingCycleID,TDC.DepBillingCycleID,TDC.BagSizeFlag,TDCD.FinalPrice,TDCD.DepositFinalPrice,TCG.G_UID,TDCD.DepConDID"
                  tmp = tmp & " From  Txn_CxTF_GSL TCG"
                  tmp = tmp & " Inner Join Txn_CxTF_Details TC on TCG.SM_Prefix=TC.SM_Prefix And TCG.CxTF_DetailsID=TC.CxTF_DetailsID And TCG.TxnType=TC.TxnType"
                  tmp = tmp & " left Join Mst_GSL GSL on TCG.Mst_GSL_ID=GSL.Mst_GSL_ID And TCG.SM_Prefix=GSL.SM_Prefix"
                  tmp = tmp & " Inner Join Txn_DepositContractGodownDetail TDCD On GSL.GodownID = TDCD.GodownId"
                  tmp = tmp & " Inner Join Txn_DepositContract TDC On TDCD.DepConID = TDC.DepConID And  TDC.SM_Prefix = GSL.SM_Prefix And GSL.ClientID = TDC.ClientID And TC.CommodityID = TDC.CommodityID And TDC.ApprovalFlag in ('S','A')  And TC.CxTFDate Between StartDate and EndDate And ContractType = 'R'  And TDC.ClosedOn is null "
                  tmp = tmp & " Where TCG.Mst_GSL_ID = " & mBGSLID_ & " And TC.SM_Prefix = '" & mBSM_Prefix & "' And TC.CxTF_DetailsID = " & TmpRs5!CxTF_DetailsID & " And TC.ExchangeTypeID=1 And  TC.TxnType='D'  And TCG.DepConDID is not null And GSL.SpillageFlag = '0' And GSL.MadeUpFlag = '0' and isnull(GSL.NContractID,0)=0 and isnull(GSL.ContractID,0)=0 And TDC.BagSizeFlag = 'N'"
                  If mLastInvDate <> "" Then
                     tmp = tmp & " And ('" & CDate(mLastInvDate) & "' <= '" & TmpRs5!CDTFDate & "' And '" & CDate(mNewTransDate) - 1 & "' >= '" & TmpRs5!CDTFDate & "')"
                  End If
                  tmp = tmp & ") A"
                  tmp = tmp & " Order By SM_Prefix,CxTF_DetailsID,Mst_GSL_ID"
                  
                  Call Tmp1Table
                
                
                  For j = 1 To TmpRs1.RecordCount
                      mDepoBillingCycleID = ""
                      mDepoBillingUnit = ""
                      mNoofDays = ""
                      mDepConDID = ""
                      
                      
                     
                      mDepoBillingCycleID = IIf(IsNull(TmpRs1!DepBillingCycleID), "", TmpRs1!DepBillingCycleID)
                      If TmpRs!BagSizeFlag = "Y" Then
                         mDepoBillingUnit = "B"
                      Else
                         mDepoBillingUnit = "Q"
                      End If
                      mDepoBillingRate = IIf(IsNull(TmpRs1!DepositFinalPrice), 0, TmpRs1!DepositFinalPrice)
                      mDepConDID = IIf(IsNull(TmpRs1!DepConDID), "", TmpRs1!DepConDID)
                     
                      If mDepoBillingCycleID = 1 Then
                         mNoofDays = a
                         mFromDate = Year(TmpRs1!CxTFDate) & "-" & Month(TmpRs1!CxTFDate) & "-01"
                         mSMonth = Month(TmpRs1!CxTFDate)
                         mSYear = Year(TmpRs1!CxTFDate)
                         mTotalDays = Day(DateSerial(mSYear, mSMonth + 1, 0))
                         mNoofDays = (Val(mTotalDays) - Val(Day(TmpRs1!CxTFDate))) + 1
                      ElseIf mDepoBillingCycleID = 2 Then
                         If Day(TmpRs1!CxTFDate) < 8 Then
                            mNoofDays = "4"
                         ElseIf Day(TmpRs1!CxTFDate) < 15 Then
                            mNoofDays = "3"
                         ElseIf Day(TmpRs1!CxTFDate) < 22 Then
                            mNoofDays = "2"
                         Else
                             mNoofDays = "1"
                         End If
                      ElseIf mDepoBillingCycleID = 3 Then
                         If Day(TmpRs1!CxTFDate) < 16 Then
                            mNoofDays = "2"
                         Else
                            mNoofDays = "1"
                         End If
                      ElseIf mDepoBillingCycleID = 4 Then
                         mNoofDays = "1"
                      End If
                    
                      If mDepoBillingCycleID <> "" And mDepoBillingUnit <> "" And Val(mDepoBillingRate) <> 0 And mNoofDays <> "" Then
                         tmp = "Update Txn_CxTF_GSL Set RateCardType = 'D',BillingCycleID = " & mDepoBillingCycleID & ",BillingUnit = '" & mDepoBillingUnit & "',Rate = " & mDepoBillingRate & ",RateForDailyPnL = " & mDepoBillingRate & ",DepConDID =  " & mDepConDID & ",NoofDays = " & mNoofDays & " Where SM_Prefix = '" & TmpRs1!SM_Prefix & "' And CxTF_DetailsID = '" & TmpRs1!CxTF_DetailsID & "' And TxnType='D' And DepConDID is not null And G_UID = '" & TmpRs1!G_UID & "' "
                         Call Tmp4Table
                      End If
                      TmpRs1.MoveNext
                  Next j
               'End If

           Else
              Call UpdateBGSLDCDetails(TmpRs5!BGSLID_Mst_GSL_ID, TmpRs5!SM_Prefix)
           End If
        End If
        TmpRs5.MoveNext
    Next ab


End Sub
Private Sub CmdUpdateProcess_Click()
Dim mLocation As Variant
ProgressBar1.Min = 0
tmp = "Select Mst_GSL_ID, SM_prefix, TransactionType, TransactionID, TransactionDate From Mst_GSL order by SM_prefix,Mst_GSL_ID "
Call Tmp1Table

tmp = "Select TCD.CxTF_DetailsID, TCD.TxnType, TCD.SM_Prefix, TCD.CxTFDate, TCG.Mst_GSL_ID from Txn_CXTF_Details  as TCD"
tmp = tmp & " Inner Join Txn_CXTF_GSL as TCG on TCG.CxTF_DetailsID = TCD.CxTF_DetailsID and TCG.TxnType = TCD.TxnType and TCG.SM_prefix = TCD.SM_prefix"
tmp = tmp & " Where TCD.TxnType='W'"
tmp = tmp & " Order by TCD.SM_Prefix, TCD.CxTFDate"
Call Tmp2Table

mLocation = ""
If Not TmpRs2.EOF Then
   TmpRs2.MoveFirst
   ProgressBar1.Max = TmpRs2.RecordCount
   For I = 1 To TmpRs2.RecordCount
       ProgressBar1.Value = I
       If mLocation <> TmpRs2!SM_Prefix Then
          TmpRs1.Filter = "SM_prefix = '" & TmpRs2!SM_Prefix & "'"
          mLocation = TmpRs2!SM_Prefix
         ' MsgBox mLocation
       End If
       Call UpdateGSLData("CITF")
       TmpRs2.MoveNext
   Next
End If

tmp = "Select Txn_AT.TxnID CxTF_DetailsID, Txn_AT.TxnDate as CxTFDate, Txn_AT.SM_Prefix, Txn_AT.TxnType, Txn_GSLT.Mst_GSL_ID  from Txn_AdjTransaction  as Txn_AT"
tmp = tmp & " Inner Join Txn_GSLADJTransaction as Txn_GSLT on Txn_GSLT.TxnID = Txn_AT.TxnID"
tmp = tmp & " Where Txn_AT.TxnType <> 'AR'"
tmp = tmp & " order by Txn_AT.SM_Prefix, Txn_AT.TxnDate"
Call Tmp2Table
mLocation = ""
If Not TmpRs2.EOF Then
   TmpRs2.MoveFirst
   ProgressBar1.Max = TmpRs2.RecordCount
   For I = 1 To TmpRs2.RecordCount
       ProgressBar1.Value = I
       If mLocation <> TmpRs2!SM_Prefix Then
          TmpRs1.Filter = "SM_prefix = '" & TmpRs2!SM_Prefix & "'"
          mLocation = TmpRs2!SM_Prefix
          'MsgBox mLocation
       End If
       Call UpdateGSLData("Adj")
       TmpRs2.MoveNext
   Next
End If
tmp = "Select MG.Mst_GSL_ID, MG.SM_Prefix, MG.TransactionType, MG.TransactionID , MG.TransactionDate, MG.BalanceWeight"
tmp = tmp & " From Mst_GSL as MG Where MG.Flag = 'E'"
tmp = tmp & " And MG.TransactionType = 'CITF' "
tmp = tmp & " Order by MG.TransactionType, MG.SM_Prefix, MG.Mst_GSL_ID"
Call Tmp1Table

tmp = "Select Txn_G.CxTF_DetailsID, Txn_G.TxnType, Txn_G.SM_Prefix, Txn_G.Mst_GSL_ID, Txn_G.WeightGainLoss, Txn_G.WeightGainLossGSL From Txn_CXTF_GSL as Txn_G"
tmp = tmp & " Where Txn_G.TxnType = 'W'"
tmp = tmp & " Order by Txn_G.SM_Prefix, Txn_G.Mst_GSL_ID, Txn_G.CxTF_DetailsID"
Call Tmp2Table

mLocation = ""
If Not TmpRs1.EOF Then
   ProgressBar1.Max = TmpRs1.RecordCount
   TmpRs1.MoveFirst
   For I = 1 To TmpRs1.RecordCount
       ProgressBar1.Value = I
       If mLocation <> TmpRs1!SM_Prefix Then
          'TmpRs2.Filter = ""
          'TmpRs2.MoveFirst
          TmpRs2.Filter = "SM_prefix = '" & TmpRs1!SM_Prefix & "' "
          mLocation = TmpRs1!SM_Prefix
          'MsgBox mLocation
       End If
       If TmpRs2.RecordCount > 0 Then
          Call UpdateGainLossData
       End If
       TmpRs1.MoveNext
   Next I
End If
tmp = "Select MG.Mst_GSL_ID, MG.SM_Prefix, MG.TransactionType, MG.TransactionID , MG.TransactionDate, MG.BalanceWeight"
tmp = tmp & " From Mst_GSL as MG Where MG.Flag = 'E'"
tmp = tmp & " And MG.TransactionType = 'Adj' "
tmp = tmp & " Order by MG.TransactionType, MG.SM_Prefix, MG.Mst_GSL_ID"
Call Tmp1Table

tmp = "Select Txn_AT.TxnID CxTF_DetailsID, Txn_AT.TxnDate as CxTFDate, Txn_AT.SM_Prefix, Txn_AT.TxnType, Txn_GSLT.Mst_GSL_ID, Txn_GSLT.WeightGainLossGSL"
tmp = tmp & " From Txn_AdjTransaction  as Txn_AT"
tmp = tmp & " Inner Join Txn_GSLADJTransaction as Txn_GSLT on Txn_GSLT.TxnID = Txn_AT.TxnID"
tmp = tmp & " Where Txn_AT.TxnType <> 'AR'"
tmp = tmp & " order by Txn_AT.SM_Prefix, Txn_AT.TxnID, Txn_GSLT.Mst_GSL_ID"
Call Tmp2Table
If Not TmpRs1.EOF Then
   ProgressBar1.Max = TmpRs1.RecordCount
   TmpRs1.MoveFirst
   For I = 1 To TmpRs1.RecordCount
       ProgressBar1.Value = I
       If mLocation <> TmpRs1!SM_Prefix Then
          'TmpRs2.Filter = ""
          'TmpRs2.MoveFirst
          TmpRs2.Filter = "SM_prefix = '" & TmpRs1!SM_Prefix & "' "
          mLocation = TmpRs1!SM_Prefix
          'MsgBox mLocation
       End If
       If TmpRs2.RecordCount > 0 Then
          Call UpdateGainLossData
       End If
       TmpRs1.MoveNext
   Next I
End If
MsgBox "Done!!!"
End Sub

Private Sub UpdateGSLData(mType As Variant)
TmpRs1.MoveFirst
TmpRs1.Find "Mst_GSL_ID = " & TmpRs2!Mst_GSL_ID
If Not TmpRs1.EOF Then
   If Format(TmpRs1!TransactionDate, "yyyy-MM-dd") < Format(TmpRs2!CxTFDate, "yyyy-MM-dd") Then
      TmpRs1!TransactionType = mType
      TmpRs1!TransactionID = TmpRs2!CxTF_DetailsID
      TmpRs1!TransactionDate = Format(TmpRs2!CxTFDate, "yyyy-MM-dd")
      TmpRs1.Update
   End If
End If
End Sub

Private Sub UpdateGainLossData()
TmpRs2.MoveFirst
TmpRs2.Find "Mst_GSL_ID = " & TmpRs1!Mst_GSL_ID
If Not TmpRs2.EOF Then
   For k = 1 To TmpRs2.RecordCount
       If TmpRs1!TransactionID = TmpRs2!CxTF_DetailsID Then
          TmpRs2!WeightGainLossGSL = TmpRs1!BalanceWeight
          TmpRs2.Update
          Exit Sub
       End If
       TmpRs2.MoveNext
       If TmpRs2.EOF Then Exit For
    Next
End If
End Sub
Private Sub Form_Load()

Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   'Call GButtonLockAD(Me)
   MsgBox "Access Denied !!!!!!!!!"
   Frame1.Enabled = False
   Frame2.Enabled = False
   Exit Sub
End If

Frame1.Enabled = True
Frame2.Enabled = True
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Call Check1_Click
Call Check2_Click
Call Check3_Click

Call TableTxn_CxTF_Details

Call Grid_Head


CmdUpdateProcess.Visible = False
CmdPendingGSL.Visible = False
If LCase(Gen_UserRole) = "power" Then
   CmdUpdateProcess.Visible = True
   CmdUpdateBillingRate.Visible = True
End If




'Commented by Parag on 10 Sept 2019
'tmp = "select * from mst_gsl where BalanceBags <= 0 and flag in  ('O','J')"
'Call TmpTable
'
'
'tmp = "Select a.sm_prefix,a.WHRNo,a.pledgeid , a.WHRNoOfBags , a.DeliveredNoOfBags from Txn_WHR_Detail a "
'tmp = tmp & " inner join Txn_PledgeRelease b on b.pledgeid = a.pledgeid and b.sm_prefix = a.sm_prefix"
'tmp = tmp & " where  (a.WHRNoOfBags -a.DeliveredNoOfBags ) > 0 "
'
'Call Tmp1Table
'
'If TmpRs.RecordCount > 0 Or TmpRs1.RecordCount > 0 Then
'   Cmd_GSLFlg.Enabled = True
'   Exit Sub
'Else
'   Cmd_GSLFlg.Enabled = False
'End If


'tmp = "Select * from mst_gsl where BalanceBags > 0 and flag = 'E'"
'Call Tmp1Table
'
'If TmpRs1.RecordCount > 0 Then
'   Cmd_GSLFlg.Enabled = True
'Else
'   Cmd_GSLFlg.Enabled = False
'End If

End Sub

Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationID <> 0 ")
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
'----
Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
   mLocationID = Null
   TxtPreFix = ""
   
   Exit Sub
End If
RsMst_Location.MoveFirst
RsMst_Location.Find "LocationName = '" & CmbLocationID.Text & "'"
If RsMst_Location.EOF Then
   MsgBox "Invalid  selection "
   CmbLocationID.SetFocus
   Exit Sub
End If
If mLocationID <> RsMst_Location!LocationID Then
   CmbCMPID.Text = ""
   TxtPreFix = ""
End If

mLocationID = RsMst_Location!LocationID
TxtPreFix = RsMst_Location!SM_Prefix
End Sub

Private Sub CmbCMPID_GotFocus()
    tmp = CmbCMPID.Text
    Call TableMst_CMP(" Where CloseDate is null and LocationID = " & mLocationID & " ")
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
   mCMPID = Null
  
   TxtCMPAddress = ""
   TxtCommodity = ""
   TxtEHoPCMPID = ""
   TxStartDate = ""
   TxtCurrentDate = ""
   CmbModuleName.Text = ""
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
   CmbAgent.Text = ""
   TxtAgent = ""
   Check1.Value = 0
   Check2.Value = 0
   Check3.Value = 0
   mAgentID = Null
   mPrevAgentID = Null
   Exit Sub
End If

RsMst_CMP.MoveFirst
RsMst_CMP.Find "CMPName = '" & CmbCMPID.Text & "'"
If RsMst_CMP.EOF Then
   MsgBox "Invalid  selection "
   CmbCMPID.SetFocus
   Exit Sub
End If
If mCMPID <> RsMst_CMP!CMPID Then
   TxtCommodity = ""
   TxtStartDate = ""
   TxtCurrentDate = ""
   CmbModuleName.Text = ""
   TxtEHoPCMPID = ""
   TxtModuleNo = ""
   TxtCommodity = ""
   TxtExchangeType = ""
End If


mCMPID = RsMst_CMP!CMPID
TxtCMPAddress = IIf(IsNull(RsMst_CMP!Address), "", RsMst_CMP!Address)
TxtStartDate = RsMst_CMP!StartDate 'Format(RsMst_CMP!StartDate, "dd/mm/yyyy")
TxtCurrentDate = RsMst_CMP!CurrentDate + 1 ' CDate(Format(RsMst_CMP!CurrentDate, "dd/mm/yyyy")) + 1
TxtEHoPCMPID = RsMst_CMP!EHOPCMPID


'TxtCxTFDate.Value = IIf(IsNull(RsMst_CMP!CurrentDate), RsMst_CMP!StartDate, RsMst_CMP!CurrentDate) + 1
'TxtCxTFDispalyDate = TxtCxTFDate.Value
'TxtPreFix = GetPreFix(RsMst_CMP!CMPID, "P")
End Sub
Private Sub TxtModuleNo_LostFocus()
Dim ans As Variant

If TxtModuleNo.Text = "" Then Exit Sub
If CmbLocationID.Text = "" Then Exit Sub
If CmbCMPID.Text = "" Then Exit Sub
If CmbModuleName.Text = "" Then Exit Sub


If ChkInvoiceFlag = False Then
   MsgBox "Bill Flag Marked . Can't update this transaction!!!"
   TxtModuleNo.SetFocus
   Exit Sub
End If

RsTxn_CxTF_Details.MoveFirst
tmp = "SM_Prefix = '" & TxtPreFix & "' And CMPID = " & mCMPID & " And "
If CmbModuleName.Text = "CDTF" Then
   mCurrentDate = Format(TxtCurrentDate, Gen_DatFormat)
   tmp = tmp & "TxnType = 'D'  And CxTF_DetailsID = " & Val(TxtModuleNo) & "" 'And CxTFDate = '" & mCurrentDate & "'
   RsTxn_CxTF_Details.Filter = tmp
   If RsTxn_CxTF_Details.RecordCount > 0 Then
      TxtModuleDate = RsTxn_CxTF_Details!CxTFDate
   End If
   
ElseIf CmbModuleName.Text = "CITF" Then
   mCurrentDate = Format(TxtCurrentDate, Gen_DatFormat)
   
   tmp = tmp & "TxnType = 'W'  And CxTF_DetailsID = " & Val(TxtModuleNo) & "" 'And CxTFDate = '" & mCurrentDate & "'
   
   RsTxn_CxTF_Details.Filter = tmp
   If RsTxn_CxTF_Details.RecordCount > 0 Then
      TxtModuleDate = RsTxn_CxTF_Details!CxTFDate
   End If
 End If

If CmbModuleName.Text = "CDTF" Or CmbModuleName.Text = "CITF" Then
   If RsTxn_CxTF_Details.RecordCount > 0 Then
      mCommodityID = RsTxn_CxTF_Details!CommodityID
      TxtCommodity = GetCommodityName(RsTxn_CxTF_Details!CommodityID)
      mExchangeTypeID = RsTxn_CxTF_Details!ExchangeTypeID
      TxtExchangeType = GetExchangeName(RsTxn_CxTF_Details!ExchangeTypeID)
      TxtAgent.Text = GetAgentName(IIf(IsNull(RsTxn_CxTF_Details!AgentID), 0, RsTxn_CxTF_Details!AgentID))
      mPrevAgentID = RsTxn_CxTF_Details!AgentID
      TxtExchangeType = GetExchangeName(RsTxn_CxTF_Details!ExchangeTypeID)
      mModuleDate = RsTxn_CxTF_Details!CxTFDate
      TxtPrevManualNo = IIf(IsNull(RsTxn_CxTF_Details!ManualCDTFNo), "", RsTxn_CxTF_Details!ManualCDTFNo)
      TxtHologramNo = IIf(IsNull(RsTxn_CxTF_Details!HologramNo), "", RsTxn_CxTF_Details!HologramNo)
   Else
      MsgBox "No Record found!!!"
      TxtModuleNo.SetFocus
      TxtCommodity = ""
      TxtExchangeType = ""
      RsTxn_CxTF_Details.Filter = ""
      TxtHologramNo = ""
      CmbAgent.Text = ""
      TxtAgent = ""
      Check1.Value = 0
      Check2.Value = 0
      Check3.Value = 0
      mAgentID = Null
      mPrevAgentID = Null
      Exit Sub
   End If
End If

End Sub

Private Function ChkInvoiceFlag() As Boolean
Dim Retval As Boolean
Retval = True

tmp = "Select * from Txn_CxTF_GSL Where SM_Prefix = '" & TxtPreFix & "' And "
tmp = tmp & " TxnType = '" & mType & "' And CxTF_DetailsID = " & TxtModuleNo & "  And BillFlag is not null  "
Call TmpTable

If TmpRs.RecordCount > 0 Then
'   MsgBox "Invoice Generated for this Transaction!!!"
   Retval = False
End If

ChkInvoiceFlag = Retval

End Function


Private Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .AllowUserResizing = flexResizeColumns
    .WordWrap = True
    .Cols = 17 '16 '15
    .Rows = 2
    .TextMatrix(0, 0) = "StateName"
    .TextMatrix(0, 1) = "LocationName"
    .TextMatrix(0, 2) = "SM_Prefix"
    .TextMatrix(0, 3) = "CMPName"
    .TextMatrix(0, 4) = "Mst_GSL_ID"
    .TextMatrix(0, 5) = "GodownNo"
    .TextMatrix(0, 6) = "Commodity"
    .TextMatrix(0, 7) = "Client"
    .TextMatrix(0, 8) = "BalanceBags"
    .TextMatrix(0, 9) = "BalanceWeight"
    .TextMatrix(0, 10) = "BagSize"
    .TextMatrix(0, 11) = "ContractID"
    .TextMatrix(0, 12) = "NContractID"
    .TextMatrix(0, 13) = "G_UID"
    .TextMatrix(0, 14) = "CDTFDate"
    .TextMatrix(0, 15) = "CDTF No"
    .TextMatrix(0, 16) = "Made up Flag"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500

    .RowHeight(0) = 500
End With
End Sub

Private Sub CmdExcel_Click()
Gen_mTimeStamp = GetTimeStamp
Call Gen_Create_Xls(Gen_mTimeStamp & Gen_UserName & "Mst_PendingGSL.xls")
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
Set oWB = oXL.Workbooks.Open(Pat & "\excel\" & Gen_mTimeStamp & Gen_UserName & "Mst_PendingGSL.xls")
Set oWS = oWB.Worksheets("Sheet1")
oWS.ClearArrows
mRow = 1
For I = 0 To MSHFlexGrid1.Rows - 1
    mRow = I + 1 ' + 6
    For C = 0 To MSHFlexGrid1.Cols - 1
        If C = 5 Then
           oWS.Cells(mRow, C + 1) = "'" & MSHFlexGrid1.TextMatrix(I, C)
        Else
           If InStr(MSHFlexGrid1.TextMatrix(I, C), "/") > 0 Then
              If IsDate(MSHFlexGrid1.TextMatrix(I, C)) Then
                 oWS.Cells(mRow, C + 1) = Format(MSHFlexGrid1.TextMatrix(I, C), "MM/dd/yyyy")
              Else
                 oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
              End If
           Else
              oWS.Cells(mRow, C + 1) = MSHFlexGrid1.TextMatrix(I, C)
           End If
        End If
    Next
Next
oWB.Save
MsgBox ("Excel file created !!! ")
oXL.Visible = True
End Sub


