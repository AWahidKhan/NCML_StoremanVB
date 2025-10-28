VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Frm_ISINRevalidation 
   Caption         =   "ISIN Revalidation"
   ClientHeight    =   10380
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10380
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame ButtonFrm 
      Height          =   1200
      Left            =   120
      TabIndex        =   42
      Top             =   8760
      Width           =   14925
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11985
         TabIndex        =   51
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11985
         TabIndex        =   50
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9045
         TabIndex        =   49
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9045
         TabIndex        =   48
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6105
         TabIndex        =   47
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6105
         TabIndex        =   46
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3165
         TabIndex        =   45
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3165
         TabIndex        =   5
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   44
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   43
         Top             =   195
         Width           =   3060
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "New ISIN"
      Height          =   2655
      Left            =   7680
      TabIndex        =   26
      Top             =   840
      Width           =   7260
      Begin VB.TextBox TxtNewISINID 
         Appearance      =   0  'Flat
         Height          =   325
         Left            =   1800
         TabIndex        =   4
         Top             =   165
         Width           =   2640
      End
      Begin VB.TextBox TxtNewCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2000
      End
      Begin MSComCtl2.DTPicker NewNE_Date 
         Height          =   360
         Left            =   1800
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48955393
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker NewFED 
         Height          =   360
         Left            =   5490
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1305
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48955393
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   960
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   1800
         Width           =   1500
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Old ISIN"
      Height          =   2655
      Left            =   240
      TabIndex        =   17
      Top             =   840
      Width           =   7335
      Begin VB.TextBox TxtOldExchangeTypeID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   12
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   3
         Top             =   195
         Width           =   2640
      End
      Begin VB.TextBox TxtOldCommodity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   325
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1755
         Width           =   5325
      End
      Begin MSComCtl2.DTPicker OldNE_Date 
         Height          =   360
         Left            =   1890
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48955393
         CurrentDate     =   39864
      End
      Begin MSComCtl2.DTPicker OldFED 
         Height          =   360
         Left            =   5925
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1335
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   48955393
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      Format          =   48955393
      CurrentDate     =   39889
   End
   Begin VB.Frame Frame1 
      Height          =   8655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   14895
      Begin VB.TextBox TxtRevalidationID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   360
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
         Width           =   1080
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
         Height          =   5160
         Left            =   120
         TabIndex        =   16
         Top             =   3420
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   9102
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
         TabIndex        =   53
         Top             =   240
         Width           =   1620
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   240
         Width           =   2145
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   10005
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
Dim Edit_Flg As Variant

Private Sub DeleteCmd_Click()

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Are you Sure you want to Delete this Record", vbYesNo)
If ans = vbYes Then
Call TableMst_GSL
For i = 2 To MSHFlexGrid2.Rows - 1
   If MSHFlexGrid2.TextMatrix(i, 1) <> "" Then
        RsMst_GSL.Find ("GSLID = " & MSHFlexGrid2.TextMatrix(i, 6))
        If Not RsMst_GSL.EOF Then
          RsMst_GSL!ISINID = RsTxn_ISINRevalidation!OldISIN
          RsMst_GSL.Update
        End If
        
   End If
Next
DeleteGridData
RsTxn_ISINRevalidation.Delete
RsTxn_ISINRevalidation.Update

If RsTxn_ISINRevalidation.RecordCount = 0 Then
   Call AddCmd_Click
   Exit Sub
End If
RsTxn_ISINRevalidation.MoveNext
If RsTxn_ISINRevalidation.EOF() Then
   RsTxn_ISINRevalidation.MoveLast
End If
Call InData
End If
Exit Sub
msgError:
MsgBox "Child Record Present "
RsTxn_ISINRevalidation.CancelUpdate

End Sub
Private Sub DeleteGridData()
Call TableTxn_ISINGSLRevalidation("Where ISINRevalID = " & TxtRevalidationID)
If Not RsTxn_ISINRevalidation.RecordCount = 0 Then
        RsTxn_ISINGSLRevalidation.Delete
        RsTxn_ISINGSLRevalidation.Update
End If
 
End Sub


Private Sub EditCmd_Click()
'Edit_Flg = "E"
'TxtRevalidationReqNo.Locked = False
'TxtOldISINID.Locked = False
'TxtNewISINID.Locked = False
'RevalidationDate.Value = Date
'MSHFlexGrid2.Enabled = True
End Sub

Private Sub FirstCmd_Click()
RsTxn_ISINRevalidation.MoveFirst
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
mExchangeTypeID = GetExchangeID("NCDEX")
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
TxtOldISINID.Locked = False
TxtNewISINID.Locked = False
RevalidationDate.Value = Date
MSHFlexGrid2.Enabled = True

TxtRevalidationID.Text = ""
TxtRevalidationReqNo.Text = ""
    
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

mExchangeTypeID = 0
mCMPID = 0
mCommodityID = 0
Call Grid_Head
Call GButtonLock(Me, False)
    
End Sub

Private Sub LastCmd_Click()
RsTxn_ISINRevalidation.MoveLast
Call InData
'TmpRs.MoveLast
PreviousCmd.Enabled = True
FirstCmd.Enabled = True
NextCmd.Enabled = False
LastCmd.Enabled = False
End Sub

Private Sub NextCmd_Click()
Dim LF_Flag As Variant
LF_Flag = "N"
RsTxn_ISINRevalidation.MoveNext
If RsTxn_ISINRevalidation.EOF Then
   RsTxn_ISINRevalidation.MoveLast
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
RsTxn_ISINRevalidation.MovePrevious

If RsTxn_ISINRevalidation.BOF Then
   RsTxn_ISINRevalidation.MoveFirst
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

If ChkSelection() = False Then
    MsgBox "Please Select Atleast One Row From Grid"
    MSHFlexGrid2.SetFocus
    Exit Sub
End If

If Edit_Flg = "A" Then
   RsTxn_ISINRevalidation.AddNew
   RsTxn_ISINRevalidation!ISINRevalID = GetISINRevalID
   TxtRevalidationID.Text = RsTxn_ISINRevalidation!ISINRevalID
Else
   RsTxn_ISINRevalidation.MoveFirst
   RsTxn_ISINRevalidation.Find " ISINRevalID= " & TxtISINRevalReqNo.Text
End If
RsTxn_ISINRevalidation!ISINRevalID = TxtRevalidationID.Text
RsTxn_ISINRevalidation!ISINRevalReqNo = TxtRevalidationReqNo.Text
RsTxn_ISINRevalidation!ISINRevalDate = RevalidationDate.Value
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
    Call TableMst_GSL
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
              RsTxn_ISINGSLRevalidation!GSLID = MSHFlexGrid2.TextMatrix(i, 6)
              RsMst_GSL.Find ("GSLID = " & MSHFlexGrid2.TextMatrix(i, 6))
              If Not RsMst_GSL.EOF Then
                 RsMst_GSL!ISINID = TxtNewISINID.Text
                 RsMst_GSL!Flag = "R"
                 RsMst_GSL.Update
              End If
              RsTxn_ISINGSLRevalidation.Update
           End If
    Next
End If
RsTxn_ISINGSLRevalidation.Requery
Call InData
End Sub
Private Sub InData()
  
TxtRevalidationID = RsTxn_ISINRevalidation!ISINRevalID
TxtRevalidationReqNo = RsTxn_ISINRevalidation!ISINRevalReqNo
RevalidationDate = RsTxn_ISINRevalidation!ISINRevalDate
TxtOldISINID = RsTxn_ISINRevalidation!OldISIN
TxtNewISINID = RsTxn_ISINRevalidation!NewISIN
Call TxtOldISINID_LostFocus
Call TxtNewISINID_LostFocus
Call GridInData
TxtRevalidationReqNo.Locked = True
TxtOldISINID.Locked = True
TxtNewISINID.Locked = True
MSHFlexGrid2.Enabled = False
Call GButtonLock(Me, True)
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
        tmp = "SELECT Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, Mst_GSL.DepositeWeight,Mst_GSL.ISINID, "
        tmp = tmp & " Mst_GSL.GSLID FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
        tmp = tmp & " WHERE (Mst_GSL.GSLId=" & RsTxn_ISINGSLRevalidation!GSLID & ")"
        Call TmpTable
        MSHFlexGrid2.TextMatrix(k, 0) = "*"
        MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!GodownNo
        MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!StackNo
        MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!LotNo
        MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!DepositeBags
        MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!DepositeWeight
        MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!GSLID
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
Private Sub TxtNewISINID_LostFocus()
If TxtNewISINID.Text = "" Then
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

TxtNewCMP = GetCMPName(RsMst_ISIN!CMPID) '(IsNull(RsMst_ISIN!CMPName), "", RsMst_ISIN!CMPName)
NewNE_Date = IIf(IsNull(RsMst_ISIN!NE_Date), "", RsMst_ISIN!NE_Date)
NewFED = IIf(IsNull(RsMst_ISIN!FED), "", RsMst_ISIN!FED)
TxtNewLocationId = GetLocationName(RsMst_ISIN!LocationID) '(IsNull(RsMst_ISIN!LocationName), "", RsMst_ISIN!LocationName)
TxtNewCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
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
If SaveCmd.Enabled = False Then Exit Sub
Call GetGridData(mExchangeTypeID, mCMPID, mCommodityID)
    

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
MSHFlexGrid2.TextMatrix(1, 5) = "Electronic Weight"
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


Private Sub GetGridData(mExchangeID_, mCMPID_, mCommodityID_)
Call Grid_Head
If SaveCmd.Enabled = False Then Exit Sub
If IsNull(mExchangeID_) Or mExchangeID_ = 0 Then
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
tmp = tmp & " WHERE (Mst_GSL.CommodityID=" & mCommodityID_ & "  AND "
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
   MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!GSLID
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

Function ChkSelection() As Boolean
Dim RetVal As Boolean
RetVal = False
If MSHFlexGrid2.Rows > 3 Then
    For i = 2 To MSHFlexGrid2.Rows - 1
      If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
        RetVal = True
        Exit For
      End If
    Next
End If
ChkSelection = RetVal
End Function
