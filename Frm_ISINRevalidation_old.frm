VERSION 5.00
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
      TabIndex        =   3
      Top             =   8760
      Width           =   14925
      Begin VB.CommandButton ExitCmd 
         Caption         =   "E&xit"
         Height          =   400
         Left            =   11985
         TabIndex        =   12
         Top             =   600
         Width           =   2820
      End
      Begin VB.CommandButton SearchCmd 
         Caption         =   "Search"
         Height          =   400
         Left            =   11985
         TabIndex        =   11
         Top             =   195
         Width           =   2820
      End
      Begin VB.CommandButton LastCmd 
         Caption         =   "&Last"
         Height          =   400
         Left            =   9045
         TabIndex        =   10
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton DeleteCmd 
         Caption         =   "&Delete"
         Height          =   400
         Left            =   9045
         TabIndex        =   9
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton FirstCmd 
         Caption         =   "&First"
         Height          =   400
         Left            =   6105
         TabIndex        =   8
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton EditCmd 
         Caption         =   "&Edit"
         Height          =   400
         Left            =   6105
         TabIndex        =   7
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton PreviousCmd 
         Caption         =   "&Previous"
         Height          =   400
         Left            =   3165
         TabIndex        =   6
         Top             =   600
         Width           =   2940
      End
      Begin VB.CommandButton SaveCmd 
         Caption         =   "Save"
         Height          =   400
         Left            =   3165
         TabIndex        =   0
         Top             =   195
         Width           =   2940
      End
      Begin VB.CommandButton NextCmd 
         Caption         =   "&Next"
         Height          =   400
         Left            =   105
         TabIndex        =   5
         Top             =   600
         Width           =   3060
      End
      Begin VB.CommandButton AddCmd 
         Caption         =   "&Add New"
         Height          =   400
         Left            =   105
         TabIndex        =   4
         Top             =   195
         Width           =   3060
      End
   End
   Begin VB.Frame Frame1 
      Height          =   8655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   15015
      Begin VB.Frame Frame4 
         Height          =   3420
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         Begin VB.Frame Frame3 
            Caption         =   "New ISIN"
            Height          =   2655
            Left            =   7515
            TabIndex        =   37
            Top             =   600
            Width           =   7215
            Begin VB.TextBox TxtPreFix1 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   5125
               Locked          =   -1  'True
               TabIndex        =   56
               TabStop         =   0   'False
               Top             =   165
               Width           =   2000
            End
            Begin VB.TextBox TxtNewExchangeTypeId 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   43
               TabStop         =   0   'False
               Top             =   2100
               Width           =   2000
            End
            Begin VB.TextBox TxtNewLocationId 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   42
               TabStop         =   0   'False
               Top             =   915
               Width           =   5565
            End
            Begin VB.TextBox TxtNewCMP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   540
               Width           =   5565
            End
            Begin VB.TextBox TxtNewGrade 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   5130
               Locked          =   -1  'True
               TabIndex        =   40
               TabStop         =   0   'False
               Top             =   2100
               Width           =   2000
            End
            Begin VB.TextBox TxtNewCommodity 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   39
               TabStop         =   0   'False
               Top             =   1725
               Width           =   5565
            End
            Begin VB.TextBox TxtNewISINID 
               Appearance      =   0  'Flat
               Height          =   325
               Left            =   1560
               TabIndex        =   38
               Top             =   165
               Width           =   2280
            End
            Begin MSComCtl2.DTPicker NewNE_Date 
               Height          =   360
               Left            =   1560
               TabIndex        =   44
               TabStop         =   0   'False
               Top             =   1350
               Width           =   1680
               _ExtentX        =   2963
               _ExtentY        =   635
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   59375617
               CurrentDate     =   39864
            End
            Begin MSComCtl2.DTPicker NewFED 
               Height          =   360
               Left            =   5490
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   1305
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   635
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   59375617
               CurrentDate     =   39864
            End
            Begin VB.Label Label3 
               Caption         =   "Location ID"
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
               TabIndex        =   57
               Top             =   240
               Width           =   1020
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
               TabIndex        =   53
               Top             =   1800
               Width           =   1500
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
               TabIndex        =   52
               Top             =   240
               Width           =   1500
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
               TabIndex        =   51
               Top             =   2115
               Width           =   780
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
               TabIndex        =   50
               Top             =   615
               Width           =   1080
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
               TabIndex        =   49
               Top             =   1380
               Width           =   1050
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
               TabIndex        =   48
               Top             =   1335
               Width           =   600
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
               TabIndex        =   47
               Top             =   960
               Width           =   1560
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
               TabIndex        =   46
               Top             =   2160
               Width           =   1920
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Old ISIN"
            Height          =   2655
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   7335
            Begin VB.TextBox TxtPreFix 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   5205
               Locked          =   -1  'True
               TabIndex        =   54
               TabStop         =   0   'False
               Top             =   195
               Width           =   2000
            End
            Begin VB.TextBox TxtOldCommodity 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   1755
               Width           =   5565
            End
            Begin VB.TextBox TxtOldISINID 
               Appearance      =   0  'Flat
               Height          =   325
               Left            =   1650
               TabIndex        =   25
               Top             =   195
               Width           =   2280
            End
            Begin VB.TextBox TxtOldGrade 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   5205
               Locked          =   -1  'True
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   2145
               Width           =   2000
            End
            Begin VB.TextBox TxtOldLocationID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   960
               Width           =   5565
            End
            Begin VB.TextBox txtOldCMP 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   585
               Width           =   5565
            End
            Begin VB.TextBox TxtOldExchangeTypeID 
               Appearance      =   0  'Flat
               BackColor       =   &H80000000&
               Height          =   325
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   2130
               Width           =   2000
            End
            Begin MSComCtl2.DTPicker OldNE_Date 
               Height          =   360
               Left            =   1650
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   1335
               Width           =   1440
               _ExtentX        =   2540
               _ExtentY        =   635
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   59375617
               CurrentDate     =   39864
            End
            Begin MSComCtl2.DTPicker OldFED 
               Height          =   360
               Left            =   5925
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   1335
               Width           =   1320
               _ExtentX        =   2328
               _ExtentY        =   635
               _Version        =   393216
               Enabled         =   0   'False
               Format          =   59375617
               CurrentDate     =   39864
            End
            Begin VB.Label Label1 
               Caption         =   "Location ID"
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
               TabIndex        =   55
               Top             =   240
               Width           =   1020
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
               TabIndex        =   36
               Top             =   1755
               Width           =   1500
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
               TabIndex        =   35
               Top             =   195
               Width           =   1500
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
               Left            =   4200
               TabIndex        =   34
               Top             =   2145
               Width           =   780
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
               TabIndex        =   33
               Top             =   585
               Width           =   1080
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
               TabIndex        =   32
               Top             =   1335
               Width           =   810
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
               TabIndex        =   31
               Top             =   1335
               Width           =   840
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
               TabIndex        =   30
               Top             =   945
               Width           =   1200
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
         End
         Begin VB.TextBox TxtRevalidationReqNo 
            Appearance      =   0  'Flat
            Height          =   360
            Left            =   5385
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   1950
         End
         Begin VB.TextBox TxtRevalidationID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Enabled         =   0   'False
            Height          =   360
            Left            =   1995
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   1080
         End
         Begin MSComCtl2.DTPicker RevalidationDate 
            Height          =   360
            Left            =   9360
            TabIndex        =   19
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   635
            _Version        =   393216
            Format          =   59375617
            CurrentDate     =   39889
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
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   2145
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
            TabIndex        =   17
            Top             =   255
            Visible         =   0   'False
            Width           =   1740
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
            TabIndex        =   16
            Top             =   240
            Width           =   1620
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid2 
         Height          =   4800
         Left            =   105
         TabIndex        =   2
         Top             =   3720
         Width           =   14700
         _ExtentX        =   25929
         _ExtentY        =   8467
         _Version        =   393216
         WordWrap        =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
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
Function CheckDelete() As Boolean
Dim RetVal As Boolean
RetVal = True
If MSHFlexGrid2.Rows > 3 Then
   For u = 2 To MSHFlexGrid2.Rows - 1
       If MSHFlexGrid2.TextMatrix(u, 8) <> "" Then
          If MSHFlexGrid2.TextMatrix(u, 8) <> TxtNewISINID Then
             RetVal = False
             Exit For
          End If
          If Val(MSHFlexGrid2.TextMatrix(u, 9)) <> 0 Then
             RetVal = False
             Exit For
          End If
          If Val(MSHFlexGrid2.TextMatrix(u, 10)) <> 0 Then
             RetVal = False
             Exit For
          End If
          
       End If
   Next
End If

CheckDelete = RetVal
End Function

Private Sub DeleteCmd_Click()
If CheckDelete = False Then
   MsgBox "You can't delete this record !!! "
   Exit Sub
End If

Dim ans As Variant
On Error GoTo msgError
ans = MsgBox("Do you really want to DELETE this record???", vbYesNo)
If ans = vbYes Then

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
MsgBox "Child record Present "
RsTxn_ISINRevalidation.CancelUpdate

End Sub

Private Sub DeleteGridData()
Call TableMst_GSL("Where sm_prefix = '" & TxtPreFix & "'")
Call TableTxn_ISINGSLRevalidation("Where sm_prefix = '" & TxtPreFix & "' And ISINRevalId = " & Val(TxtRevalidationID))

For i = 1 To RsTxn_ISINGSLRevalidation.RecordCount
    RsMst_GSL.Find ("GSLID = " & RsTxn_ISINGSLRevalidation!GSLID)
    If Not RsMst_GSL.EOF Then
       RsMst_GSL!ISINID = RsTxn_ISINRevalidation!OldISIN
       If RsMst_GSL!ISINID = RsMst_GSL!FirstISINID Then
          RsMst_GSL!Flag = "O"
       End If
       RsMst_GSL.Update
    End If
    RsTxn_ISINGSLRevalidation.MoveNext
Next
tmp = "Delete from Txn_ISINGSLRevalidation Where sm_prefix = '" & TxtPreFix & "' And ISINRevalID = " & TxtRevalidationID
Call TmpTable

RsMst_GSL.Close

End Sub
Private Sub EditCmd_Click()
Edit_Flg = "E"
Frame4.Enabled = False
MSHFlexGrid2.Enabled = True
Call GButtonLock(Me, False)
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
Dim FrmLoadAccess() As Boolean
FrmLoadAccess = GetFrmLoadAccess(Me.Name)
If FrmLoadAccess(0) = False Then
   Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

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
Frame4.Enabled = True
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
'If TxtRevalidationReqNo.Text = "" Then
'    MsgBox "Blank Revalidation Req No Not Allowed !!! "
'    TxtRevalidationReqNo.SetFocus
'    Exit Sub
'End If

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
'Commented on 26 march 2010 110 pm Parag and Mayur sir.
'If txtOldCMP.Text <> TxtNewCMP.Text Then
'    MsgBox "Old CMP and New CMP Name must be same !!! "
'    TxtNewISINID.SetFocus
'    Exit Sub
'End If

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

tmp = " OldISIN = '" & TxtOldISINID.Text & "' And NewISIN = '" & TxtNewISINID & "' and sm_prefix = '" & TxtPreFix & "'"
RsTxn_ISINRevalidation.Filter = tmp
 
If RsTxn_ISINRevalidation.RecordCount = 0 Then
   RsTxn_ISINRevalidation.AddNew
   RsTxn_ISINRevalidation!ISINRevalID = GetISINRevalID
   RsTxn_ISINRevalidation!G_UID = GetGUID
   TxtRevalidationID.Text = RsTxn_ISINRevalidation!ISINRevalID
Else
'   RsTxn_ISINRevalidation.MoveFirst
'   RsTxn_ISINRevalidation.Find " ISINRevalID= " & TxtISINRevalReqNo.Text
   TxtRevalidationID.Text = RsTxn_ISINRevalidation!ISINRevalID
   If Edit_Flg = "E" Then
      Call DeleteGridData
   End If
End If
'RsTxn_ISINRevalidation!ISINRevalID = TxtRevalidationID.Text
RsTxn_ISINRevalidation!ISINRevalReqNo = Val(TxtRevalidationReqNo.Text)
RsTxn_ISINRevalidation!ISINRevalDate = RevalidationDate.Value
RsTxn_ISINRevalidation!OldISIN = TxtOldISINID.Text
RsTxn_ISINRevalidation!NewISIN = TxtNewISINID.Text
RsTxn_ISINRevalidation!SM_Prefix = TxtPreFix
RsTxn_ISINRevalidation.Update
SaveGridData
RsTxn_ISINRevalidation.Filter = ""
RsTxn_ISINRevalidation.Requery
If RsTxn_ISINRevalidation.RecordCount > 0 Then
   RsTxn_ISINRevalidation.MoveLast
End If

tmp = "select * from mst_gsl"
Call TmpTable
TmpRs.Close
Call InData
End Sub
Private Sub SaveGridData()
If MSHFlexGrid2.Rows > 3 Then
    Call TableMst_GSL(" Where sm_prefix = '" & TxtPreFix & "'")
    Call TableTxn_ISINGSLRevalidation
    For i = 2 To MSHFlexGrid2.Rows - 1
    If MSHFlexGrid2.TextMatrix(i, 1) = "" Then Exit For
       If MSHFlexGrid2.TextMatrix(i, 0) = "*" Then
          RsTxn_ISINGSLRevalidation.AddNew
          RsTxn_ISINGSLRevalidation!ISINGSLRevalId = GetISINGSLRevalId
          RsTxn_ISINGSLRevalidation!G_UID = GetGUID
          RsTxn_ISINGSLRevalidation!SM_Prefix = TxtPreFix
          RsTxn_ISINGSLRevalidation!ISINRevalID = Trim(TxtRevalidationID.Text)
          RsTxn_ISINGSLRevalidation!GSLID = MSHFlexGrid2.TextMatrix(i, 6)
          RsMst_GSL.Find ("GSLID = " & MSHFlexGrid2.TextMatrix(i, 6))
          If Not RsMst_GSL.EOF Then
             RsMst_GSL!ISINID = TxtNewISINID.Text
             RsMst_GSL!Flag = "R"
             RsMst_GSL.Update
          End If
          RsTxn_ISINGSLRevalidation.Update
          RsTxn_ISINGSLRevalidation.Requery
       End If
    Next
    RsMst_GSL.Close
    RsTxn_ISINGSLRevalidation.Close
End If


End Sub
Private Sub ExitCmd_Click()
    Unload Me
End Sub
Private Sub MSHFlexGrid2_Click()
If SaveCmd.Enabled = False Then Exit Sub  ' 1 / End
If MSHFlexGrid2.Rows > 3 Then ' 2
   If MSHFlexGrid2.RowSel <> "0" Then '3
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 1) = "" Then Exit Sub ' 4 / End
      If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "" Then ' 5
         If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 9)) = 0 And Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 10)) = 0 Then ' 7 ' 5A
            MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
         Else ' 5A
             MsgBox "Warning : Selected Lot is partially withdrawn or Pledged !!!!!"
             MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = "*"
             'Exit Sub
         End If ' 5A
      Else ' 5
         If MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 8) = TxtNewISINID Or Edit_Flg = "A" Then ' 6
            If Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 9)) = 0 And Val(MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 10)) = 0 Then ' 7
               MSHFlexGrid2.TextMatrix(MSHFlexGrid2.RowSel, 0) = ""
            Else ' 7
               MsgBox "Selected Lot is not available for de-selection !!!!!"
               Exit Sub
            End If ' 7
         Else ' 6
            MsgBox "You can't Edit this recored !!! "
         End If ' 6
      End If ' 5
   End If ' 3
End If ' 2
End Sub

Private Sub InData()
TxtRevalidationID = RsTxn_ISINRevalidation!ISINRevalID
TxtRevalidationReqNo = RsTxn_ISINRevalidation!ISINRevalReqNo
RevalidationDate = RsTxn_ISINRevalidation!ISINRevalDate
TxtOldISINID = RsTxn_ISINRevalidation!OldISIN
TxtNewISINID = RsTxn_ISINRevalidation!NewISIN
TxtPreFix = RsTxn_ISINRevalidation!SM_Prefix
Call GButtonLock(Me, True)
Call TxtOldISINID_LostFocus
Call TxtNewISINID_LostFocus
Call GridInData
TxtRevalidationReqNo.Locked = True
TxtOldISINID.Locked = True
TxtNewISINID.Locked = True
End Sub

Private Sub GridInData()
Call Grid_Head

tmp = "SELECT Mst_GSL.GSLID, Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, " & _
      "Mst_GSL.DepositeWeight  , Mst_GSL.ISINID,Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight,Mst_GSL.CDTFBags,Mst_GSL.CDTFWeight FROM " & _
      "(Mst_CMP INNER JOIN Mst_Godown ON Mst_CMP.CMPID = Mst_Godown.CMPID) INNER JOIN ((Txn_ISINRevalidation INNER JOIN Txn_ISINGSLRevalidation ON Txn_ISINRevalidation.ISINRevalID = Txn_ISINGSLRevalidation.ISINRevalId) INNER JOIN Mst_GSL ON Txn_ISINGSLRevalidation.GSLId = Mst_GSL.GSLID) ON (Mst_Godown.GodownID = Mst_GSL.GodownID) AND (Mst_CMP.CMPID = Mst_GSL.CMPID) " & _
      " WHERE Mst_GSL.sm_prefix = '" & TxtPreFix & "' And ((Txn_ISINGSLRevalidation.ISINRevalID)= " & TxtRevalidationID & ")  "

Call TmpTable
If TmpRs.RecordCount > 0 Then
    TmpRs.MoveFirst
    k = 1
    For i = 0 To TmpRs.RecordCount - 1
        k = k + 1
        MSHFlexGrid2.Rows = MSHFlexGrid2.Rows + 1
'        tmp = "SELECT Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, Mst_GSL.DepositeWeight,Mst_GSL.ISINID, "
'        tmp = tmp & " Mst_GSL.GSLID FROM Mst_Godown INNER JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
'        tmp = tmp & " WHERE (Mst_GSL.GSLId= " & TxtRevalidationID & ")"
'        Call TmpTable
        MSHFlexGrid2.TextMatrix(k, 0) = "*"
        MSHFlexGrid2.TextMatrix(k, 1) = TmpRs!GodownNo
        MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!StackNo
        MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!LotNo
        MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!DepositeBags
        MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!DepositeWeight
        MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!GSLID
        MSHFlexGrid2.TextMatrix(k, 8) = IIf(IsNull(TmpRs!ISINID), "", TmpRs!ISINID)
        MSHFlexGrid2.TextMatrix(k, 9) = TmpRs!PledgeBags + TmpRs!CDTFBags
        MSHFlexGrid2.TextMatrix(k, 10) = TmpRs!PledgeWeight + TmpRs!CDTFWeight
        TmpRs.MoveNext
    Next
End If
MSHFlexGrid2.Enabled = False
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
TxtNewExchangeTypeId = GetExchangeName(RsMst_ISIN!ExchangeTypeId) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
TxtNewGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
TxtPreFix1 = GetPreFixByLocation(RsMst_ISIN!LocationID, "P")


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
TxtPreFix = GetPreFixByLocation(RsMst_ISIN!LocationID, "P")
mCommodityID = IIf(IsNull(RsMst_ISIN!CommodityID), "", RsMst_ISIN!CommodityID)
TxtOldCommodity = GetCommodityName(RsMst_ISIN!CommodityID) 'IsNull(RsMst_ISIN!Commodity), "", RsMst_ISIN!Commodity)
mExchangeTypeID = IIf(IsNull(RsMst_ISIN!ExchangeTypeId), "", RsMst_ISIN!ExchangeTypeId)
TxtOldExchangeTypeID = GetExchangeName(RsMst_ISIN!ExchangeTypeId) '(IsNull(RsMst_ISIN!ExchangeType), "", RsMst_ISIN!ExchangeType)
TxtOldGrade = IIf(IsNull(RsMst_ISIN!Grade), "", RsMst_ISIN!Grade)
If SaveCmd.Enabled = False Then Exit Sub
Call GetGridData(mExchangeTypeID, mCMPID, mCommodityID)
    

End Sub

Public Sub Grid_Head()
MSHFlexGrid2.Clear
MSHFlexGrid2.Cols = 11
MSHFlexGrid2.Rows = 3
MSHFlexGrid2.FixedRows = 2

MSHFlexGrid2.TextMatrix(1, 0) = "   "
MSHFlexGrid2.TextMatrix(1, 1) = "CMP Name  "
MSHFlexGrid2.TextMatrix(1, 2) = "Godown  "
MSHFlexGrid2.TextMatrix(1, 3) = "Stack No"
MSHFlexGrid2.TextMatrix(1, 4) = "Lot No"
MSHFlexGrid2.TextMatrix(1, 5) = "No of Bags"
MSHFlexGrid2.TextMatrix(1, 6) = "Gross Weight"
MSHFlexGrid2.TextMatrix(1, 7) = "GSLID"
MSHFlexGrid2.TextMatrix(1, 8) = "ISINGSLRevalId"
MSHFlexGrid2.TextMatrix(1, 9) = "Current ISIN"
MSHFlexGrid2.TextMatrix(1, 10) = "Pledged or Withdrawn Bags"
MSHFlexGrid2.TextMatrix(1, 11) = "Pledged or Withdrawn Weight"

MSHFlexGrid2.ColWidth(1) = 1800
MSHFlexGrid2.ColWidth(2) = 1800
MSHFlexGrid2.ColWidth(3) = 1600
MSHFlexGrid2.ColWidth(4) = 1600
MSHFlexGrid2.ColWidth(5) = 1500
MSHFlexGrid2.ColWidth(6) = 1500
MSHFlexGrid2.ColWidth(7) = 0
MSHFlexGrid2.ColWidth(8) = 0
MSHFlexGrid2.ColWidth(9) = 1800
MSHFlexGrid2.ColWidth(10) = 1800
MSHFlexGrid2.ColWidth(11) = 1800

MSHFlexGrid2.RowHeight(1) = 500
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

tmp = "SELECT MstCMP.CMPName,Mst_Godown.GodownNo, Mst_GSL.StackNo, Mst_GSL.LotNo, Mst_GSL.DepositeBags, Mst_GSL.DepositeWeight,Mst_GSL.ISINID, "
tmp = tmp & " Mst_GSL.GSLID,Mst_GSL.ISINID,Mst_GSL.PledgeBags,Mst_GSL.PledgeWeight,Mst_GSL.CDTFBags,Mst_GSL.CDTFWeight  FROM Mst_Godown INNER "
tmp = tmp & " JOIN Mst_GSL ON Mst_Godown.GodownID = Mst_GSL.GodownID "
tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_GSL.CMPID "
tmp = tmp & " WHERE sm_prefix = '" & TxtPreFix & "' And ( " & Gen_ExpireGSLFlag & " ) And (Mst_GSL.CommodityID=" & mCommodityID_ & "  AND "
tmp = tmp & " Mst_GSL.ExchangeTypeID=" & mExchangeID_ & ") AND " ' AND Mst_Godown.CMPID= " & mCMPID_ & ") AND  "
tmp = tmp & " (Mst_GSL.ISINID= '" & TxtOldISINID & "')"

'If SaveCmd.Enabled = False Then
'   tmp = tmp & " (Mst_GSL.ISINID= '" & TxtOldISINID & "')"
'Else
   'tmp = tmp & " (Mst_GSL.ISINID Is Null  OR Mst_GSL.ISINID= '' OR Mst_GSL.ISINID= '" & TxtOldISINID & "')"
'   tmp = tmp & " (Mst_GSL.ISINID= '" & TxtOldISINID & "')"
'End If
  
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
   MSHFlexGrid2.TextMatrix(k, 1) = "CMP Name"
   MSHFlexGrid2.TextMatrix(k, 2) = TmpRs!GodownNo
   MSHFlexGrid2.TextMatrix(k, 3) = TmpRs!StackNo
   MSHFlexGrid2.TextMatrix(k, 4) = TmpRs!LotNo
   MSHFlexGrid2.TextMatrix(k, 5) = TmpRs!DepositeBags
   MSHFlexGrid2.TextMatrix(k, 6) = TmpRs!DepositeWeight
   MSHFlexGrid2.TextMatrix(k, 7) = TmpRs!GSLID
   MSHFlexGrid2.TextMatrix(k, 8) = TmpRs!ISINID
   MSHFlexGrid2.TextMatrix(k, 9) = TmpRs!PledgeBags + TmpRs!CDTFBags
   MSHFlexGrid2.TextMatrix(k, 10) = TmpRs!PledgeWeight + TmpRs!CDTFWeight
   
   TmpRs.MoveNext
Next

End Sub

Private Function GetISINRevalID() As Double
Dim RetVal As Double
tmp = "Select max(ISINRevalID) from Txn_ISINRevalidation where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
TmpRs.Close
GetISINRevalID = RetVal
End Function

Private Function GetISINGSLRevalId() As Double
Dim RetVal As Double
tmp = "Select max(ISINGSLRevalId) from Txn_ISINGSLRevalidation where sm_prefix = '" & TxtPreFix & "'"
Call TmpTable
RetVal = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
RetVal = RetVal + 1
TmpRs.Close
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
