VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmTxn_InvoiceGeneration 
   Caption         =   "Invoice Generation"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   9615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox CmbYear 
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
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   825
         Width           =   2280
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13515
         TabIndex        =   5
         Top             =   8685
         Width           =   1335
      End
      Begin VB.CommandButton CmdSave 
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
         Height          =   495
         Left            =   12165
         TabIndex        =   4
         Top             =   8685
         Width           =   1335
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   7215
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   12726
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
         Left            =   4920
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   7320
      End
      Begin VB.Label Label1 
         Caption         =   "Year"
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
         Left            =   225
         TabIndex        =   7
         Top             =   870
         Width           =   1440
      End
      Begin VB.Label LblLocationID 
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
         Left            =   4080
         TabIndex        =   2
         Top             =   870
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmTxn_InvoiceGeneration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mYr1 As Double
Dim mYr2 As Double
Dim mLocationID  As Double
Const strChecked = "þ"
Const strUnChecked = "q"
Private Sub CmbLocationID_GotFocus()
tmp = CmbLocationID.Text
Call TableMst_Location(" Where LocationID <> 0 ")
CmbLocationID.Clear
If RsMst_Location.RecordCount = 0 Then
   MsgBox "No Record found !!!! "
   Exit Sub
End If
For i = 1 To RsMst_Location.RecordCount
    CmbLocationID.AddItem RsMst_Location!LocationName
    RsMst_Location.MoveNext
Next
CmbLocationID.Text = tmp

End Sub

Private Sub CmbLocationID_LostFocus()
If CmbLocationID.Text = "" Then
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


Private Sub CmbYear_GotFocus()
tmp = CmbYear.Text
CmbYear.Clear
For i = 2010 To Year(Date)
    If i = Year(Date) Then
       If Month(Date) > 3 Then
          CmbYear.AddItem i & "-" & i + 1
       End If
    Else
       CmbYear.AddItem i & "-" & i + 1
    End If
Next
CmbYear.Text = tmp
End Sub

Private Sub CmbYear_LostFocus()
If CmbYear.Text = "" Then Exit Sub
Dim mFlag As Boolean
Dim mYr() As String

mYr = Split(CmbYear.Text, "-")

If UBound(mYr) <> 1 Then
   MsgBox "Invalid Selection"
   CmbYear.SetFocus
   Exit Sub
End If
mYr1 = mYr(0)
mYr2 = mYr(1)
mFlag = False
For i = 2010 To Year(Now)
    If mYr1 = i Then
       mFlag = True
       Exit For
    End If
Next
If mYr2 <> mYr1 + 1 Then mFlag = False
If mFlag = False Then
   MsgBox "Invalid Selection"
   CmbYear.SetFocus
   Exit Sub
End If

Call Grid_Head
End Sub

Public Sub Grid_Head()
With MSHFlexGrid1
    .Clear
    .Rows = 4
    .Cols = 13
    .Font.Size = 10
    .FixedCols = 1
    .FixedRows = 1
    .AllowUserResizing = flexResizeColumns
    .TextMatrix(0, 0) = "Module"
    .TextMatrix(0, 1) = "Apr " & mYr1
    .TextMatrix(0, 2) = "May " & mYr1
    .TextMatrix(0, 3) = "Jun " & mYr1
    .TextMatrix(0, 4) = "Jul " & mYr1
    .TextMatrix(0, 5) = "Aug " & mYr1
    .TextMatrix(0, 6) = "Sep " & mYr1
    .TextMatrix(0, 7) = "Oct " & mYr1
    .TextMatrix(0, 8) = "Nov " & mYr1
    .TextMatrix(0, 9) = "Dec " & mYr1
    .TextMatrix(0, 10) = "Jan " & mYr2
    .TextMatrix(0, 11) = "Feb " & mYr2
    .TextMatrix(0, 12) = "Mar " & mYr2

    .ColWidth(0) = 2700
    .TextMatrix(1, 0) = "Opening Balance"
    .TextMatrix(2, 0) = "Monthly Transaction Process"
    .TextMatrix(3, 0) = "Invoice Generation"
    .RowHeightMin = 400
    .RowHeight(0) = 600
End With
End Sub

