VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmSelectDatabase 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Select Database"
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   FillStyle       =   3  'Vertical Line
   Icon            =   "FrmSelectDatabase.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   6588
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
End
Attribute VB_Name = "FrmSelectDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Gen_DbPathFromAzure = "Y"


Call Gen_ReadSysConfigINI
Call Gen_ReadEmailConfigINI
Call Gen_ReadEmailGodownViability


Call Gen_SysTimeINI

Gen_ExcelPassword = "SM9172#9145"
Gen_ExpireGSLFlag = " (MST_GSL.Status <> 'E' And MST_GSL.Status <> 'Z') "
Dim Retval, Stmp As String
Retval = ""

Dim Pth As Variant
Pth = App.Path
Set cn = New ADODB.Connection
'MSHFlexGrid1.Cols = 3
'
'cns = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Pth & "\Store-ManDBCon.DAT; " & _
'"Jet OLEDB:Database Password=M710211#S600607;"
'cn.Open cns
'
'Stmp = "select * from Mst_Config "
'Set TmpRs0 = New ADODB.Recordset
'TmpRs0.Open Stmp, cn, adOpenKeyset, adLockOptimistic
'MSHFlexGrid1.TextMatrix(0, 0) = "Sr No"
'MSHFlexGrid1.TextMatrix(0, 1) = "Database"
'MSHFlexGrid1.TextMatrix(0, 2) = "Database Name"
'
'MSHFlexGrid1.ColWidth(0) = 1200
'MSHFlexGrid1.ColWidth(1) = 2000
'MSHFlexGrid1.ColWidth(2) = 4000
'
'If TmpRs0.RecordCount > 0 Then
'   If TmpRs0.RecordCount > 1 Then
'      MSHFlexGrid1.Rows = TmpRs0.RecordCount + 1
'      For I = 1 To TmpRs0.RecordCount
'          MSHFlexGrid1.TextMatrix(I, 0) = IIf(IsNull(TmpRs0!ID), "", TmpRs0!ID)
'          MSHFlexGrid1.TextMatrix(I, 1) = IIf(IsNull(TmpRs0!DBType), "", TmpRs0!DBType)
'          MSHFlexGrid1.TextMatrix(I, 2) = IIf(IsNull(TmpRs0!DBName), "", TmpRs0!DBName)
'          TmpRs0.MoveNext
'      Next
'      TmpRs0.Close
'      cn.Close
'      cns = ""
'   Else
      Gen_DBID = 1 ' IIf(IsNull(TmpRs0!ID), "", TmpRs0!ID)
      Gen_DBType = "SQL" 'IIf(IsNull(TmpRs0!DBType), "", TmpRs0!DBType)
      Unload Me
      frmCheckVersion.Show
      Unload frmCheckVersion
      FrmStart.Show
'   End If
'End If



End Sub

Private Sub MSHFlexGrid1_Click()
If MSHFlexGrid1.RowSel = 0 Then Exit Sub
Dim I As Variant
I = MSHFlexGrid1.RowSel
Gen_DBID = MSHFlexGrid1.TextMatrix(I, 0)
Gen_DBType = MSHFlexGrid1.TextMatrix(I, 1)
Unload Me
frmCheckVersion.Show
Unload frmCheckVersion
FrmStart.Show
End Sub
