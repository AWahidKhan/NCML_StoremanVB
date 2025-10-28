VERSION 5.00
Begin VB.Form Frm_ExcelImport 
   Caption         =   "Excel Testing"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Frm_ExcelImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo msgError
tmp = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & _
"\dpsep09.xls" & "; Extended Properties=Excel 8.0;"
Set cn = New ADODB.Connection
cn.Open tmp
tmp = "SELECT distinct ISIN FROM [Sheet1$] "
Set TmpRs = New ADODB.Recordset
TmpRs.Open tmp, cn, adOpenKeyset, adLockOptimistic
MsgBox TmpRs.Fields.Count
If TmpRs.RecordCount <= 0 Then
   Exit Sub
End If

For i = 1 To TmpRs.RecordCount
    
    tmp = "update [Sheet1$] set CLID = " & i & "where ISIN = '" & TmpRs.Fields(0) & "'"
    Set TmpRs = New ADODB.Recordset
    TmpRs.Open tmp, cn, adOpenKeyset, adLockOptimistic
    
    TmpRs.MoveNext
Next

Dim ans As Boolean
Exit Sub

msgError:

ans = Check_File
If ans = True Then
   MsgBox "Import Successfull !!!! "
Else
   MsgBox "Error in Import !!!! "
End If


End Sub

