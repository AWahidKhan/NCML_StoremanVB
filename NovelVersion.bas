Attribute VB_Name = "NovelVersion"
Dim sql As String
Dim yesNo As VbMsgBoxStyle
Dim isMatched As Boolean
Public vVersionNumber As Variant ' Add Get Version No
'Interest on SD
'Debit Note
'Credit Note
'Milling Charges

'''****---- Function used to store Version Info -----******
''************************* --------  Add On Date :- 03/07/2007 --------- *************************''
Public Function GetVersionNo()
vVersionNumber = "332"
End Function
''----******  END-----*****

Public Sub CreateQuery(strDBPath As String, _
                strsql As String, _
                strQryName As String)
   Dim catDB As ADOX.Catalog
   Dim cmd   As ADODB.Command
   Set catDB = New ADOX.Catalog
   ' Open the catalog.
   catDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & strDBPath & "; Jet OLEDB:Database Password=Mypwd;"

   Set cmd = New ADODB.Command
   ' Define a Command object to contain the query's SQL, and then
   ' save it to the database's Views collection. StrSQL must
   ' contain only a SELECT statement.
   cmd.CommandText = strsql
   catDB.Views.Append strQryName, cmd
   Set catDB = Nothing
End Sub

Public Function CompareVersion(verEXE As String, verDB As String) As Boolean
    If verEXE = verDB Then
        'MsgBox "Update completed", vbInformation, "Normal Execution of Application"
        CompareVersion = True
    Else
        verDB = Trim(Str(Val(verDB) + 1))
        Select Case verDB
'            Case "1"
'                Version1
'            Case "2"
'                Version2
'            Case "3"
'                Version3
'            Case "4"
'                Version4
'            Case "5"
'                Version5
'            Case "6"
'                Version6
'            Case "7"
'                Version7
'            Case "8"
'                Version8
'            Case "9"
'                Version9
'            Case "10"
'                Version10
'            Case "11"
'                Version11
'            Case "12"
'                Version12
'            Case "13"
'                Version13
'            Case "14"
'                Version14
'            Case "15"
'                Version15
'            Case "16"
'                Version16
'            Case "17"
'                Version17
'            Case "18"
'                Version18
'            Case "19"
'                Version19
'            Case "20"
'                Version20
'            Case "21"
'                Version21
'            Case "22"
'                Version22
'            Case "23"
'                Version23
'            Case "24"
'                Version24
'            Case "25"
'                Version25
'            Case "26"
'                Version26
'            Case "27"
'                Version27
'            Case "28"
'                Version28
'            Case "29"
'                Version29
'            Case "30"
'                Version30
'            Case "31"
'                Version31
'            Case "32"
'                Version32
'            Case "33"
'                Version33
'            Case "34"
'                Version34
'            Case "35"
'                Version35
'            Case "36"
'                Version36
'            Case "37"
'                Version37
'            Case "38"
'                Version38
'            Case "39"
'                Version39
'            Case "40"
'                Version40
'            Case "41"
'                Version41
'            Case "42"
'                Version42
'            Case "43"
'                Version43
'            Case "44"
'                Version44
'             Case "45"
'                Version45
'             Case "46"
'                Version46
'             Case "47"
'                Version47
'             Case "48"
'                Version48
'             Case "49"
'                Version49
'             Case "50"
'                Version50
'             Case "51"
'                Version51
'             Case "52"
'                Version52
'             Case "53"
'                Version53
'             Case "54"
'                Version54
'             Case "55"
'                Version55
'             Case "56"
'                Version56
'             Case "57"
'                Version57
'             Case "58"
'                Version58
'             Case "59"
'                Version59
'             Case "60"
'                Version60
'             Case "61"
'                Version61
'             Case "62"
'                Version62
'             Case "63"
'                Version63
'             Case "64"
'                Version64
'             Case "65"
'                Version65
'             Case "66"
'                Version66
'             Case "67"
'                Version67
'             Case "68"
'                Version68
'             Case "69"
'                Version69
'             Case "70"
'                Version70
'             Case "71"
'                Version71
'             Case "72"
'                Version72
'             Case "73"
'                Version73
'             Case "74"
'                Version74
'             Case "75"
'                Version75
'             Case "76"
'                Version76
'             Case "77"
'                Version77
'             Case "78"
'                Version78
'             Case "79"
'                Version79
'             Case "80"
'                Version80
'             Case "81"
'                Version81
'             Case "82"
'                Version82
'             Case "83"
'                Version83
'             Case "84"
'                Version84
'             Case "85"
'                Version85
'              Case "86"
'                Version86
'              Case "87"
'                Version87
'              Case "88"
'                Version88
'              Case "89"
'                Version89
'              Case "90"
'                Version90
'              Case "91"
'                Version91
'              Case "92"
'                Version92
'              Case "93"
'                Version93
'              Case "94"
'                Version94
'              Case "95"
'                Version95
'              Case "96"
'                Version96
'              Case "97"
'                Version97
'              Case "98"
'                Version98
'              Case "99"
'                Version99
'              Case "100"
'                Version100
'              Case "101"
'                Version101
'              Case "102"
'                Version102
'              Case "103"
'                Version103
'              Case "104"
'                Version104
'              Case "105"
'                Version105
'              Case "106"
'                Version106
'              Case "107"
'                Version107
'              Case "108"
'                Version108
'              Case "109"
'                Version109
'              Case "110"
'                Version110
'              Case "111"
'                Version111
'              Case "112"
'                Version112
'              Case "113"
'                Version113
'              Case "114"
'                Version114
'              Case "115"
'                Version115
'              Case "116"
'                Version116
'              Case "117"
'                Version117
'              Case "118"
'                Version118
'              Case "119"
'                Version119
'              Case "120"
'                Version120
'              Case "121"
'                Version121
'              Case "122"
'                Version122
'              Case "123"
'                Version123
'              Case "124"
'                Version124
'              Case "125"
'                Version125
'              Case "126"
'                Version126
'              Case "127"
'                Version127
'              Case "128"
'                Version128
'              Case "129"
'                Version129
'              Case "130"
'                Version130
'              Case "131"
'                Version131
'              Case "132"
'                Version132
'              Case "133"
'                Version133
'              Case "134"
'                Version134
'              Case "135"
'                Version135
'              Case "136"
'                Version136
'              Case "137"
'                Version137
'              Case "138"
'                Version138
'              Case "139"
'                Version139
'              Case "140"
'                Version140
'              Case "141"
'                Version141
'              Case "142"
'                Version142
'              Case "143"
'                Version143
'              Case "144"
'                Version144
'              Case "145"
'                Version145
'              Case "146"
'                Version146
'              Case "147"
'                Version147
'              Case "148"
'                Version148
'              Case "149"
'                Version149
'              Case "150"
'                Version150
'              Case "151"
'                Version151
'              Case "152"
'                Version152
'              Case "153"
'                Version153
'              Case "154"
'                Version154
'              Case "155"
'                Version155
              Case "156"
                Version156
              Case "157"
                Version157
              Case "158"
                Version158
              Case "159"
                Version159
              Case "160"
                Version160
              Case "161"
                Version161
              Case "162"
                Version162
              Case "163"
                Version163
              Case "164"
                Version164
              Case "165"
                Version165
              Case "166"
                Version166
              Case "167"
                Version167
              Case "168"
                Version168
              Case "169"
                Version169
              Case "170"
                Version170
              Case "171"
                Version171
              Case "172"
                Version172
              Case "173"
                Version173
              Case "173"
                Version173
              Case "174"
                Version174
              Case "175"
                Version175
              Case "176"
                Version176
              Case "177"
                Version177
              Case "178"
                Version178
              Case "179"
                version179
              Case "180"
                version180
              Case "181"
                version181
              Case "182"
                version182
              Case "183"
                version183
              Case "184"
                Version184
              Case "185"
                Version185
              Case "186"
                Version186
              Case "187"
                version187
              Case "188"
                version188
              Case "189"
                version189
              Case "190"
                version190
              Case "191"
                version191
              Case "192"
                version192
              Case "193"
                version193
              Case "194"
                version194
              Case "195"
                version195
              Case "196"
                version196
              Case "197"
                version197
              Case "198"
                version198
              Case "199"
                version199
              Case "200"
                version200
              Case "201"
                version201
              Case "202"
                version202
              Case "203"
                version203
              Case "204"
                version204
              Case "205"
                version205
              Case "206"
                version206
              Case "207"
                version207
              Case "208"
                version208
              Case "209"
                version209
              Case "210"
                version210
              Case "211"
                version211
              Case "212"
                version212
              Case "213"
                version213
              Case "214"
                version214
              Case "215"
                version215
              Case "216"
                version216
              Case "217"
                version217
              Case "218"
                version218
              Case "219"
                version219
              Case "220"
                version220
              Case "221"
                version221
              Case "222"
                version222
              Case "223"
                version223
              Case "224"
                version224
              Case "225"
                version225
              Case "226"
                version226
              Case "227"
                version227
              Case "228"
                version228
              Case "229"
                version229
              Case "230"
                version230
              Case "231"
                version231
              Case "232"
                version232
              Case "233"
                version233
              Case "234"
                version234
              Case "235"
                version235
              Case "236"
                version236
              Case "237"
                version237
              Case "238"
                version238
              Case "239"
                version239
              Case "240"
                version240
              Case "241"
                version241
              Case "242"
                version242
              Case "243"
                version243
              Case "244"
                version244
              Case "245"
                version245
              Case "246"
                version246
              Case "247"
                version247
              Case "248"
                version248
              Case "249"
                version249
              Case "250"
                version250
              Case "251"
                version251
              Case "252"
                version252
              Case "253"
                version253
              Case "254"
                version254
              Case "255"
                version255
              Case "256"
                version256
              Case "257"
                version257
              Case "258"
                version258
              Case "259"
                version259
              Case "260"
                Version260
              Case "261"
                Version261
              Case "262"
                Version262
              Case "263"
                Version263
              Case "264"
                Version264
              Case "265"
                Version265
              Case "266"
                Version266
              Case "267"
                Version267
              Case "268"
                Version268
              Case "269"
                Version269
              Case "270"
                Version270
              Case "271"
                Version271
              Case "272"
                Version272
              Case "273"
                Version273
              Case "274"
                Version274
              Case "275"
                Version275
              Case "276"
                Version276
              Case "277"
                Version277
              Case "278"
                Version278
              Case "279"
                Version279
              Case "280"
                version280
              Case "281"
                Version281
              Case "282"
                Version282
              Case "283"
                Version283
              Case "284"
                Version284
              Case "285"
                version285
              Case "286"
                version286
              Case "287"
                version287
              Case "288"
                version288
              Case "289"
                Version289
              Case "290"
                Version290
              Case "291"
                Version291
              Case "292"
                Version292
              Case "293"
                Version293
              Case "294"
                Version294
              Case "295"
                Version295
              Case "296"
                Version296
              Case "297"
                Version297
              Case "298"
                Version298
              Case "299"
                Version299
              Case "300"
                Version300
              Case "301"
                Version301
              Case "302"
                Version302
              Case "303"
                Version303
              Case "304"
                Version304
              Case "305"
                Version305
              Case "306"
                Version306
              Case "307"
                Version307
              Case "308"
                Version308
              Case "309"
                Version309
              Case "310"
                Version310
              Case "311"
                Version311
              Case "312"
                Version312
              Case "313"
                Version313
              Case "314"
                Version314
              Case "315"
                version315
              Case "316"
                Version316
              Case "317"
                Version317
              Case "318"
                Version318
              Case "319"
                Version319
              Case "320"
                Version320
              Case "321"
                Version321
              Case "322"
                Version322
              Case "323"
                Version323
              Case "324"
                Version324
              Case "325"
                Version325
              Case "326"
                Version326
              Case "327"
                Version327
              Case "328"
                Version328
              Case "329"
                Version329
              Case "330"
                Version330
              Case "331"
                Version331
              Case "332"
                Version332
              Case Else
                VersionUnknown
                End
        End Select
        CompareVersion = False
    End If
End Function
Public Function Version332()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Alter Table Mst_StorageContract  Add StockManagmentContract Varchar(1)"
Call TmpTable


UpdateConfig ("332")
MsgBox "Version 332 Installed successfully"
End Function

Public Function Version331()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Alter Table Txn_InsurancePolicyData Alter Column PolicyDayPremium Decimal(18,5)"
Call TmpTable

tmp = " Alter Table Txn_ArchiveInsurancePolicyData Alter Column PolicyDayPremium Decimal(18,5)"
Call TmpTable

UpdateConfig ("331")
MsgBox "Version 331 Installed successfully"
End Function

Public Function Version330()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Txn_GodownWiseInsurancePremiumDetail("
tmp = tmp & " ID      Decimal(18,0) IDENTITY(1,1) NOT NULL,"
tmp = tmp & " ProcessDate DateTime,"
tmp = tmp & " GodownID Decimal(18,0) NULL,"
tmp = tmp & " Amount   Decimal(18,2) NOT NULL,"
tmp = tmp & " ProcessFlag  " & TextType & "(30),"
tmp = tmp & " G_UID   " & TextType & "(36),"
tmp = tmp & " Remark " & TextType & "(255),"
tmp = tmp & " Concurrency  DATETIME,"
tmp = tmp & " CreatedDate   DATETIME,"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate  DATETIME,"
tmp = tmp & " UpdatedBy  " & TextType & "(30))"

Call TmpTable


tmp = " ALTER TABLE Txn_GodownWiseInsurancePremiumDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_GodownWiseInsurancePremiumDetail PRIMARY KEY CLUSTERED (ID),"
tmp = tmp & " CONSTRAINT DF_Txn_GodownWiseInsurancePremiumDetail DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Txn_GodownWiseInsurancePremiumDetail_Mst_Godown_GodownID FOREIGN KEY (GodownID) REFERENCES Mst_Godown(GodownID)"

Call TmpTable

UpdateConfig ("330")
MsgBox "Version 330 Installed successfully"
End Function

Public Function Version329()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Txn_GodownWisePnLExceptionLocation("
tmp = tmp & " GodownWisePnLExceptionLocationID      Decimal(18,0) IDENTITY(1,1) NOT NULL,"
tmp = tmp & " MonthYear DateTime,"
tmp = tmp & " LocationID Decimal(18,0) NULL,"
tmp = tmp & " GodownID Decimal(18,0) NULL,"
tmp = tmp & " Amount   Decimal(18,2) NOT NULL,"
tmp = tmp & " Flag        VARCHAR(30),"
tmp = tmp & " G_UID   VARCHAR(36),"
tmp = tmp & " Remark VARCHAR(255),"
tmp = tmp & " Concurrency  DATETIME,"
tmp = tmp & " CreatedDate   DATETIME,"
tmp = tmp & " CreatedBy   VARCHAR(30),"
tmp = tmp & " UpdatedDate  DATETIME,"
tmp = tmp & " UpdatedBy  VARCHAR(30))"

Call TmpTable

tmp = " ALTER TABLE Txn_GodownWisePnLExceptionLocation WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_GodownWisePnLExceptionLocation PRIMARY KEY CLUSTERED (GodownWisePnLExceptionLocationID),"
tmp = tmp & " CONSTRAINT DF_Txn_GodownWisePnLExceptionLocation DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Txn_GodownWisePnLExceptionLocation_Mst_Location_LocationID FOREIGN KEY (LocationID) REFERENCES Mst_Location(LocationID),"
tmp = tmp & " CONSTRAINT FK_Txn_GodownWisePnLExceptionLocation_Mst_Godown_GodownID FOREIGN KEY (GodownID) REFERENCES Mst_Godown(GodownID)"

Call TmpTable

UpdateConfig ("329")
MsgBox "Version 329 Installed successfully"
End Function

Public Function Version328()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


''--  Alter Txn_ArchiveInsurancePolicyData table
tmp = " ALTER TABLE Txn_ArchiveInsurancePolicyData Add PolicyDayPremium Decimal(18,2),ProcessFlag Varchar(1) "
Call TmpTable

''--  Alter Txn_InsurancePolicyData table
tmp = " ALTER TABLE Txn_InsurancePolicyData Add PolicyDayPremium Decimal(18,2),ProcessFlag Varchar(1)  "
Call TmpTable

''--  Alter Txn_ArchiveInsurancePolicyData table
tmp = " ALTER TABLE Txn_ArchiveInsurancePolicyData Alter Column RowID BigInt "
Call TmpTable


UpdateConfig ("328")
MsgBox "Version 328 Installed successfully"
End Function


Public Function Version327()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Create Insurance Premium table
tmp = " Create Table Txn_InsurancePremium ("
tmp = tmp & " InsurancePremiumID              Decimal(18,0) Not Null,"
tmp = tmp & " InsurancePremiumDate            Date,"
tmp = tmp & " InsurancePolicyID               Decimal(18,0) Not Null,"
tmp = tmp & " PolicyPremium                   Decimal(18,2),"
tmp = tmp & " Flag                            Varchar(1),"
tmp = tmp & " Remark                          Varchar(1000),"
tmp = tmp & " G_UID                           Varchar(36),"
tmp = tmp & " Concurrency                     DATETIME,"
tmp = tmp & " CreatedBy                       Varchar(30),"
tmp = tmp & " CreatedDate                     DATETIME,"
tmp = tmp & " UpdatedBy                       Varchar(30),"
tmp = tmp & " UpdatedDate                     DateTime )"
Call TmpTable

tmp = " ALTER TABLE Txn_InsurancePremium WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_InsurancePremium PRIMARY KEY CLUSTERED (InsurancePremiumID),"
tmp = tmp & " CONSTRAINT DF_Txn_InsurancePremium DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Txn_InsurancePremium_InsurancePolicyID FOREIGN KEY (InsurancePolicyID)        REFERENCES Mst_InsuranceDetails (InsurancePolicyID)"
Call TmpTable

''--  Create Insurance Premium Detail table
tmp = " Create Table Txn_InsurancePremiumDetail ("
tmp = tmp & " InsurancePremiumDetailID                Decimal(18,0) Not Null,"
tmp = tmp & " InsurancePremiumID                      Decimal(18,0) Not Null,"
tmp = tmp & " PremiumMonth                            Date,"
tmp = tmp & " PolicyPremiumMonthly                    Decimal(18,2),"
tmp = tmp & " AdditionalPremium                       Decimal(18,2),"
tmp = tmp & " FromDate                                Date,"
tmp = tmp & " ToDate                                  Date,"
tmp = tmp & " AdditionalPremiumMonthly                Decimal(18,2),"
tmp = tmp & " TotalMonthlyPremium                     Decimal(18,2),"
tmp = tmp & " Flag                                    Varchar(1),"
tmp = tmp & " Remark                                  Varchar(8000),"
tmp = tmp & " G_UID                                   Varchar(36),"
tmp = tmp & " Concurrency                             DATETIME,"
tmp = tmp & " CreatedBy                               Varchar(30),"
tmp = tmp & " CreatedDate                             DATETIME,"
tmp = tmp & " UpdatedBy                               Varchar(30),"
tmp = tmp & " UpdatedDate                             DateTime)"
Call TmpTable

tmp = " ALTER TABLE Txn_InsurancePremiumDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_InsurancePremiumDetail PRIMARY KEY CLUSTERED (InsurancePremiumDetailID),"
tmp = tmp & " CONSTRAINT DF_InsurancePremiumDetailID DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Txn_InsurancePremiumDetail_InsurancePremiumID FOREIGN KEY (InsurancePremiumID)        REFERENCES Txn_InsurancePremium (InsurancePremiumID)"
Call TmpTable

''--  Create Log Table for Insurance Premium Detail
tmp = " Create Table Log_Txn_InsurancePremiumDetail ("
tmp = tmp & " LogInsurancePremiumDetailID             Decimal(18,0) Not Null Identity(1,1),"
tmp = tmp & " LogConcurrency                          DATETIME DEFAULT GETDATE(),"
tmp = tmp & " InsurancePremiumDetailID                Decimal(18,0) Not Null,"
tmp = tmp & " InsurancePremiumID                      Decimal(18,0) Not Null,"
tmp = tmp & " PremiumMonth                            Date,"
tmp = tmp & " PolicyPremiumMonthly                    Decimal(18,2),"
tmp = tmp & " AdditionalPremium                       Decimal(18,2),"
tmp = tmp & " FromDate                                Date,"
tmp = tmp & " ToDate                                  Date,"
tmp = tmp & " AdditionalPremiumMonthly                Decimal(18,2),"
tmp = tmp & " TotalMonthlyPremium                     Decimal(18,2),"
tmp = tmp & " Flag                                    Varchar(1),"
tmp = tmp & " Remark                                  Varchar(8000),"
tmp = tmp & " G_UID                                   Varchar(36),"
tmp = tmp & " Concurrency                             DATETIME,"
tmp = tmp & " CreatedBy                               Varchar(30),"
tmp = tmp & " CreatedDate                             DATETIME,"
tmp = tmp & " UpdatedBy                               Varchar(30),"
tmp = tmp & " UpdatedDate                             DATETIME,"
tmp = tmp & " LogFlag                                 VarChar(2))"
Call TmpTable

''--  Create TRIGGER for Log Insurance Premium Detail
tmp = " CREATE TRIGGER [dbo].[Tr_Log_Txn_InsurancePremiumDetail]"
tmp = tmp & " ON [dbo].[Txn_InsurancePremiumDetail] After Insert, Update, Delete AS"
tmp = tmp & " If exists (Select * from inserted) and not exists(Select * from deleted)"
tmp = tmp & " Begin"
tmp = tmp & " Insert Log_Txn_InsurancePremiumDetail(InsurancePremiumDetailID, InsurancePremiumID, PremiumMonth, PolicyPremiumMonthly, AdditionalPremium, FromDate, ToDate, AdditionalPremiumMonthly, TotalMonthlyPremium, Flag, Remark, G_UID, Concurrency, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, LogFlag)"
tmp = tmp & " SELECT InsurancePremiumDetailID,InsurancePremiumID,PremiumMonth,PolicyPremiumMonthly,AdditionalPremium,FromDate,ToDate,AdditionalPremiumMonthly,TotalMonthlyPremium,Flag,Remark,G_UID,Concurrency,CreatedBy,CreatedDate,UpdatedBy,UpdatedDate,'I' FROM inserted"
tmp = tmp & " End"
tmp = tmp & " If exists(SELECT * from inserted) and exists (SELECT * from deleted)"
tmp = tmp & " Begin"
tmp = tmp & " Insert Log_Txn_InsurancePremiumDetail(InsurancePremiumDetailID, InsurancePremiumID, PremiumMonth, PolicyPremiumMonthly, AdditionalPremium, FromDate, ToDate, AdditionalPremiumMonthly, TotalMonthlyPremium, Flag, Remark, G_UID, Concurrency, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, LogFlag)"
tmp = tmp & " SELECT i.InsurancePremiumDetailID,i.InsurancePremiumID,i.PremiumMonth,i.PolicyPremiumMonthly,i.AdditionalPremium,i.FromDate,i.ToDate,i.AdditionalPremiumMonthly,i.TotalMonthlyPremium,i.Flag,i.Remark,i.G_UID,i.Concurrency,i.CreatedBy,i.CreatedDate,i.UpdatedBy,i.UpdatedDate,'U' FROM inserted i INNER JOIN deleted d ON i.InsurancePremiumDetailID = d.InsurancePremiumDetailID"
tmp = tmp & " End"
tmp = tmp & " If exists(select * from deleted) and not exists(Select * from inserted)"
tmp = tmp & " Begin"
tmp = tmp & " Insert Log_Txn_InsurancePremiumDetail(InsurancePremiumDetailID, InsurancePremiumID, PremiumMonth, PolicyPremiumMonthly, AdditionalPremium, FromDate, ToDate, AdditionalPremiumMonthly, TotalMonthlyPremium, Flag, Remark, G_UID, Concurrency, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, LogFlag)"
tmp = tmp & " SELECT InsurancePremiumDetailID,InsurancePremiumID,PremiumMonth,PolicyPremiumMonthly,AdditionalPremium,FromDate,ToDate,AdditionalPremiumMonthly,TotalMonthlyPremium,Flag,Remark,G_UID,Concurrency,CreatedBy,CreatedDate,UpdatedBy,UpdatedDate, 'D' FROM deleted"
tmp = tmp & " End"
Call TmpTable


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_InsurancePremium.frm','Insurance Premium','FrmTxn_InsurancePremium','Y','T',0,228,1"
Call TmpTable


UpdateConfig ("327")
MsgBox "Version 327 Installed successfully"

End Function


Public Function Version326()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Alter Commodity Price table
tmp = " ALTER TABLE Mst_CommodityPrice ADD G_UID VARCHAR(36),CreatedBy VARCHAR(30),CreatedDate DATETIME,UpdatedBy VARCHAR(30),UpdatedDate DATETIME"
Call TmpTable

UpdateConfig ("326")
MsgBox "Version 326 Installed successfully"
End Function

Public Function Version325()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Txn_ArchiveInsurancePolicyData("
tmp = tmp & " RowID                           Int Not NUll Identity(1,1),"
tmp = tmp & " ClientIDRow                     Decimal(18,0),"
tmp = tmp & " CommodityID                     Decimal(18,0),"
tmp = tmp & " CMPID                           Decimal(18,0),"
tmp = tmp & " GodownID                        Decimal(18,0),"
tmp = tmp & " LocationID                      Decimal(18,0),"
tmp = tmp & " BalanceBags                     Decimal(18,0),"
tmp = tmp & " BalanceWeight                   Decimal(18,4),"
tmp = tmp & " Rate                            Decimal(18,2),"
tmp = tmp & " YearMonth                       Varchar(10),"
tmp = tmp & " AUM                             Decimal(18,2),"
tmp = tmp & " InsuranceCovered                Decimal(18,2),"
tmp = tmp & " LinkByPolicyID                  Varchar(500),"
tmp = tmp & " CoveredByPolicyID               Varchar(500),"
tmp = tmp & " ClientNCMLFlagID                Decimal(18,0),"
tmp = tmp & " StructureWiseCMPLimitOver       Varchar(1),"
tmp = tmp & " ClientNCMLStockFlag             Varchar(1),"
tmp = tmp & " Flag                            Varchar(1),"
tmp = tmp & " G_UID                           Varchar(36),"
tmp = tmp & " Concurrency                     DATETIME DEFAULT GETDATE(),"
tmp = tmp & " CreatedBy                       Varchar(30),"
tmp = tmp & " CreatedDate                     DATETIME,"
tmp = tmp & " UpdatedBy                       Varchar(30),"
tmp = tmp & " UpdatedDate                     DateTime"
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table Txn_InsurancePolicyData ("
tmp = tmp & " RowID                           Int Not NUll Identity(1,1),"
tmp = tmp & " ClientIDRow                     Decimal(18,0),"
tmp = tmp & " CommodityID                     Decimal(18,0),"
tmp = tmp & " CMPID                           Decimal(18,0),"
tmp = tmp & " GodownID                        Decimal(18,0),"
tmp = tmp & " LocationID                      Decimal(18,0),"
tmp = tmp & " BalanceBags                     Decimal(18,0),"
tmp = tmp & " BalanceWeight                   Decimal(18,4),"
tmp = tmp & " Rate                            Decimal(18,2),"
tmp = tmp & " YearMonth                       Varchar(10),"
tmp = tmp & " AUM                             Decimal(18,2),"
tmp = tmp & " InsuranceCovered                Decimal(18,2),"
tmp = tmp & " LinkByPolicyID                  Varchar(500),"
tmp = tmp & " CoveredByPolicyID               Varchar(500),"
tmp = tmp & " ClientNCMLFlagID                Decimal(18,0),"
tmp = tmp & " StructureWiseCMPLimitOver       Varchar(1),"
tmp = tmp & " ClientNCMLStockFlag             Varchar(1),"
tmp = tmp & " Flag                            Varchar(1),"
tmp = tmp & " G_UID                           Varchar(36),"
tmp = tmp & " Concurrency                     DATETIME DEFAULT GETDATE(),"
tmp = tmp & " CreatedBy                       Varchar(30),"
tmp = tmp & " CreatedDate                     DATETIME,"
tmp = tmp & " UpdatedBy                       Varchar(30),"
tmp = tmp & " UpdatedDate DateTime"
tmp = tmp & " )"
Call TmpTable





''--  Insert Menu for CLient List for Navision Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_InsuranceProcess.frm','Insurance Process','FrmTxn_InsuranceProcess','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("325")
MsgBox "Version 325 Installed successfully"
End Function


Public Function Version324()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Insert Menu for CLient List for Navision Report
tmp = " Alter Table Mst_Location Add SiloLocation Bit Default 0 "
Call TmpTable

tmp = " Alter Table Txn_CxTF_TruckDetail Add ForeignMatter Decimal(18,3) "
Call TmpTable

UpdateConfig ("324")
MsgBox "Version 324 Installed successfully"
End Function


Public Function Version323()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Log_WHR_RevalidationDetail ("
tmp = tmp & " WHRNo Decimal(18,0),"
tmp = tmp & " SM_Prefix Varchar(5),"
tmp = tmp & " WHRDate DateTime,"
tmp = tmp & " OldStorageTo DateTime,"
tmp = tmp & " OldGradeValidTill DateTime,"
tmp = tmp & " NewStorageTo DateTime,"
tmp = tmp & " NewGradeValidTill DateTime,"
tmp = tmp & " ReasonForRevalidation Varchar(5000),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy Varchar(30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " G_UID VarChar(36)"
tmp = tmp & " Primary Key (G_UID))"
Call TmpTable

tmp = " ALTER TABLE Log_WHR_RevalidationDetail WITH NOCHECK ADD CONSTRAINT DF_Log_WHR_RevalidationDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_WHRRevalidation.frm','WHR Revalidation','FrmTxn_WHRRevalidation','Y','T',0,228,1"

Call TmpTable


''--  Insert Menu for CLient List for Navision Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_StateWiseGSTNforNavision.frm','Pending Client List for Navision Report','FrmRpt_StateWiseGSTNforNavision','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("323")
MsgBox "Version 323 Installed successfully"
End Function

Public Function Version322()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table log_Mst_GSL Add BillingCycleID Decimal(18,0),BillingUnit Varchar(1),Rate Decimal(18,2),RateForDailyPnL Decimal(18,2),DepConDID Decimal(18,0),NoofDays Decimal(18,0)"
Call TmpTable


tmp = " ALTER TRIGGER [dbo].[tr_Mst_GSL] ON [dbo].[Mst_GSL]"
tmp = tmp & " AFTER"
tmp = tmp & " UPDATE AS DECLARE @mGSLID decimal(18, 0), @mSM_Prefix Varchar(5), @mFlag Varchar(1), @mWHRBags Decimal(18,0), @mWHRWeight Decimal(18,3),"
tmp = tmp & " @mCDTFBags Decimal(18,0), @mCDTFWeight Decimal(18,3), @mBalanceBags Decimal(18,0), @mBalanceWeight Decimal(18,3), @mRemarks Varchar (8000),"
tmp = tmp & " @mUpdatedBy  Varchar (30), @mUpdatedOn DateTime, @BillingCycleID Decimal(18,0),@BillingUnit Varchar(1),@Rate Decimal(18,2),@RateForDailyPnL Decimal(18,2),@DepConDID Decimal(18,0),@NoofDays Decimal(18,0)"
tmp = tmp & " SELECT @mGSLID = d.GSLID, @mSM_Prefix = d.SM_Prefix, @mFlag = d.Flag, @mWHRBags = d.WHRBags, @mWHRWeight = d.WHRWeight, @mCDTFBags = d.CDTFBags,"
tmp = tmp & " @mCDTFWeight = d.CDTFWeight, @mBalanceBags = d.BalanceBags, @mBalanceWeight = d.BalanceWeight, @mRemarks = d.Remarks, @mUpdatedBy    = d.UpdatedBy,"
tmp = tmp & " @mUpdatedOn    = d.UpdatedDate,  @BillingCycleID = d.BillingCycleID, @BillingUnit =  d.BillingUnit, @Rate = d.Rate, @RateForDailyPnL = d.RateForDailyPnL,"
tmp = tmp & " @DepConDID = d.DepConDID , @NoofDays = d.NoOfDays"
tmp = tmp & "  FROM inserted i INNER JOIN deleted d ON i.GSLID = d.GSLID and i.SM_Prefix = d.SM_Prefix"
tmp = tmp & " insert into log_Mst_GSL (GSLID, SM_Prefix, Flag,  WHRBags, WHRWeight, CDTFBags, CDTFWeight, BalanceBags, BalanceWeight, Remarks, UpdatedBy, UpdatedOn,BillingCycleID,BillingUnit,Rate,RateForDailyPnL,DepConDID,NoofDays)"
tmp = tmp & " values (@mGSLID, @mSM_Prefix, @mFlag, @mWHRBags, @mWHRWeight, @mCDTFBags, @mCDTFWeight, @mBalanceBags, @mBalanceWeight, @mRemarks, @mUpdatedBy, @mUpdatedOn,@BillingCycleID,@BillingUnit,@Rate,@RateForDailyPnL,@DepConDID,@NoofDays)"
Call TmpTable

UpdateConfig ("322")
MsgBox "Version 322 Installed successfully"
End Function
Public Function Version321()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''-- Alter StateIndia for StateIndiaCode Code Length 2 to Length 3
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_InsuranceMasterReport.frm','Insurance Policy Master Report','FrmRpt_InsuranceMasterReport','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("321")
MsgBox "Version 321 Installed successfully"
End Function


Public Function Version320()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''-- Alter StateIndia for StateIndiaCode Code Length 2 to Length 3
tmp = " ALTER TABLE Mst_StateIndia ALTER COLUMN StateIndiaCode " & TextType & "(3) "
Call TmpTable

''--  Insert Menu for Insurance By Master
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_InsuranceBy.frm','Insurance By','FrmMst_InsuranceBy','Y','M',0,227,1"
Call TmpTable

UpdateConfig ("320")
MsgBox "Version 320 Installed successfully"
End Function

Public Function Version319()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Insert Menu for  Update Lessor Agreement Receive Flag
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_UpdateLessorAgreementFlag.frm','Update Lessor Agreement Receive Flag','FrmTxn_UpdateLessorAgreementFlag','Y','T',0,228,1"
Call TmpTable

''--  Insert Menu for N-Fin Pledge Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NFinPledge.frm','NFin Pledge Creation And Pledge Release Report','FrmRpt_NFinPledge','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("319")
MsgBox "Version 319 Installed successfully"

End Function


Public Function Version318()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''-- Create InsuranceRiskCovered Table
tmp = " Create Table Mst_InsuranceRiskCovered ( "
tmp = tmp & " InsuranceRiskCoveredID  Decimal(18,0) Not Null, "
tmp = tmp & " InsuranceRiskCovered    " & TextType & "(2000), "
tmp = tmp & " Status                  " & TextType & "(1), "
tmp = tmp & " G_UID                   " & TextType & "(36), "
tmp = tmp & " Remark                  " & TextType & "(2000), "
tmp = tmp & " Concurrency             Datetime, "
tmp = tmp & " CreatedBy               " & TextType & "(30), "
tmp = tmp & " CreatedDate             Datetime, "
tmp = tmp & " UpdatedBy               " & TextType & "(30), "
tmp = tmp & " UpdatedDate             DateTime ) "
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceRiskCovered WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_InsuranceRiskCovered PRIMARY KEY CLUSTERED (InsuranceRiskCoveredID),"
tmp = tmp & " CONSTRAINT DF_Mst_InsuranceRiskCovered DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable

''-- Insert Values in InsuranceRiskCovered
tmp = " Insert Into Mst_InsuranceRiskCovered (InsuranceRiskCoveredID,InsuranceRiskCovered,Status,G_UID) Select Isnull(MAX(InsuranceRiskCoveredID),0)+1,'STANDARD FIRE','A',NEWID() From Mst_InsuranceRiskCovered "
Call TmpTable
tmp = " Insert Into Mst_InsuranceRiskCovered (InsuranceRiskCoveredID,InsuranceRiskCovered,Status,G_UID) Select Isnull(MAX(InsuranceRiskCoveredID),0)+1,'BURGLARY','A',NEWID() From Mst_InsuranceRiskCovered "
Call TmpTable
tmp = " Insert Into Mst_InsuranceRiskCovered (InsuranceRiskCoveredID,InsuranceRiskCovered,Status,G_UID) Select Isnull(MAX(InsuranceRiskCoveredID),0)+1,'FIRE AND BURGLARY','A',NEWID() From Mst_InsuranceRiskCovered "
Call TmpTable
tmp = " Insert Into Mst_InsuranceRiskCovered (InsuranceRiskCoveredID,InsuranceRiskCovered,Status,G_UID) Select Isnull(MAX(InsuranceRiskCoveredID),0)+1,'STANDARD FIRE AND SPECIAL PERILS','A',NEWID() From Mst_InsuranceRiskCovered "
Call TmpTable

''-- DROP Relation With CDC Insurance & InsuranceDetails And Add InsurancePolicyID column
tmp = " ALTER TABLE Txn_CDC_Insurance_Detail DROP CONSTRAINT FK_Txn_CDC_Insurance_Detail_Mst_InsuranceDetails "
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceDetails DROP CONSTRAINT PK_Mst_InsuranceDetails "
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceDetails ADD InsurancePolicyID Decimal (18,0),DeleteDate Datetime ,DeleteBy Varchar(30) "
Call TmpTable

''-- Update InsurancePolicyID ID With Concurrency
tmp = " Update Mst_InsuranceDetails "
tmp = tmp & " Set Mst_InsuranceDetails.InsurancePolicyID = A.AutoNum "
tmp = tmp & " From "
tmp = tmp & " ( "
tmp = tmp & "     SELECT ROW_NUMBER() OVER (ORDER BY Concurrency) AS AutoNum,ISNULL(Concurrency,0)Concurrency "
tmp = tmp & "     From Mst_InsuranceDetails "
tmp = tmp & " ) A "
tmp = tmp & " Where IsNull(Mst_InsuranceDetails.Concurrency, 0) = A.Concurrency "
Call TmpTable

''-- Alter InsurancePolicyID as Not Null and Primary Key and Create Unique Columns to PolicyNo,CoverNoteNo
tmp = " ALTER TABLE Mst_InsuranceDetails ALTER Column InsurancePolicyID Decimal (18,0) NOT NULL "
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceDetails WITH NOCHECK ADD CONSTRAINT PK_Mst_InsuranceDetails PRIMARY KEY CLUSTERED (InsurancePolicyID) "
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceDetails ADD CONSTRAINT UC_PolicyNo_CoverNoteNo UNIQUE (PolicyNo,CoverNoteNo) "
Call TmpTable

''-- Alter Mst_InsuranceDetails table to Add InsuranceRiskCoveredID,InsuranceByID,Remark,StructureWiseLimit
tmp = " ALTER TABLE Mst_InsuranceDetails ADD InsuranceRiskCoveredID Decimal (18,0),InsuranceByID Decimal(18,0),Remark " & TextType & "(4000),StructureWiseLimit Decimal (18,2) "
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceDetails WITH NOCHECK ADD "
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceDetails_InsuranceRiskCoveredID FOREIGN KEY (InsuranceRiskCoveredID)    REFERENCES Mst_InsuranceRiskCovered (InsuranceRiskCoveredID), "
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceDetails_InsuranceByID          FOREIGN KEY (InsuranceByID)             REFERENCES Mst_InsuranceBy (InsuranceByID) "
Call TmpTable

''--->>>>>>>
''----- Mst_InsuranceForUnitTypeDetail
tmp = " Create Table Mst_InsuranceForUnitTypeDetail ("
tmp = tmp & " InsuranceForUnitTypeDetailID        decimal(18,0) Not null,"
tmp = tmp & " InsurancePolicyID                   decimal(18,0) Not null,"
tmp = tmp & " UnitTypeID                          decimal(18,0),"
tmp = tmp & " Flag                                " & TextType & "(1),"
tmp = tmp & " G_UID                               " & TextType & "(36),"
tmp = tmp & " Remarks                             " & TextType & "(1000),"
tmp = tmp & " Concurrency                         datetime,"
tmp = tmp & " CreatedDate                         datetime,"
tmp = tmp & " CreatedBy                           " & TextType & "(30),"
tmp = tmp & " UpdatedDate                         datetime,"
tmp = tmp & " UpdatedBy                           " & TextType & "(30))"
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceForUnitTypeDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_InsuranceForUnitTypeDetail PRIMARY KEY CLUSTERED (InsuranceForUnitTypeDetailID),"
tmp = tmp & " CONSTRAINT DF_Mst_InsuranceForUnitTypeDetail DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceForUnitTypeDetail_InsurancePolicyID  FOREIGN KEY (InsurancePolicyID) REFERENCES Mst_InsuranceDetails (InsurancePolicyID),"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceForUnitTypeDetail_UnitTypeID FOREIGN KEY (UnitTypeID)    REFERENCES Mst_UnitType(UnitTypeID)"
Call TmpTable

''----- Mst_InsuranceValidFor
tmp = " Create Table Mst_InsuranceValidFor("
tmp = tmp & " InsuranceValidForID               decimal(18,0) Not null,"
tmp = tmp & " InsurancePolicyID                 decimal(18,0) Not null,"
tmp = tmp & " ValidForType                      " & TextType & "(1),"
tmp = tmp & " ValidForTypeID                    decimal(18,0),"
tmp = tmp & " Flag                              " & TextType & "(1),"
tmp = tmp & " G_UID                             " & TextType & "(36),"
tmp = tmp & " Remarks                           " & TextType & "(1000),"
tmp = tmp & " Concurrency                       datetime,"
tmp = tmp & " CreatedDate                       datetime,"
tmp = tmp & " CreatedBy                         " & TextType & "(30),"
tmp = tmp & " UpdatedDate                       datetime,"
tmp = tmp & " UpdatedBy                         " & TextType & "(30))"
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceValidFor WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_InsuranceValidFor PRIMARY KEY CLUSTERED (InsuranceValidForID),"
tmp = tmp & " CONSTRAINT DF_Mst_InsuranceValidFor DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceValidFor_InsurancePolicyID   FOREIGN KEY (InsurancePolicyID) REFERENCES Mst_InsuranceDetails (InsurancePolicyID)"
Call TmpTable

''----- Mst_InsuranceGodownMapping
tmp = " Create Table Mst_InsuranceGodownMapping ("
tmp = tmp & " InsuranceGodownMappingID        decimal(18,0) Not null,"
tmp = tmp & " InsurancePolicyID               decimal(18,0) Not null,"
tmp = tmp & " GodownID                        decimal(18,0),"
tmp = tmp & " Connected                       " & TextType & "(1),"
tmp = tmp & " Flag                            " & TextType & "(1),"
tmp = tmp & " G_UID                           " & TextType & "(36),"
tmp = tmp & " Remarks                         " & TextType & "(1000),"
tmp = tmp & " Concurrency                     datetime,"
tmp = tmp & " CreatedDate                     datetime,"
tmp = tmp & " CreatedBy                       " & TextType & "(30),"
tmp = tmp & " UpdatedDate                     datetime,"
tmp = tmp & " UpdatedBy                       " & TextType & "(30))"
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceGodownMapping WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_InsuranceGodownMapping PRIMARY KEY CLUSTERED (InsuranceGodownMappingID),"
tmp = tmp & " CONSTRAINT DF_Mst_InsuranceGodownMapping DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceGodownMapping_InsurancePolicyID  FOREIGN KEY (InsurancePolicyID) REFERENCES Mst_InsuranceDetails (InsurancePolicyID),"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceGodownMapping_GodownID   FOREIGN KEY (GodownID)  REFERENCES Mst_Godown (GodownID)"
Call TmpTable

''----- Mst_InsuranceCommodityDetail
tmp = " Create Table Mst_InsuranceCommodityDetail("
tmp = tmp & " InsuranceCommodityDetailID      decimal(18,0) Not null,"
tmp = tmp & " InsurancePolicyID               decimal(18,0) Not null,"
tmp = tmp & " CommodityID                     decimal(18,0),"
tmp = tmp & " EntryType                       " & TextType & "(1),"
tmp = tmp & " Flag                            " & TextType & "(1),"
tmp = tmp & " G_UID                           " & TextType & "(36),"
tmp = tmp & " Remarks                         " & TextType & "(1000),"
tmp = tmp & " Concurrency                     datetime,"
tmp = tmp & " CreatedDate                     datetime,"
tmp = tmp & " CreatedBy                       " & TextType & "(30),"
tmp = tmp & " UpdatedDate                     datetime,"
tmp = tmp & " UpdatedBy                       " & TextType & "(30))"
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceCommodityDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_InsuranceCommodityDetail PRIMARY KEY CLUSTERED (InsuranceCommodityDetailID),"
tmp = tmp & " CONSTRAINT DF_Mst_InsuranceCommodityDetail DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceCommodityDetail_InsurancePolicyID    FOREIGN KEY (InsurancePolicyID) REFERENCES Mst_InsuranceDetails (InsurancePolicyID),"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceCommodityDetail_CommodityID  FOREIGN KEY (CommodityID)   REFERENCES Mst_Commodity(CommodityID)"
Call TmpTable

''----- Mst_InsuranceClientDetail
tmp = " Create Table Mst_InsuranceClientDetail("
tmp = tmp & " InsuranceClientDetailID       decimal(18,0) Not null,"
tmp = tmp & " InsurancePolicyID             decimal(18,0) Not null,"
tmp = tmp & " ClientIDRow                   decimal(18,0),"
tmp = tmp & " Flag                          " & TextType & "(1),"
tmp = tmp & " G_UID                         " & TextType & "(36),"
tmp = tmp & " Remarks                       " & TextType & "(1000),"
tmp = tmp & " Concurrency                   datetime,"
tmp = tmp & " CreatedDate                   datetime,"
tmp = tmp & " CreatedBy                     " & TextType & "(30),"
tmp = tmp & " UpdatedDate                   datetime,"
tmp = tmp & " UpdatedBy                     " & TextType & "(30))"
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceClientDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_InsuranceClientDetail PRIMARY KEY CLUSTERED (InsuranceClientDetailID),"
tmp = tmp & " CONSTRAINT DF_Mst_InsuranceClientDetail DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_InsuranceClientDetail_InsurancePolicyID   FOREIGN KEY (InsurancePolicyID) REFERENCES Mst_InsuranceDetails (InsurancePolicyID)"
Call TmpTable

''----- Log Insurance Details
tmp = " Create table Log_Mst_InsuranceDetails("
tmp = tmp & " LogID                       Decimal(18,0) IDENTITY(1,1),"
tmp = tmp & " InsurancePolicyID           Decimal(18,0) NOT NULL,"
tmp = tmp & " PolicyNo                    " & TextType & "(255),"
tmp = tmp & " CoverNoteNo                 " & TextType & "(255),"
tmp = tmp & " InsuranceCoID               Decimal(18,0),"
tmp = tmp & " Risks                       " & TextType & "(255),"
tmp = tmp & " Amount                      Decimal(18,2),"
tmp = tmp & " ValidFrom                   Smalldatetime,"
tmp = tmp & " ValidTo                     Smalldatetime,"
tmp = tmp & " Flag                        " & TextType & "(1),"
tmp = tmp & " InsuranceRiskCoveredID      Decimal(18,0),"
tmp = tmp & " InsuranceByID               Decimal(18,0),"
tmp = tmp & " StructureWiseLimit          Decimal(18,2),"
tmp = tmp & " Remark                      " & TextType & "(4000),"
tmp = tmp & " G_UID                       " & TextType & "(36),"
tmp = tmp & " Concurrency                 Datetime,"
tmp = tmp & " CreatedBy                   " & TextType & "(30),"
tmp = tmp & " CreatedDate                 Datetime,"
tmp = tmp & " UpdatedDate                 Datetime,"
tmp = tmp & " UpdatedBy                   " & TextType & "(30),"
tmp = tmp & " DeleteDate                  Datetime,"
tmp = tmp & " DeleteBy                    " & TextType & "(30),"
tmp = tmp & " LogConcurrency              Datetime,"
tmp = tmp & " LogFlag                     " & TextType & "(1))"
Call TmpTable

''----- Log InsuranceGodownMapping
tmp = " Create Table Log_Mst_InsuranceGodownMapping ("
tmp = tmp & " LogID                           Decimal(18,0) IDENTITY(1,1),"
tmp = tmp & " InsuranceGodownMappingID        Decimal(18,0) Not null,"
tmp = tmp & " InsurancePolicyID               Decimal(18,0) Not null,"
tmp = tmp & " GodownID                        Decimal(18,0),"
tmp = tmp & " Connected                       " & TextType & "(1),"
tmp = tmp & " Flag                            " & TextType & "(1),"
tmp = tmp & " G_UID                           " & TextType & "(36),"
tmp = tmp & " Remarks                         " & TextType & "(1000),"
tmp = tmp & " Concurrency                     Datetime,"
tmp = tmp & " CreatedDate                     Datetime,"
tmp = tmp & " CreatedBy                       " & TextType & "(30),"
tmp = tmp & " UpdatedDate                     Datetime,"
tmp = tmp & " UpdatedBy                       " & TextType & "(30),"
tmp = tmp & " LogConcurrency                  Datetime,"
tmp = tmp & " LogFlag                         " & TextType & "(1))"
Call TmpTable


''---- Create Trigger for Log_Mst_InsuranceDetails ON Insert / UPDATE / Delete
tmp = " Create TRIGGER Tr_Log_Mst_InsuranceDetails ON Mst_InsuranceDetails"
tmp = tmp & " After Insert, Update, Delete"
tmp = tmp & " AS"
tmp = tmp & " If exists (Select * from inserted) and not exists(Select * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Mst_InsuranceDetails(InsurancePolicyID, PolicyNo, CoverNoteNo, InsuranceCoID, Risks, Amount, ValidFrom, ValidTo, Flag, InsuranceRiskCoveredID, InsuranceByID, StructureWiseLimit, Remark, G_UID, Concurrency, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, DeleteDate, DeleteBy, LogConcurrency, LogFlag)"
tmp = tmp & "     SELECT InsurancePolicyID,PolicyNo,CoverNoteNo,InsuranceCoID,Risks,Amount,ValidFrom,ValidTo,Flag,InsuranceRiskCoveredID,InsuranceByID,StructureWiseLimit,Remark,G_UID,Concurrency,CreatedBy,CreatedDate,UpdatedDate,UpdatedBy,DeleteDate,DeleteBy,GETDATE(),'I' FROM inserted"
tmp = tmp & " End"
tmp = tmp & " if exists(SELECT * from inserted) and exists (SELECT * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Mst_InsuranceDetails(InsurancePolicyID, PolicyNo, CoverNoteNo, InsuranceCoID, Risks, Amount, ValidFrom, ValidTo, Flag, InsuranceRiskCoveredID, InsuranceByID, StructureWiseLimit, Remark, G_UID, Concurrency, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, DeleteDate, DeleteBy, LogConcurrency, LogFlag)"
tmp = tmp & "     SELECT i.InsurancePolicyID,i.PolicyNo,i.CoverNoteNo,i.InsuranceCoID,i.Risks,i.Amount,i.ValidFrom,i.ValidTo,i.Flag,i.InsuranceRiskCoveredID,i.InsuranceByID,i.StructureWiseLimit,i.Remark,i.G_UID,i.Concurrency,i.CreatedBy,i.CreatedDate,i.UpdatedDate,i.UpdatedBy,i.DeleteDate,i.DeleteBy,GETDATE(),'U'"
tmp = tmp & "     FROM inserted i INNER JOIN deleted d ON i.InsurancePolicyID = d.InsurancePolicyID"
tmp = tmp & " End"
tmp = tmp & " If exists(select * from deleted) and not exists(Select * from inserted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Mst_InsuranceDetails(InsurancePolicyID, PolicyNo, CoverNoteNo, InsuranceCoID, Risks, Amount, ValidFrom, ValidTo, Flag, InsuranceRiskCoveredID, InsuranceByID, StructureWiseLimit, Remark, G_UID, Concurrency, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, DeleteDate, DeleteBy, LogConcurrency, LogFlag)"
tmp = tmp & "     SELECT InsurancePolicyID,PolicyNo,CoverNoteNo,InsuranceCoID,Risks,Amount,ValidFrom,ValidTo,Flag,InsuranceRiskCoveredID,InsuranceByID,StructureWiseLimit,Remark,G_UID,Concurrency,CreatedBy,CreatedDate,UpdatedDate,UpdatedBy,DeleteDate,DeleteBy,GETDATE(),'D' FROM deleted"
tmp = tmp & " End"
Call TmpTable

''---- Create Trigger for Log_Mst_InsuranceGodownMapping ON Insert / UPDATE / Delete
tmp = " Create TRIGGER Tr_Log_Mst_InsuranceGodownMapping ON Mst_InsuranceGodownMapping"
tmp = tmp & " After Insert, Update, Delete"
tmp = tmp & " AS"
tmp = tmp & " If exists (Select * from inserted) and not exists(Select * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Mst_InsuranceGodownMapping(InsuranceGodownMappingID, InsurancePolicyID, GodownID, Connected, Flag, G_UID, Remarks, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy, LogConcurrency, LogFlag)"
tmp = tmp & "     SELECT InsuranceGodownMappingID,InsurancePolicyID,GodownID,Connected,Flag,G_UID,Remarks,Concurrency,CreatedDate,CreatedBy,UpdatedDate,UpdatedBy,GETDATE(),'I' FROM inserted"
tmp = tmp & " End"
tmp = tmp & " if exists(SELECT * from inserted) and exists (SELECT * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Mst_InsuranceGodownMapping(InsuranceGodownMappingID, InsurancePolicyID, GodownID, Connected, Flag, G_UID, Remarks, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy, LogConcurrency, LogFlag)"
tmp = tmp & "     SELECT i.InsuranceGodownMappingID,i.InsurancePolicyID,i.GodownID,i.Connected,i.Flag,i.G_UID,i.Remarks,i.Concurrency,i.CreatedDate,i.CreatedBy,i.UpdatedDate,i.UpdatedBy,GETDATE(),'U'"
tmp = tmp & "     FROM inserted i INNER JOIN deleted d ON i.InsuranceGodownMappingID = d.InsuranceGodownMappingID"
tmp = tmp & " End"
tmp = tmp & " If exists(select * from deleted) and not exists(Select * from inserted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Mst_InsuranceGodownMapping(InsuranceGodownMappingID, InsurancePolicyID, GodownID, Connected, Flag, G_UID, Remarks, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy, LogConcurrency, LogFlag)"
tmp = tmp & "     SELECT InsuranceGodownMappingID,InsurancePolicyID,GodownID,Connected,Flag,G_UID,Remarks,Concurrency,CreatedDate,CreatedBy,UpdatedDate,UpdatedBy,GETDATE(),'D' FROM deleted"
tmp = tmp & " End"
Call TmpTable

''--  Insert Menu for GSL Pending Report for DC
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_InsuranceRiskCovered.frm','Insurance Risk Covered','FrmMst_InsuranceRiskCovered','Y','M',0,227,1"
Call TmpTable


UpdateConfig ("318")
MsgBox "Version 318 Installed successfully"

End Function


Public Function Version317()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Insert Menu for GSL Pending Report for DC
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_PendingGSLforDC.frm','PendingGSLforRateCard','FrmRpt_PendingGSLforDC','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("317")
MsgBox "Version 317 Installed successfully"

End Function

Public Function Version316()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Txn_DepositContract Add ContractType " & TextType & "(1)"
Call TmpTable

tmp = " Alter Table Log_Txn_DepositContract Add Verified " & TextType & "(1),VerifiedBy " & TextType & "(30),VerifiedDate DATETIME,ContractType " & TextType & "(1)"
Call TmpTable

tmp = " ALTER TRIGGER [dbo].[Tr_Log_Txn_DepositContract] ON [dbo].[Txn_DepositContract]"
tmp = tmp & " After Insert, Update, Delete AS"
tmp = tmp & " If exists (Select * from inserted) and not exists(Select * from deleted)"
tmp = tmp & "  Begin"
tmp = tmp & "       Insert Log_Txn_DepositContract(DepConID, Date, StateID, LocationID, ClientID, CommodityID, QuantumMT, StartDate, EndDate, BillingCycleID, DepBillingCycleID, WBillingCycleID, RatePer, TentativeDur, LockPeriod, EmployeeId, InsuranceByID, DepMonMTRate, BagSizeFlag, SM_Prefix, ApprovalFlag, ApprovalRemarks, ApproveOn, ApproveBy, ClosedRemark, ClosedOn, ClosedBy, G_UID, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy,Verified,VerifiedBy,VerifiedDate,ContractType, LogFlag)"
tmp = tmp & "       SELECT DepConID,Date,StateID,LocationID,ClientID,CommodityID,QuantumMT,StartDate,EndDate,BillingCycleID,DepBillingCycleID,WBillingCycleID,RatePer,TentativeDur,LockPeriod,EmployeeId,InsuranceByID,DepMonMTRate,BagSizeFlag,SM_Prefix,ApprovalFlag,ApprovalRemarks,ApproveOn,ApproveBy,ClosedRemark,ClosedOn,ClosedBy,G_UID,Concurrency,CreatedDate,CreatedBy,UpdatedDate,UpdatedBy,Verified,VerifiedBy,VerifiedDate, ContractType, 'I' FROM inserted"
tmp = tmp & "  End"
tmp = tmp & " If exists(SELECT * from inserted) and exists (SELECT * from deleted)"
tmp = tmp & "  Begin"
tmp = tmp & "       Insert Log_Txn_DepositContract(DepConID, Date, StateID, LocationID, ClientID, CommodityID, QuantumMT, StartDate, EndDate, BillingCycleID, DepBillingCycleID, WBillingCycleID, RatePer, TentativeDur, LockPeriod, EmployeeId, InsuranceByID, DepMonMTRate, BagSizeFlag, SM_Prefix, ApprovalFlag, ApprovalRemarks, ApproveOn, ApproveBy, ClosedRemark, ClosedOn, ClosedBy, G_UID, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy,Verified,VerifiedBy,VerifiedDate, ContractType, LogFlag)"
tmp = tmp & "       SELECT i.DepConID,i.Date,i.StateID,i.LocationID,i.ClientID,i.CommodityID,i.QuantumMT,i.StartDate,i.EndDate,i.BillingCycleID,i.DepBillingCycleID,i.WBillingCycleID,i.RatePer,i.TentativeDur,i.LockPeriod,i.EmployeeId,i.InsuranceByID,i.DepMonMTRate,i.BagSizeFlag,i.SM_Prefix,i.ApprovalFlag,i.ApprovalRemarks,i.ApproveOn,i.ApproveBy,i.ClosedRemark,i.ClosedOn,i.ClosedBy,i.G_UID,i.Concurrency,i.CreatedDate,i.CreatedBy,i.UpdatedDate,i.UpdatedBy,i.Verified,i.VerifiedBy,i.VerifiedDate, i.ContractType, 'U' FROM inserted i INNER JOIN deleted d ON i.DepConID = d.DepConID"
tmp = tmp & "  End"
tmp = tmp & " If exists(select * from deleted) and not exists(Select * from inserted)"
tmp = tmp & "  Begin"
tmp = tmp & "       Insert Log_Txn_DepositContract(DepConID, Date, StateID, LocationID, ClientID, CommodityID, QuantumMT, StartDate, EndDate, BillingCycleID, DepBillingCycleID, WBillingCycleID, RatePer, TentativeDur, LockPeriod, EmployeeId, InsuranceByID, DepMonMTRate, BagSizeFlag, SM_Prefix, ApprovalFlag, ApprovalRemarks, ApproveOn, ApproveBy, ClosedRemark, ClosedOn, ClosedBy, G_UID, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy,Verified,VerifiedBy,VerifiedDate, ContractType, LogFlag)"
tmp = tmp & "       SELECT DepConID,Date,StateID,LocationID,ClientID,CommodityID,QuantumMT,StartDate,EndDate,BillingCycleID,DepBillingCycleID,WBillingCycleID,RatePer,TentativeDur,LockPeriod,EmployeeId,InsuranceByID,DepMonMTRate,BagSizeFlag,SM_Prefix,ApprovalFlag,ApprovalRemarks,ApproveOn,ApproveBy,ClosedRemark,ClosedOn,ClosedBy,G_UID,Concurrency,CreatedDate,CreatedBy,UpdatedDate,UpdatedBy,Verified,VerifiedBy,VerifiedDate, ContractType, 'D' FROM deleted"
tmp = tmp & "  End"
Call TmpTable

tmp = "Update Txn_DepositContract Set ContractType = 'N'"
Call TmpTable


UpdateConfig ("316")
MsgBox "Version 316 Installed successfully"

End Function

Public Function version315()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If
 
If Gen_DBType <> "MDB" Then
   tmp = " Alter Table Txn_DepositContract add Verified VarChar(1),VerifiedBy VarChar(30),VerifiedDate Datetime "
   Call TmpTable
End If

UpdateConfig ("315")
MsgBox "Version 315 Installed successfully"

End Function


Public Function Version314()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Insert Menu for State Wise GST Data Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_DepositContractChangeFlag.frm','Deposit Contract Change Flag','FrmTxn_DepositContractChangeFlag','Y','T',0,228,1"
Call TmpTable

UpdateConfig ("314")
MsgBox "Version 314 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function



Public Function Version313()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create Table Mst_InsuranceBy("
tmp = tmp & " InsuranceByID   decimal(18,0) Not null,"
tmp = tmp & " InsuranceByName varchar(100),"
tmp = tmp & " Flag            varchar(1),"
tmp = tmp & " G_UID           varchar(36),"
tmp = tmp & " Remarks         varchar(1000),"
tmp = tmp & " Concurrency     datetime,"
tmp = tmp & " CreatedDate     datetime,"
tmp = tmp & " CreatedBy       varchar(30),"
tmp = tmp & " UpdatedDate     datetime,"
tmp = tmp & " UpdatedBy       varchar(30))"
Call TmpTable

tmp = " ALTER TABLE Mst_InsuranceBy WITH NOCHECK ADD CONSTRAINT PK_Mst_InsuranceBy PRIMARY KEY CLUSTERED (InsuranceByID), CONSTRAINT DF_Mst_InsuranceBy DEFAULT (GETDATE()) FOR Concurrency "
Call TmpTable


tmp = " Create Table Txn_DepositContract("
tmp = tmp & " DepConID            Decimal(18,0) NOT NULL,"
tmp = tmp & " Date                Datetime,"
tmp = tmp & " StateID             Decimal(18,0),"
tmp = tmp & " LocationID          Decimal(18,0),"
tmp = tmp & " ClientID            Decimal(18,0),"
tmp = tmp & " CommodityID         Decimal(18,0),"
tmp = tmp & " QuantumMT           Decimal(18,0),"
tmp = tmp & " StartDate           Datetime,"
tmp = tmp & " EndDate             Datetime,"
tmp = tmp & " BillingCycleID      Decimal(18,0),"
tmp = tmp & " DepBillingCycleID   Decimal(18,0),"
tmp = tmp & " WBillingCycleID     Decimal(18,0),"
tmp = tmp & " RatePer             Varchar(1),"
tmp = tmp & " TentativeDur        Decimal(18,0),"
tmp = tmp & " LockPeriod          Decimal(18,0),"
tmp = tmp & " EmployeeId          Decimal(18,0),"
tmp = tmp & " InsuranceByID       Decimal(18,0),"
tmp = tmp & " DepMonMTRate        Decimal(18,4),"
tmp = tmp & " BagSizeFlag         Varchar(1),"
tmp = tmp & " SM_Prefix           Varchar(5),"
tmp = tmp & " ApprovalFlag        Varchar(1),"
tmp = tmp & " ApprovalRemarks     Varchar(4000),"
tmp = tmp & " ApproveOn           Datetime,"
tmp = tmp & " ApproveBy           Varchar(30),"
tmp = tmp & " ClosedRemark        Varchar(8000),"
tmp = tmp & " ClosedOn            Datetime,"
tmp = tmp & " ClosedBy            Varchar(30),"
tmp = tmp & " G_UID               Varchar(36),"
tmp = tmp & " Concurrency         Datetime,"
tmp = tmp & " CreatedDate         Datetime,"
tmp = tmp & " CreatedBy           Varchar(30),"
tmp = tmp & " UpdatedDate         Datetime,"
tmp = tmp & " UpdatedBy VarChar(30))"
Call TmpTable

tmp = " ALTER TABLE Txn_DepositContract WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_DepositContract                   PRIMARY KEY CLUSTERED (DepConID),"
tmp = tmp & " CONSTRAINT DF_Txn_DepositContract                   DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_StateID           FOREIGN KEY (StateID)           REFERENCES Mst_State (StateID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_LocationID        FOREIGN KEY (LocationID)        REFERENCES Mst_Location (LocationID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_Mst_Commodity     FOREIGN KEY (CommodityID)       REFERENCES Mst_Commodity (CommodityID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_BillingCycleID    FOREIGN KEY (BillingCycleID)    REFERENCES Mst_BillingCycle (BillingCycleID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_DepBillingCycleID FOREIGN KEY (DepBillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_WBillingCycleID   FOREIGN KEY (WBillingCycleID)   REFERENCES Mst_BillingCycle (BillingCycleID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_EmployeeId        FOREIGN KEY (EmployeeId)        REFERENCES Mst_Employee (EmployeeId),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContract_InsuranceByID     FOREIGN KEY (InsuranceByID)     REFERENCES Mst_InsuranceBy (InsuranceByID)"
Call TmpTable


tmp = " Create table Log_Txn_DepositContract("
tmp = tmp & " LogDepConID         Decimal(18,0) IDENTITY(1,1),"
tmp = tmp & " DepConID            Decimal(18,0) Not null,"
tmp = tmp & " Date                Datetime,"
tmp = tmp & " StateID             Decimal(18,0),"
tmp = tmp & " LocationID          Decimal(18,0),"
tmp = tmp & " ClientID            Decimal(18,0),"
tmp = tmp & " CommodityID         Decimal(18,0),"
tmp = tmp & " QuantumMT           Decimal(18,0),"
tmp = tmp & " StartDate           Datetime,"
tmp = tmp & " EndDate             Datetime,"
tmp = tmp & " BillingCycleID      Decimal(18,0),"
tmp = tmp & " DepBillingCycleID   Decimal(18,0),"
tmp = tmp & " WBillingCycleID     Decimal(18,0),"
tmp = tmp & " RatePer             Varchar(1),"
tmp = tmp & " TentativeDur        Decimal(18,0),"
tmp = tmp & " LockPeriod          Decimal(18,0),"
tmp = tmp & " EmployeeId          Decimal(18,0),"
tmp = tmp & " InsuranceByID       Decimal(18,0),"
tmp = tmp & " DepMonMTRate        Decimal(18,4),"
tmp = tmp & " BagSizeFlag         Varchar(1),"
tmp = tmp & " SM_Prefix           Varchar(5),"
tmp = tmp & " ApprovalFlag        Varchar(1),"
tmp = tmp & " ApprovalRemarks     Varchar(4000),"
tmp = tmp & " ApproveOn           Datetime,"
tmp = tmp & " ApproveBy           Varchar(30),"
tmp = tmp & " ClosedRemark        VARCHAR(8000),"
tmp = tmp & " ClosedOn            DATETIME,"
tmp = tmp & " ClosedBy            VARCHAR(30),"
tmp = tmp & " G_UID               Varchar(36),"
tmp = tmp & " Concurrency         Datetime,"
tmp = tmp & " CreatedDate         Datetime,"
tmp = tmp & " CreatedBy           Varchar(30),"
tmp = tmp & " UpdatedDate         Datetime,"
tmp = tmp & " UpdatedBy           Varchar(30),"
tmp = tmp & " LogFlag VarChar(1))"
Call TmpTable

''---- Create Trigger for Log_Txn_DepositContract ON Insert / UPDATE / Delete
tmp = " Create TRIGGER Tr_Log_Txn_DepositContract ON Txn_DepositContract"
tmp = tmp & " After Insert, Update, Delete"
tmp = tmp & " AS"
tmp = tmp & " If exists (Select * from inserted) and not exists(Select * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Txn_DepositContract(DepConID, Date, StateID, LocationID, ClientID, CommodityID, QuantumMT, StartDate, EndDate, BillingCycleID, DepBillingCycleID, WBillingCycleID, RatePer, TentativeDur, LockPeriod, EmployeeId, InsuranceByID, DepMonMTRate, BagSizeFlag, SM_Prefix, ApprovalFlag, ApprovalRemarks, ApproveOn, ApproveBy, ClosedRemark, ClosedOn, ClosedBy, G_UID, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy, LogFlag)"
tmp = tmp & "     SELECT DepConID,Date,StateID,LocationID,ClientID,CommodityID,QuantumMT,StartDate,EndDate,BillingCycleID,DepBillingCycleID,WBillingCycleID,RatePer,TentativeDur,LockPeriod,EmployeeId,InsuranceByID,DepMonMTRate,BagSizeFlag,SM_Prefix,ApprovalFlag,ApprovalRemarks,ApproveOn,ApproveBy,ClosedRemark,ClosedOn,ClosedBy,G_UID,Concurrency,CreatedDate,CreatedBy,UpdatedDate,UpdatedBy,'I' FROM inserted"
tmp = tmp & " End"
tmp = tmp & " if exists(SELECT * from inserted) and exists (SELECT * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Txn_DepositContract(DepConID, Date, StateID, LocationID, ClientID, CommodityID, QuantumMT, StartDate, EndDate, BillingCycleID, DepBillingCycleID, WBillingCycleID, RatePer, TentativeDur, LockPeriod, EmployeeId, InsuranceByID, DepMonMTRate, BagSizeFlag, SM_Prefix, ApprovalFlag, ApprovalRemarks, ApproveOn, ApproveBy, ClosedRemark, ClosedOn, ClosedBy, G_UID, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy, LogFlag)"
tmp = tmp & "     SELECT i.DepConID,i.Date,i.StateID,i.LocationID,i.ClientID,i.CommodityID,i.QuantumMT,i.StartDate,i.EndDate,i.BillingCycleID,i.DepBillingCycleID,i.WBillingCycleID,i.RatePer,i.TentativeDur,i.LockPeriod,i.EmployeeId,i.InsuranceByID,i.DepMonMTRate,i.BagSizeFlag,i.SM_Prefix,i.ApprovalFlag,i.ApprovalRemarks,i.ApproveOn,i.ApproveBy,i.ClosedRemark,i.ClosedOn,i.ClosedBy,i.G_UID,i.Concurrency,i.CreatedDate,i.CreatedBy,i.UpdatedDate,i.UpdatedBy,'U'"
tmp = tmp & "     FROM inserted i INNER JOIN deleted d ON i.DepConID = d.DepConID"
tmp = tmp & " End"
tmp = tmp & " If exists(select * from deleted) and not exists(Select * from inserted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Txn_DepositContract(DepConID, Date, StateID, LocationID, ClientID, CommodityID, QuantumMT, StartDate, EndDate, BillingCycleID, DepBillingCycleID, WBillingCycleID, RatePer, TentativeDur, LockPeriod, EmployeeId, InsuranceByID, DepMonMTRate, BagSizeFlag, SM_Prefix, ApprovalFlag, ApprovalRemarks, ApproveOn, ApproveBy, ClosedRemark, ClosedOn, ClosedBy, G_UID, Concurrency, CreatedDate, CreatedBy, UpdatedDate, UpdatedBy, LogFlag)"
tmp = tmp & "     SELECT DepConID,Date,StateID,LocationID,ClientID,CommodityID,QuantumMT,StartDate,EndDate,BillingCycleID,DepBillingCycleID,WBillingCycleID,RatePer,TentativeDur,LockPeriod,EmployeeId,InsuranceByID,DepMonMTRate,BagSizeFlag,SM_Prefix,ApprovalFlag,ApprovalRemarks,ApproveOn,ApproveBy,ClosedRemark,ClosedOn,ClosedBy,G_UID,Concurrency,CreatedDate,CreatedBy,UpdatedDate,UpdatedBy,'D' FROM deleted"
tmp = tmp & " End"
Call TmpTable


tmp = " Create Table Txn_DepositContractGodownDetail ("
tmp = tmp & " DepConDID                   Decimal(18,0) Not Null,"
tmp = tmp & " DepConID                    Decimal(18,0),"
tmp = tmp & " GodownId                    Decimal(18,0),"
tmp = tmp & " FinalPrice                  Decimal(18,4),"
tmp = tmp & " DepositFinalPrice           Decimal(18,4),"
tmp = tmp & " WithdrawalFinalPrice        Decimal(18,4),"
tmp = tmp & " BaseRate                    Decimal(18,4),"
tmp = tmp & " DepositBaseRate             Decimal(18,4),"
tmp = tmp & " WithdrawalBaseRate          Decimal(18,4),"
tmp = tmp & " OldBillingRate              Decimal(18,4),"
tmp = tmp & " OldDepositBillingRate       Decimal(18,4),"
tmp = tmp & " OldWithdrawalBillingRate    Decimal(18,4),"
tmp = tmp & " MinBagSize                  Decimal(18,4),"
tmp = tmp & " MaxBagSize                  Decimal(18,4),"
tmp = tmp & " BaseRateColor               Varchar(10),"
tmp = tmp & " DepositBaseRateColor        Varchar(10),"
tmp = tmp & " WithdrawalBaseRateColor     Varchar(10),"
tmp = tmp & " SM_Prefix                   Varchar(5),"
tmp = tmp & " RatePerMonthPerMT           Decimal(18,4),"
tmp = tmp & " Rent                        Decimal(18,4),"
tmp = tmp & " G_UID                       Varchar(36),"
tmp = tmp & " Concurrency                 Datetime,"
tmp = tmp & " CreatedBy                   Varchar(30),"
tmp = tmp & " CreatedDate                 Datetime,"
tmp = tmp & " UpdatedBy                   Varchar(30),"
tmp = tmp & " UpdatedDate DateTime)"
Call TmpTable

tmp = " ALTER TABLE Txn_DepositContractGodownDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_DepositContractGodownDetail PRIMARY KEY CLUSTERED (DepConDID),"
tmp = tmp & " CONSTRAINT DF_Txn_DepositContractGodownDetail DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContractGodownDetail_DepConID FOREIGN KEY (DepConID) REFERENCES Txn_DepositContract (DepConID),"
tmp = tmp & " CONSTRAINT FK_Txn_DepositContractGodownDetail_GodownId FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID)"
Call TmpTable


tmp = " Create Table Log_Txn_DepositContractGodownDetail("
tmp = tmp & " LogDepConDID                Decimal(18,0) IDENTITY(1,1),"
tmp = tmp & " DepConDID                   Decimal(18,0),"
tmp = tmp & " DepConID                    Decimal(18,0),"
tmp = tmp & " GodownId                    Decimal(18,0),"
tmp = tmp & " FinalPrice                  Decimal(18,4),"
tmp = tmp & " DepositFinalPrice           Decimal(18,4),"
tmp = tmp & " WithdrawalFinalPrice        Decimal(18,4),"
tmp = tmp & " BaseRate                    Decimal(18,4),"
tmp = tmp & " DepositBaseRate             Decimal(18,4),"
tmp = tmp & " WithdrawalBaseRate          Decimal(18,4),"
tmp = tmp & " OldBillingRate              Decimal(18,4),"
tmp = tmp & " OldDepositBillingRate       Decimal(18,4),"
tmp = tmp & " OldWithdrawalBillingRate    Decimal(18,4),"
tmp = tmp & " MinBagSize                  Decimal(18,4),"
tmp = tmp & " MaxBagSize                  Decimal(18,4),"
tmp = tmp & " BaseRateColor               Varchar(10),"
tmp = tmp & " DepositBaseRateColor        Varchar(10),"
tmp = tmp & " WithdrawalBaseRateColor     Varchar(10),"
tmp = tmp & " SM_Prefix                   Varchar(5),"
tmp = tmp & " RatePerMonthPerMT           Decimal(18,4),"
tmp = tmp & " Rent                        Decimal(18,4),"
tmp = tmp & " Log_Concurrency             Datetime,"
tmp = tmp & " G_UID                       Varchar(36),"
tmp = tmp & " Concurrency                 Datetime,"
tmp = tmp & " CreatedBy                   Varchar(30),"
tmp = tmp & " CreatedDate                 Datetime,"
tmp = tmp & " UpdatedBy                   Varchar(30),"
tmp = tmp & " UpdatedDate                 Datetime,"
tmp = tmp & " LogFlag VarChar(1))"
Call TmpTable

tmp = " ALTER TABLE Log_Txn_DepositContractGodownDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Log_Txn_DepositContractGodownDetail PRIMARY KEY CLUSTERED (LogDepConDID),"
tmp = tmp & " CONSTRAINT DF_Log_Txn_DepositContractGodownDetail DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable


''---- Create Trigger for Log_DepositContractGodownDetail ON Insert / UPDATE / Delete
tmp = " Create TRIGGER Tr_Log_Txn_DepositContractGodownDetail ON Txn_DepositContractGodownDetail"
tmp = tmp & " After Insert, Update, Delete"
tmp = tmp & " AS"
tmp = tmp & " If exists (Select * from inserted) and not exists(Select * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Txn_DepositContractGodownDetail(DepConDID, DepConID, GodownId, FinalPrice, DepositFinalPrice, WithdrawalFinalPrice, BaseRate, DepositBaseRate, WithdrawalBaseRate, OldBillingRate, OldDepositBillingRate, OldWithdrawalBillingRate, MinBagSize, MaxBagSize, BaseRateColor, DepositBaseRateColor, WithdrawalBaseRateColor, SM_Prefix, RatePerMonthPerMT, Rent, Log_Concurrency, G_UID, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, LogFlag)"
tmp = tmp & "     SELECT DepConDID,DepConID,GodownId,FinalPrice,DepositFinalPrice,WithdrawalFinalPrice,BaseRate,DepositBaseRate,WithdrawalBaseRate,OldBillingRate,OldDepositBillingRate,OldWithdrawalBillingRate,MinBagSize,MaxBagSize,BaseRateColor,DepositBaseRateColor,WithdrawalBaseRateColor,SM_Prefix,RatePerMonthPerMT,Rent,Concurrency,G_UID,CreatedBy,CreatedDate,UpdatedBy,UpdatedDate,'I' FROM inserted"
tmp = tmp & " End"
tmp = tmp & " if exists(SELECT * from inserted) and exists (SELECT * from deleted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Txn_DepositContractGodownDetail(DepConDID, DepConID, GodownId, FinalPrice, DepositFinalPrice, WithdrawalFinalPrice, BaseRate, DepositBaseRate, WithdrawalBaseRate, OldBillingRate, OldDepositBillingRate, OldWithdrawalBillingRate, MinBagSize, MaxBagSize, BaseRateColor, DepositBaseRateColor, WithdrawalBaseRateColor, SM_Prefix, RatePerMonthPerMT, Rent, Log_Concurrency, G_UID, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, LogFlag)"
tmp = tmp & "     SELECT i.DepConDID,i.DepConID,i.GodownId,i.FinalPrice,i.DepositFinalPrice,i.WithdrawalFinalPrice,i.BaseRate,i.DepositBaseRate,i.WithdrawalBaseRate,i.OldBillingRate,i.OldDepositBillingRate,i.OldWithdrawalBillingRate,i.MinBagSize,i.MaxBagSize,i.BaseRateColor,i.DepositBaseRateColor,i.WithdrawalBaseRateColor,i.SM_Prefix,i.RatePerMonthPerMT,i.Rent,i.Concurrency,i.G_UID,i.CreatedBy,i.CreatedDate,i.UpdatedBy,i.UpdatedDate,'U'"
tmp = tmp & "     FROM inserted i INNER JOIN deleted d ON i.DepConID = d.DepConID"
tmp = tmp & " End"
tmp = tmp & " If exists(select * from deleted) and not exists(Select * from inserted)"
tmp = tmp & " begin"
tmp = tmp & "     Insert Log_Txn_DepositContractGodownDetail(DepConDID, DepConID, GodownId, FinalPrice, DepositFinalPrice, WithdrawalFinalPrice, BaseRate, DepositBaseRate, WithdrawalBaseRate, OldBillingRate, OldDepositBillingRate, OldWithdrawalBillingRate, MinBagSize, MaxBagSize, BaseRateColor, DepositBaseRateColor, WithdrawalBaseRateColor, SM_Prefix, RatePerMonthPerMT, Rent, Log_Concurrency, G_UID, CreatedBy, CreatedDate, UpdatedBy, UpdatedDate, LogFlag)"
tmp = tmp & "     SELECT DepConDID,DepConID,GodownId,FinalPrice,DepositFinalPrice,WithdrawalFinalPrice,BaseRate,DepositBaseRate,WithdrawalBaseRate,OldBillingRate,OldDepositBillingRate,OldWithdrawalBillingRate,MinBagSize,MaxBagSize,BaseRateColor,DepositBaseRateColor,WithdrawalBaseRateColor,SM_Prefix,RatePerMonthPerMT,Rent,Concurrency,G_UID,CreatedBy,CreatedDate,UpdatedBy,UpdatedDate,'D' FROM deleted"
tmp = tmp & " End"
Call TmpTable

tmp = " Insert Into Mst_InsuranceBy(InsuranceByID,InsuranceByName,Flag,G_UID,CreatedDate,CreatedBy,Remarks) SELECT ISNULL(MAX(InsuranceByID),0)+1,'NCML','A',NEWID(),GETDATE(),'2047','' FROM Mst_InsuranceBy"
Call TmpTable
tmp = " Insert Into Mst_InsuranceBy(InsuranceByID,InsuranceByName,Flag,G_UID,CreatedDate,CreatedBy,Remarks) SELECT ISNULL(MAX(InsuranceByID),0)+1,'Client','A',NEWID(),GETDATE(),'2047','' FROM Mst_InsuranceBy"
Call TmpTable

''---- Mst_BillingCycle for Add New Ratio Column
tmp = " Alter Table Mst_BillingCycle Add Ratio decimal(18,2)"
Call TmpTable

tmp = " UPDATE Mst_BillingCycle SET Ratio = 15 ,Flag = 'N' Where BillingCycleID = 1 "
Call TmpTable

tmp = " UPDATE Mst_BillingCycle SET Ratio = 10 ,Flag = 'N' Where BillingCycleID = 2 "
Call TmpTable

tmp = " UPDATE Mst_BillingCycle SET Ratio = 5 ,Flag = 'N' Where BillingCycleID = 3 "
Call TmpTable

tmp = " UPDATE Mst_BillingCycle SET Ratio = 0 ,Flag = 'N' Where BillingCycleID = 4 "
Call TmpTable

tmp = " UPDATE Mst_BillingCycle SET Ratio = 0 ,Flag = 'R' Where BillingCycleID = 5 "
Call TmpTable

tmp = " UPDATE Mst_BillingCycle SET Ratio = 0 ,Flag = 'R' Where BillingCycleID = 6 "
Call TmpTable

tmp = " ALTER TABLE Mst_Godown ADD ApprovalLimit DECIMAL(18,2),RentMinimumValue DECIMAL(18,2) "
Call TmpTable

tmp = " Alter Table Txn_CxTF_GSL Add DepConDID  Decimal (18,0)"
Call TmpTable

tmp = " ALTER TABLE Txn_CxTF_GSL ADD CONSTRAINT FK_Txn_CxTF_GSL_Txn_DepositContractGodownDetail_DepConDID FOREIGN KEY (DepConDID) REFERENCES dbo.Txn_DepositContractGodownDetail(DepConDID)"
Call TmpTable

tmp = " Alter Table Txn_CxTF_GSL Add NoofDays Decimal (18,0)"
Call TmpTable
 
tmp = " Alter Table Mst_GSL Add DepConDID  Decimal (18,0) "
Call TmpTable

tmp = " ALTER TABLE Mst_GSL ADD CONSTRAINT FK_Mst_GSL_Txn_DepositContractGodownDetail_DepConDID FOREIGN KEY (DepConDID) REFERENCES dbo.Txn_DepositContractGodownDetail(DepConDID) "
Call TmpTable

tmp = " Alter Table Mst_GSL Add NoofDays Decimal (18,0) "
Call TmpTable

tmp = " Alter Table Mst_Config Add DepositContractProcessUpto DateTime"
Call TmpTable


''-- Deposit Contract Attachment
tmp = " Create Table Txn_DepositContractAttachment("
tmp = tmp & " AttachmentID    Decimal(18,0) Not null,"
tmp = tmp & " DepConID        Decimal(18,0),"
tmp = tmp & " [FileName]      Varchar(500),"
tmp = tmp & " [File]          image,"
tmp = tmp & " Remarks         Varchar(8000),"
tmp = tmp & " G_UID           Varchar(36),"
tmp = tmp & " Concurrency     Datetime,"
tmp = tmp & " CreatedBy       Varchar(30),"
tmp = tmp & " CreatedDate     Datetime,"
tmp = tmp & " UpdatedBy       Varchar(30),"
tmp = tmp & " UpdatedDate     DateTime)"
Call TmpTableSMAttachment

tmp = " ALTER TABLE Txn_DepositContractAttachment WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_DepositContractAttachment PRIMARY KEY CLUSTERED (AttachmentID),"
tmp = tmp & " CONSTRAINT DF_Txn_DepositContractAttachment DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTableSMAttachment


''--  Insert Menus
tmp = " INSERT INTO Mst_MenuDetail (MenuID,FormFileName,MenuName,FormName,Flag,HMenuFlag,SelfMenu,MasterMenuID,SequenceNo) Select MAX(MenuID)+1,'FrmTxn_DepositContract.frm','Deposit Contract','FrmTxn_DepositContract','Y','T',0,228,1 From Mst_MenuDetail "
Call TmpTable
tmp = " INSERT INTO Mst_MenuDetail (MenuID,FormFileName,MenuName,FormName,Flag,HMenuFlag,SelfMenu,MasterMenuID,SequenceNo) Select MAX(MenuID)+1,'FrmTxn_GodownWiseApprovalMargin.frm','Godown Wise Approval Limit','FrmTxn_GodownWiseApprovalMargin','Y','T',0,228,1 From Mst_MenuDetail "
Call TmpTable
tmp = " INSERT INTO Mst_MenuDetail (MenuID,FormFileName,MenuName,FormName,Flag,HMenuFlag,SelfMenu,MasterMenuID,SequenceNo) Select MAX(MenuID)+1,'FrmTxn_DepositContractExtend.frm','Extend Deposit Contract','FrmTxn_DepositContractExtend','Y','T',0,228,1 From Mst_MenuDetail "
Call TmpTable
tmp = " INSERT INTO Mst_MenuDetail (MenuID,FormFileName,MenuName,FormName,Flag,HMenuFlag,SelfMenu,MasterMenuID,SequenceNo) Select MAX(MenuID)+1,'FrmRpt_DepositContract.frm','Deposit Contract Report','FrmRpt_DepositContract','Y','R',0,229,1 From Mst_MenuDetail "
Call TmpTable

UpdateConfig ("313")
MsgBox "Version 313 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version312()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Alter Table Txn_CDTFBOE Add GSTGroupCode Varchar (25),HSNNo Varchar(20),AccountType Varchar(10)"
Call TmpTable
tmp = " Alter Table Txn_MonthlyCommtrackInvoice Add GSTGroupCode Varchar (25),HSNNo Varchar(20),AccountType Varchar(10)"
Call TmpTable
tmp = " Alter Table Txn_ReservationInvoice Add HSNNo Varchar(20),AccountType Varchar(10)"
Call TmpTable
tmp = " Alter Table Txn_TempInvoice Add HSNNo Varchar(20),AccountType Varchar(10)"
Call TmpTable


tmp = "Alter Table Txn_LessorAgreement Add BRAgreement bit"
Call TmpTable

tmp = " Update Txn_LessorAgreement Set BRAgreement = 0 "
Call TmpTable

UpdateConfig ("312")
MsgBox "Version 312 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version311()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "ALter Table Mst_ExpenseType Add HSNNo Varchar(20)"
Call TmpTable

UpdateConfig ("311")
MsgBox "Version 311 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version310()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " ALTER TABLE Txn_GodownRent Add BillingStateIndiaID DECIMAL(18,0) "
Call TmpTable

tmp = " Alter Table Txn_GodownRent With NoCheck Add Constraint FK_Txn_GodownRent_Mst_StateIndia_BillingStateIndiaID FOREIGN KEY (BillingStateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID) "
Call TmpTable


UpdateConfig ("310")
MsgBox "Version 310 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version309()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " ALTER TABLE Txn_SecurityBilling Add BillingStateIndiaID DECIMAL(18,0)  "
Call TmpTable

tmp = " Alter Table Txn_SecurityBilling With NoCheck Add Constraint FK_Txn_SecurityBilling_Mst_StateIndia_BillingStateIndiaID FOREIGN KEY (BillingStateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID)"
Call TmpTable

tmp = " ALTER TABLE Txn_FumigationApproval Add BillingStateIndiaID DECIMAL(18,0)  "
Call TmpTable

tmp = " Alter Table Txn_FumigationApproval With NoCheck Add Constraint FK_Txn_FumigationApproval_Mst_StateIndia_BillingStateIndiaID FOREIGN KEY (BillingStateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID)"
Call TmpTable

tmp = " ALTER TABLE Txn_FranchiseeRent Add BillingStateIndiaID DECIMAL(18,0)"
Call TmpTable

tmp = " Alter Table Txn_FranchiseeRent With NoCheck Add Constraint FK_Txn_FranchiseeRent_Mst_StateIndia_BillingStateIndiaID FOREIGN KEY (BillingStateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID)"
Call TmpTable


UpdateConfig ("309")
MsgBox "Version 309 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function
Public Function Version308()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Mst_Lessor Add Status Varchar(1)"
Call TmpTable

tmp = " Alter Table Mst_Depository Add Status Varchar(1)"
Call TmpTable

UpdateConfig ("308")
MsgBox "Version 308 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version307()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_NewCashReceiptClientDetail Add NAVCode varchar(25)"
Call TmpTable


UpdateConfig ("307")
MsgBox "Version 307 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version306()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Mst_GLAccType Add GroupType VARCHAR(1)"
Call TmpTable

tmp = " Update Mst_GLAccType SET GroupType = 'I'"
Call TmpTable

tmp = "Alter Table Mst_GLAccountIncomeTypeMapping Add ExpenseTypeID Decimal(18,0)"
Call TmpTable

tmp = "Alter Table Mst_GLAccountIncomeTypeMapping Alter Column IETypeID Decimal(18,0) Null"
Call TmpTable

tmp = "Alter Table Mst_GLAccountIncomeTypeMapping With NoCheck Add Constraint FK_Mst_GLAccountIncomeTypeMapping_Mst_ExpenseType_ExpenseTypeID FOREIGN KEY (ExpenseTypeID) REFERENCES Mst_ExpenseType(ExpenseTypeID)"
Call TmpTable

tmp = "Alter Table Txn_SecurityBilling Add StateIndiaID Decimal(18,0)"
Call TmpTable

tmp = "Alter Table Txn_SecurityBilling With NoCheck Add Constraint FK_Txn_SecurityBilling_Mst_StateIndia_StateIndiaID FOREIGN KEY (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID)"
Call TmpTable



UpdateConfig ("306")
MsgBox "Version 306 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version305()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

''--  Insert Menu for State Wise GST Data Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_StateWiseGSTNo.frm','State Wise GST No Report','FrmRpt_StateWiseGSTNo','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("305")
MsgBox "Version 305 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version304()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Alter Table Mst_GLAccType Add HSNNo Varchar(20), AccountType Varchar(10)"
Call TmpTable

tmp = "Alter Table Mst_Commodity Add HSNNo Varchar(20), AccountType Varchar(10)"
Call TmpTable

tmp = "Alter Table Txn_ProInvoice Add HSNNo Varchar(20), AccountType Varchar(10)"
Call TmpTable

tmp = "Alter Table Mst_ServiceTax Add InvoiceType Varchar(20)"
Call TmpTable

tmp = "Create Table Mst_GLAccountIncomeTypeMapping("
tmp = tmp & " GLAccountIncomeTypeMappingID      Decimal(18,0) NOT NULL,"
tmp = tmp & " GLAccTypeID Decimal(18,0) NOT NULL,"
tmp = tmp & " IETypeID    Decimal(18,0) NOT NULL,"
tmp = tmp & " Flag        VARCHAR(1),"
tmp = tmp & " G_UID   VARCHAR(36),"
tmp = tmp & " Remark VARCHAR(1000),"
tmp = tmp & " Concurrency  DATETIME,"
tmp = tmp & " CreatedDate   DATETIME,"
tmp = tmp & " CreatedBy   VARCHAR(30),"
tmp = tmp & " UpdatedDate  DATETIME,"
tmp = tmp & " UpdatedBy  VARCHAR(30))"

Call TmpTable


tmp = "ALTER TABLE Mst_GLAccountIncomeTypeMapping WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_GLAccountIncomeTypeMapping PRIMARY KEY CLUSTERED (GLAccountIncomeTypeMappingID),"
tmp = tmp & " CONSTRAINT DF_Mst_GLAccountIncomeTypeMapping DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_GLAccountIncomeTypeMapping_Mst_GLAccType_GLAccTypeID FOREIGN KEY (GLAccTypeID) REFERENCES Mst_GLAccType(GLAccTypeID),"
tmp = tmp & " CONSTRAINT FK_Mst_GLAccountIncomeTypeMapping_Mst_IncomeExpenseType_IETypeID FOREIGN KEY (IETypeID) REFERENCES Mst_IncomeExpenseType(IETypeID)"

Call TmpTable


UpdateConfig ("304")
MsgBox "Version 304 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version303()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Mst_ClientStateWiseGST ADD RGSTNavCode " & TextType & "(25) "
Call TmpTable


UpdateConfig ("303")
MsgBox "Version 303 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version302()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_MonthlyCommtrackFileDetail Add GST Decimal(18,2)"
Call TmpTable

tmp = "Alter Table Txn_MonthlyCommtrackInvoice Add CPID " & TextType & "(50) ,CPName " & TextType & "(200)"
Call TmpTable


tmp = " ALTER TABLE Mst_Client ADD TANNO " & TextType & "(10)"
Call TmpTable

tmp = " ALTER TABLE Mst_Lessor ADD TANNO " & TextType & "(10)"
Call TmpTable

tmp = " ALTER TABLE Mst_Depository ADD TANNO " & TextType & "(10)"
Call TmpTable

tmp = " ALTER TABLE Mst_SecurityAgency ADD TANNO " & TextType & "(10)"
Call TmpTable

tmp = " ALTER TABLE Mst_Vendor ADD TANNO " & TextType & "(10)"
Call TmpTable

UpdateConfig ("302")
MsgBox "Version 302 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version301()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

If Gen_DBType <> "MDB" Then
    ''---- Alter Mst_StateIndia for GST No
    tmp = "Alter Table Mst_StateIndia Add TINNo " & TextType & "(2),StateIndiaCode " & TextType & "(2),GSTIN " & TextType & "(15),NAVCode " & TextType & "(25),StateIndiaOfficeAddress " & TextType & "(500)"
    Call TmpTable
End If

''---- Client State Wise GST
tmp = "Create Table Mst_ClientStateWiseGST("
tmp = tmp & "ClientStateWiseGSTID        Decimal(18,0) NOT NULL,"
tmp = tmp & "ClientIDRow                 Decimal(10,0),"
tmp = tmp & "StateIndiaID                Decimal(18,0),"
tmp = tmp & "ClientStateOfficeAddress    " & TextType & "(500),"
tmp = tmp & "GSTNo                       " & TextType & "(15),"
tmp = tmp & "GSTNavCode                  " & TextType & "(25),"
tmp = tmp & "Flag                        " & TextType & "(1),"
tmp = tmp & "G_UID                       " & TextType & "(36),"
tmp = tmp & "Remarks                     " & TextType & "(100),"
tmp = tmp & "Concurrency                 DATETIME,"
tmp = tmp & "CreatedDate                 DATETIME,"
tmp = tmp & "CreatedBy                   " & TextType & "(30),"
tmp = tmp & "UpdatedDate                 DATETIME,"
tmp = tmp & "UpdatedBy                   " & TextType & "(30))"
Call TmpTable

If Gen_DBType <> "MDB" Then
tmp = "ALTER TABLE Mst_ClientStateWiseGST WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_ClientStateWiseGST PRIMARY KEY CLUSTERED (ClientStateWiseGSTID),"
tmp = tmp & " CONSTRAINT DF_Mst_ClientStateWiseGST DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_ClientStateWiseGST_Mst_StateIndia_StateIndiaID FOREIGN KEY (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID)"
Call TmpTable
End If

''---- Depository State Wise GST
tmp = "Create Table Mst_DepositoryStateWiseGST("
tmp = tmp & "DepositoryStateWiseGSTID        Decimal(18,0) NOT NULL,"
tmp = tmp & "DepositoryID                    " & TextType & "(50) ,"
tmp = tmp & "ExchangeTypeID                  Decimal(18,0),"
tmp = tmp & "StateIndiaID                    Decimal(18,0),"
tmp = tmp & "DepositoryStateOfficeAddress    " & TextType & "(500),"
tmp = tmp & "GSTNo                           " & TextType & "(15),"
tmp = tmp & "GSTNavCode                      " & TextType & "(25),"
tmp = tmp & "Flag                            " & TextType & "(1),"
tmp = tmp & "G_UID                           " & TextType & "(36),"
tmp = tmp & "Remarks                         " & TextType & "(100),"
tmp = tmp & "Concurrency                     DATETIME,"
tmp = tmp & "CreatedDate                     DATETIME,"
tmp = tmp & "CreatedBy                       " & TextType & "(30),"
tmp = tmp & "UpdatedDate                     DATETIME,"
tmp = tmp & "UpdatedBy                       " & TextType & "(30))"
Call TmpTable

If Gen_DBType <> "MDB" Then
tmp = "ALTER TABLE Mst_DepositoryStateWiseGST WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_DepositoryStateWiseGST PRIMARY KEY CLUSTERED (DepositoryStateWiseGSTID),"
tmp = tmp & " CONSTRAINT DF_Mst_DepositoryStateWiseGST DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_DepositoryStateWiseGST_Mst_StateIndia_StateIndiaID FOREIGN KEY (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID),"
tmp = tmp & " CONSTRAINT FK_Mst_DepositoryStateWiseGST_Mst_ExchangeType_ExchangeTypeID FOREIGN KEY (ExchangeTypeID) REFERENCES Mst_ExchangeType(ExchangeTypeID),"
tmp = tmp & " CONSTRAINT FK_Mst_DepositoryStateWiseGST_Mst_Depository_DepositoryIDExchangeTypeID FOREIGN KEY (DepositoryID,ExchangeTypeID) REFERENCES Mst_Depository(DepositoryID,ExchangeTypeID)"
Call TmpTable
End If

''---- Lessor State Wise GST
tmp = "Create Table Mst_LessorStateWiseGST("
tmp = tmp & "LessorStateWiseGSTID            Decimal(18,0) NOT NULL,"
tmp = tmp & "LessorID                        Decimal(18,0),"
tmp = tmp & "StateIndiaID                    Decimal(18,0),"
tmp = tmp & "LessorStateOfficeAddress        " & TextType & "(500),"
tmp = tmp & "GSTNo                           " & TextType & "(15),"
tmp = tmp & "GSTNavCode                      " & TextType & "(25),"
tmp = tmp & "Flag                            " & TextType & "(1),"
tmp = tmp & "G_UID                           " & TextType & "(36),"
tmp = tmp & "Remarks                         " & TextType & "(100),"
tmp = tmp & "Concurrency                     DATETIME,"
tmp = tmp & "CreatedDate                     DATETIME,"
tmp = tmp & "CreatedBy                       " & TextType & "(30),"
tmp = tmp & "UpdatedDate                     DATETIME,"
tmp = tmp & "UpdatedBy                       " & TextType & "(30))"
Call TmpTable

If Gen_DBType <> "MDB" Then
tmp = "ALTER TABLE Mst_LessorStateWiseGST WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_LessorStateWiseGST PRIMARY KEY CLUSTERED (LessorStateWiseGSTID),"
tmp = tmp & " CONSTRAINT DF_Mst_LessorStateWiseGST DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_LessorStateWiseGST_Mst_StateIndia_StateIndiaID FOREIGN KEY (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID),"
tmp = tmp & " CONSTRAINT FK_Mst_LessorStateWiseGST_Mst_Lessor_LessorID FOREIGN KEY (LessorID) REFERENCES Mst_Lessor(LessorID)"
Call TmpTable
End If

''---- Security Agency State Wise GST
tmp = "Create Table Mst_SecurityAgencyStateWiseGST("
tmp = tmp & "SecurityAgencyStateWiseGSTID            Decimal(18,0) NOT NULL,"
tmp = tmp & "AgencyID                                Decimal(18,0),"
tmp = tmp & "StateIndiaID                            Decimal(18,0),"
tmp = tmp & "SecurityAgencyStateOfficeAddress        " & TextType & "(500),"
tmp = tmp & "GSTNo                                   " & TextType & "(15),"
tmp = tmp & "GSTNavCode                              " & TextType & "(25),"
tmp = tmp & "Flag                                    " & TextType & "(1),"
tmp = tmp & "G_UID                                   " & TextType & "(36),"
tmp = tmp & "Remarks                                 " & TextType & "(100),"
tmp = tmp & "Concurrency                             DATETIME,"
tmp = tmp & "CreatedDate                             DATETIME,"
tmp = tmp & "CreatedBy                               " & TextType & "(30),"
tmp = tmp & "UpdatedDate                             DATETIME,"
tmp = tmp & "UpdatedBy                               " & TextType & "(30))"
Call TmpTable

If Gen_DBType <> "MDB" Then
tmp = "ALTER TABLE Mst_SecurityAgencyStateWiseGST WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_SecurityAgencyStateWiseGST PRIMARY KEY CLUSTERED (SecurityAgencyStateWiseGSTID),"
tmp = tmp & " CONSTRAINT DF_Mst_SecurityAgencyStateWiseGST DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_SecurityAgencyStateWiseGST_Mst_StateIndia_StateIndiaID FOREIGN KEY (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID),"
tmp = tmp & " CONSTRAINT FK_Mst_SecurityAgencyStateWiseGST_Mst_SecurityAgency_AgencyID FOREIGN KEY (AgencyID) REFERENCES Mst_SecurityAgency(AgencyID)"
Call TmpTable
End If

''---- Security Agency State Wise GST
tmp = "Create Table Mst_VendorStateWiseGST("
tmp = tmp & "VendorStateWiseGSTID            Decimal(18,0) NOT NULL,"
tmp = tmp & "VendorID                        Decimal(18,0),"
tmp = tmp & "StateIndiaID                    Decimal(18,0),"
tmp = tmp & "VendorStateOfficeAddress        " & TextType & "(500),"
tmp = tmp & "GSTNo                           " & TextType & "(15),"
tmp = tmp & "GSTNavCode                      " & TextType & "(25),"
tmp = tmp & "Flag                            " & TextType & "(1),"
tmp = tmp & "G_UID                           " & TextType & "(36),"
tmp = tmp & "Remarks                         " & TextType & "(100),"
tmp = tmp & "Concurrency                     DATETIME,"
tmp = tmp & "CreatedDate                     DATETIME,"
tmp = tmp & "CreatedBy                       " & TextType & "(30),"
tmp = tmp & "UpdatedDate                     DATETIME,"
tmp = tmp & "UpdatedBy                       " & TextType & "(30))"
Call TmpTable

If Gen_DBType <> "MDB" Then
tmp = "ALTER TABLE Mst_VendorStateWiseGST WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Mst_VendorStateWiseGST PRIMARY KEY CLUSTERED (VendorStateWiseGSTID),"
tmp = tmp & " CONSTRAINT DF_Mst_VendorStateWiseGST DEFAULT (GETDATE()) FOR Concurrency,"
tmp = tmp & " CONSTRAINT FK_Mst_VendorStateWiseGST_Mst_StateIndia_StateIndiaID FOREIGN KEY (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID),"
tmp = tmp & " CONSTRAINT FK_Mst_VendorStateWiseGST_Mst_Vendor_VendorID FOREIGN KEY (VendorID) REFERENCES Mst_Vendor(VendorID)"
Call TmpTable
End If


''--  Insert Menu for Import Client State Wise GST
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ImportClientStateWiseGST.frm','Import Client State Wise GST','FrmTxn_ImportClientStateWiseGST','Y','M',0,227,1"
Call TmpTable

''--  Insert Menu for Import Lessor State Wise GST
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ImportLessorStateWiseGST.frm','Import Lessor State Wise GST','FrmTxn_ImportLessorStateWiseGST','Y','M',0,227,1"
Call TmpTable

''--  Insert Menu for Import Depository State Wise GST
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ImportDepositoryStateWiseGST.frm','Import Depository State Wise GST','FrmTxn_ImportDepositoryStateWiseGST','Y','M',0,227,1"
Call TmpTable

''--  Insert Menu for Import Security Agency State Wise GST
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ImportSecurityAgencyStateWiseGST.frm','Import Security Agency State Wise GST','FrmTxn_ImportSecurityAgencyStateWiseGST','Y','M',0,227,1"
Call TmpTable

''--  Insert Menu for Import Vendor State Wise GST
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ImportVendorStateWiseGST.frm','Import Vendor State Wise GST','FrmTxn_ImportVendorStateWiseGST','Y','M',0,227,1"
Call TmpTable

UpdateConfig ("301")
MsgBox "Version 301 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version300()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Mst_StorageContract Add BlockDeposit Varchar(1)"
Call TmpTable

tmp = "Update Mst_StorageContract Set BlockDeposit = '0'"
Call TmpTable

UpdateConfig ("300")
MsgBox "Version 300 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version299()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If



tmp = " ALTER TABLE Mst_GSL ADD RateCardType Varchar(1), BillingCycleID DECIMAL(18,0), BillingUnit VARCHAR(1),Rate DECIMAL(18,2),RateForDailyPnL DECIMAL(18,2)"
Call TmpTable


tmp = " ALTER TABLE MST_GSL WITH NOCHECK ADD CONSTRAINT FK_MST_GSL_BillingCycleID FOREIGN KEY (BillingCycleID) REFERENCES Mst_BillingCycle(BillingCycleID)"
Call TmpTable

tmp = " ALTER TABLE Txn_CxTF_GSL ADD RateCardType Varchar(1), BillingCycleID DECIMAL(18,0), BillingUnit VARCHAR(1),Rate DECIMAL(18,2),RateForDailyPnL DECIMAL(18,2)"
Call TmpTable

tmp = " ALTER TABLE Txn_CxTF_GSL WITH NOCHECK ADD CONSTRAINT FK_Txn_CxTF_GSL_BillingCycleID FOREIGN KEY (BillingCycleID) REFERENCES Mst_BillingCycle(BillingCycleID)"
Call TmpTable

UpdateConfig ("299")
MsgBox "Version 299 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version298()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo) "
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_VendorClaimSettledMultiple.frm','Settle Multiple Vendor Claim','FrmTxn_VendorClaimSettledMultiple','Y','T',0,228,1"

Call TmpTable


UpdateConfig ("298")
MsgBox "Version 298 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version297()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


''-- Alter FAG Bank Master TO add Payment Limitation
tmp = " Alter table Mst_Fagbank Add PaymentLimitation Decimal(18,0)"
Call TmpTable


''-- Alter Security Agency Master for Bank Details
tmp = " Alter table Mst_SecurityAgency Add BankID Decimal (18,0),BranchID Decimal(18,0),BankAcNo " & TextType & "(100),IFSCCode " & TextType & "(50)"
Call TmpTable

tmp = " ALTER TABLE Mst_SecurityAgency WITH NOCHECK ADD CONSTRAINT FK_Mst_SecurityAgency_Bank FOREIGN KEY (BankID) REFERENCES Mst_Bank (BankID)"
Call TmpTable

tmp = " ALTER TABLE Mst_SecurityAgency WITH NOCHECK ADD CONSTRAINT FK_Mst_SecurityAgency_Branch FOREIGN KEY (BranchID) REFERENCES Mst_Branch (BranchID)"
Call TmpTable


''-- Alter Vendor Master for Bank Details
tmp = " Alter Table Mst_Vendor Add BankID Decimal (18,0),BranchID Decimal(18,0),BankAcNo " & TextType & "(100),IFSCCode " & TextType & "(50)"
Call TmpTable

tmp = " ALTER TABLE Mst_Vendor WITH NOCHECK ADD CONSTRAINT FK_Mst_Vendor_Bank FOREIGN KEY (BankID) REFERENCES Mst_Bank (BankID)"
Call TmpTable

tmp = " ALTER TABLE Mst_Vendor WITH NOCHECK ADD CONSTRAINT FK_Mst_Vendor_Branch FOREIGN KEY (BranchID) REFERENCES Mst_Branch (BranchID)"
Call TmpTable


''-- Alter Security Billing for Payment
tmp = " Alter Table Txn_SecurityBilling Add DisbuDate DateTime,DisbuRemark " & TextType & "(255),DeductionRemark " & TextType & "(1000),DisbuAmount Decimal(18,2),ModeOfPay " & TextType & "(50),FAGBankID Decimal(18,0),PBatchID Decimal(18,0),DocNo " & TextType & "(50),EntryNo Decimal(18,0),FAGPayExportedDate DateTime,LetterPrintFlag " & TextType & "(1)"
Call TmpTable
''-- Alter Security Billing for Payment
tmp = " Alter Table Txn_FumigationApproval Add DisbuDate DateTime,DisbuRemark " & TextType & "(255),DeductionRemark " & TextType & "(1000),DisbuAmount Decimal(18,2),ModeOfPay " & TextType & "(50),FAGBankID Decimal(18,0),PBatchID Decimal(18,0),DocNo " & TextType & "(50),EntryNo Decimal(18,0),FAGPayExportedDate DateTime,LetterPrintFlag " & TextType & "(1)"
Call TmpTable
''-- Alter Security Billing for Payment
tmp = " Alter Table Txn_GodownRent Add DisbuDate DateTime,DisbuRemark " & TextType & "(255),DeductionRemark " & TextType & "(1000),DisbuAmount Decimal(18,2),ModeOfPay " & TextType & "(50),FAGBankID Decimal(18,0),PBatchID Decimal(18,0),DocNo " & TextType & "(50),EntryNo Decimal(18,0),FAGPayExportedDate DateTime,LetterPrintFlag " & TextType & "(1)"
Call TmpTable
''-- Alter Security Billing for Payment
tmp = " Alter Table Txn_FranchiseeRent Add DisbuDate DateTime,DisbuRemark " & TextType & "(255),DeductionRemark " & TextType & "(1000),DisbuAmount Decimal(18,2),ModeOfPay " & TextType & "(50),FAGBankID Decimal(18,0),PBatchID Decimal(18,0),DocNo " & TextType & "(50),EntryNo Decimal(18,0),FAGPayExportedDate DateTime,LetterPrintFlag " & TextType & "(1)"
Call TmpTable

''-- Update Txn_GodownRent / Security / Fumigation / Franchisee for Settle whose Batch is Exported
tmp = " UPDATE Txn_GodownRent SET Flag = 'S',PBatchID = 1 Where BatchID > 0 And Flag = 'A'"
Call TmpTable
tmp = " UPDATE Txn_SecurityBilling SET Status = 'S',PBatchID = 1 Where BatchID > 0 And Status = 'A'"
Call TmpTable
tmp = " UPDATE Txn_FumigationApproval SET Flag = 'S',PBatchID = 1 Where BatchID > 0 And Flag = 'A'"
Call TmpTable
tmp = " UPDATE Txn_FranchiseeRent SET Flag = 'S',PBatchID = 1 Where BatchID > 0 And Flag = 'A'"
Call TmpTable

''-- Insert Claim Payment Export Menu
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo) Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportClaimPayment.frm','Export Claim Payment','Frm_ExportClaimPayment','Y','T',0,228,1"

UpdateConfig ("297")
MsgBox "Version 297 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function





Public Function Version296()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " CREATE proc sp_clientledger"
tmp = tmp & " @CustomerNo varchar(10)=NULL,"
tmp = tmp & "  @CustomerName varchar(500)=NULL,"
tmp = tmp & " @EntryNo decimal(18, 0)=NULL,"
tmp = tmp & " @EHOPInvoiceNo varchar(50) =NULL,"
tmp = tmp & " @DocumentType varchar(50) =NULL,"
tmp = tmp & " @DocumentNo varchar(50) =NULL,"
tmp = tmp & " @VoucherNo varchar(50) =NULL,"
tmp = tmp & " @PostingDate datetime =NULL,"
tmp = tmp & " @InvoiceAmount decimal(18, 2) =NULL,"
tmp = tmp & " @BaseAmount decimal(18, 2) =NULL,"
tmp = tmp & " @TaxAmount decimal(18, 2) =NULL,"
tmp = tmp & " @RemainingAmount decimal(18, 2) =NULL,"
tmp = tmp & " @Narration varchar(2000) =NULL,"
tmp = tmp & " @ActualDatePosted datetime =NULL,"
tmp = tmp & " @ActualTimePosted varchar(20) =NULL,"
tmp = tmp & " @RegionCode varchar(30) =NULL,"
tmp = tmp & " @DepartmentCode varchar(30) =NULL,"
tmp = tmp & " @CustomerPostingGroup varchar(20) =NULL,"
tmp = tmp & " @CreatedBy varchar(30) =NULL,"
tmp = tmp & " @CreatedDate datetime =NULL,"
tmp = tmp & " @Concurrency datetime =NULL,"
tmp = tmp & " @G_UID varchar(36) =NULL,"
tmp = tmp & " @LastUpdatedDate datetime =NULL"
tmp = tmp & " as"
tmp = tmp & " if not exists (select EntryNo from txn_clientledger where EntryNo=@EntryNo)"
tmp = tmp & " begin"
tmp = tmp & " insert into txn_clientledger"
tmp = tmp & " (CustomerNo,CustomerName,EntryNo,EHOPInvoiceNo,DocumentType,DocumentNo,VoucherNo,PostingDate,InvoiceAmount,BaseAmount,TaxAmount,RemainingAmount,Narration,ActualDatePosted,ActualTimePosted,RegionCode,DepartmentCode,CustomerPostingGroup,CreatedBy,CreatedDate,Concurrency,G_UID)"
tmp = tmp & " values ("
tmp = tmp & " @CustomerNo, @CustomerName, @EntryNo, @EHOPInvoiceNo, @DocumentType, @DocumentNo, @VoucherNo, @PostingDate, @InvoiceAmount, @BaseAmount, @TaxAmount, @RemainingAmount, @Narration, @ActualDatePosted, @ActualTimePosted, @RegionCode, @DepartmentCode, @CustomerPostingGroup, @CreatedBy, getdate(), getdate(), newid()"
tmp = tmp & " )"
tmp = tmp & " End"

Call TmpTable




''--- Invoices For NAVISION
tmp = " Create Table  Txn_InvoicesForNAVISION ("
tmp = tmp & " Date                    Datetime,"
tmp = tmp & " Invoice                 " & TextType & "(100),"
tmp = tmp & " Customer                " & TextType & "(100),"
tmp = tmp & " NavCode                 " & TextType & "(100),"
tmp = tmp & " Amount                  Decimal(18,2),"
tmp = tmp & " GLAccount               " & TextType & "(100),"
tmp = tmp & " GLACNAVCode             " & TextType & "(200),"
tmp = tmp & " Narration               " & TextType & "(1000),"
tmp = tmp & " LocationNAVCode         " & TextType & "(100),"
tmp = tmp & " DeptDim                 " & TextType & "(100),"
tmp = tmp & " EmpDim                  " & TextType & "(100),"
tmp = tmp & " PINo                    " & TextType & "(200),"
tmp = tmp & " STCode                  " & TextType & "(100),"
tmp = tmp & " AgreeNonAgree           " & TextType & "(50),"
tmp = tmp & " TINo                    " & TextType & "(50),"
tmp = tmp & " TIDate                  DateTime,"
tmp = tmp & " STAmt                   Decimal(18,2),"
tmp = tmp & " TIAmt                   Decimal(18,2),"
tmp = tmp & " PostInvoiceNo           " & TextType & "(50),"
tmp = tmp & " EntryNo                 Decimal(18,0),"
tmp = tmp & " BatchID                 Decimal(18,0),"
tmp = tmp & " DataTransfer            " & TextType & "(1),"
tmp = tmp & " InvoiceType             " & TextType & "(100),"
tmp = tmp & " journalBatchType        " & TextType & "(50),"
tmp = tmp & " journalBatchName        " & TextType & "(100),"
tmp = tmp & " ReverseInvoice          " & TextType & "(1),"
tmp = tmp & " EntryNoTransfer         " & TextType & "(1),"
tmp = tmp & " EntryNoTransferRemark   " & TextType & "(1000),"
tmp = tmp & " UserBatchID             " & TextType & "(36),"
tmp = tmp & " [LineNo]                Bigint IDENTITY(1,1),"
tmp = tmp & " Concurrency             Datetime,"
tmp = tmp & " G_UID                   " & TextType & "(36),"
tmp = tmp & " CreatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedDate             Datetime )"
Call TmpTable

tmp = " ALTER TABLE Txn_InvoicesForNAVISION WITH NOCHECK ADD CONSTRAINT DF_Txn_InvoicesForNAVISION DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable

''--- Cash Repeipts For NAVISION
tmp = " Create Table Txn_CashRepeiptsForNAVISION ("
tmp = tmp & " PostingDate             DATETIME,"
tmp = tmp & " DocumentType            " & TextType & "(100),"
tmp = tmp & " AccountType             " & TextType & "(100),"
tmp = tmp & " AccountNo               " & TextType & "(100),"
tmp = tmp & " Amount                  DECIMAL(18,2),"
tmp = tmp & " InvoiceNo               " & TextType & "(100),"
tmp = tmp & " Narration               " & TextType & "(1000),"
tmp = tmp & " RegionDim               " & TextType & "(100),"
tmp = tmp & " DeptDim                 " & TextType & "(100),"
tmp = tmp & " EmpDim                  " & TextType & "(100),"
tmp = tmp & " ChequeNo                " & TextType & "(100),"
tmp = tmp & " ChequeDate              DATETIME,"
tmp = tmp & " DepositDate             DATETIME,"
tmp = tmp & " PINo                    " & TextType & "(200),"
tmp = tmp & " PostInvoiceNo           " & TextType & "(50),"
tmp = tmp & " EntryNo                 Decimal(18,0),"
tmp = tmp & " CRADJEntryNo            Decimal(18,0),"
tmp = tmp & " BatchID                 DECIMAL(18,0),"
tmp = tmp & " DataTransfer            " & TextType & "(1),"
tmp = tmp & " InvoiceType             " & TextType & "(100),"
tmp = tmp & " JournalBatchType        " & TextType & "(50),"
tmp = tmp & " JournalBatchName        " & TextType & "(100),"
tmp = tmp & " ReverseInvoice          " & TextType & "(1),"
tmp = tmp & " STCODE                  " & TextType & "(100),"
tmp = tmp & " BalAccType              " & TextType & "(100),"
tmp = tmp & " BalAccountNo            " & TextType & "(100),"
tmp = tmp & " EntryNoTransfer         " & TextType & "(1),"
tmp = tmp & " EntryNoTransferRemark   " & TextType & "(1000),"
tmp = tmp & " UserBatchID             " & TextType & "(36),"
tmp = tmp & " [LineNo]                BIGINT IDENTITY(1,1),"
tmp = tmp & " Concurrency             DATETIME,"
tmp = tmp & " G_UID                   " & TextType & "(36),"
tmp = tmp & " CreatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedDate             DATETIME    )"
Call TmpTable

tmp = " ALTER TABLE Txn_CashRepeiptsForNAVISION WITH NOCHECK ADD CONSTRAINT DF_Txn_CashRepeiptsForNAVISION DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable


''----- Invoice Allocation For NAVISION
tmp = " Create Table Txn_InvoiceAllocationForNAVISION ("
tmp = tmp & " PaymentEntryNo          " & TextType & "(100),"
tmp = tmp & " InvoiceEntryNo          " & TextType & "(100),"
tmp = tmp & " Amount                  DECIMAL(18,2),"
tmp = tmp & " BatchID                 DECIMAL(18,0),"
tmp = tmp & " DataTransfer            " & TextType & "(1),"
tmp = tmp & " InvoiceType             " & TextType & "(100),"
tmp = tmp & " JournalBatchType        " & TextType & "(50),"
tmp = tmp & " JournalBatchName        " & TextType & "(100),"
tmp = tmp & " EntryNoTransfer         " & TextType & "(1),"
tmp = tmp & " EntryNoTransferRemark   " & TextType & "(1000),"
tmp = tmp & " UserBatchID             " & TextType & "(36),"
tmp = tmp & " [LineNo]                BIGINT IDENTITY(1,1),"
tmp = tmp & " Concurrency             DATETIME,"
tmp = tmp & " G_UID                   " & TextType & "(36),"
tmp = tmp & " CreatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedDate             DATETIME )"
Call TmpTable

tmp = " ALTER TABLE Txn_InvoiceAllocationForNAVISION WITH NOCHECK ADD CONSTRAINT DF_Txn_InvoiceAllocationForNAVISION DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable


''----- Security Bill For Navision
tmp = " Create Table Txn_SecurityBillForNavision("
tmp = tmp & " PostingDate             DATETIME,"
tmp = tmp & " DocumentType            " & TextType & "(200),"
tmp = tmp & " AccountType             " & TextType & "(100),"
tmp = tmp & " AccountNo               " & TextType & "(100),"
tmp = tmp & " Amount                  decimal(18,2),"
tmp = tmp & " TDSNatureofDeduction    " & TextType & "(100),"
tmp = tmp & " ServiceTaxGroupCode     " & TextType & "(100),"
tmp = tmp & " Narration               " & TextType & "(250),"
tmp = tmp & " RegionDim               " & TextType & "(100),"
tmp = tmp & " DeptDim                 " & TextType & "(100),"
tmp = tmp & " VendorDim_EmpDim        " & TextType & "(100),"
tmp = tmp & " ExternalDocumentNo      " & TextType & "(200),"
tmp = tmp & " InvoiceDate             DATETIME,"
tmp = tmp & " InvoiceNo               " & TextType & "(100),"
tmp = tmp & " BankName                " & TextType & "(100),"
tmp = tmp & " ChequeNo                " & TextType & "(100),"
tmp = tmp & " PaidAmount              decimal(18,2),"
tmp = tmp & " TDS                     decimal(18,2),"
tmp = tmp & " PaidDate                DATETIME,"
tmp = tmp & " PINo                    " & TextType & "(200),"
tmp = tmp & " PostInvoiceNo           " & TextType & "(50),"
tmp = tmp & " EntryNo                 Decimal(18,0),"
tmp = tmp & " BatchID                 decimal(18,0),"
tmp = tmp & " DataTransfer            " & TextType & "(1),"
tmp = tmp & " InvoiceType             " & TextType & "(100),"
tmp = tmp & " journalBatchType        " & TextType & "(50),"
tmp = tmp & " journalBatchName        " & TextType & "(100),"
tmp = tmp & " DocumentNo              " & TextType & "(36),"
tmp = tmp & " EntryNoTransfer         " & TextType & "(1),"
tmp = tmp & " EntryNoTransferRemark   " & TextType & "(1000),"
tmp = tmp & " UserBatchID             " & TextType & "(36),"
tmp = tmp & " [LineNo]                BIGINT IDENTITY(1,1),"
tmp = tmp & " Concurrency             DATETIME,"
tmp = tmp & " G_UID                   " & TextType & "(36),"
tmp = tmp & " CreatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedBy               " & TextType & "(35),"
tmp = tmp & " UpdatedDate             DATETIME )"
Call TmpTable

tmp = " ALTER TABLE Txn_SecurityBillForNavision WITH NOCHECK ADD CONSTRAINT DF_Txn_SecurityBillForNavision DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable


tmp = "ALTER TABLE Txn_InvoicesForNAVISION Add Remark " & TextType & "(2000)"
Call TmpTable

tmp = "ALTER TABLE Txn_CashRepeiptsForNAVISION Add Remark " & TextType & "(2000)"
Call TmpTable

tmp = "ALTER TABLE Txn_InvoiceAllocationForNAVISION Add Remark " & TextType & "(2000)"
Call TmpTable

tmp = "ALTER TABLE Txn_SecurityBillForNavision Add Remark " & TextType & "(2000)"
Call TmpTable

''--  Insert Menu for Invoice For Navision Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_InvoiceForNavision.frm','Invoice For Navision Report','FrmRpt_InvoiceForNavision','Y','R',0,229,1 "
Call TmpTable

''--  Insert Menu for Cash Receipt For Navision Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CashReceiptForNavision.frm','Cash Receipt For Navision Report','FrmRpt_CashReceiptForNavision','Y','R',0,229,1 "
Call TmpTable

''--  Insert Menu for Cash Receipt Allocation For Navision Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CashReceiptAllocationForNAVISION.frm','Cash Receipt Allocation For Navision Report','FrmRpt_CashReceiptAllocationForNAVISION','Y','R',0,229,1 "
Call TmpTable

''--  Insert Menu for Security Bill For Navision Report
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_SecurityBillForNavision.frm','Security Bill For Navision Report','FrmRpt_SecurityBillForNavision','Y','R',0,229,1 "
Call TmpTable

UpdateConfig ("296")
MsgBox "Version 296 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version295()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_WHR_Detail Add LotwiseWHR " & TextType & "(1)"
Call TmpTable

tmp = " Update Txn_WHR_Detail Set LotwiseWHR = '0'"
Call TmpTable

tmp = "ALTER TABLE Mst_GSL Add OriginalDepositFlag " & TextType & "(1)"
Call TmpTable

tmp = "UPDATE Mst_GSL SET OriginalDepositFlag = 'Y' WHERE CONVERT(VARCHAR,GSLID)+''+LTrim(RTrim(SM_Prefix)) IN (SELECT Distinct CONVERT(VARCHAR,GSLID)+''+LTrim(RTrim(SM_Prefix)) FROM Txn_CXTF_GSL WHERE TxnType = 'D')"
Call TmpTable

tmp = "Alter Table Mst_GSL Add WHRNo decimal(18,0)"
Call TmpTable

tmp = "ALTER TABLE Txn_GSL_Transfer ADD MadeupBags Decimal(18,0)"
Call TmpTable

tmp = "ALTER TABLE Txn_GSL_Transfer ADD MadeupNewGSLID Decimal(18,0)"
Call TmpTable

tmp = "ALTER TABLE Txn_GSL_Transfer ADD MadeupOLDGSLID Decimal(18,0)"
Call TmpTable

tmp = "ALTER TABLE Txn_PledgeCreation Add MarketYardFlag " & TextType & "(1) "
Call TmpTable

tmp = "Alter Table Mst_GSL Add NFinFlag " & TextType & "(1)"
Call TmpTable

UpdateConfig ("295")
MsgBox "Version 295 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function
Public Function Version294()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Mst_CommodityGroup Add HRCommodityID Decimal(18,0)"
Call TmpTable

tmp = "Alter Table Mst_Commodity Add HRVarietyID Decimal(18,0)"
Call TmpTable

tmp = "Alter Table Mst_Bank Add HRBankID Decimal(18,0)"
Call TmpTable

tmp = "Alter Table Mst_Branch Add HRBranchID Decimal(18,0)"
Call TmpTable

tmp = "Alter Table Mst_Bank Add MarketYard Varchar(1)"
Call TmpTable

UpdateConfig ("294")
MsgBox "Version 294 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function Version293()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Txn_MonthlyCommtrackFileDetail Add EDD DateTime,MultiplyingFactor Decimal(18,4)"
Call TmpTable

tmp = "CREATE TABLE Txn_MonthlyCommtrackFileDetailCMSP"
tmp = tmp & " (MonthYear DATETIME,"
tmp = tmp & "  LotId  Varchar (25),"
 tmp = tmp & " EDD DateTime,"
 tmp = tmp & " ClientID  Varchar (50),"
 tmp = tmp & " ClientName  Varchar (200),"
 tmp = tmp & " Qty DECIMAL(18,3),"
 tmp = tmp & " StartDate DATETIME,"
 tmp = tmp & " EndDate DATETIME,"
 tmp = tmp & " HoldDuration DECIMAL(18,0),"
 tmp = tmp & " MultiplyingFactor Decimal(18,4),"
 tmp = tmp & " Amount DECIMAL(18,4),"
 tmp = tmp & " AmtIncST DECIMAL(18,4),"
 tmp = tmp & " WarehouseName  Varchar (1000),"
 tmp = tmp & " Commodity  Varchar (500),"
 tmp = tmp & " WarehouseCode  Varchar (1000),"
 tmp = tmp & " WSPName  Varchar (500),"
 tmp = tmp & " Concurrency DATETIME,"
 tmp = tmp & " G_UID  Varchar (36),"
 tmp = tmp & " CreatedBy  Varchar (30),"
 tmp = tmp & " CreatedDate DATETIME,"
 tmp = tmp & " UpdatedBy  Varchar (30),"
 tmp = tmp & " UpdatedDate DateTime)"
 
Call TmpTable
 
If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Txn_MonthlyCommtrackFileDetailCMSP WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyCommtrackFileDetailCMSP_Concurrency DEFAULT (GETDATE()) FOR Concurrency"
   
   Call TmpTable
End If

tmp = "Alter Table Txn_MonthlyCommtrackFileDetailCMSP Add CommodityCode Varchar(200),ChargeableQty Decimal(18,4),TranxType Varchar(150),TranxNo Varchar(150),TranxDate DateTime,ChargeableDuration Decimal(18,0),WSPCode Varchar(200),Rate Decimal(18,4)"
Call TmpTable

UpdateConfig ("293")
MsgBox "Version 293 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function
Public Function Version292()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Txn_CxTF_Details Add SpotPrice Decimal(18,2)"
Call TmpTable

tmp = "Alter Table Txn_CDTFBOE_GSL Add NCDEXSpaceCharges decimal (18,2)"
Call TmpTable
    
UpdateConfig ("292")
MsgBox "Version 292 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version291()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


If Gen_DBType <> "MDB" Then
        ''--  Alter Table Mst_MAW for 3 Decimal
        tmp = " ALTER TABLE Mst_MAW Alter COLUMN MinimumRange Decimal(18,3) "
        Call TmpTable

        tmp = " ALTER TABLE Mst_MAW Alter COLUMN MaximumRange Decimal(18,3) "
        Call TmpTable

        ''--  Alter Table Mst_MAWDetails for 3 Decimal
        tmp = " ALTER TABLE Mst_MAWDetails Alter COLUMN MinimumMoisture Decimal(18,3) "
        Call TmpTable

        tmp = " ALTER TABLE Mst_MAWDetails Alter COLUMN MaximumMoisture Decimal(18,3) "
        Call TmpTable
End If
    
tmp = " Alter Table Txn_MonGodownWiseProfitability Add LeaseRentAccrued decimal(18,2) "
Call TmpTable
    
    
UpdateConfig ("291")
MsgBox "Version 291 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function




Public Function Version290()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NNNRBillingAndCurrent.frm','NNNR Invoice Data Vs Current Rate Card','FrmRpt_NNNRBillingAndCurrent','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_NavisionInvoiceCrossCheck.frm','Navision Invoice Verify','FrmTxn_NavisionInvoiceCrossCheck','Y','T',0,228,1"
Call TmpTable


    
UpdateConfig ("290")
MsgBox "Version 290 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function
Public Function Version289()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Mst_Bank Add NCMLFinance Varchar (1)"
Call TmpTable

tmp = "Update Mst_Bank Set NCMLFinance = '0'"
Call TmpTable

tmp = "Alter Table Txn_PledgeCreation Add NCMLFinance Varchar (1)"
Call TmpTable

tmp = "Update Txn_PledgeCreation Set NCMLFinance = '0'"
Call TmpTable

    
UpdateConfig ("289")
MsgBox "Version 289 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function


Public Function version288()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_GodownInspection Add NoOfStack Decimal(18,0),WeightPerStack Decimal(18,3)"
Call TmpTable

tmp = "Alter Table Mst_Godown Add NoOfStack Decimal(18,0),WeightPerStack Decimal(18,3)"
Call TmpTable


tmp = "Create Table Txn_ClientFeedbackDoc"
tmp = tmp & " (TrancationID        Decimal(18, 0) NOT NULL,"
tmp = tmp & "     TrancationDate      Datetime       NULL,"
tmp = tmp & " ClientIDRow         Decimal(18, 0) NOT NULL,"
tmp = tmp & " Flag                Varchar(1)     NULL,"
tmp = tmp & " Remarks             Varchar(1000)  NULL,"
tmp = tmp & " G_UID               Varchar(36)    NULL,"
tmp = tmp & " Concurrency         Datetime       NULL,"
tmp = tmp & "     CreatedBy           Varchar(30)    NULL,"
tmp = tmp & " CreatedDate         Datetime       NULL,"
tmp = tmp & "     UpdatedBy           Varchar(30)    NULL,"
tmp = tmp & " UpdatedDate         Datetime       NULL)"

Call TmpTable

tmp = "ALTER TABLE Txn_ClientFeedbackDoc WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_TrancationID PRIMARY KEY CLUSTERED (TrancationID),"
tmp = tmp & " CONSTRAINT DF_Txn_ClientFeedbackDoc DEFAULT (GETDATE()) FOR Concurrency"
Call TmpTable

tmp = "Create Table Txn_ClientFeedbackDocAttachment"
tmp = tmp & " (    AttachmentID        Decimal(18, 0) NOT NULL,"
tmp = tmp & " TrancationID        Decimal(18, 0) NOT NULL,"
tmp = tmp & " [FileName]          Varchar(500) NULL,"
tmp = tmp & " [File]              Image,"
tmp = tmp & " Remarks             Varchar(1000) NULL,"
tmp = tmp & " G_UID               Varchar(36) NULL,"
tmp = tmp & " Concurrency         Datetime NULL,"
tmp = tmp & " CreatedBy           Varchar(30) NULL,"
tmp = tmp & " CreatedDate         Datetime NULL,"
tmp = tmp & " UpdatedBy           Varchar(30) NULL,"
tmp = tmp & " UpdatedDate         Datetime NULL)"

Call TmpTableSMAttachment

tmp = "ALTER TABLE Txn_ClientFeedbackDocAttachment WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT PK_Txn_ClientFeedbackDocAttachment PRIMARY KEY CLUSTERED (AttachmentID),"
tmp = tmp & " CONSTRAINT DF_Txn_ClientFeedbackDocAttachment DEFAULT (GETDATE()) FOR Concurrency"

Call TmpTableSMAttachment

UpdateConfig ("288")
MsgBox "Version 288 Installed successfully"

End Function


Public Function version287()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_SecLocMappingDetail Add OldRatePerMonth " & TextType & "(300)"
Call TmpTable


UpdateConfig ("287")
MsgBox "Version 287 Installed successfully"

End Function

Public Function version286()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_InvMonGSLDetail Add WtNoOfDay Decimal(18,0),WtQty Decimal(18,3),CapacityNoOfDay Decimal(18,0)"
Call TmpTable


tmp = "ALTER TABLE Txn_InvMonGSLDetail WITH NOCHECK ADD "
tmp = tmp & " CONSTRAINT DF_Txn_InvMonGSLDetail_WtNoOfDay DEFAULT (0) FOR WtNoOfDay,"
tmp = tmp & " CONSTRAINT DF_Txn_InvMonGSLDetail_WtQty DEFAULT (0) FOR WtQty ,"
tmp = tmp & " CONSTRAINT DF_Txn_InvMonGSLDetail_CapacityNoOfDay DEFAULT (0) FOR CapacityNoOfDay"
Call TmpTable


tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownUtilizationNew.frm','Godown Capacity Utilization New','FrmRpt_GodownUtilizationNew','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("286")
MsgBox "Version 286 Installed successfully"

End Function

Public Function version285()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownViabilityVsAgreement.frm','Godown Viability V/s Agreement','FrmRpt_GodownViabilityVsAgreement','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("285")
MsgBox "Version 285 Installed successfully"

End Function

Public Function Version284()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ReservationInvoiceDetails.frm','Reservation Invoice Details','FrmRpt_ReservationInvoiceDetails','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("284")
MsgBox "Version 284 Installed successfully"


End Function

Public Function Version283()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Mst_BagEquivalent Add ProjectEquivalent numeric(18,3)"
Call TmpTable

tmp = "Update Mst_BagEquivalent Set ProjectEquivalent = Equivalent"
Call TmpTable

UpdateConfig ("283")
MsgBox "Version 283 Installed successfully"


End Function
Public Function Version282()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Mst_CMP Add BlockDays Decimal(18,0)"
Call TmpTable

tmp = "Update Mst_CMP Set BlockDays = 1"
Call TmpTable

tmp = " Create Table Log_UpdateCMPBlockDays ("
tmp = tmp & " CMPUpdateID Decimal(18,0),CMPID Decimal(18,0),"
tmp = tmp & " OldBlockDays Decimal(18,0),"
tmp = tmp & " NewBlockDays Decimal(18,0),"
tmp = tmp & " Remark " & TextType & "(3500),"
tmp = tmp & " UpdatedBy Varchar(35),"
tmp = tmp & " UpdatedOn DateTime)"
Call TmpTable



UpdateConfig ("282")
MsgBox "Version 282 Installed successfully"


End Function


Public Function Version281()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create Table Mst_Stack("
tmp = tmp & " StackID Decimal(18,0) NOT NULL,"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " StackNo " & TextType & "(10),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " Remark " & TextType & "(1000),"
tmp = tmp & " G_UID  " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30))"

Call TmpTable

If Gen_DBType <> "MDB" Then
    tmp = " ALTER TABLE Mst_Stack WITH NOCHECK ADD "
    tmp = tmp & " CONSTRAINT PK_Mst_Stack PRIMARY KEY CLUSTERED (StackID), "
    tmp = tmp & " CONSTRAINT FK_Mst_Stack_GodownID FOREIGN KEY (GodownID) REFERENCES Mst_Godown(GodownID), "
    tmp = tmp & " CONSTRAINT DF_Mst_Stack DEFAULT (GETDATE()) FOR Concurrency "
    Call TmpTable
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_Stack.frm','Stack Master','FrmMst_Stack','Y','M',0,227,1"
Call TmpTable

    
UpdateConfig ("281")
MsgBox "Version 281 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function version280()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Mst_GSL Alter Column MoistureValue decimal(18,3)"
Call TmpTable


UpdateConfig ("280")
MsgBox "Version 280 Installed successfully.", vbOKOnly + vbInformation, App.Title


End Function


Public Function Version279()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Mst_Commodity ADD MAWFlag " & TextType & "(1)  "
Call TmpTable
    
tmp = " ALTER TABLE Mst_GSL ADD MoistureValue decimal(18,0)"
Call TmpTable
    
tmp = " ALTER TABLE Mst_GSL ADD MAWWt decimal(18,3)"
Call TmpTable


tmp = " Create Table Mst_MAW("
tmp = tmp & " MAWID Decimal(18,0) NOT NULL,"
tmp = tmp & " CommodityID Decimal(18,0) NOT NULL,"
tmp = tmp & " MinimumRange Decimal(18,2) NOT NULL,"
tmp = tmp & " MaximumRange Decimal(18,2) NOT NULL,"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " Remarks " & TextType & "(200),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency Datetime Not Null,"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate Datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"

Call TmpTable
    
If Gen_DBType <> "MDB" Then
    ''--  Alter Table for Primary Key
    tmp = " ALTER TABLE Mst_MAW WITH NOCHECK ADD CONSTRAINT PK_Mst_MAW PRIMARY KEY CLUSTERED (MAWID)"
    Call TmpTable
    ''--  Alter Table for foreign Key Reference
    tmp = " ALTER TABLE Mst_MAW ADD CONSTRAINT FK_Mst_MAW_Mst_Commodity_CommodityID FOREIGN KEY (CommodityID) REFERENCES dbo.Mst_Commodity(CommodityID)"
    Call TmpTable
    ''--  Alter Table for Default Value
    tmp = " ALTER TABLE Mst_MAW ADD CONSTRAINT DF_Mst_MAW_Concurrency DEFAULT (GETDATE()) FOR Concurrency"
    Call TmpTable
End If
    
    
tmp = " Create Table Mst_MAWDetails("
tmp = tmp & " MAWDetailID Decimal(18,0) NOT NULL,"
tmp = tmp & " MAWID Decimal(18,0) NOT NULL,"
tmp = tmp & " MinimumMoisture Decimal(18,2) NOT NULL,"
tmp = tmp & " MaximumMoisture Decimal(18,2) NOT NULL,"
tmp = tmp & " MAWD Decimal(18,2) NOT NULL,"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " Remarks " & TextType & "(200),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency Datetime Not Null,"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate Datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"
    
Call TmpTable

If Gen_DBType <> "MDB" Then
    ''--  Alter Table for Primary Key
    tmp = " ALTER TABLE Mst_MAWDetails WITH NOCHECK ADD CONSTRAINT PK_Mst_MAWDetails PRIMARY KEY CLUSTERED (MAWDetailID)"
    Call TmpTable
    ''--  Alter Table for foreign Key Reference
    tmp = " ALTER TABLE Mst_MAWDetails ADD CONSTRAINT FK_Mst_MAWDetails_Mst_MAW_MAWID FOREIGN KEY (MAWID) REFERENCES dbo.Mst_MAW(MAWID)"
    Call TmpTable
    ''--  Alter Table for Default Value
    tmp = " ALTER TABLE Mst_MAWDetails ADD CONSTRAINT DF_Mst_MAWDetails_Concurrency DEFAULT (GETDATE()) FOR Concurrency"
    Call TmpTable
End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_MAW.frm','MAW Commodity Mapping','FrmMst_MAW','Y','M',0,227,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ImportGSLWiseMoisture.frm','Import GSL wise Moisture','Frm_ImportGSLWiseMoisture','Y','T',0,228,1"
Call TmpTable


    
UpdateConfig ("279")
MsgBox "Version 279 Installed successfully.", vbOKOnly + vbInformation, App.Title

End Function

Public Function Version278()


Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Txn_ImportNcdexLotPleadge"
tmp = tmp & " ("
tmp = tmp & " SrNo  decimal(18,0) Not null,"
tmp = tmp & " NCDEXCode " & TextType & " (1000),"
tmp = tmp & " CMSELotId " & TextType & " (25),"
tmp = tmp & " TransNo   " & TextType & " (25),"
tmp = tmp & " CMSEClientId  decimal(18,0),"
tmp = tmp & " CMSEClientName    " & TextType & " (1000),"
tmp = tmp & " CMSPLotId " & TextType & " (25),"
tmp = tmp & " ActualQuantity    decimal(18,3),"
tmp = tmp & " UOM   " & TextType & " (3),"
tmp = tmp & " FED   datetime,"
tmp = tmp & " Remarks   " & TextType & " (4000),"
tmp = tmp & " Type  " & TextType & " (1),"
tmp = tmp & " BatchID1  " & TextType & " (1000),"
tmp = tmp & " BatchID2  " & TextType & " (1000),"
tmp = tmp & " Flag  " & TextType & " (1),"
tmp = tmp & " ErrorMsg  " & TextType & " (4000),"
tmp = tmp & " RatePerUOM    decimal(18,3),"
tmp = tmp & " ValueCommodity    decimal(18,3),"
tmp = tmp & " ValidityDate  datetime,"
tmp = tmp & " ActualGrade   " & TextType & " (1000),"
tmp = tmp & " GradePremium  " & TextType & " (1000),"
tmp = tmp & " LotRefNo  " & TextType & " (1000),"
tmp = tmp & " PreviousColumnsDet    " & TextType & " (1),"
tmp = tmp & " Pledge    " & TextType & " (1),"
tmp = tmp & " Code  " & TextType & " (1000),"
tmp = tmp & " SerialNo  decimal(18,0),"
tmp = tmp & " SM_Prefix " & TextType & " (5),"
tmp = tmp & " PledgeID  decimal(18,0),"
tmp = tmp & " CARNo decimal(10,0),"
tmp = tmp & " GSLID decimal(18,0),"
tmp = tmp & " Bags  decimal(18,0),"
tmp = tmp & " Quantity  decimal(18,3),"
tmp = tmp & " G_UID " & TextType & " (36),"
tmp = tmp & " Concurrency   datetime,"
tmp = tmp & " CreatedBy " & TextType & " (30),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & " (30),"
tmp = tmp & " UpdatedDate   datetime,"
tmp = tmp & " ImportDate    datetime,"
tmp = tmp & " GodownID  decimal(18,0),"
tmp = tmp & " BranchID  decimal(18,0)"
tmp = tmp & " )"

Call TmpTable

tmp = "Alter Table Log_Txn_WHR_Detail Add Code " & TextType & "(20)"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_ImportNcdexLotPleadge WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_ImportNcdexLotPleadge PRIMARY KEY CLUSTERED (SrNo),"
   tmp = tmp & " CONSTRAINT DF_Txn_ImportNcdexLotPleadge DEFAULT (GETDATE()) FOR Concurrency"

  Call TmpTable
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ImportNCDEXLotPleadge.frm','Import NCDEX Lot For Pledge','Frm_ImportNCDEXLotPleadge','Y','T',0,228,1"
Call TmpTable


'-------



UpdateConfig ("278")
MsgBox "Version 278 Installed successfully"

End Function

Public Function Version277()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Create Table Mst_InspectionAreaDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionAreaID decimal(18,0) Not Null,"
tmp = tmp & " InspectionID decimal(18,0),"
tmp = tmp & " Length    decimal(18,2),"
tmp = tmp & " Width decimal(18,2),"
tmp = tmp & " Height    decimal(18,2),"
tmp = tmp & " Area    decimal(18,2), "
tmp = tmp & " Remarks Varchar(1000),"
tmp = tmp & " G_UID Varchar(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy Varchar(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy Varchar(30))"

Call TmpTable

If Gen_DBType <> "MDB" Then

   tmp = " ALTER TABLE Mst_InspectionAreaDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_InspectionAreaDetail PRIMARY KEY CLUSTERED (InspectionAreaID),"
   tmp = tmp & " CONSTRAINT DF_Mst_InspectionAreaDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionAreaDetail_InspectionID FOREIGN KEY (InspectionID) REFERENCES Mst_GodownInspection(InspectionID)"

   Call TmpTable

End If


UpdateConfig ("277")
MsgBox "Version 277 Installed successfully"


End Function

Public Function Version276()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Create Table Mst_VehicleType"
tmp = tmp & " (VehicleTypeID Decimal(18,0) Not Null,"
tmp = tmp & " VehicleType " & TextType & "(50),"
tmp = tmp & " MinTareWt decimal(18,3),"
tmp = tmp & " MaxTareWt decimal(18,3),"
tmp = tmp & " Remarks " & TextType & "(1000),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(36),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(36),"
tmp = tmp & " UpdatedDate DateTime)"

Call TmpTable


tmp = "Alter Table Txn_CxTF_TruckDetail Add VehicleTypeID Decimal(18,0)"
Call TmpTable


If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Mst_VehicleType WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_VehicleType PRIMARY KEY CLUSTERED (VehicleTypeID),"
   tmp = tmp & " CONSTRAINT DF_Mst_VehicleType_Concurrency DEFAULT (getdate()) FOR Concurrency"
   Call TmpTable

   tmp = "ALTER TABLE Txn_CxTF_TruckDetail ADD CONSTRAINT FK_Txn_CxTF_TruckDetail_Mst_VehicleType "
   tmp = tmp & " FOREIGN KEY (VehicleTypeID) REFERENCES Mst_VehicleType (VehicleTypeID)"
   Call TmpTable
   
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_VehicleType.frm','Vehicle Model','FrmMst_VehicleType','Y','M',0,227,1"
Call TmpTable


UpdateConfig ("276")
MsgBox "Version 276 Installed successfully"

End Function

Public Function Version275()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_InvMonGSLDetail Add GodownYearlyBillingFlg " & TextType & "(1),CommodityYearlyBillingFlg " & TextType & "(1)"
Call TmpTable


UpdateConfig ("275")
MsgBox "Version 275 Installed successfully"

End Function
Public Function Version274()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Mst_Godown Add YearlyBillingFlg " & TextType & "(1)"
Call TmpTable

tmp = "Alter Table Mst_Commodity Add YearlyBillingFlg " & TextType & "(1)"
Call TmpTable


tmp = "CREATE TABLE Mst_YearlyStorageRateCard("
tmp = tmp & " YSRCID decimal(18, 0) NOT NULL ,"
tmp = tmp & " LocationID decimal(18, 0) ,"
tmp = tmp & " CommodityID decimal(18, 0) ,"
tmp = tmp & " G_UID " & TextType & " (36)  ,"
tmp = tmp & " Flag " & TextType & " (1)  ,"
tmp = tmp & " Concurrency datetime NOT NULL ,"
tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
tmp = tmp & " CreatedDate datetime  ,"
tmp = tmp & " UpdatedDate datetime  ,"

If Gen_DBType = "MDB" Then
   tmp = tmp & " UpdatedBy " & TextType & " (30), "
   tmp = " PRIMARY KEY (YSRCID))"
Else
   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
End If
Call TmpTable


tmp = "CREATE TABLE Mst_YearlyStorageRateCardDetail("
tmp = tmp & " YSRCDID decimal(18, 0) NOT NULL ,"
tmp = tmp & " YSRCID decimal(18, 0) NOT NULL ,"
tmp = tmp & " GodownID decimal(18, 0) ,"
tmp = tmp & " Rate decimal(18, 4), "
tmp = tmp & " G_UID " & TextType & " (36)  ,"
tmp = tmp & " Flag " & TextType & " (1)  ,"
tmp = tmp & " Concurrency datetime NOT NULL ,"
tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
tmp = tmp & " CreatedDate datetime  ,"
tmp = tmp & " UpdatedDate datetime  ,"

If Gen_DBType = "MDB" Then
   tmp = tmp & " UpdatedBy " & TextType & " (30), "
   tmp = " PRIMARY KEY (YSRCDID))"
Else
   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
End If
Call TmpTable

If Gen_DBType <> "MDB" Then
   
   tmp = "ALTER TABLE Mst_Godown WITH NOCHECK ADD CONSTRAINT DF_Mst_Godown_YearlyBillingFlg DEFAULT (0) FOR YearlyBillingFlg"
   Call TmpTable

   tmp = "ALTER TABLE Mst_Commodity WITH NOCHECK ADD CONSTRAINT DF_Mst_Commodity_YearlyBillingFlg DEFAULT (0) FOR YearlyBillingFlg"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_YearlyStorageRateCard WITH NOCHECK ADD "
   tmp = tmp & " CONSTRAINT PK_Mst_YearlyStorageRateCard PRIMARY KEY  CLUSTERED (YSRCID)"

   Call TmpTable

   tmp = " ALTER TABLE Mst_YearlyStorageRateCard ADD "
   tmp = tmp & " CONSTRAINT DF_Mst_YearlyStorageRateCard_Concurrency DEFAULT (getdate()) FOR Concurrency"

   Call TmpTable

   tmp = " ALTER TABLE Mst_YearlyStorageRateCard ADD "
   tmp = tmp & " CONSTRAINT FK_Mst_YearlyStorageRateCard_Mst_Location FOREIGN KEY "
   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Mst_YearlyStorageRateCard_Mst_Commodity FOREIGN KEY "
   tmp = tmp & " (CommodityID) REFERENCES Mst_Commodity (CommodityID)"
   Call TmpTable

   tmp = " ALTER TABLE Mst_YearlyStorageRateCardDetail WITH NOCHECK ADD "
   tmp = tmp & " CONSTRAINT PK_Mst_YearlyStorageRateCardDetail PRIMARY KEY  CLUSTERED (YSRCDID)"

   Call TmpTable

   tmp = " ALTER TABLE Mst_YearlyStorageRateCardDetail ADD "
   tmp = tmp & " CONSTRAINT DF_Mst_YearlyStorageRateCardDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"

   Call TmpTable

   tmp = " ALTER TABLE Mst_YearlyStorageRateCardDetail ADD "
   tmp = tmp & " CONSTRAINT FK_Mst_YearlyStorageRateCardDetail_Mst_Godown FOREIGN KEY "
   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
   tmp = tmp & " CONSTRAINT FK_Mst_YearlyStorageRateCardDetail_Mst_YearlyStorageRateCard FOREIGN KEY "
   tmp = tmp & " (YSRCID) REFERENCES Mst_YearlyStorageRateCard (YSRCID)"


   Call TmpTable

End If

tmp = "Update Mst_Godown Set YearlyBillingFlg = 0"
Call TmpTable

tmp = "Update Mst_Commodity Set YearlyBillingFlg = 0"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_StorageYearlyRateCard.frm','Yearly Storage Rate Card','FrmMst_StorageYearlyRateCard','Y','T',0,228,1"
Call TmpTable


UpdateConfig ("274")
MsgBox "Version 274 Installed successfully"

End Function


Public Function Version273()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_QualityInformation Add DetailsEnteredFlg " & TextType & "(1)"
Call TmpTable

tmp = "Update Txn_QualityInformation Set DetailsEnteredFlg = 'L'"
Call TmpTable

tmp = "Alter Table Txn_QualityInformationDetail Add BalanceBags Decimal(18,0)"
Call TmpTable

UpdateConfig ("273")
MsgBox "Version 273 Installed successfully"

End Function

Public Function Version272()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_RentServiceTaxDetails.frm','Rent Service Tax Details','FrmRpt_RentServiceTaxDetails','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("272")
MsgBox "Version 272 Installed successfully"

End Function


Public Function Version271()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_LotWiseDepositWithdrawalDetail.frm','Lot Wise Deposit Withdrawal Detail','FrmRpt_LotWiseDepositWithdrawalDetail','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("271")
MsgBox "Version 271 Installed successfully"
End Function

'-----Version 270 For Store Man.

Public Function Version270()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Mst_GSL Add CITF_MoisturePer Decimal(18,3)"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownAgreeDocCycle.frm','Godown Agreement Document Cycle Details','FrmRpt_GodownAgreeDocCycle','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("270")
MsgBox "Version 270 Installed successfully"
End Function


Public Function Version269()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Log_FranchiseeFeeRejectHistory Alter Column LotNo " & TextType & "(25)"
Call TmpTable

UpdateConfig ("269")
MsgBox "Version 269 Installed successfully"
End Function
Public Function Version268()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Txn_PledgeCreation Add SCMFlag Varchar(1)"
Call TmpTable

tmp = "Alter Table Txn_PledgeRelease Add SCMFlag Varchar(1)"
Call TmpTable


tmp = "Create Table Txn_SCMDO"
tmp = tmp & " ("
tmp = tmp & " DOID Decimal(18,0) Not null,"
tmp = tmp & " DODate DateTime ,"
tmp = tmp & " LocationID Decimal(18,0),"
tmp = tmp & " WHRNo Decimal(18,0),"
tmp = tmp & " DOBags Decimal(18,0),"
tmp = tmp & " DOQty Decimal(18,3),"
tmp = tmp & " PrevDOBags Decimal(18,0),"
tmp = tmp & " PrevDOQty Decimal(18,3),"
tmp = tmp & " BalanceBags Decimal(18,0),"
tmp = tmp & " BalanceQty Decimal(18,3),"
tmp = tmp & " Flag Varchar(1),"
tmp = tmp & " Remarks Varchar(1500),"
tmp = tmp & " ApprovalRemarks Varchar(1500),"
tmp = tmp & " G_UID Varchar(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy Varchar(30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedBy Varchar(30),"
tmp = tmp & " UpdatedDate DateTime)"

Call TmpTable

If Gen_DBType <> "MDB" Then

   tmp = " ALTER TABLE Txn_SCMDO WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_SCMDO PRIMARY KEY CLUSTERED (DOID),"
   tmp = tmp & " CONSTRAINT DF_Txn_SCMDO_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_SCMDO_LocationID FOREIGN KEY (LocationID) REFERENCES Mst_Location(LocationID)"

   Call TmpTable

End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_SCMDO.frm','SCM DO','FrmTxn_SCMDO','Y','T',0,228,1"

Call TmpTable


UpdateConfig ("268")
MsgBox "Version 268 Installed successfully"


End Function

Public Function Version267()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Txn_ProInvoice Drop Constraint PK_Txn_ProInvoice"
Call TmpTable

tmp = " ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD CONSTRAINT PK_Txn_ProInvoice PRIMARY KEY CLUSTERED (InvoiceNo)"
Call TmpTable

UpdateConfig ("267")
MsgBox "Version 267 Installed successfully"
End Function
Public Function Version266()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_GADCGodownDetail Alter column GodownNO " & TextType & " (25)"
Call TmpTable

UpdateConfig ("266")
MsgBox "Version 266 Installed successfully"
End Function

Public Function Version265()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_LicenseWiseGodownDetail.frm','License Wise Godown Detail','FrmRpt_LicenseWiseGodownDetail','Y','R',0,229,1"
Call TmpTable
tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_RRNCxtfDetail.frm','RRN CITF Detail','FrmRpt_RRNCxtfDetail','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("265")
MsgBox "Version 265 Installed successfully"
End Function
Public Function Version264()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create Table Mst_LicenseType "
tmp = tmp & " ("
tmp = tmp & " LicenseTypeID decimal(18,0),"
tmp = tmp & " LicenseTypeName " & TextType & "(100),"
tmp = tmp & " Flag    " & TextType & "(1),"
tmp = tmp & " G_UID   " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30) Primary Key (LicenseTypeID)"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "Alter Table Mst_LicenseType with Nocheck Add"
   tmp = tmp & " Constraint DF_Mst_LicenseType default (getdate()) For Concurrency"
   Call TmpTable
End If

tmp = " Alter Table Mst_License Add EmployeeID Decimal(18,0), LicenseTypeID Decimal(18,0) ,Capacity Decimal(18,0)"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "Alter Table Mst_License with Nocheck Add"
   tmp = tmp & " CONSTRAINT FK_Mst_LicenseMst_Employee FOREIGN KEY (EmployeeID) REFERENCES Mst_Employee (EmployeeID),"
   tmp = tmp & " CONSTRAINT FK_Mst_LicenseMst_LicenseType FOREIGN KEY (LicenseTypeID) REFERENCES Mst_LicenseType (LicenseTypeID)"
   Call TmpTable
End If

tmp = " Create Table Txn_LicenseAllocationDetail"
tmp = tmp & " ("
tmp = tmp & " AllocateID  decimal(18,0),"
tmp = tmp & " LicenseID   decimal(18,0),"
tmp = tmp & " Flag    " & TextType & "(1),"
tmp = tmp & " SLGID   decimal(18,0),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " G_UID   " & TextType & "(36) Primary Key (AllocateID)"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "Alter Table Txn_LicenseAllocationDetail with Nocheck Add"
   tmp = tmp & " Constraint DF_Txn_LicenseAllocationDetail default (getdate()) For Concurrency,"
   tmp = tmp & " Constraint FK_Txn_LicenseAllocationDetailMst_License FOREIGN KEY (LicenseID) REFERENCES Mst_License (LicenseID)"
   Call TmpTable
End If

tmp = " Create Table Txn_LicenseGodownDetail"
tmp = tmp & " ("
tmp = tmp & " ID  decimal(18,0),"
tmp = tmp & " LicenseID decimal(18,0),"
tmp = tmp & " GodownID decimal(18,0),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " UpdateBy    " & TextType & "(30),"
tmp = tmp & " UpdateDate  " & TextType & "(30),"
tmp = tmp & " G_UID   " & TextType & "(36),"
tmp = tmp & " Concurrency datetime Primary Key (ID)"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "Alter Table Txn_LicenseGodownDetail with Nocheck Add"
   tmp = tmp & " Constraint DF_Txn_LicenseGodownDetail default (getdate()) For Concurrency,"
   tmp = tmp & " Constraint FK_Txn_LicenseGodownDetailMst_License FOREIGN KEY (LicenseID) REFERENCES Mst_License (LicenseID)"
   Call TmpTable
End If

tmp = "Alter Table Mst_GSL Add MoisturePer Decimal(18,3)"
Call TmpTable

tmp = "Alter Table Txn_CxTF_TruckDetail Add MoisturePer Decimal(18,3)"
Call TmpTable

tmp = "Alter Table Txn_CxTF_GSL Add MoisturePer Decimal(18,3)"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_Mst_LicenseType.frm','License Type','Frm_Mst_LicenseType','Y','M',0,227,1"
Call TmpTable

UpdateConfig ("264")
MsgBox "Version 264 Installed successfully"

End Function


Public Function Version263()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_Client Add CRFID decimal(18,0)"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_QQSGodownMapping.frm','QQS Godown Mapping Report','FrmRpt_QQSGodownMapping','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("263")
MsgBox "Version 263 Installed successfully"

End Function

Public Function Version262()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_SCMQualityParameter.frm','SCM Quality Parameter Detail Report','FrmRpt_SCMQualityParameter','Y','R',0,229,1"

Call TmpTable


UpdateConfig ("262")
MsgBox "Version 262 Installed successfully"


End Function

Public Function Version261()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Log_GADCForwardToHistory Add DocumentTypeID Decimal(18,0)"
Call TmpTable


UpdateConfig ("261")
MsgBox "Version 261 Installed successfully"


End Function

Public Function Version260()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create Table Mst_DocumentType"
tmp = tmp & " ("
tmp = tmp & " DocumentTypeID Decimal(18,0) Not Null,"
tmp = tmp & " DocumentType  " & TextType & "(500),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " G_UID  " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy  " & TextType & " (30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedBy  " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime Primary Key(DocumentTypeID)"
tmp = tmp & ")"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Mst_DocumentType WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Mst_DocumentType DEFAULT (GETDATE()) FOR Concurrency"
   
   Call TmpTable
End If

tmp = " create table Txn_GodownAgreeDocCycle"
tmp = tmp & " ("
tmp = tmp & " AgreementID decimal(18,0) Not Null,"
tmp = tmp & " EntryDate Datetime,"
tmp = tmp & " StartDate Datetime,"
tmp = tmp & " ExpiryDate Datetime,"
tmp = tmp & " AgreementActualStartDate Datetime,"
tmp = tmp & " LocationID decimal (18,0),"
tmp = tmp & " UnitTypeID decimal (18,0),"
tmp = tmp & " Status " & TextType & "(1),"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime Primary Key (AgreementID)"

tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Txn_GodownAgreeDocCycle WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_GodownAgreeDocCycle_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_GodownAgreeDocCycle_LocationID FOREIGN KEY (LocationID) REFERENCES Mst_Location(LocationID),"
   tmp = tmp & " CONSTRAINT FK_Txn_GodownAgreeDocCycle_UnitTypeID FOREIGN KEY (UnitTypeID) REFERENCES Mst_UnitType(UnitTypeID)"
   Call TmpTable
End If

tmp = " create table Txn_GADCDocumentAttachDetail"
tmp = tmp & " ("
tmp = tmp & " DocDID decimal(18,0) Not Null,"
tmp = tmp & " AgreementID decimal(18,0) ,"
tmp = tmp & " DocumentTypeID decimal(18,0) ,"
tmp = tmp & " [FileName] " & TextType & "(4000),"
tmp = tmp & " [File] image,"
tmp = tmp & " [Desc] " & TextType & "(4000),"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " ForwardID decimal(18,0) ,"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime  Primary Key (DocDID)"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_GADCDocumentAttachDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_GADCDocumentAttachDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_GADCDocumentAttachDetail_AgreementID FOREIGN KEY (AgreementID) REFERENCES Txn_GodownAgreeDocCycle(AgreementID),"
   tmp = tmp & " CONSTRAINT FK_Txn_GADCDocumentAttachDetail_DocumentTypeID FOREIGN KEY (DocumentTypeID) REFERENCES Mst_DocumentType(DocumentTypeID),"
   tmp = tmp & " CONSTRAINT FK_Txn_GADCDocumentAttachDetail_ForwardID FOREIGN KEY (ForwardID) REFERENCES Mst_Employee(EmployeeID)"
   Call TmpTable
End If

tmp = " create table Txn_GADCGodownDetail"
tmp = tmp & " ("
tmp = tmp & " GodownDID decimal(18,0) Not Null,"
tmp = tmp & " AgreementID decimal(18,0) ,"
tmp = tmp & " InspectionID decimal(18,0) ,"
tmp = tmp & " GodownNo decimal(18,0) ,"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime  Primary Key (GodownDID)"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Txn_GADCGodownDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_GADCGodownDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_GADCGodownDetail_AgreementID FOREIGN KEY (AgreementID) REFERENCES Txn_GodownAgreeDocCycle(AgreementID)"
   Call TmpTable
End If

tmp = "create table Txn_GADCAgreementAttachDetail"
tmp = tmp & " ("
tmp = tmp & " AgreementDID decimal(18,0) Not Null,"
tmp = tmp & " AgreementID decimal(18,0) ,"
tmp = tmp & " [FileName] " & TextType & "(4000),"
tmp = tmp & " [File] image,"
tmp = tmp & " [Desc] " & TextType & "(4000),"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime  Primary Key (AgreementDID)"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Txn_GADCAgreementAttachDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_GADCAgreementAttachDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_GADCAgreementAttachDetail_AgreementID FOREIGN KEY (AgreementID) REFERENCES Txn_GodownAgreeDocCycle(AgreementID)"
   Call TmpTable
End If

tmp = "CREATE TABLE Log_GADCForwardToHistory ("
tmp = tmp & " AgreementRowID decimal(18, 0) NOT NULL ,"
tmp = tmp & " OldForwardToID decimal(18, 0) NULL ,"
tmp = tmp & " NewForwardToID decimal(18, 0) NULL ,"
tmp = tmp & " UpdatedBy " & TextType & " (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,"
tmp = tmp & " Concurrency datetime NOT NULL CONSTRAINT DF_Log_GADCForwardToHistory DEFAULT (getdate()))"

Call TmpTable

'tmp = "CREATE TRIGGER tr_Txn_GodownAgreeDocCycle_U ON Txn_GodownAgreeDocCycle AFTER UPDATE"
'tmp = tmp & " AS"
'tmp = tmp & " DECLARE @mAgreementRowID  decimal(18, 0),"
'tmp = tmp & " @mOldForwardToID decimal(18, 0),"
'tmp = tmp & " @mNewForwardToID decimal(18, 0),"
'tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
'tmp = tmp & " SELECT  @mAgreementRowID= d.AgreementID,"
'tmp = tmp & " @mOldForwardToID = d.ForwardEmpolyeeID,"
'tmp = tmp & " @mNewForwardToID = i.ForwardEmpolyeeID,"
'tmp = tmp & " @mUpdatedBy= i.UpdatedBy"
'tmp = tmp & " FROM    inserted i"
'tmp = tmp & " INNER JOIN deleted d ON i.AgreementID = d.AgreementID"
'tmp = tmp & " insert into Log_GADCForwardToHistory (AgreementRowID ,OldForwardToID ,NewForwardToID ,UpdatedBy)"
'tmp = tmp & " values(@mAgreementRowID ,@mOldForwardToID ,@mNewForwardToID ,@mUpdatedBy)"
'
'Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_GodownAgreeDoc.frm','Godown Agreement Document Cycle','FrmTxn_GodownAgreeDoc','Y','T',0,228,1"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_DocumentType.frm','Document Type','FrmMst_DocumentType','Y','M',0,227,1"

Call TmpTable

UpdateConfig ("260")
MsgBox "Version 260 Installed successfully"

End Function


Public Function version259()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_NewCashReceipt Add FlagApproval " & TextType & "(3),ApprovalRemark " & TextType & "(4000),ApproveDate Datetime,ApproveBy " & TextType & "(30)"
Call TmpTable

UpdateConfig ("259")
MsgBox "Version 259 Installed successfully"

End Function
Public Function version258()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Create table Txn_QQSGodownMapping"
tmp = tmp & "("
tmp = tmp & " QQSMapID decimal(18,0) Not Null,"
tmp = tmp & " EntryDate Datetime,"
tmp = tmp & " EmployeeID decimal (18,0),"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " FreeFlag " & TextType & "(1),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime Primary Key (QQSMapID)"
tmp = tmp & ")"

Call TmpTable
   
If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_QQSGodownMapping WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_QQSGodownMapping_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_QQSGodownMapping_EmpolyeeID FOREIGN KEY (EmployeeID) REFERENCES Mst_Employee(EmployeeID)"
   Call TmpTable
End If

tmp = " Create table Txn_QQSGodownMappingDetail"
tmp = tmp & " ("
tmp = tmp & " QQSMapDID decimal(18,0) ,"
tmp = tmp & " QQSMapID decimal(18,0) ,"
tmp = tmp & " GodownID decimal (18,0),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime Primary Key (QQSMapDID)"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Txn_QQSGodownMappingDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_QQSGodownMappingDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_QQSGodownMappingDetail_QQSMapID FOREIGN KEY (QQSMapID) REFERENCES Txn_QQSGodownMapping(QQSMapID),"
   tmp = tmp & " CONSTRAINT FK_Txn_QQSGodownMappingDetail_GodownID FOREIGN KEY (GodownID) REFERENCES Mst_Godown(GodownID)"
   Call TmpTable
End If

tmp = " CREATE TABLE Log_Txn_QQSGodownMappingDetail ("
tmp = tmp & " ID decimal(18, 0) NOT NULL ,"
tmp = tmp & " QQSMAPID decimal(18, 0) NOT NULL ,"
tmp = tmp & " GodownID  " & TextType & "(50) NOT NULL ,"
tmp = tmp & " UpdatedBy  " & TextType & "(30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,"
tmp = tmp & " Concurrency datetime NOT NULL CONSTRAINT DF_Log_Txn_QQSGodownMappingDetail DEFAULT (getdate()))"

Call TmpTable

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_QQSGodownMapping.frm','QQS Godown Mapping','FrmTxn_QQSGodownMapping','Y','T',0,228,1"

Call TmpTable

UpdateConfig ("258")
MsgBox "Version 258 Installed successfully"

End Function

Public Function version257()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Create Table Txn_ScmQualityParameter"
tmp = tmp & " ("
tmp = tmp & " ID decimal(18,0) Not Null,"
tmp = tmp & " EntryDate Datetime,"
tmp = tmp & " LocationID decimal (18,0),"
tmp = tmp & " SM_Prefix " & TextType & "(5),"
tmp = tmp & " Status " & TextType & "(1),"
tmp = tmp & " CxTFNo decimal (18,0),"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime Primary Key (ID)"
tmp = tmp & " )"

Call TmpTable
   
If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_ScmQualityParameter WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_ScmQualityParameter_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_ScmQualityParameter_LocationID FOREIGN KEY (LocationID) REFERENCES Mst_Location(LocationID)"
   Call TmpTable
End If

tmp = " Create Table Txn_ScmQualityParameterDetail"
tmp = tmp & " ("
tmp = tmp & " DID decimal(18,0) ,"
tmp = tmp & " ID decimal(18,0) ,"
tmp = tmp & " TruckNo " & TextType & "(50),"
tmp = tmp & " WeighBridgeSlipno  " & TextType & "(50),"
tmp = tmp & " TotalWeight decimal(18,0) ,"
tmp = tmp & " TareWeight decimal(18,0) ,"
tmp = tmp & " GrossWeight decimal(18,0) ,"
tmp = tmp & " NoOfBags decimal(18,0) ,"
tmp = tmp & " Remarks " & TextType & "(4000),"
tmp = tmp & " QualityParameterID decimal(18,0) ,"
tmp = tmp & " Value decimal(18,2) ,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(35),"
tmp = tmp & " CreatedDate   datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(35),"
tmp = tmp & " UpdatedDate datetime Primary Key (DID)"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Txn_ScmQualityParameterDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_ScmQualityParameterDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_ScmQualityParameterDetail_ID FOREIGN KEY (ID) REFERENCES Txn_ScmQualityParameter(ID)"
   Call TmpTable
End If

tmp = "CREATE TABLE Log_TxnScmQualityParameterDetail ("
tmp = tmp & " ID decimal(18, 0) NOT NULL ,"
tmp = tmp & " TruckNo " & TextType & "(50) NOT NULL ,"
tmp = tmp & " WeighBridgeSlipno " & TextType & "(50) NOT NULL ,"
tmp = tmp & " QualityParameterID decimal(18, 0) NULL ,"
tmp = tmp & " OldValue decimal(18, 2) NULL ,"
tmp = tmp & " UpdatedBy " & TextType & "(30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,"
tmp = tmp & " Concurrency datetime NOT NULL CONSTRAINT DF_Log_TxnScmQualityParameterDetail DEFAULT (getdate()))"

Call TmpTable

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ScmDepQualityParameter.frm','SCM Depositor Quality Parameter Mapping','FrmTxn_ScmDepQualityParameter','Y','T',0,228,1"

Call TmpTable

UpdateConfig ("257")
MsgBox "Version 257 Installed successfully"

End Function

Public Function version256()


Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If



tmp = " Create Table Txn_MonGodownWiseProfitabilityFreezed"
tmp = tmp & " ("
tmp = tmp & " MGWPID Decimal(18,0) not null,"
tmp = tmp & " MonthYear DateTime,"
tmp = tmp & " StateID   Decimal(18,0),"
tmp = tmp & " LocationID  Decimal(18,0),"
tmp = tmp & " AM " & TextType & " (5000),"
tmp = tmp & " RM " & TextType & " (5000),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " Capacity Decimal(18,3),"
tmp = tmp & " NcdexStock Decimal(18,3),"
tmp = tmp & " NNStock Decimal(18,3),"
tmp = tmp & " NNNRIncomeAdhoc Decimal(18,2),"
tmp = tmp & " NNNRIncome Decimal(18,2),"
tmp = tmp & " ReservationIncomeAdhoc Decimal(18,2),"
tmp = tmp & " ReservationIncome Decimal(18,2),"
tmp = tmp & " SevenDaysIncome Decimal(18,2),"
tmp = tmp & " DPIncomeAdHoc Decimal(18,2),"
tmp = tmp & " FinalDPIncome Decimal(18,2),"
tmp = tmp & " DPIncomeGSL Decimal(18,2),"
tmp = tmp & " Diff Decimal(18,2),"
tmp = tmp & " DPIncomeInvoice Decimal(18,2),"
tmp = tmp & " RejectedIncome Decimal(18,2),"
tmp = tmp & " SpotIncomeAdHoc Decimal(18,2),"
tmp = tmp & " OtherIncome Decimal(18,2),"
tmp = tmp & " SCMIncome Decimal(18, 2),"
tmp = tmp & " CMUnbilledIncome Decimal(18,2),"
tmp = tmp & " CMPendingIncome Decimal(18,2),"
tmp = tmp & " CMApprovedIncome Decimal(18,2),"
tmp = tmp & " LeftOverStock Decimal(18,2),"
tmp = tmp & " PledgeIncome Decimal(18,2),"
tmp = tmp & " LeaseRentPaid Decimal(18,2),"
tmp = tmp & " LeaseRentFAG Decimal(18,2),"
tmp = tmp & " LeaseRentDRC Decimal(18,2),"
tmp = tmp & " EmpClaimFAG Decimal(18,2),"
tmp = tmp & " EmpClaimAccrual Decimal(18,2),"
tmp = tmp & " VendorExpenses Decimal(18,2),"
tmp = tmp & " VendorExpensesAccrued Decimal(18,2),"
tmp = tmp & " SecurityFeesPaid Decimal(18,2),"
tmp = tmp & " SecurityFeesAccrual Decimal(18,2),"
tmp = tmp & " FumigationFeesVendor Decimal(18,2),"
tmp = tmp & " FumigationFeesPaid Decimal(18,2),"
tmp = tmp & " FumigationFeesAccrual Decimal(18,2),"
tmp = tmp & " FranchiseeFeesPaid Decimal(18,2),"
tmp = tmp & " FranchiseeFeesFAG Decimal(18,2),"
tmp = tmp & " FranchiseeFeesCMG Decimal(18,2),"
tmp = tmp & " Salary Decimal(18,2),"
tmp = tmp & " NCDEXCharges Decimal(18,2),"
tmp = tmp & " CoExpense Decimal(18,2),"
tmp = tmp & " InsurancePremium Decimal(18,2),"
tmp = tmp & " InterestCharges Decimal(18,2),"
tmp = tmp & " ReportFlag  " & TextType & " (1),"
tmp = tmp & " G_UID  " & TextType & " (36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy  " & TextType & " (30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy  " & TextType & " (30),"
tmp = tmp & " SalaryFieldStaff Decimal(18,2),"
tmp = tmp & " PartiallyReversedIncome Decimal(18,2)) "

Call TmpTable
   
If Gen_DBType <> "MDB" Then

   tmp = " ALTER TABLE Txn_MonGodownWiseProfitabilityFreezed WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_MonGodownWiseProfitabilityFreezed PRIMARY KEY CLUSTERED (MGWPID),"
   tmp = tmp & " CONSTRAINT DF_Txn_MonGodownWiseProfitabilityFreezed_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_MonGodownWiseProfitabilityFreezed_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Txn_MonGodownWiseProfitabilityFreezed_Mst_State FOREIGN KEY (StateID) REFERENCES Mst_State (StateID),"
   tmp = tmp & " CONSTRAINT FK_Txn_MonGodownWiseProfitabilityFreezed_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID)"

   Call TmpTable

End If


tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_GodownPnLFreezing.frm','Godown P&L Freezing','FrmTxn_GodownPnLFreezing','Y','T',0,228,1"

Call TmpTable


UpdateConfig ("256")
MsgBox "Version 256 Installed successfully"

End Function



Public Function version255()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If


tmp = "Alter Table Mst_Config Add BlockValueNNTransfer decimal(18,0),BlockValueNCDEXTransfer decimal(18,0)"
Call TmpTable

tmp = "Update Mst_Config Set BlockValueNNTransfer = 5,BlockValueNCDEXTransfer = 5"
Call TmpTable



UpdateConfig ("255")
MsgBox "Version 255 Installed successfully"

End Function


Public Function version254()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_LogViewInvoice.frm','Invoice Printed Log','FrmRpt_LogViewInvoice','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("254")
MsgBox "Version 254 Installed successfully"

End Function
Public Function version253()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_MonthlyCommtrackFileDetail Add CommodityCode Varchar(200),ChargeableQty Decimal(18,4),TranxType Varchar(150),TranxNo Varchar(150),TranxDate DateTime,ChargeableDuration Decimal(18,0),WSPCode Varchar(200),Rate Decimal(18,4)"
Call TmpTable


UpdateConfig ("253")
MsgBox "Version 253 Installed successfully"

End Function
Public Function version252()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GSLDiscrepancies.frm','GSL Discrepancies','FrmRpt_GSLDiscrepancies','Y','R',0,229,1"

Call TmpTable


UpdateConfig ("252")
MsgBox "Version 252 Installed successfully"

End Function
Public Function version251()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GSLCDTFvsTRFDate.frm','GSL CDTF vs TRF vs KR Date','FrmRpt_GSLCDTFvsTRFDate','Y','R',0,229,1"

Call TmpTable


UpdateConfig ("251")
MsgBox "Version 251 Installed successfully"

End Function

Public Function version250()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Txn_MonGodownWiseProfitability Add PartiallyReversedIncome Decimal(18,2)"
Call TmpTable

tmp = " Alter Table Txn_Spraying Add LastSprayingDate DateTime"
Call TmpTable

tmp = " Update Txn_Spraying Set LastSprayingDate = SprayingDate"
Call TmpTable


UpdateConfig ("250")
MsgBox "Version 250 Installed successfully"

End Function

Public Function version249()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Create table Txn_GodownRentST ("
tmp = tmp & " GodownRentSTID decimal(18,0) not null,"
tmp = tmp & " GodownRentSTDate Datetime,"
tmp = tmp & " LocationID decimal(18,0),"
tmp = tmp & " LessorID decimal(18,0),"
tmp = tmp & " MonthYear Datetime,"
tmp = tmp & " BillDate Datetime,"
tmp = tmp & " BillNo " & TextType & "(100),"
tmp = tmp & " BillAmount Decimal(18,2),"
tmp = tmp & " STRate Decimal(18,2),"
tmp = tmp & " STAmount Decimal(18,2),"
tmp = tmp & " Flag " & TextType & "(2),"
tmp = tmp & " BatchID decimal(18,0),"
tmp = tmp & " Remarks " & TextType & "(255),"
tmp = tmp & " FAGRemarks  " & TextType & " (255),"
tmp = tmp & " FAGPostingDate  datetime,"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " G_UID   " & TextType & "(36),"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate datetime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (GodownRentSTID)"
End If
tmp = tmp & " )"


Call TmpTable

tmp = " Create table Txn_GodownRentSTClaimDetail("
tmp = tmp & " GSTCID decimal(18,0) Not null,"
tmp = tmp & " GodownRentSTID decimal(18,0),"
tmp = tmp & " ClaimID Decimal(18,0),"
tmp = tmp & " Flag " & TextType & "(2),"
tmp = tmp & " Remarks " & TextType & "(255),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " G_UID   " & TextType & "(36)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (GSTCID)"
End If
tmp = tmp & " )"

Call TmpTable

tmp = " Create table Txn_GodownRentSTAgreeDetail("
tmp = tmp & " GSTAID decimal(18,0) not null,"
tmp = tmp & " GodownRentSTID decimal(18,0),"
tmp = tmp & " AgreementID Decimal(18,0),"
tmp = tmp & " AmountWithoutST Decimal(18,2),"
tmp = tmp & " AmountWithST Decimal(18,2),"
tmp = tmp & " Flag " & TextType & "(2),"
tmp = tmp & " Remarks " & TextType & "(255),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " G_UID   " & TextType & "(36)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (GSTAID)"
End If
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then

    tmp = " ALTER TABLE Txn_GodownRentST WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Txn_GodownRentST PRIMARY KEY CLUSTERED (GodownRentSTID),"
    tmp = tmp & " CONSTRAINT DF_Txn_GodownRentST DEFAULT (GETDATE()) FOR Concurrency,"
    tmp = tmp & " CONSTRAINT FK_Txn_GodownRentST_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location(LocationID),"
    tmp = tmp & " CONSTRAINT FK_Txn_GodownRentST_Mst_Lessor FOREIGN KEY (LessorID) REFERENCES Mst_Lessor(LessorID)"
       
    Call TmpTable
    
    tmp = " ALTER TABLE Txn_GodownRentSTClaimDetail WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Txn_GodownRentSTClaimDetail PRIMARY KEY CLUSTERED (GSTCID),"
    tmp = tmp & " CONSTRAINT DF_Txn_GodownRentSTClaimDetail DEFAULT (GETDATE()) FOR Concurrency,"
    tmp = tmp & " CONSTRAINT FK_Txn_GodownRentSTClaimDetail_Txn_GodownRent FOREIGN KEY (ClaimID) REFERENCES Txn_GodownRent(ClaimID)"
    
    Call TmpTable
    
    tmp = " ALTER TABLE Txn_GodownRentSTAgreeDetail WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Txn_GodownRentSTAgreeDetail PRIMARY KEY CLUSTERED (GSTAID),"
    tmp = tmp & " CONSTRAINT DF_Txn_GodownRentSTAgreeDetail DEFAULT (GETDATE()) FOR Concurrency,"
    tmp = tmp & " CONSTRAINT FK_Txn_GodownRentSTAgreeDetail_Txn_LessorAgreement FOREIGN KEY (AgreementID) REFERENCES Txn_LessorAgreement(LAID)"
    
    Call TmpTable
End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_RentServiceTax.frm','Godown Rent Service Tax','FrmTxn_RentServiceTax','Y','T',0,228,1"

Call TmpTable


UpdateConfig ("249")
MsgBox "Version 249 Installed successfully"

End Function

Public Function version248()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Create Table Txn_MonthlyInsuranceCharges"
tmp = tmp & " ("
tmp = tmp & " TxnID Decimal(18,0) Not Null,"
tmp = tmp & " TxnDate DateTime,"
tmp = tmp & " YearMonth DateTime,"
tmp = tmp & " InsuranceAmount Decimal(18,2),"
tmp = tmp & " ActualAmount Decimal(18,2),"
tmp = tmp & " Remarks " & TextType & " (1000),"
tmp = tmp & " G_UID " & TextType & " (36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy " & TextType & " (30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & " (30),"
tmp = tmp & " UpdatedDate DateTime)"

Call TmpTable

tmp = " Create Table Txn_MonthlyInsuranceChargesDetail"
tmp = tmp & " (DetailID Decimal(18,0) Not Null,"
tmp = tmp & " TxnID Decimal(18,0),"
tmp = tmp & " SM_Prefix " & TextType & " (5),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " CommodityID Decimal(18,0),"
tmp = tmp & " BalanceBags Decimal(18,0),"
tmp = tmp & " BalanceWeight Decimal(18,3),"
tmp = tmp & " CommodityRate Decimal(18,3),"
tmp = tmp & " AUM Decimal(18,2),"
tmp = tmp & " InsuranceAmount Decimal(18,2),"
tmp = tmp & " ActualAmount Decimal(18,2),"
tmp = tmp & " Remarks " & TextType & " (1000),"
tmp = tmp & " G_UID " & TextType & " (36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy " & TextType & " (30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & " (30),"
tmp = tmp & " UpdatedDate DateTime)"

Call TmpTable

If Gen_DBType = "MDB" Then

   tmp = " Alter Table Txn_MonthlyInsuranceCharges With Nocheck Add"
   tmp = tmp & " CONSTRAINT PK_Txn_MonthlyInsuranceCharges Primary key clustered(TxnID),"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyInsuranceCharges DEFAULT (GETDATE()) FOR Concurrency"
   Call TmpTable

   tmp = " Alter Table Txn_MonthlyInsuranceChargesDetail With Nocheck Add"
   tmp = tmp & " CONSTRAINT PK_Txn_MonthlyInsuranceChargesDetail Primary key clustered(DetailID),"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyInsuranceChargesDetail DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyInsuranceChargesDetailActualAmt DEFAULT (0) FOR ActualAmount,"
   tmp = tmp & " CONSTRAINT FK_Txn_MonthlyInsuranceChargesDetail_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID),"
   tmp = tmp & " CONSTRAINT FK_Txn_MonthlyInsuranceChargesDetail_Txn_MonthlyInsuranceCharges FOREIGN KEY (TxnID) REFERENCES Txn_MonthlyInsuranceCharges (TxnID),"
   tmp = tmp & " CONSTRAINT FK_Txn_MonthlyInsuranceChargesDetail_Mst_Commodity FOREIGN KEY (CommodityID) REFERENCES Mst_Commodity (CommodityID)"

   Call TmpTable

End If

tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_GenerateMonthlyInsuranceCharges.frm','Generate Monthly Insurance Charges','FrmTxn_GenerateMonthlyInsuranceCharges.frm','Y','T',0,228,1"

Call TmpTable


tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_GodownwiseInsuranceCharges.frm','Monthly Insurance Charges','FrmTxn_GodownwiseInsuranceCharges','Y','T',0,228,1"

Call TmpTable

UpdateConfig ("248")
MsgBox "Version 248 Installed successfully"


End Function

Public Function version247()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Create table Txn_TempInvGodownDetails"
tmp = tmp & " (TIGodownMapDetailsID decimal(18,0) Not null,"
tmp = tmp & " PINo Decimal(18,0),"
tmp = tmp & " GodownID decimal(18,0),"
tmp = tmp & " ReservationQty decimal(18,2),"
tmp = tmp & " G_UID  " & TextType & " (36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy  " & TextType & " (30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy  " & TextType & " (30),"
tmp = tmp & " InvoiceAmt decimal(18,2)"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType = "MDB" Then

   tmp = " ALTER TABLE Txn_TempInvGodownDetails WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_TempInvGodownDetails PRIMARY KEY CLUSTERED (TIGodownMapDetailsID),"
   tmp = tmp & " CONSTRAINT DF_Txn_TempInvGodownDetails DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_TempInvGodownDetails_Txn_TempInvoice FOREIGN KEY (PINo) REFERENCES Txn_TempInvoice (PINo)"

   Call TmpTable

End If

UpdateConfig ("247")
MsgBox "Version 247 Installed successfully"


End Function
Public Function version246()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_DR Add CommtrackFlag " & TextType & "(1)"
Call TmpTable


tmp = "Update Mst_DR Set CommtrackFlag = '0' Where NewISINID like '%INC%'"
Call TmpTable

tmp = "Update Mst_DR Set CommtrackFlag = '1' Where NewISINID not like '%INC%'"
Call TmpTable

tmp = " Alter Table Txn_InvMonGSLDetail add RRNDate datetime"
Call TmpTable

tmp = "Alter Table Mst_Config Add BlockingSecurityBilling decimal(18,0),BlockingFumigationBilling decimal(18,0)"
Call TmpTable

tmp = " Update Mst_Config  Set BlockingSecurityBilling = 30,BlockingFumigationBilling = 45"
Call TmpTable

tmp = "Alter Table Txn_Fumigation Add PrintedDate DateTime"
Call TmpTable

tmp = "Alter Table Txn_Spraying Add PrintedDate DateTime"
Call TmpTable

tmp = "Update Txn_Fumigation Set PrintedDate = FumigationDate"
Call TmpTable

tmp = "Update Txn_Spraying Set PrintedDate = SprayingDate"
Call TmpTable


UpdateConfig ("246")
MsgBox "Version 246 Installed successfully"

End Function
Public Function version245()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_Spraying Add BillableFlag " & TextType & "(1)"
Call TmpTable

tmp = "Update Txn_Spraying Set BillableFlag = 'Y'"
Call TmpTable

UpdateConfig ("245")
MsgBox "Version 245 Installed successfully"

End Function
Public Function version244()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Log_Print_History Alter Column DocType " & TextType & "(30) "
Call TmpTable

UpdateConfig ("244")
MsgBox "Version 244 Installed successfully"

End Function

Public Function version243()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Create Table Txn_SecurityAgreement("
tmp = tmp & " SAID  decimal(18,0) Not null,"
tmp = tmp & " EntryDate datetime,"
tmp = tmp & " SecurityAgencyID decimal(18,0),"
tmp = tmp & " StateID  decimal(18,0),"
tmp = tmp & " AgreeStartDate datetime,"
tmp = tmp & " AgreeExpiryDate   datetime,"
tmp = tmp & " ServiceTax " & TextType & "(1),"
tmp = tmp & " Remark  " & TextType & "(2000),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(36),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(36))"

Call TmpTable

tmp = " Create Table Txn_SecurityAgreementDetail"
tmp = tmp & " (SADID decimal(18,0) Not null,"
tmp = tmp & " SAID decimal(18,0),"
tmp = tmp & " SecurityTypeID decimal(18,0),"
tmp = tmp & " SecurityDesignationID decimal(18,0),"
tmp = tmp & " MonthlyRate decimal(18,2),"
tmp = tmp & " Remark " & TextType & "(2000),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy   " & TextType & "(36),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(36))"

Call TmpTable


If Gen_DBType <> "MDB" Then

   tmp = " ALTER TABLE Txn_SecurityAgreement WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_SecurityAgreement PRIMARY KEY CLUSTERED (SAID),"
   tmp = tmp & " CONSTRAINT DF_Txn_SecurityAgreement_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_SecurityAgreement_Mst_State FOREIGN KEY (StateID) REFERENCES Mst_State (StateID),"
   tmp = tmp & " CONSTRAINT FK_Txn_SecurityAgreement_Mst_SecurityAgency FOREIGN KEY (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID)"
   
   Call TmpTable

   tmp = " ALTER TABLE Txn_SecurityAgreementDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_SecurityAgreementDetailt PRIMARY KEY CLUSTERED (SADID),"
   tmp = tmp & " CONSTRAINT DF_Txn_SecurityAgreementDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_SecurityAgreementDetail_Txn_SecurityAgreement FOREIGN KEY (SAID) REFERENCES Txn_SecurityAgreement (SAID),"
   tmp = tmp & " CONSTRAINT FK_Txn_SecurityAgreementDetail_Mst_SecurityDesignation FOREIGN KEY (SecurityDesignationID) REFERENCES Mst_SecurityDesignation (SecurityDesignationID),"
   tmp = tmp & " CONSTRAINT FK_Txn_SecurityAgreementDetail_Mst_SecurityShifts FOREIGN KEY (SecurityTypeID) REFERENCES Mst_SecurityShifts (SecurityShiftsID)"
   
   Call TmpTable

End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_SecurityAgreement.frm','Security Agreement','FrmTxn_SecurityAgreement','Y','T',0,228,1"

Call TmpTable

UpdateConfig ("243")
MsgBox "Version 243 Installed successfully"

End Function


Public Function version242()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Mst_Client Add CorpClient " & TextType & "(1)"
Call TmpTable

tmp = "Update Mst_Client Set CorpClient = 0"
Call TmpTable

UpdateConfig ("242")
MsgBox "Version 242 Installed successfully"

End Function

Public Function version241()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If


tmp = "Alter Table Log_FranchiseeFeeRejectHistory Add InvoiceAmount Decimal(18,2)"

Call TmpTable


UpdateConfig ("241")
MsgBox "Version 241 Installed successfully"

End Function

Public Function version240()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If


tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_FranchiseeFeesApproval.frm','Franchisee Fees Approval Report','FrmRpt_FranchiseeFeesApproval','Y','R',0,229,1"

Call TmpTable


UpdateConfig ("240")
MsgBox "Version 240 Installed successfully"

End Function

Public Function version239()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownAggreementExpiry.frm','Godown Agreement Expiry','FrmRpt_GodownAggreementExpiry','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("239")
MsgBox "Version 239 Installed successfully"
End Function
Public Function version238()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Txn_LessorAgreement add AgreemenActualtStartDate DateTime"
Call TmpTable

tmp = "Update Txn_LessorAgreement set AgreemenActualtStartDate = StartDate"
Call TmpTable

UpdateConfig ("238")
MsgBox "Version 238 Installed successfully"

End Function
Public Function version237()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

'----Mst_Location
tmp = "Alter table Mst_Location Add RiskType " & TextType & "(10)"
Call TmpTable

'---Mst_Depository
tmp = "Alter table Mst_Depository Add RiskType " & TextType & "(10)"
Call TmpTable

'---Mst_Commodity
tmp = "Alter table Mst_Commodity Add RiskType " & TextType & "(10)"
Call TmpTable


'---Mst_Client
tmp = "Alter table Mst_Client Add RiskType " & TextType & "(10)"
Call TmpTable


UpdateConfig ("237")
MsgBox "Version 237 Installed successfully"
End Function

Public Function version236()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmRpt_StorageContractDetail.frm', 'Storage Contract Detail','FrmRpt_StorageContractDetail','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("236")
MsgBox "Version 236 Installed successfully"
End Function


Public Function version235()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create Table Txn_InsuranceExceptionMapping"
tmp = tmp & " ("
tmp = tmp & " IEMID Decimal(18,0) Not Null,"
tmp = tmp & " LocationID Decimal(18,0) Not Null ,"
tmp = tmp & " CommodityID Decimal(18,0),"
tmp = tmp & " ClientIDRow Decimal(18,0),"
tmp = tmp & " StartDate datetime,"
tmp = tmp & " EndDate datetime,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy  " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy  " & TextType & "(30)"
    
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,Primary Key(IEMID)"
End If

tmp = tmp & " )"

Call TmpTable


If Gen_DBType <> "MDB" Then
   tmp = " Alter Table Txn_InsuranceExceptionMapping With Nocheck Add"
   tmp = tmp & " CONSTRAINT PK_Txn_InsuranceExceptionMapping Primary key clustered(IEMID),"
   tmp = tmp & " CONSTRAINT DF_Txn_InsuranceExceptionMapping DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_InsuranceExceptionMapping_LocationID FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Txn_InsuranceExceptionMapping_CommodityID FOREIGN KEY (CommodityID) REFERENCES Mst_Commodity (CommodityID)"

   Call TmpTable

End If


tmp = " Alter Table Mst_GSL Add TransactionType " & TextType & "(10)"
Call TmpTable

tmp = " Alter Table Mst_GSL Add TransactionID Decimal(18,0)"
Call TmpTable

tmp = " Alter Table Mst_GSL Add  TransactionDate DateTime"
Call TmpTable

tmp = " Alter table Txn_CXTF_GSL add WeightGainLossGSL Decimal(18,3)"
Call TmpTable

tmp = " Alter table Txn_GSLADJTransaction add WeightGainLossGSL Decimal(18,3)"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_InsuranceExceptionMapping.frm','Insurance Exception Mapping','FrmTxn_InsuranceExceptionMapping','Y','T',0,228,1"

Call TmpTable



UpdateConfig ("235")
MsgBox "Version 235 Installed successfully"


End Function

Public Function version234()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If


tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_LogInvoiceDeletion.frm','Log Details for Deleted Invoices','FrmRpt_LogInvoiceDeletion','Y','R',0,229,1"

Call TmpTable


UpdateConfig ("234")
MsgBox "Version 234 Installed successfully"

End Function
Public Function version233()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Txn_MonGodownWiseProfitability Add SalaryFieldStaff Decimal(18,2)"
Call TmpTable
tmp = "Alter Table Txn_GSL_Transfer Add OldGodownID Decimal(18,0),NewGodownID Decimal(18,0)"
Call TmpTable
tmp = " Update Txn_GSL_Transfer Set OldGodownID = OldGSLID ,NewGodownID = NewGSLID Where Flag = 'G'"
Call TmpTable


UpdateConfig ("233")
MsgBox "Version 233 Installed successfully"

End Function

Public Function version232()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If


tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_EstimatedNCDXBOE.frm','Estimated Storage Charges For NCDEX Withdrawal','FrmTxn_EstimatedNCDXBOE','Y','T',0,228,1"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_InsuranceDataGodownWise.frm', 'Stock Details for given date Godown wise', 'FrmRpt_InsuranceDataGodownWise', 'Y', 'R',0,229,1"

Call TmpTable



UpdateConfig ("232")
MsgBox "Version 232 Installed successfully"

End Function
Public Function version231()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_Config Add BlockGodRentProcessDay Decimal(18,0)"
Call TmpTable

tmp = "Update Mst_Config Set BlockGodRentProcessDay = 25 "
Call TmpTable


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_OutStWithClientInvoiceDetails.frm', 'OutStanding Invoice Report', 'FrmRpt_OutStWithClientInvoiceDetails', 'Y', 'R',0,229,1"

Call TmpTable



UpdateConfig ("231")
MsgBox "Version 231 Installed successfully"

End Function
Public Function version230()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_Config Add BlockingSecurityDay Decimal(18,0)"
Call TmpTable

tmp = "Update Mst_Config Set BlockingSecurityDay = 25 "
Call TmpTable

UpdateConfig ("230")
MsgBox "Version 230 Installed successfully"

End Function
Public Function version229()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Mst_CA Add " & C & " ISINID " & TextType & "(25)"
Call TmpTable

tmp = "Create table Log_InvoiceDeletetion"
tmp = tmp & "("
tmp = tmp & "DocumentType  " & TextType & " (50),"
tmp = tmp & "DcoumentID Decimal(18,0),"
tmp = tmp & "SM_Prefix  " & TextType & " (20),"
tmp = tmp & "ClientIDRow Decimal(18,0),"
tmp = tmp & "Amount Decimal(18,2),"
tmp = tmp & "DeletetionRemark  " & TextType & " (2000),"
tmp = tmp & "DeletedBy  " & TextType & " (30),"
tmp = tmp & "DeletedOn DateTime"
tmp = tmp & ")"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CommTrackvsGSLIncome.frm','CommTrack vs GSL Income','FrmRpt_CommTrackvsGSLIncome','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("229")
MsgBox "Version 229 Installed successfully"

End Function

Public Function version228()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Txn_GodownRent Add PaymentdetailsUpdatedBy " & TextType & "(30),PaymentdetailsUpdatedOn DateTime "
Call TmpTable
tmp = " Alter Table Txn_FranchiseeRent Add PaymentdetailsUpdatedBy " & TextType & "(30),PaymentdetailsUpdatedOn DateTime"
Call TmpTable
tmp = " Alter Table Txn_FumigationApproval Add BankName " & TextType & "(100),ChequeNo " & TextType & "(100),PaidAmount decimal(18,2),TDS decimal(18,2),PaidDate datetime,PaymentdetailsUpdatedBy " & TextType & "(30),PaymentdetailsUpdatedOn DateTime"
Call TmpTable
tmp = " Alter Table Txn_SecurityBilling Add BankName " & TextType & "(100),ChequeNo " & TextType & "(100),PaidAmount decimal(18,2),TDS decimal(18,2),PaidDate datetime,PaymentdetailsUpdatedBy " & TextType & "(30),PaymentdetailsUpdatedOn DateTime"
Call TmpTable

tmp = "Alter Table Txn_InvMonGSLDetail Alter Column CommtrackIncome Decimal(18,4)"
Call TmpTable

UpdateConfig ("228")
MsgBox "Version 228 Installed successfully"
End Function

Public Function version227()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_InvMonGSLDetail Add CommtrackIncome Decimal(18,2)"
Call TmpTable

UpdateConfig ("227")
MsgBox "Version 227 Installed successfully"
End Function

Public Function version226()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " CREATE TABLE Mst_RRN_LotDetail"
tmp = tmp & " ("
tmp = tmp & " LotDetailID DECIMAL(18,0) Not Null,"
tmp = tmp & " RRNID DECIMAL(18,0),"
tmp = tmp & " ISINID " & TextType & "(25),"
tmp = tmp & " RematQty DECIMAL(18,6),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (DetailID)"
End If
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Mst_RRN_LotDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_RRN_LotDetail PRIMARY KEY CLUSTERED (LotDetailID),"
   tmp = tmp & " CONSTRAINT DF_Mst_RRN_LotDetail_CConcurrency DEFAULT (GETDATE()) FOR Concurrency"
   
   Call TmpTable
End If

tmp = "Alter Table Mst_RRN Add " & C & " CommTrackFlag Decimal(18,0)"
Call TmpTable


UpdateConfig ("226")
MsgBox "Version 226 Installed successfully"
End Function


Public Function version225()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmRpt_NCDEXCodewiseStock.frm', 'NCDEX Code wise Stock','FrmRpt_NCDEXCodewiseStock','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmTxn_UpdateStockOpBalance.frm', 'Update Stock Op Balance Access for Power User Only','FrmTxn_UpdateStockOpBalance','Y','T',0,228,1"
Call TmpTable


UpdateConfig ("225")
MsgBox "Version 225 Installed successfully"
End Function

Public Function version224()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = "Alter Table Txn_TempInvoice Add " & C & " ContractID Decimal(18,0)"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_TempInvoice WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Txn_TempInvoice_Mst_StorageContract FOREIGN KEY (ContractID) REFERENCES Mst_StorageContract (ContractID)"
   
   Call TmpTable
End If

UpdateConfig ("224")
MsgBox "Version 224 Installed successfully"

End Function

Public Function version223()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Create table Txn_TempInvSRGodownDetails"
tmp = tmp & " (TIGodownMapDetailsID decimal(18,0) Not null,"
tmp = tmp & " PINo Decimal(18,0),"
tmp = tmp & " GodownID decimal(18,0),"
tmp = tmp & " ReservationQty decimal(18,2),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " InvoiceAmt decimal(18,2)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (TIGodownMapDetailsID)"
End If
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_TempInvSRGodownDetails WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_TempInvSRGodownDetails PRIMARY KEY CLUSTERED (TIGodownMapDetailsID),"
   tmp = tmp & " CONSTRAINT DF_Txn_TempInvSRGodownDetails DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_TempInvSRGodownDetails_Txn_TempInvoice FOREIGN KEY (PINo) REFERENCES Txn_TempInvoice (PINo)"
   
   Call TmpTable
End If

UpdateConfig ("223")
MsgBox "Version 223 Installed successfully"
End Function
Public Function version222()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmRpt_CRTDSDetails.frm', 'CR TDS Adjustment Details','FrmRpt_CRTDSDetails','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("222")
MsgBox "Version 222 Installed successfully"
End Function
Public Function version221()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Txn_RRNPWRGSL Add NewTransferSM_Prefix " & TextType & "(5),NewTransferGSLID decimal (18,0) "
Call TmpTable


UpdateConfig ("221")
MsgBox "Version 221 Installed successfully"
End Function



Public Function version220()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Delete From Txn_MonthlyCommtrackInvoice "
Call TmpTable

tmp = " Delete From Txn_MonthlyCommtrackFileDetail"
Call TmpTable

tmp = " Alter Table Txn_MonthlyCommtrackFileDetail Alter Column Amount Decimal (18,4)"
Call TmpTable

tmp = " Alter Table Txn_MonthlyCommtrackFileDetail Alter Column AmtIncST Decimal (18,4)"
Call TmpTable

tmp = " Alter Table Txn_MonthlyCommtrackInvoice Alter Column Amt Decimal (18,4)"
Call TmpTable

tmp = " Alter Table Txn_MonthlyCommtrackInvoice Alter Column STAmt Decimal (18,4)"
Call TmpTable

tmp = " Alter Table Txn_MonthlyCommtrackInvoice Alter Column TotalAmt Decimal (18,4)"
Call TmpTable


UpdateConfig ("220")
MsgBox "Version 220 Installed successfully"
End Function

Public Function version219()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " UPDATE Mst_GodownInspection Set Status ='G' WHERE Status = 'A'"
Call TmpTable

tmp = " ALTER TABLE Mst_GodownInspection ADD " & C & " LicenseID DECIMAL(18,0)"
Call TmpTable

tmp = " ALTER TABLE LogMst_GodownInspection ADD " & C & " LicenseID DECIMAL(18,0)"
Call TmpTable

tmp = " ALTER TABLE LogMst_GodownInspection ADD Concurrency DATETIME "
Call TmpTable


If Gen_DBType <> "MDB" Then
   
   tmp = " ALTER TABLE Mst_GodownInspection WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Mst_GodownInspection_Mst_License FOREIGN KEY (LicenseID) REFERENCES Mst_License(LicenseID)"
   Call TmpTable
   
   tmp = " ALTER TABLE LogMst_GodownInspection WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_LogMst_GodownInspection_Concurrency DEFAULT (getdate()) FOR Concurrency"
   
   Call TmpTable
   
   
   tmp = " ALTER TRIGGER tr_Mst_GodownInspection ON Mst_GodownInspection"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mInspectionID Decimal(18,0),"
   tmp = tmp & " @mType " & TextType & "(2),"
   tmp = tmp & " @mLocationID Decimal(18,0),"
   tmp = tmp & " @mCMPID Decimal(18,0),"
   tmp = tmp & " @mGodownID Decimal(18,0),"
   tmp = tmp & " @mGodownNo " & TextType & "(25),"
   tmp = tmp & " @mGodownTypeID Decimal(18,0),"
   tmp = tmp & " @mGodownCategoryID Decimal(18,0),"
   tmp = tmp & " @mLicenseID Decimal(18,0),"
   tmp = tmp & " @mCapacity Decimal(18,3),"
   tmp = tmp & " @mAddress " & TextType & "(255),"
   tmp = tmp & " @mArea Decimal(18,2),"
   tmp = tmp & " @mStatus " & TextType & "(2),"
   tmp = tmp & " @mStartDate DateTime,"
   tmp = tmp & " @mCloseDate DateTime,"
   tmp = tmp & " @mnoofstaff decimal(18,0),"
   tmp = tmp & " @mstaffsalary decimal(18,2),"
   tmp = tmp & " @mnoofsec decimal(18,0),"
   tmp = tmp & " @msecsalary decimal(18,2),"
   tmp = tmp & " @mnoofstaffuser decimal(18,0),"
   tmp = tmp & " @mstaffsalaryuser decimal(18,2),"
   tmp = tmp & " @mnoofsecuser decimal(18,0),"
   tmp = tmp & " @msecsalaryuser decimal(18,2),"
   tmp = tmp & " @mRemarks " & TextType & "(5000),"
   tmp = tmp & " @mtotalrevenue decimal(18,2),"
   tmp = tmp & " @mdirectexpense decimal(18,2),"
   tmp = tmp & " @mgrossprofit decimal(18,2),"
   tmp = tmp & " @mindirectexpense decimal(18,2),"
   tmp = tmp & " @mnetprofit decimal(18,2),"
   tmp = tmp & " @mCapacityUtilisation Decimal(18,3),"
   tmp = tmp & " @mNCDEXStock Decimal(18,3),"
   tmp = tmp & " @mNNNRStock Decimal(18,3),"
   tmp = tmp & " @mReservationStock Decimal(18,3),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mInspectionID = d.InspectionID,"
   tmp = tmp & " @mType = d.Type,"
   tmp = tmp & " @mLocationID = d.LocationID,"
   tmp = tmp & " @mCMPID = d.CMPID,"
   tmp = tmp & " @mGodownID = d.GodownID,"
   tmp = tmp & " @mGodownNo = d.GodownNo,"
   tmp = tmp & " @mGodownTypeID = d.GodownTypeID,"
   tmp = tmp & " @mGodownCategoryID = d.GodownCategoryID,"
   tmp = tmp & " @mLicenseID = d.LicenseID,"
   tmp = tmp & " @mCapacity = d.Capacity,"
   tmp = tmp & " @mAddress = d.Address,"
   tmp = tmp & " @mArea = d.Area,"
   tmp = tmp & " @mStatus = d.Status,"
   tmp = tmp & " @mStartDate = d.StartDate,"
   tmp = tmp & " @mCloseDate = d.CloseDate,"
   tmp = tmp & " @mnoofstaff = d.noofstaff,"
   tmp = tmp & " @mstaffsalary = d.staffsalary,"
   tmp = tmp & " @mnoofsec = d.noofsec,"
   tmp = tmp & " @msecsalary = d.secsalary,"
   tmp = tmp & " @mnoofstaffuser = d.noofstaffuser,"
   tmp = tmp & " @mstaffsalaryuser = d.staffsalaryuser,"
   tmp = tmp & " @mnoofsecuser = d.noofsecuser,"
   tmp = tmp & " @msecsalaryuser = d.secsalaryuser,"
   tmp = tmp & " @mRemarks = d.Remarks,"
   tmp = tmp & " @mtotalrevenue = d.totalrevenue,"
   tmp = tmp & " @mdirectexpense = d.directexpense,"
   tmp = tmp & " @mgrossprofit = d.grossprofit,"
   tmp = tmp & " @mindirectexpense = d.indirectexpense,"
   tmp = tmp & " @mnetprofit = d.netprofit,"
   tmp = tmp & " @mCapacityUtilisation = d.CapacityUtilisation,"
   tmp = tmp & " @mNCDEXStock = d.NCDEXStock,"
   tmp = tmp & " @mNNNRStock = d.NNNRStock,"
   tmp = tmp & " @mReservationStock = d.ReservationStock,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.InspectionID = d.InspectionID"
   tmp = tmp & " insert into logMst_GodownInspection (InspectionID, Type, LocationID ,CMPID ,GodownID ,GodownNo ,GodownTypeID ,GodownCategoryID, LicenseID "
   tmp = tmp & " ,Capacity ,Address ,Area ,Status ,StartDate ,CloseDate ,noofstaff ,staffsalary ,noofsec ,secsalary ,noofstaffuser ,staffsalaryuser"
   tmp = tmp & " ,noofsecuser ,secsalaryuser ,Remarks ,totalrevenue ,directexpense ,grossprofit ,indirectexpense ,netprofit ,CapacityUtilisation ,NCDEXStock"
   tmp = tmp & " ,NNNRStock ,ReservationStock ,UpdatedOn ,UpdatedBy)"
   tmp = tmp & " values (@mInspectionID, @mType  ,@mLocationID ,@mCMPID ,@mGodownID ,@mGodownNo ,@mGodownTypeID ,@mGodownCategoryID ,@mLicenseID ,@mCapacity ,@mAddress"
   tmp = tmp & " ,@mArea ,@mStatus ,@mStartDate  ,@mCloseDate ,@mnoofstaff ,@mstaffsalary ,@mnoofsec ,@msecsalary ,@mnoofstaffuser ,@mstaffsalaryuser"
   tmp = tmp & " ,@mnoofsecuser ,@msecsalaryuser ,@mRemarks ,@mtotalrevenue ,@mdirectexpense ,@mgrossprofit ,@mindirectexpense ,@mnetprofit"
   tmp = tmp & " ,@mCapacityUtilisation ,@mNCDEXStock ,@mNNNRStock ,@mReservationStock ,@mUpdatedOn ,@mUpdatedBy)"
   
   Call TmpTable


End If

UpdateConfig ("219")
MsgBox "Version 219 Installed successfully"
End Function

Public Function version218()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmTxn_GenerateCommtrackInvoice.frm', 'Generate Commtrack Invoice','FrmTxn_GenerateCommtrackInvoice','Y','T',0,228,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'Frm_ExportCommtrackInvoice.frm', 'Export Commtrack Invoice','Frm_ExportCommtrackInvoice','Y','C',0,230,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_CommtrackInvoice.frm','Commtrack Invoice','FrmTxn_CommtrackInvoice','Y','T',0,228,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'Frm_UpdateCommtrackInvoiceNo.frm', 'Update Commtrack InvoiceNo','Frm_UpdateCommtrackInvoiceNo','Y','C',0,230,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'Frm_ExportReverseCommtrackInvoice.frm', 'Export Reverse Commtrack Invoice','Frm_ExportReverseCommtrackInvoice','Y','C',0,230,1"
Call TmpTable


tmp = " CREATE TABLE Txn_MonthlyCommtrackFileDetail"
tmp = tmp & " ("
tmp = tmp & " MonthYear DATETIME,"
tmp = tmp & " LotId " & TextType & "(25),"
tmp = tmp & " CPID " & TextType & "(50),"
tmp = tmp & " CPName " & TextType & "(200),"
tmp = tmp & " ClientID " & TextType & "(50),"
tmp = tmp & " ClientName " & TextType & "(200),"
tmp = tmp & " Qty DECIMAL(18,3),"
tmp = tmp & " StartDate DATETIME,"
tmp = tmp & " EndDate DATETIME,"
tmp = tmp & " HoldDuration DECIMAL(18,0),"
tmp = tmp & " Amount DECIMAL(18,2),"
tmp = tmp & " AmtIncST DECIMAL(18,2),"
tmp = tmp & " WarehouseName " & TextType & "(1000),"
tmp = tmp & " Commodity " & TextType & "(500),"
tmp = tmp & " WarehouseCode " & TextType & "(1000),"
tmp = tmp & " WSPName " & TextType & "(500),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
tmp = tmp & " )"
Call TmpTable

tmp = " CREATE TABLE Txn_MonthlyCommtrackInvoice"
tmp = tmp & " ("
tmp = tmp & " InvoiceID DECIMAL(18,0) NOT NULL,"
tmp = tmp & " MonthYear DATETIME,"
tmp = tmp & " LocationID DECIMAL(18,0),"
tmp = tmp & " CPNavCode " & TextType & "(25),"
tmp = tmp & " Amt DECIMAL(18,2),"
tmp = tmp & " STAmt DECIMAL(18,2),"
tmp = tmp & " TotalAmt DECIMAL(18,2),"
tmp = tmp & " PINo " & TextType & "(50),"
tmp = tmp & " BatchID DECIMAL(18,0),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " InvoiceNo " & TextType & "(25),"
tmp = tmp & " InvoiceDate DATETIME,"
tmp = tmp & " InvSTAmt DECIMAL(18,2),"
tmp = tmp & " InvoiceAmt DECIMAL(18,2),"
tmp = tmp & " ReverseRemark " & TextType & "(2000),"
tmp = tmp & " ReverseFlag " & TextType & "(1),"
tmp = tmp & " ReverseDate DATETIME,"
tmp = tmp & " RBatchID DECIMAL(18,0),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_MonthlyCommtrackFileDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyCommtrackFileDetail_Concurrency DEFAULT (GETDATE()) FOR Concurrency"
   Call TmpTable
   
   tmp = " ALTER TABLE Txn_MonthlyCommtrackInvoice WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_MonthlyCommtrackInvoice PRIMARY KEY CLUSTERED (InvoiceID),"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyCommtrackInvoice_Concurrency DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyCommtrackInvoice_BatchID DEFAULT (0) FOR BatchID,"
   tmp = tmp & " CONSTRAINT DF_Txn_MonthlyCommtrackInvoice_RBatchID DEFAULT (0) FOR RBatchID"
   Call TmpTable
End If

UpdateConfig ("218")
MsgBox "Version 218 Installed successfully"
End Function



Public Function version217()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmTxn_StorageContractExpireGslMapping.frm', 'Storage Contract Expired GSL Mapping','FrmTxn_StorageContractExpireGslMapping','Y','T',0,228,1"
Call TmpTable

tmp = " ALTER TABLE Log_ContractGSL ADD " & C & " EntryFlag " & TextType & "(1)"
Call TmpTable

tmp = " UPDATE Log_ContractGSL SET EntryFlag = 'O'"
Call TmpTable

tmp = "Alter table Txn_InvMonGSLDetail add MadeUpFlag " & TextType & "(1)"
Call TmpTable

UpdateConfig ("217")
MsgBox "Version 217 Installed successfully"
End Function

Public Function version216()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " Alter Table Mst_StorageContract Add SCMFLag " & TextType & "(1) "
Call TmpTable

tmp = " Update Mst_StorageContract Set SCMFlag = 0 Where SCMFlag is null"
Call TmpTable

UpdateConfig ("216")
MsgBox "Version 216 Installed successfully"

End Function

Public Function version215()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " CREATE TABLE Txn_TRF"
tmp = tmp & " ("
tmp = tmp & " TRFID DECIMAL(18,0) Not Null,"
tmp = tmp & " TRFNo " & TextType & "(50),"
tmp = tmp & " SamplingDate DATETIME,"
tmp = tmp & " ExchangeTypeID DECIMAL(18,0),"
tmp = tmp & " SM_Prefix " & TextType & "(5),"
tmp = tmp & " CMPID DECIMAL(18,0),"
tmp = tmp & " CommodityID DECIMAL(18,0),"
tmp = tmp & " ClientIDRow DECIMAL(10,0),"
tmp = tmp & " SamplerName " & TextType & "(500),"
tmp = tmp & " CMPRepresentive " & TextType & "(500),"
tmp = tmp & " Remarks " & TextType & "(500),"
tmp = tmp & " G_UID  " & TextType & " (36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy  " & TextType & " (30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy  " & TextType & " (30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (TRFID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " CREATE TABLE Txn_TRFDetail"
tmp = tmp & " ("
tmp = tmp & " DetailID DECIMAL(18,0) Not Null,"
tmp = tmp & " TRFID DECIMAL(18,0),"
tmp = tmp & " GSLID DECIMAL(18,0),"
tmp = tmp & " SM_Prefix " & TextType & "(5),"
tmp = tmp & " BalanceBags DECIMAL(18,0),"
tmp = tmp & " BalanceWeight DECIMAL(18,3),"
tmp = tmp & " G_UID  " & TextType & " (36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy  " & TextType & " (30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy  " & TextType & " (30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (DetailID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " CREATE TABLE LogTxn_TRFDetail"
tmp = tmp & " ("
tmp = tmp & " DetailID DECIMAL(18,0) Not Null,"
tmp = tmp & " TRFID DECIMAL(18,0),"
tmp = tmp & " GSLID DECIMAL(18,0),"
tmp = tmp & " SM_Prefix " & TextType & "(5),"
tmp = tmp & " BalanceBags DECIMAL(18,0),"
tmp = tmp & " BalanceWeight DECIMAL(18,3),"
tmp = tmp & " UpdatedBy  " & TextType & " (30),"
tmp = tmp & " UpdatedOn DateTime"
tmp = tmp & " )"
Call TmpTable


If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_TRF WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_TRF PRIMARY KEY CLUSTERED (TRFID),"
   tmp = tmp & " CONSTRAINT DF_Txn_TRF_Concurrency DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_TRF_Mst_ExchangeType FOREIGN KEY (ExchangeTypeID) REFERENCES Mst_ExchangeType (ExchangeTypeID),"
   tmp = tmp & " CONSTRAINT FK_Txn_TRF_Mst_CMP FOREIGN KEY (CMPID) REFERENCES Mst_CMP (CMPID),"
   tmp = tmp & " CONSTRAINT FK_Txn_TRF_Mst_Commodity FOREIGN KEY (CommodityID) REFERENCES Mst_Commodity (CommodityID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Txn_TRFDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_TRFDetail PRIMARY KEY CLUSTERED (DetailID),"
   tmp = tmp & " CONSTRAINT DF_Txn_TRFDetail_Concurrency DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_TRFDetail_Txn_TRF FOREIGN KEY (TRFID) REFERENCES Txn_TRF (TRFID),"
   tmp = tmp & " CONSTRAINT FK_Txn_TRFDetail_Mst_GSL FOREIGN KEY (GSLID, SM_Prefix) REFERENCES Mst_GSL (GSLID, SM_Prefix)"
   Call TmpTable

   tmp = " CREATE TRIGGER tr_Txn_TRFDetail ON Txn_TRFDetail"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mDetailID DECIMAL(18,0),"
   tmp = tmp & " @mTRFID DECIMAL(18,0),"
   tmp = tmp & " @mGSLID DECIMAL(18,0),"
   tmp = tmp & " @mSM_Prefix " & TextType & "(5),"
   tmp = tmp & " @mBalanceBags DECIMAL(18,0),"
   tmp = tmp & " @mBalanceWeight DECIMAL(18,3),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mDetailID = d.DetailID,"
   tmp = tmp & " @mTRFID = d.TRFID,"
   tmp = tmp & " @mGSLID = d.GSLID,"
   tmp = tmp & " @mSM_Prefix = d.SM_Prefix,"
   tmp = tmp & " @mBalanceBags = d.BalanceBags,"
   tmp = tmp & " @mBalanceWeight = d.BalanceWeight,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.DetailID = d.DetailID"
   tmp = tmp & " insert into LogTxn_TRFDetail (DetailID, TRFID, GSLID, SM_Prefix, BalanceBags, BalanceWeight, UpdatedOn, UpdatedBy)"
   tmp = tmp & " values (@mDetailID, @mTRFID, @mGSLID, @mSM_Prefix, @mBalanceBags, @mBalanceWeight, @mUpdatedOn, @mUpdatedBy)"
   Call TmpTable
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmTxn_TRF.frm', 'TRF','FrmTxn_TRF','Y','T',0,228,1"
Call TmpTable


UpdateConfig ("215")
MsgBox "Version 215 Installed successfully"
End Function

Public Function version214()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " ALTER TABLE Mst_StorageContract ALTER COLUMN Remarks " & TextType & "(2000)"
Call TmpTable

UpdateConfig ("214")
MsgBox "Version 214 Installed successfully"
End Function
Public Function version213()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Txn_InvMonGSLDetail add CommTrackFlag " & TextType & "(1)"
Call TmpTable

UpdateConfig ("213")
MsgBox "Version 213 Installed successfully"
End Function


Public Function version212()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If


tmp = " ALTER TABLE Txn_Spraying ADD " & C & " Area Decimal(18,2),  " & C & "  ActualArea Decimal(18,2),  " & C & " AgreementDetailID Decimal(18,0)"
Call TmpTable

tmp = " ALTER TABLE Txn_FumigationDetail ADD  " & C & " AgreementDetailID Decimal(18,0),  " & C & " Amount Decimal(18,2),"
tmp = tmp & C & "  MappingID DECIMAL(18,0),  " & C & " VendorQuantity DECIMAL(18,3),  " & C & " VendorRate DECIMAL(18,2)"
Call TmpTable

tmp = " CREATE TABLE Mst_TreatmentType"
tmp = tmp & " ("
tmp = tmp & " TreatmentTypeID DECIMAL(18,0) Not Null,"
tmp = tmp & " TreatmentType " & TextType & "(500),"
tmp = tmp & " Status " & TextType & "(2),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (TreatmentTypeID)"
End If
tmp = tmp & " )"
Call TmpTable


tmp = " CREATE TABLE Mst_FumigationAgreement"
tmp = tmp & " ("
tmp = tmp & " AgreementID DECIMAL(18,0) Not Null,"
tmp = tmp & " EntryDate DATETIME,"
tmp = tmp & " StartDate DATETIME,"
tmp = tmp & " ExpiryDate DATETIME,"
tmp = tmp & " Type " & TextType & "(2),"
tmp = tmp & " Status " & TextType & "(2),"
tmp = tmp & " LocationID DECIMAL(18,0),"
tmp = tmp & " VendorID DECIMAL(18,0),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (TreatmentTypeID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " CREATE TABLE Mst_FumigationTreatmentDetail"
tmp = tmp & " ("
tmp = tmp & " DetailID DECIMAL(18,0) Not Null,"
tmp = tmp & " AgreementID DECIMAL(18,0),"
tmp = tmp & " TreatmentTypeID DECIMAL(18,0),"
tmp = tmp & " RatePerMT DECIMAL(18,2),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (TreatmentTypeID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " CREATE TABLE Mst_FumigationQuotationDetail"
tmp = tmp & " ("
tmp = tmp & " DetailID DECIMAL(18,0) Not Null,"
tmp = tmp & " AgreementID DECIMAL(18,0),"
tmp = tmp & " VendorID DECIMAL(18,0),"
tmp = tmp & " TreatmentTypeID DECIMAL(18,0),"
tmp = tmp & " RatePerMT DECIMAL(18,2),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (TreatmentTypeID)"
End If
tmp = tmp & " )"
Call TmpTable


If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Mst_TreatmentType WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_TreatmentType PRIMARY KEY CLUSTERED (TreatmentTypeID),"
   tmp = tmp & " CONSTRAINT DF_Mst_TreatmentType_CConcurrency DEFAULT (GETDATE()) FOR Concurrency"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_FumigationAgreement WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_FumigationAgreement PRIMARY KEY CLUSTERED (AgreementID),"
   tmp = tmp & " CONSTRAINT DF_Mst_FumigationAgreement_Concurrency DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationAgreement_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationAgreement_Mst_Vendor FOREIGN KEY (VendorID) REFERENCES Mst_Vendor (VendorID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_FumigationTreatmentDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_FumigationTreatmentDetail PRIMARY KEY CLUSTERED (DetailID),"
   tmp = tmp & " CONSTRAINT DF_Mst_FumigationTreatmentDetail_Concurrency DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationTreatmentDetail_Mst_FumigationAgreement FOREIGN KEY (AgreementID) REFERENCES Mst_FumigationAgreement (AgreementID),"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationTreatmentDetail_Mst_TreatmentType FOREIGN KEY (TreatmentTypeID) REFERENCES Mst_TreatmentType (TreatmentTypeID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_FumigationQuotationDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_FumigationQuotationDetail PRIMARY KEY CLUSTERED (DetailID),"
   tmp = tmp & " CONSTRAINT DF_Mst_FumigationQuotationDetail_Concurrency DEFAULT (GETDATE()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationQuotationDetail_Mst_FumigationAgreement FOREIGN KEY (AgreementID) REFERENCES Mst_FumigationAgreement (AgreementID),"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationQuotationDetail_Mst_TreatmentType FOREIGN KEY (TreatmentTypeID) REFERENCES Mst_TreatmentType (TreatmentTypeID),"
   tmp = tmp & " CONSTRAINT FK_Mst_FumigationQuotationDetail_Mst_Vendor FOREIGN KEY (VendorID) REFERENCES Mst_Vendor (VendorID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Txn_FumigationDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Txn_FumigationDetail_Mst_FumigationTreatmentDetail FOREIGN KEY (AgreementDetailID) REFERENCES Mst_FumigationTreatmentDetail (DetailID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Txn_Spraying WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Txn_Spraying_Mst_FumigationTreatmentDetail FOREIGN KEY (AgreementDetailID) REFERENCES Mst_FumigationTreatmentDetail (DetailID)"
   Call TmpTable
End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_FumigationAgreement.frm','Fumigation Agreement','FrmMst_FumigationAgreement','Y','M',0,227,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_TreatmentType.frm','Treatment Type','FrmMst_TreatmentType','Y','M',0,227,1"
Call TmpTable

UpdateConfig ("212")
MsgBox "Version 212 Installed successfully"
End Function


Public Function version211()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

tmp = " ALTER TABLE Mst_Commodity ADD " & C & " CommtrackNCDEXCode " & TextType & "(100), NComNCDEXCode " & TextType & "(100)"
Call TmpTable

tmp = " ALTER TABLE Mst_Godown ADD " & C & " CommtrackNCDEXCode " & TextType & "(100), NComNCDEXCode " & TextType & "(100)"
Call TmpTable

tmp = " ALTER TABLE Mst_Location ADD " & C & " CommtrackNCDEXCode " & TextType & "(100), NComNCDEXCode " & TextType & "(100)"
Call TmpTable

tmp = " Alter Table Mst_GSL Add " & C & " CMSPBatchID Decimal (18,0) , " & C & " CMSPBatchGeneratedBy " & TextType & "(30)," & C & " CMSPBatchGeneratedDate DateTime, " & C & " CMSPUpdatedBy " & TextType & "(30)," & C & " CMSPUpdatedDate DateTime "
Call TmpTable

tmp = "Update Mst_GSL Set CMSPBatchID = 0 Where CMSPBatchID is null "
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = "ALTER TABLE Mst_GSL WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Mst_GSL_CMSPBatchID DEFAULT (0) FOR CMSPBatchID"
   
   Call TmpTable
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CMSPPending.frm','CMS-P Pending','FrmRpt_CMSPPending','Y','',0,280,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_CMSPUpload.frm','CMS-P Upload','FrmTxn_CMSPUpload','Y','',0,280,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_CMSEUpload.frm','CMS-E Upload','FrmTxn_CMSEUpload','Y','',0,280,1"
Call TmpTable



UpdateConfig ("211")
MsgBox "Version 211 Installed successfully"


End Function



Public Function version210()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "VARCHAR"
   C = " "
End If

'--take all form coexpense, inspection_expense_type, godown_inspection, storage_ratecard, ncdex_storageratecard, unittype
'--take gen_godowninspection variable declaration

tmp = " ALTER Table mst_coexpense ADD " & C & " BudgetIncome Decimal(18,2) ,  " & C & " BankInterest Decimal(18,2),  " & C & " AOROCMGExpense Decimal(18,2)"
Call TmpTable
tmp = " ALTER TABLE Mst_UnitType ADD  " & C & " Status " & TextType & "(2)"
Call TmpTable
tmp = " Update Mst_UnitType set Status = 'A'"
Call TmpTable

tmp = " Create Table Mst_InspectionExpenseType"
tmp = tmp & " ("
tmp = tmp & " InsExpTypeID Decimal(18,0) NOT NULL,"
tmp = tmp & " UnitTypeID Decimal(18,0),"
tmp = tmp & " ExpenseType " & TextType & "(255),"
tmp = tmp & " CalculationOn " & TextType & "(255),"
tmp = tmp & " Amount Decimal(18,2),"
tmp = tmp & " UserInputAmount " & TextType & "(1),"
tmp = tmp & " UserInputCalc " & TextType & "(1),"
tmp = tmp & " ClearFlag " & TextType & "(1),"
tmp = tmp & " Status " & TextType & "(2),"
tmp = tmp & " Type " & TextType & "(2),"
tmp = tmp & " Formula " & TextType & "(3000),"
tmp = tmp & " Remarks " & TextType & "(3000),"
tmp = tmp & " OrderNo Decimal(18,0),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (InsExpTypeID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table Mst_GodownInspection"
tmp = tmp & " ("
tmp = tmp & " InspectionID Decimal(18,0) NOT NULL,"
tmp = tmp & " Type " & TextType & "(2),"
tmp = tmp & " LocationID Decimal(18,0),"
tmp = tmp & " CMPID Decimal(18,0),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " GodownNo " & TextType & "(25),"
tmp = tmp & " GodownTypeID Decimal(18,0),"
tmp = tmp & " GodownCategoryID Decimal(18,0),"
tmp = tmp & " Capacity Decimal(18,3),"
tmp = tmp & " Address " & TextType & "(255),"
tmp = tmp & " Area Decimal(18,2),"
tmp = tmp & " Status " & TextType & "(2),"
tmp = tmp & " StartDate DateTime,"
tmp = tmp & " CloseDate DateTime,"
tmp = tmp & " noofstaff decimal(18,0),"
tmp = tmp & " staffsalary decimal(18,2),"
tmp = tmp & " noofsec decimal(18,0),"
tmp = tmp & " secsalary decimal(18,2),"
tmp = tmp & " noofstaffuser decimal(18,0),"
tmp = tmp & " staffsalaryuser decimal(18,2),"
tmp = tmp & " noofsecuser decimal(18,0),"
tmp = tmp & " secsalaryuser decimal(18,2),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " totalrevenue decimal(18,2),"
tmp = tmp & " directexpense decimal(18,2),"
tmp = tmp & " grossprofit decimal(18,2),"
tmp = tmp & " indirectexpense decimal(18,2),"
tmp = tmp & " netprofit decimal(18,2),"
tmp = tmp & " CapacityUtilisation Decimal(18,3),"
tmp = tmp & " NCDEXStock Decimal(18,3),"
tmp = tmp & " NNNRStock Decimal(18,3),"
tmp = tmp & " ReservationStock Decimal(18,3),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (InspectionID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table Mst_InspectionGodownDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionGodownID Decimal(18,0) NOT NULL,"
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " Capacity Decimal(18,3),"
tmp = tmp & " Rent Decimal(18,2),"
tmp = tmp & " Stock Decimal(18,3),"
tmp = tmp & " Area Decimal(18,2),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (InspectionGodownID)"
End If
tmp = tmp & " )"
Call TmpTable
    
tmp = " Create Table Mst_InspectionCommodityDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionCommodityID Decimal(18,0) NOT NULL,"
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " Category " & TextType & "(2),"
tmp = tmp & " CommodityID Decimal(18,0),"
tmp = tmp & " ExpQtyP Decimal(18,3),"
tmp = tmp & " ExpQtyMT Decimal(18,3),"
tmp = tmp & " FumigableQty Decimal(18,3),"
tmp = tmp & " StorageRate Decimal(18,2),"
tmp = tmp & " Price Decimal(18,2),"
tmp = tmp & " Revenue Decimal(18,2),"
tmp = tmp & " AUM Decimal(18,2),"
tmp = tmp & " StorageRateRemarks " & TextType & "(5000),"
tmp = tmp & " CMShareP Decimal(18,3),"
tmp = tmp & " CMShareR Decimal(18,2),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (InspectionCommodityID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table Mst_InspectionExpenseDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionExpenseID Decimal(18,0) NOT NULL,"
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " InsExpTypeID Decimal(18,0),"
tmp = tmp & " QtySQ Decimal(18,3),"
tmp = tmp & " Amount Decimal(18,2),"
tmp = tmp & " AmountPer Decimal(18,2),"
tmp = tmp & " Formula " & TextType & "(3000),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (InspectionExpenseID)"
End If
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table LogMst_InspectionExpenseType"
tmp = tmp & " ("
tmp = tmp & " InsExpTypeID Decimal(18,0),"
tmp = tmp & " UnitTypeID Decimal(18,0),"
tmp = tmp & " ExpenseType " & TextType & "(255),"
tmp = tmp & " CalculationOn " & TextType & "(255),"
tmp = tmp & " Amount Decimal(18,2),"
tmp = tmp & " UserInputAmount " & TextType & "(1),"
tmp = tmp & " UserInputCalc " & TextType & "(1),"
tmp = tmp & " ClearFlag " & TextType & "(1),"
tmp = tmp & " Status " & TextType & "(2),"
tmp = tmp & " Type " & TextType & "(2),"
tmp = tmp & " Formula " & TextType & "(3000),"
tmp = tmp & " Remarks " & TextType & "(3000),"
tmp = tmp & " OrderNo Decimal(18,0),"
tmp = tmp & " UpdatedOn DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"
Call TmpTable
   
tmp = " Create Table LogMst_GodownInspection"
tmp = tmp & " ("
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " Type " & TextType & "(2),"
tmp = tmp & " LocationID Decimal(18,0),"
tmp = tmp & " CMPID Decimal(18,0),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " GodownNo " & TextType & "(25),"
tmp = tmp & " GodownTypeID Decimal(18,0),"
tmp = tmp & " GodownCategoryID Decimal(18,0),"
tmp = tmp & " Capacity Decimal(18,3),"
tmp = tmp & " Address " & TextType & "(255),"
tmp = tmp & " Area Decimal(18,2),"
tmp = tmp & " Status " & TextType & "(2),"
tmp = tmp & " StartDate DateTime,"
tmp = tmp & " CloseDate DateTime,"
tmp = tmp & " noofstaff decimal(18,0),"
tmp = tmp & " staffsalary decimal(18,2),"
tmp = tmp & " noofsec decimal(18,0),"
tmp = tmp & " secsalary decimal(18,2),"
tmp = tmp & " noofstaffuser decimal(18,0),"
tmp = tmp & " staffsalaryuser decimal(18,2),"
tmp = tmp & " noofsecuser decimal(18,0),"
tmp = tmp & " secsalaryuser decimal(18,2),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " totalrevenue decimal(18,2),"
tmp = tmp & " directexpense decimal(18,2),"
tmp = tmp & " grossprofit decimal(18,2),"
tmp = tmp & " indirectexpense decimal(18,2),"
tmp = tmp & " netprofit decimal(18,2),"
tmp = tmp & " CapacityUtilisation Decimal(18,3),"
tmp = tmp & " NCDEXStock Decimal(18,3),"
tmp = tmp & " NNNRStock Decimal(18,3),"
tmp = tmp & " ReservationStock Decimal(18,3),"
tmp = tmp & " UpdatedOn DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table LogMst_InspectionGodownDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionGodownID Decimal(18,0),"
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " Capacity Decimal(18,3),"
tmp = tmp & " Rent Decimal(18,2),"
tmp = tmp & " Stock Decimal(18,3),"
tmp = tmp & " Area Decimal(18,2),"
tmp = tmp & " UpdatedOn DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table LogMst_InspectionCommodityDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionCommodityID Decimal(18,0),"
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " Category " & TextType & "(2),"
tmp = tmp & " CommodityID Decimal(18,0),"
tmp = tmp & " ExpQtyP Decimal(18,3),"
tmp = tmp & " ExpQtyMT Decimal(18,3),"
tmp = tmp & " FumigableQty Decimal(18,3),"
tmp = tmp & " StorageRate Decimal(18,2),"
tmp = tmp & " Price Decimal(18,2),"
tmp = tmp & " Revenue Decimal(18,2),"
tmp = tmp & " AUM Decimal(18,2),"
tmp = tmp & " StorageRateRemarks " & TextType & "(5000),"
tmp = tmp & " CMShareP Decimal(18,3),"
tmp = tmp & " CMShareR Decimal(18,2),"
tmp = tmp & " UpdatedOn DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"
Call TmpTable

tmp = " Create Table LogMst_InspectionExpenseDetail"
tmp = tmp & " ("
tmp = tmp & " InspectionExpenseID Decimal(18,0),"
tmp = tmp & " InspectionID Decimal(18,0),"
tmp = tmp & " InsExpTypeID Decimal(18,0),"
tmp = tmp & " QtySQ Decimal(18,3),"
tmp = tmp & " Amount Decimal(18,2),"
tmp = tmp & " AmountPer Decimal(18,2),"
tmp = tmp & " Formula " & TextType & "(3000),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " UpdatedOn DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Mst_InspectionExpenseType WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_InspectionExpenseType PRIMARY KEY CLUSTERED (InsExpTypeID),"
   tmp = tmp & " CONSTRAINT DF_Mst_InspectionExpenseType_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionExpenseType_Mst_UnitType FOREIGN KEY (UnitTypeID) REFERENCES Mst_UnitType (UnitTypeID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_GodownInspection WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_GodownInspection PRIMARY KEY CLUSTERED (InspectionID),"
   tmp = tmp & " CONSTRAINT DF_Mst_GodownInspection_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_GodownInspection_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Mst_GodownInspection_Mst_CMP FOREIGN KEY (CMPID) REFERENCES Mst_CMP (CMPID),"
   tmp = tmp & " CONSTRAINT FK_Mst_GodownInspection_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID),"
   tmp = tmp & " CONSTRAINT FK_Mst_GodownInspection_Mst_UnitType FOREIGN KEY (GodownTypeID) REFERENCES Mst_UnitType (UnitTypeID),"
   tmp = tmp & " CONSTRAINT FK_Mst_GodownInspection_Mst_Category FOREIGN KEY (GodownCategoryID) REFERENCES Mst_Category (CategoryID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_InspectionGodownDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_InspectionGodownDetail PRIMARY KEY CLUSTERED (InspectionGodownID),"
   tmp = tmp & " CONSTRAINT DF_Mst_InspectionGodownDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionGodownDetail_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_InspectionCommodityDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_InspectionCommodityDetail PRIMARY KEY CLUSTERED (InspectionCommodityID),"
   tmp = tmp & " CONSTRAINT DF_Mst_InspectionCommodityDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionCommodityDetail_Mst_GodownInspection FOREIGN KEY (InspectionID) REFERENCES Mst_GodownInspection (InspectionID),"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionCommodityDetail_Mst_Commodity FOREIGN KEY (CommodityID) REFERENCES Mst_Commodity (CommodityID)"
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_InspectionExpenseDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_InspectionExpenseDetail PRIMARY KEY CLUSTERED (InspectionExpenseID),"
   tmp = tmp & " CONSTRAINT DF_Mst_InspectionExpenseDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionExpenseDetail_Mst_GodownInspection FOREIGN KEY (InspectionID) REFERENCES Mst_GodownInspection (InspectionID),"
   tmp = tmp & " CONSTRAINT FK_Mst_InspectionExpenseDetail_Mst_InspectionExpenseType FOREIGN KEY (InsExpTypeID) REFERENCES Mst_InspectionExpenseType (InsExpTypeID)"
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_InspectionExpenseType ON Mst_InspectionExpenseType"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mInsExpTypeID Decimal(18,0),"
   tmp = tmp & " @mUnitTypeID Decimal(18,0),"
   tmp = tmp & " @mExpenseType " & TextType & "(255),"
   tmp = tmp & " @mCalculationOn " & TextType & "(255),"
   tmp = tmp & " @mAmount Decimal(18,2),"
   tmp = tmp & " @mUserInputAmount " & TextType & "(1),"
   tmp = tmp & " @mUserInputCalc " & TextType & "(1),"
   tmp = tmp & " @mClearFlag " & TextType & "(1),"
   tmp = tmp & " @mStatus " & TextType & "(2),"
   tmp = tmp & " @mType " & TextType & "(2),"
   tmp = tmp & " @mFormula " & TextType & "(3000),"
   tmp = tmp & " @mRemarks " & TextType & "(3000),"
   tmp = tmp & " @mOrderNo Decimal(18,0),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mInsExpTypeID = d.InsExpTypeID,"
   tmp = tmp & " @mUnitTypeID = d.UnitTypeID,"
   tmp = tmp & " @mExpenseType = d.ExpenseType,"
   tmp = tmp & " @mCalculationOn = d.CalculationOn,"
   tmp = tmp & " @mAmount = d.Amount,"
   tmp = tmp & " @mUserInputAmount = d.UserInputAmount,"
   tmp = tmp & " @mUserInputCalc = d.UserInputCalc,"
   tmp = tmp & " @mClearFlag = d.ClearFlag,"
   tmp = tmp & " @mStatus = d.Status,"
   tmp = tmp & " @mType = d.Type,"
   tmp = tmp & " @mFormula = d.Formula,"
   tmp = tmp & " @mRemarks = d.Remarks,"
   tmp = tmp & " @mOrderNo = d.OrderNo,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.InsExpTypeID = d.InsExpTypeID"
   tmp = tmp & " insert into LogMst_InspectionExpenseType (InsExpTypeID, UnitTypeID, ExpenseType, CalculationOn, Amount, UserInputAmount, UserInputCalc, ClearFlag, Status, Type, Formula, Remarks, OrderNo, UpdatedOn, UpdatedBy)"
   tmp = tmp & " values (@mInsExpTypeID, @mUnitTypeID, @mExpenseType, @mCalculationOn, @mAmount, @mUserInputAmount, @mUserInputCalc, @mClearFlag, @mStatus, @mType, @mFormula, @mRemarks, @mOrderNo, @mUpdatedOn, @mUpdatedBy)"
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_GodownInspection ON Mst_GodownInspection"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mInspectionID Decimal(18,0),"
   tmp = tmp & " @mType " & TextType & "(2),"
   tmp = tmp & " @mLocationID Decimal(18,0),"
   tmp = tmp & " @mCMPID Decimal(18,0),"
   tmp = tmp & " @mGodownID Decimal(18,0),"
   tmp = tmp & " @mGodownNo " & TextType & "(25),"
   tmp = tmp & " @mGodownTypeID Decimal(18,0),"
   tmp = tmp & " @mGodownCategoryID Decimal(18,0),"
   tmp = tmp & " @mCapacity Decimal(18,3),"
   tmp = tmp & " @mAddress " & TextType & "(255),"
   tmp = tmp & " @mArea Decimal(18,2),"
   tmp = tmp & " @mStatus " & TextType & "(2),"
   tmp = tmp & " @mStartDate DateTime,"
   tmp = tmp & " @mCloseDate DateTime,"
   tmp = tmp & " @mnoofstaff decimal(18,0),"
   tmp = tmp & " @mstaffsalary decimal(18,2),"
   tmp = tmp & " @mnoofsec decimal(18,0),"
   tmp = tmp & " @msecsalary decimal(18,2),"
   tmp = tmp & " @mnoofstaffuser decimal(18,0),"
   tmp = tmp & " @mstaffsalaryuser decimal(18,2),"
   tmp = tmp & " @mnoofsecuser decimal(18,0),"
   tmp = tmp & " @msecsalaryuser decimal(18,2),"
   tmp = tmp & " @mRemarks " & TextType & "(5000),"
   tmp = tmp & " @mtotalrevenue decimal(18,2),"
   tmp = tmp & " @mdirectexpense decimal(18,2),"
   tmp = tmp & " @mgrossprofit decimal(18,2),"
   tmp = tmp & " @mindirectexpense decimal(18,2),"
   tmp = tmp & " @mnetprofit decimal(18,2),"
   tmp = tmp & " @mCapacityUtilisation Decimal(18,3),"
   tmp = tmp & " @mNCDEXStock Decimal(18,3),"
   tmp = tmp & " @mNNNRStock Decimal(18,3),"
   tmp = tmp & " @mReservationStock Decimal(18,3),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mInspectionID = d.InspectionID,"
   tmp = tmp & " @mType = d.Type,"
   tmp = tmp & " @mLocationID = d.LocationID,"
   tmp = tmp & " @mCMPID = d.CMPID,"
   tmp = tmp & " @mGodownID = d.GodownID,"
   tmp = tmp & " @mGodownNo = d.GodownNo,"
   tmp = tmp & " @mGodownTypeID = d.GodownTypeID,"
   tmp = tmp & " @mGodownCategoryID = d.GodownCategoryID,"
   tmp = tmp & " @mCapacity = d.Capacity,"
   tmp = tmp & " @mAddress = d.Address,"
   tmp = tmp & " @mArea = d.Area,"
   tmp = tmp & " @mStatus = d.Status,"
   tmp = tmp & " @mStartDate = d.StartDate,"
   tmp = tmp & " @mCloseDate = d.CloseDate,"
   tmp = tmp & " @mnoofstaff = d.noofstaff,"
   tmp = tmp & " @mstaffsalary = d.staffsalary,"
   tmp = tmp & " @mnoofsec = d.noofsec,"
   tmp = tmp & " @msecsalary = d.secsalary,"
   tmp = tmp & " @mnoofstaffuser = d.noofstaffuser,"
   tmp = tmp & " @mstaffsalaryuser = d.staffsalaryuser,"
   tmp = tmp & " @mnoofsecuser = d.noofsecuser,"
   tmp = tmp & " @msecsalaryuser = d.secsalaryuser,"
   tmp = tmp & " @mRemarks = d.Remarks,"
   tmp = tmp & " @mtotalrevenue = d.totalrevenue,"
   tmp = tmp & " @mdirectexpense = d.directexpense,"
   tmp = tmp & " @mgrossprofit = d.grossprofit,"
   tmp = tmp & " @mindirectexpense = d.indirectexpense,"
   tmp = tmp & " @mnetprofit = d.netprofit,"
   tmp = tmp & " @mCapacityUtilisation = d.CapacityUtilisation,"
   tmp = tmp & " @mNCDEXStock = d.NCDEXStock,"
   tmp = tmp & " @mNNNRStock = d.NNNRStock,"
   tmp = tmp & " @mReservationStock = d.ReservationStock,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.InspectionID = d.InspectionID"
   tmp = tmp & " insert into logMst_GodownInspection (InspectionID, Type, LocationID ,CMPID ,GodownID ,GodownNo ,GodownTypeID ,GodownCategoryID"
   tmp = tmp & " ,Capacity ,Address ,Area ,Status ,StartDate ,CloseDate ,noofstaff ,staffsalary ,noofsec ,secsalary ,noofstaffuser ,staffsalaryuser"
   tmp = tmp & " ,noofsecuser ,secsalaryuser ,Remarks ,totalrevenue ,directexpense ,grossprofit ,indirectexpense ,netprofit ,CapacityUtilisation ,NCDEXStock"
   tmp = tmp & " ,NNNRStock ,ReservationStock ,UpdatedOn ,UpdatedBy)"
   tmp = tmp & " values (@mInspectionID, @mType  ,@mLocationID ,@mCMPID ,@mGodownID ,@mGodownNo ,@mGodownTypeID ,@mGodownCategoryID ,@mCapacity ,@mAddress"
   tmp = tmp & " ,@mArea ,@mStatus ,@mStartDate  ,@mCloseDate ,@mnoofstaff ,@mstaffsalary ,@mnoofsec ,@msecsalary ,@mnoofstaffuser ,@mstaffsalaryuser"
   tmp = tmp & " ,@mnoofsecuser ,@msecsalaryuser ,@mRemarks ,@mtotalrevenue ,@mdirectexpense ,@mgrossprofit ,@mindirectexpense ,@mnetprofit"
   tmp = tmp & " ,@mCapacityUtilisation ,@mNCDEXStock ,@mNNNRStock ,@mReservationStock ,@mUpdatedOn ,@mUpdatedBy)"
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_InspectionGodownDetail ON Mst_InspectionGodownDetail"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mInspectionGodownID Decimal(18,0),"
   tmp = tmp & " @mInspectionID Decimal(18,0),"
   tmp = tmp & " @mGodownID Decimal(18,0),"
   tmp = tmp & " @mCapacity Decimal(18,3),"
   tmp = tmp & " @mRent Decimal(18,2),"
   tmp = tmp & " @mStock Decimal(18,3),"
   tmp = tmp & " @mArea Decimal(18,2),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mInspectionGodownID = d.InspectionGodownID,"
   tmp = tmp & " @mInspectionID = d.InspectionID,"
   tmp = tmp & " @mGodownID = d.GodownID,"
   tmp = tmp & " @mCapacity = d.Capacity,"
   tmp = tmp & " @mRent = d.Rent,"
   tmp = tmp & " @mStock = d.Stock,"
   tmp = tmp & " @mArea = d.Area,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.InspectionGodownID = d.InspectionGodownID"
   tmp = tmp & " insert into LogMst_InspectionGodownDetail (InspectionGodownID ,InspectionID ,GodownID ,Capacity ,Rent ,Stock ,Area ,UpdatedOn ,UpdatedBy)"
   tmp = tmp & " values (@mInspectionGodownID ,@mInspectionID ,@mGodownID ,@mCapacity ,@mRent ,@mStock ,@mArea ,@mUpdatedOn ,@mUpdatedBy)"
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_InspectionCommodityDetail ON Mst_InspectionCommodityDetail"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mInspectionCommodityID Decimal(18,0),"
   tmp = tmp & " @mInspectionID Decimal(18,0),"
   tmp = tmp & " @mCategory " & TextType & "(2),"
   tmp = tmp & " @mCommodityID Decimal(18,0),"
   tmp = tmp & " @mExpQtyP Decimal(18,3),"
   tmp = tmp & " @mExpQtyMT Decimal(18,3),"
   tmp = tmp & " @mFumigableQty Decimal(18,3),"
   tmp = tmp & " @mStorageRate Decimal(18,2),"
   tmp = tmp & " @mPrice Decimal(18,2),"
   tmp = tmp & " @mRevenue Decimal(18,2),"
   tmp = tmp & " @mAUM Decimal(18,2),"
   tmp = tmp & " @mStorageRateRemarks " & TextType & "(5000),"
   tmp = tmp & " @mCMShareP Decimal(18,3),"
   tmp = tmp & " @mCMShareR Decimal(18,2),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mInspectionCommodityID = d.InspectionCommodityID,"
   tmp = tmp & " @mInspectionID = d.InspectionID,"
   tmp = tmp & " @mCategory = d.Category,"
   tmp = tmp & " @mCommodityID = d.CommodityID,"
   tmp = tmp & " @mExpQtyP = d.ExpQtyP,"
   tmp = tmp & " @mExpQtyMT = d.ExpQtyMT,"
   tmp = tmp & " @mFumigableQty = d.FumigableQty,"
   tmp = tmp & " @mStorageRate = d.StorageRate,"
   tmp = tmp & " @mPrice = d.Price,"
   tmp = tmp & " @mRevenue = d.Revenue,"
   tmp = tmp & " @mAUM = d.AUM,"
   tmp = tmp & " @mStorageRateRemarks = d.StorageRateRemarks,"
   tmp = tmp & " @mCMShareP = d.CMShareP,"
   tmp = tmp & " @mCMShareR = d.CMShareR,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.InspectionCommodityID = d.InspectionCommodityID"
   tmp = tmp & " insert into LogMst_InspectionCommodityDetail (InspectionCommodityID, InspectionID, Category, CommodityID ,ExpQtyP ,ExpQtyMT ,FumigableQty ,StorageRate ,Price ,Revenue ,AUM ,StorageRateRemarks ,CMShareP, CMShareR ,UpdatedOn ,UpdatedBy)"
   tmp = tmp & " values (@mInspectionCommodityID ,@mInspectionID ,@mCategory ,@mCommodityID  ,@mExpQtyP ,@mExpQtyMT ,@mFumigableQty ,@mStorageRate ,@mPrice ,@mRevenue ,@mAUM ,@mStorageRateRemarks, @mCMShareP, @mCMShareR ,@mUpdatedOn ,@mUpdatedBy)"
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_InspectionExpenseDetail ON Mst_InspectionExpenseDetail"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mInspectionExpenseID Decimal(18,0),"
   tmp = tmp & " @mInspectionID Decimal(18,0),"
   tmp = tmp & " @mInsExpTypeID Decimal(18,0),"
   tmp = tmp & " @mQtySQ Decimal(18,3),"
   tmp = tmp & " @mAmount Decimal(18,2),"
   tmp = tmp & " @mAmountPer Decimal(18,2),"
   tmp = tmp & " @mFormula " & TextType & "(3000),"
   tmp = tmp & " @mRemarks " & TextType & "(5000),"
   tmp = tmp & " @mUpdatedOn DateTime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mInspectionExpenseID = d.InspectionExpenseID,"
   tmp = tmp & " @mInspectionID = d.InspectionID,"
   tmp = tmp & " @mInsExpTypeID = d.InsExpTypeID,"
   tmp = tmp & " @mQtySQ = d.QtySQ,"
   tmp = tmp & " @mAmount = d.Amount,"
   tmp = tmp & " @mAmountPer = d.AmountPer,"
   tmp = tmp & " @mFormula = d.Formula,"
   tmp = tmp & " @mRemarks = d.Remarks,"
   tmp = tmp & " @mUpdatedOn = isnull(d.UpdatedDate, d.CreatedDate),"
   tmp = tmp & " @mUpdatedBy = isnull(d.UpdatedBy, d.CreatedBy)"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.InspectionExpenseID = d.InspectionExpenseID"
   tmp = tmp & " insert into LogMst_InspectionExpenseDetail (InspectionExpenseID ,InspectionID ,InsExpTypeID ,QtySQ ,Amount ,AmountPer ,Formula ,Remarks ,UpdatedOn ,UpdatedBy)"
   tmp = tmp & " values (@mInspectionExpenseID ,@mInspectionID ,@mInsExpTypeID ,@mQtySQ ,@mAmount ,@mAmountPer ,@mFormula ,@mRemarks ,@mUpdatedOn ,@mUpdatedBy)"
   Call TmpTable
End If

tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID)  Values (1,'Leased Rent','Area(per Sq.Ft)',5,'A',newid(),'godownarea*amountperunit','Area of the Godown in Sq.ft x Rate per Unit','D',1,1,1,1,1)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID)  Values (2,'Franchisee Fees based on Qty','Qty(per MT)',20,'A',newid(),'godowncapacity*amountperunit','franchisee fees is calculated based on capacity multiply by amount per unit','D',1,1,1,2,2)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (3,'Franchisee fees on revenue sharing basis','Fixed',60,'A',newid(),'godownincome*amountperunit/100','calculation is based on revenue','D',1,0,1,3,2)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (4,'Spraying Expense','Area(per Sq.Ft)',0.2,'A',newid(),'godownarea*amountperunit*2','Spraying rate is assumed to for a fortnight. Hence the formula is :  area of the godown x Rate per Unit *2','D',1,1,1,4,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (5,'Fumigation - once every 45 days','Fumigation Qty(per MT)',20,'A',newid(),'fumigableqty*amountperunit*2/3','Total Fumigable Qty in MT in the godown x Rate x 2 / 3','D',1,0,1,5,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (6,'Fumigation - once every 3 months','Fumigation Qty(per MT)',20,'A',newid(),'fumigableqty*amountperunit/3','fumigable quantity x fumigation rate / 3','D',1,0,1,6,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (7,'Dunnage','Area(per Sq.Ft)',2,'A',newid(),'godownarea*amountperunit/month','Area of the godown in Sq.ft x Rate per unit / lease period of the godown','D',1,1,1,7,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (8,'Insurance','per AUM',0.07,'A',newid(),'(aum*amountperunit/100)/12','AUM of the godown of all commodities in the godown  x 0.07% / 12','D',1,1,1,8,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (9,'Fire Extinguiser','Qty(per MT)',1500,'A',newid(),'(godowncapacity/1000*amountperunit)/month','Ceiling (Capacity of Godown in MT / 1000) * Rate per unit (i.e. 1500) /   number of lease period','D',1,1,1,9,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (10,'Communication Cost','Fixed',5000,'A',newid(),'godowncapacity*amountperunit/totalcapacity','This is a per location monthly communication cost. The formula for the godown is : godown capacity in MT x 5000 / Total Capacity of all live godowns in that location.','D',1,0,1,10,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (11,'Stationery','Fixed',3000,'A',newid(),'godowncapacity*amountperunit/totalcapacity','This is a per location monthly Stationery cost. The formula for the godown is : godown capacity in MT x 3000 (Rate per unit) / Total Capacity of all live godowns in that location.','D',1,0,1,11,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (12,'Power/Water/Godown Cleaning','Fixed',1000,'A',newid(),'','Since this a fixed amount per month per godown, no further formula is required.','D',1,0,0,12,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (13,'NCDEX Fees','Fixed',2,'A',newid(),'ncdexincome*amountperunit/100','Monthly NCDEX income x Rate per unit (i.e. 2) / 100 ','D',1,0,1,13,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (14,'Labour Charges','Fixed',0,'A',newid(),'','Any specific employee cost or labour cost (total monthly cost) which are specific only to this godown can be captured here...','D',1,0,0,14,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (15,'Risk Factor as percentage of AUM','Fixed',0.01,'A',newid(),'(aum*amountperunit/100)/month','AUM of the godown x amount per unit  / 100 / lease tenor of the godown','D',1,0,1,15,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (16,'Miscellaneous Expenses','Fixed',1500,'A',newid(),'','Since this is a fixed expense per godown per month, no further formula is required.','D',1,0,0,16,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (17,'Bank Interest','Fixed',0,'A',newid(),'(cmpbankinterest/cmpbudgetincome*godownincome)','Total Bank Interest pertaining to S&P segment / Total budgeted income of S&P segment x monthly income of the godown','I',0,0,1,17,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (18,'Fidelity Insurance','Fixed',500,'A',newid(),'','Rs.500 per godown is  a fixed amount towards Fidelity Insurance.','I',1,0,0,18,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (19,'CO Expnese','Fixed',0,'A',newid(),'(coexpense/cmpbudgetincome*godownincome)','Total CO expense / Total budgeted income of S&P segment x Monthly income of the godown','I',0,0,1,19,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (20,'AO - RO - CMG Expenses','Fixed',0,'A',newid(),'(aorocmgexpense/cmpbudgetincome*godownincome)','AO-cum-RO-cum-CMG expense for the year / Budgeted income for the year x monthly income of the godown','I',0,0,1,20,NULL)"
Call TmpTable
tmp = " insert into  mst_inspectionexpensetype (InsExpTypeID, ExpenseType, CalculationOn, Amount, Status, G_UID, Formula, Remarks, Type, UserInputAmount, UserInputCalc, ClearFlag, OrderNo, UnitTypeID) Values (21,'NCDEX warehouse renewal fee','Fixed',56000,'A',newid(),'','if capacity <5000, fee is 10,500  if capacity < 10000, fee is 21,000  for others fee is 52500','D',1,0,1,21,NULL)"
Call TmpTable


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_GodownInspection.frm','Godown Viability','FrmMst_GodownInspection','Y','M',0,227,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_InspectionExpenseType.frm','Inspection Expense Type','FrmMst_InspectionExpenseType','Y','M',0,227,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownwiseMonthlyStockAUM.frm','Godown wise Monthly Stock AUM','FrmRpt_GodownwiseMonthlyStockAUM','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("210")
MsgBox "Version 210 Installed successfully"
End Function


Public Function version209()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALter Table Mst_Client Add CMSPCode " & TextType & "(50),KYCFlag " & TextType & "(1)"
Call TmpTable

tmp = " ALter Table Mst_Commodity Add CommTrackFlag " & TextType & "(1)"
Call TmpTable

tmp = " Alter Table Txn_CxTF_Details Add RecAnnexure " & TextType & "(1)"
Call TmpTable

tmp = " Update Mst_Client Set KYCFlag = 0 Where KYCFlag is null"
Call TmpTable

tmp = " Update Mst_Commodity Set CommTrackFlag = 0 Where CommTrackFlag is null"
Call TmpTable

tmp = " Update Txn_CxTF_Details Set RecAnnexure = 0 Where RecAnnexure is null And TxnType = 'D'"
Call TmpTable


UpdateConfig ("209")
MsgBox "Version 209 Installed successfully"
End Function

Public Function version208()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Mst_GSL Add BSM_Prefix " & TextType & "(5),FSM_Prefix " & TextType & "(5)"
Call TmpTable


UpdateConfig ("208")
MsgBox "Version 208 Installed successfully"
End Function

Public Function version207()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "CREATE TABLE Txn_ReservationDailyStock("
tmp = tmp & " RDSID decimal(18, 0) Not Null,"
tmp = tmp & " MonthYear datetime ,"
tmp = tmp & " TxnDate datetime ,"
tmp = tmp & " LocationID decimal(18, 0) ,"
tmp = tmp & " ContractID decimal(18, 0) ,"
tmp = tmp & " BalanceBags decimal(18, 0) ,"
tmp = tmp & " BalanceWeight decimal(18, 3) ,"
tmp = tmp & " Flag " & TextType & "(2) ,"
tmp = tmp & " G_UID " & TextType & "(36) ,"
tmp = tmp & " Concurrency datetime ,"
tmp = tmp & " CreatedBy " & TextType & "(30) ,"
tmp = tmp & " CreatedDate datetime ,"
tmp = tmp & " UpdatedDate datetime ,"
tmp = tmp & " UpdatedBy " & TextType & "(30) ,"
tmp = tmp & " OpnBalBags decimal(18, 0) ,"
tmp = tmp & " OpnBalQty decimal(18, 3) ,"
tmp = tmp & " DepoBags decimal(18, 0) ,"
tmp = tmp & " DepoQty decimal(18, 3) ,"
tmp = tmp & " WidraBags decimal(18, 0) ,"
tmp = tmp & " WidraQty decimal(18, 3) ,"
tmp = tmp & " CxTFNo " & TextType & "(500)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (RDSID)"
End If
tmp = tmp & " )"

Call TmpTable

tmp = " CREATE TABLE Txn_ReservationInvoiceDetail("
tmp = tmp & " RIDID decimal(18, 0) Not Null ,"
tmp = tmp & " InvoiceNo decimal(18, 0) ,"
tmp = tmp & " BillingCycleID decimal(18, 0) ,"
tmp = tmp & " FromDate datetime ,"
tmp = tmp & " ToDate datetime ,"
tmp = tmp & " PeakValue decimal(18, 3) ,"
tmp = tmp & " Rate decimal(18, 2) ,"
tmp = tmp & " Amount decimal(18, 2) ,"
tmp = tmp & " Flag " & TextType & "(2) ,"
tmp = tmp & " G_UID " & TextType & "(36) ,"
tmp = tmp & " Concurrency datetime ,"
tmp = tmp & " CreatedBy " & TextType & "(30) ,"
tmp = tmp & " CreatedDate datetime ,"
tmp = tmp & " UpdatedDate datetime ,"
tmp = tmp & " UpdatedBy " & TextType & "(30) ,"
tmp = tmp & " ContractID decimal(18, 0)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (RIDID)"
End If
tmp = tmp & " )"

Call TmpTable

tmp = " Alter table Txn_reservationInvoice add Discount decimal(18,2), Premium decimal(18,2)"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_ReservationDailyStock WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_ReservationDailyStock PRIMARY KEY CLUSTERED (RDSID),"
   tmp = tmp & " CONSTRAINT DF_Txn_ReservationDailyStock_Concurrency DEFAULT (getdate()) FOR Concurrency"
   
   Call TmpTable
   
   tmp = " ALTER TABLE Txn_ReservationInvoiceDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_ReservationInvoiceDetail PRIMARY KEY CLUSTERED (RIDID),"
   tmp = tmp & " CONSTRAINT DF_Txn_ReservationInvoiceDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_ReservationInvoiceDetail_Txn_ReservationInvoice FOREIGN KEY (InvoiceNo) REFERENCES Txn_ReservationInvoice(InvoiceNo),"
   tmp = tmp & " CONSTRAINT FK_Txn_ReservationInvoiceDetail_Mst_BillingCycle FOREIGN KEY (BillingCycleID) REFERENCES Mst_BillingCycle(BillingCycleID)"
   
   Call TmpTable
   
End If

UpdateConfig ("207")
MsgBox "Version 207 Installed successfully"
End Function

Public Function version206()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Mst_Config ADD " & C & " BlockingDate DateTime"
Call TmpTable

tmp = " Update Mst_Config Set BlockingDate = '2012-12-31' "
Call TmpTable

UpdateConfig ("206")
MsgBox "Version 206 Installed successfully"
End Function
Public Function version205()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ReservationStockStatement.frm','Reservation Stock Statement','FrmRpt_ReservationStockStatement','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("205")
MsgBox "Version 205 Installed successfully"
End Function

Public Function version204()
'Godown Wise Profitability
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Txn_MonGodownWiseProfitability"
tmp = tmp & " ("
tmp = tmp & " MGWPID Decimal(18,0) not null,"
tmp = tmp & " MonthYear DateTime,"
tmp = tmp & " StateID   Decimal(18,0),"
tmp = tmp & " LocationID  Decimal(18,0),"
tmp = tmp & " AM   " & TextType & "(5000),"
tmp = tmp & " RM   " & TextType & "(5000),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " Capacity Decimal(18,3),"
tmp = tmp & " NcdexStock Decimal(18,3),"
tmp = tmp & " NNStock Decimal(18,3),"
tmp = tmp & " NNNRIncomeAdhoc Decimal(18,2),"
tmp = tmp & " NNNRIncome Decimal(18,2),"
tmp = tmp & " ReservationIncomeAdhoc Decimal(18,2),"
tmp = tmp & " ReservationIncome Decimal(18,2),"
tmp = tmp & " SevenDaysIncome Decimal(18,2),"
tmp = tmp & " DPIncomeAdHoc Decimal(18,2),"
tmp = tmp & " FinalDPIncome Decimal(18,2),"
tmp = tmp & " DPIncomeGSL Decimal(18,2),"
tmp = tmp & " Diff Decimal(18,2),"
tmp = tmp & " DPIncomeInvoice Decimal(18,2),"
tmp = tmp & " RejectedIncome Decimal(18,2),"
tmp = tmp & " SpotIncomeAdHoc Decimal(18,2),"
tmp = tmp & " OtherIncome Decimal(18,2),"
tmp = tmp & " SCMIncome Decimal(18, 2),"
tmp = tmp & " CMUnbilledIncome Decimal(18,2),"
tmp = tmp & " CMPendingIncome Decimal(18,2),"
tmp = tmp & " CMApprovedIncome Decimal(18,2),"
tmp = tmp & " LeftOverStock Decimal(18,2),"
tmp = tmp & " PledgeIncome Decimal(18,2),"
tmp = tmp & " LeaseRentPaid Decimal(18,2),"
tmp = tmp & " LeaseRentFAG Decimal(18,2),"
tmp = tmp & " LeaseRentDRC Decimal(18,2),"
tmp = tmp & " EmpClaimFAG Decimal(18,2),"
tmp = tmp & " EmpClaimAccrual Decimal(18,2),"
tmp = tmp & " VendorExpenses Decimal(18,2),"
tmp = tmp & " VendorExpensesAccrued Decimal(18,2),"
tmp = tmp & " SecurityFeesPaid Decimal(18,2),"
tmp = tmp & " SecurityFeesAccrual Decimal(18,2),"
tmp = tmp & " FumigationFeesVendor Decimal(18,2),"
tmp = tmp & " FumigationFeesPaid Decimal(18,2),"
tmp = tmp & " FumigationFeesAccrual Decimal(18,2),"
tmp = tmp & " FranchiseeFeesPaid Decimal(18,2),"
tmp = tmp & " FranchiseeFeesFAG Decimal(18,2),"
tmp = tmp & " FranchiseeFeesCMG Decimal(18,2),"
tmp = tmp & " Salary Decimal(18,2),"
tmp = tmp & " NCDEXCharges Decimal(18,2),"
tmp = tmp & " CoExpense Decimal(18,2),"
tmp = tmp & " InsurancePremium Decimal(18,2),"
tmp = tmp & " InterestCharges Decimal(18,2),"
tmp = tmp & " ReportFlag " & TextType & "(1),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = tmp & " ,PRIMARY KEY (MGWPID)"
End If
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   
   tmp = " ALTER TABLE Txn_MonGodownWiseProfitability WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_MonGodownWiseProfitability PRIMARY KEY CLUSTERED (MGWPID),"
   tmp = tmp & " CONSTRAINT DF_Txn_MonGodownWiseProfitability_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_MonGodownWiseProfitability_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Txn_MonGodownWiseProfitability_Mst_State FOREIGN KEY (StateID) REFERENCES Mst_State (StateID),"
   tmp = tmp & " CONSTRAINT FK_Txn_MonGodownWiseProfitability_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID)"
   
   Call TmpTable
End If

tmp = " ALTER TABLE MST_GODOWN ADD " & C & " CPEN " & TextType & "(3000), " & C & " CPED DATETIME, "
tmp = tmp & C & " COPEN " & TextType & "(3000), " & C & " COPED DATETIME"
Call TmpTable

tmp = " Alter Table Txn_GSL_Transfer Add NewSM_Prefix " & TextType & "(5) "
Call TmpTable

tmp = " Update Txn_GSL_Transfer Set NewSM_Prefix = SM_Prefix "
Call TmpTable

tmp = " Alter Table Txn_ReservationInvoice Add " & C & " PrinciplePINo Decimal(18,0)"
Call TmpTable

tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1424 Where InvoiceNo =1444"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1457 Where InvoiceNo =1458"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1477 Where InvoiceNo =1478"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1477 Where InvoiceNo =1479"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1491 Where InvoiceNo =1490"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1529 Where InvoiceNo =1530"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1546 Where InvoiceNo =1547"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1564 Where InvoiceNo =1565"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1597 Where InvoiceNo =1598"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1597 Where InvoiceNo =1599"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1738 Where InvoiceNo =1737"
Call TmpTable
tmp = " Update Txn_ReservationInvoice set PrinciplePINo = 1757 Where InvoiceNo =1756"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownWiseProfitability.frm','Godown Wise Profitability','FrmRpt_GodownWiseProfitability','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownWiseProfitabilityConsolidated.frm','Godown Wise Profitability Consolidated','FrmRpt_GodownWiseProfitabilityConsolidated','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("204")
MsgBox "Version 204 Installed successfully"

End Function



Public Function version203()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_InvMonGSLDetail Alter Column LotNo " & TextType & " (25) "
Call TmpTable

UpdateConfig ("203")
MsgBox "Version 203 Installed successfully"
End Function
Public Function version202()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "ALTER TABLE Mst_Hologram ADD CONSTRAINT UC_NewHologramNo UNIQUE (NewHologramNo)"
Call TmpTable

tmp = "Alter Table Mst_Client Add SCMFlag " & TextType & " (1)"
Call TmpTable

tmp = "Update Mst_Client Set SCMFlag = 0"
Call TmpTable

tmp = "Alter Table Txn_WHR_Detail Add SCMFlag " & TextType & " (1)"
Call TmpTable

tmp = "Update Txn_WHR_Detail Set SCMFlag = 0"
Call TmpTable


UpdateConfig ("202")
MsgBox "Version 202 Installed successfully"
End Function

Public Function version201()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_TruckwiseCxtfDetail.frm','Truck wise Cxtf Detail','FrmRpt_TruckwiseCxtfDetail','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("201")
MsgBox "Version 201 Installed successfully"
End Function
Public Function version200()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_UnApprovedCRorTDS.frm','Unapproved CR or TDS','FrmRpt_UnApprovedCRorTDS','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("200")
MsgBox "Version 200 Installed successfully"
End Function
Public Function version199()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Txn_InvMonGSLDetail ADD " & C & " DPIncome decimal(18,2)"
Call TmpTable


UpdateConfig ("199")
MsgBox "Version 199 Installed successfully"

End Function

Public Function version198()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Txn_WHR_Detail ADD " & C & " ConvertedClient " & TextType & "(1)"
Call TmpTable

tmp = " Update Txn_WHR_Detail Set ConvertedClient = 0 "
Call TmpTable

UpdateConfig ("198")
MsgBox "Version 198 Installed successfully"
End Function

Public Function version197()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Mst_SpecialRateCard ADD " & C & " ConvertedClient " & TextType & "(2)"
Call TmpTable

tmp = " UPDATE Mst_SpecialRateCard SET ConvertedClient = 0"
Call TmpTable

tmp = " Create Table Mst_LicenseAttachment"
tmp = tmp & " ("
tmp = tmp & " TxnID        DECIMAL(18, 0),"
tmp = tmp & " LicenseID    DECIMAL(18, 0),"
tmp = tmp & " FileName     " & TextType & "(500),"
tmp = tmp & " [File]       IMAGE,"
tmp = tmp & " Remarks      " & TextType & "(2000),"
tmp = tmp & " Concurrency  DATETIME,"
tmp = tmp & " G_UID        " & TextType & "(36),"
tmp = tmp & " CreatedBy    " & TextType & "(36),"
tmp = tmp & " CreatedDate  DATETIME,"
tmp = tmp & " UpdatedDate  DATETIME,"
tmp = tmp & " UpdatedBy    " & TextType & "(36)"
tmp = tmp & " )"
Call TmpTableSMAttachment

tmp = " ALTER TABLE Mst_LicenseAttachment WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT DF_Mst_LicenseAttachment_Concurrency DEFAULT (getdate()) FOR Concurrency"
Call TmpTableSMAttachment

UpdateConfig ("197")
MsgBox "Version 197 Installed successfully"
End Function


Public Function version196()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_RRNPWR.frm','RRN PWR Mapping','FrmTxn_RRNPWR','Y','T',0,228,1"
Call TmpTable

tmp = " Alter Table Mst_Client Add " & C & " ConvertedClientIDRow Decimal (18,0)"
Call TmpTable

tmp = " ALTER TABLE MST_GSL ALTER COLUMN LOTNO " & TextType & "(25)"
Call TmpTable

tmp = " CREATE TABLE Txn_RRNPWR"
tmp = tmp & " ("
tmp = tmp & " ConvertedID    DECIMAl(18,0) NOT NULL,"
tmp = tmp & " PWRDate    DateTime,"
tmp = tmp & " LocationID     DECIMAl(18,0),"
tmp = tmp & " CommodityID    DECIMAl(18,0),"
tmp = tmp & " RRNID       DECIMAl(18,0),"
tmp = tmp & " ConvertedClientIDRow DECIMAl(18,0),"
tmp = tmp & " AdjstQty    DECIMAl(18,3),"
tmp = tmp & " Remarks        " & TextType & "(2000),"
tmp = tmp & " Flag  " & TextType & "(1),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " G_UID       " & TextType & "(36),"
tmp = tmp & " CreatedBy      " & TextType & "(30),"
tmp = tmp & " CreatedDate    DATETIME,"
tmp = tmp & " UpdatedBy      " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
tmp = tmp & " )"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (ConvertedID)"
End If
Call TmpTable

tmp = " CREATE TABLE Txn_RRNPWRGSL"
tmp = tmp & " ("
tmp = tmp & " ConvertedGSLID DECIMAl(18,0) NOT NULL,"
tmp = tmp & " ConvertedID DECIMAl(18,0),"
tmp = tmp & " GSLID    DECIMAl(18,0),"
tmp = tmp & " SM_Prefix   " & TextType & "(5),"
tmp = tmp & " AdjstBags   DECIMAl(18,0),"
tmp = tmp & " AdjstQty DECIMAl(18,3),"
tmp = tmp & " NewGSLID DECIMAl(18,0),"
tmp = tmp & " Concurrency DATETIME,"
tmp = tmp & " G_UID    " & TextType & "(36),"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " CreatedDate DATETIME,"
tmp = tmp & " UpdatedBy   " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (ConvertedGSLID)"
End If
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
    tmp = " ALTER TABLE Txn_RRNPWR WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Txn_RRNPWR PRIMARY KEY CLUSTERED (ConvertedID),"
    tmp = tmp & " CONSTRAINT DF_Txn_RRNPWR_Concurrency DEFAULT (getdate()) FOR Concurrency,"
    tmp = tmp & " CONSTRAINT FK_Txn_RRNPWR_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
    tmp = tmp & " CONSTRAINT FK_Txn_RRNPWR_Mst_Commodity FOREIGN KEY (CommodityID) REFERENCES Mst_Commodity (CommodityID)"
    Call TmpTable
    
    tmp = " ALTER TABLE Txn_RRNPWRGSL WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Txn_RRNPWRGSL PRIMARY KEY CLUSTERED (ConvertedGSLID),"
    tmp = tmp & " CONSTRAINT DF_Txn_RRNPWRGSL_Concurrency DEFAULT (getdate()) FOR Concurrency,"
    tmp = tmp & " CONSTRAINT FK_Txn_RRNPWRGSL_Txn_RRNPWR FOREIGN KEY (ConvertedID) REFERENCES Txn_RRNPWR (ConvertedID),"
    tmp = tmp & " CONSTRAINT FK_Txn_RRNPWRGSL_Mst_GSL FOREIGN KEY (GSLID, SM_Prefix) REFERENCES Mst_GSL (GSLID, SM_Prefix)"
    Call TmpTable
End If

UpdateConfig ("196")
MsgBox "Version 196 Installed successfully"
End Function

Public Function version195()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_StorageContractGodMapping.frm','Expired Storage Contract Godown Mapping','FrmRpt_StorageContractGodMapping','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("195")
MsgBox "Version 195 Installed successfully"
End Function

Public Function version194()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Txn_SRGodownMapDetails Add WtQty decimal(18,3)"
Call TmpTable

tmp = "ALTER TABLE Txn_SRGodownMapDetails WITH NOCHECK ADD CONSTRAINT DF_Txn_SRGodownMapDetails_WtQty DEFAULT (0) FOR WtQty"
Call TmpTable

tmp = "Update Txn_SRGodownMapDetails Set WtQty = 0 "
Call TmpTable

tmp = " Alter Table Mst_Client Add Converted Varchar(1) "
Call TmpTable

tmp = " Update Mst_Client Set Converted = 0"
Call TmpTable

UpdateConfig ("194")
MsgBox "Version 194 Installed successfully"

End Function

Public Function version193()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_UploadTIInvoicebyFAG.frm','Upload TI from NAVISION by FAG (Excel)','Frm_UploadTIInvoicebyFAG','Y','A',0,232,1"

Call TmpTable

UpdateConfig ("193")
MsgBox "Version 193 Installed successfully"
End Function

Public Function version192()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = "Alter Table Mst_SContractBillingDetail Add ResQtyforMIS decimal(18,3) , WtQty decimal(18,3)"
Call TmpTable

tmp = "ALTER TABLE Mst_SContractBillingDetail WITH NOCHECK ADD"
tmp = tmp & " CONSTRAINT DF_Mst_SContractBillingDetail_ResQtyforMIS DEFAULT (0) FOR ResQtyforMIS,"
tmp = tmp & " CONSTRAINT DF_Mst_SContractBillingDetail_WtQty DEFAULT (0) FOR WtQty"

Call TmpTable

tmp = "Update Mst_SContractBillingDetail Set ResQtyforMIS = BillingQty,WtQty = 0"

Call TmpTable


UpdateConfig ("192")
MsgBox "Version 192 Installed successfully"
End Function


Public Function version191()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Txn_GSL_Transfer ADD " & C & " CommtrackBatchID Decimal(18,0),  " & C & " CommtrackBatchDate DateTime"
Call TmpTable

tmp = " ALTER TABLE Txn_CxTF_GSL ADD  " & C & " CommtrackBatchID Decimal(18,0),  " & C & " CommtrackBatchDate DateTime"
Call TmpTable

tmp = " ALTER TABLE MST_LOCATION ADD  " & C & " NCDEXCODE " & TextType & "(100)"
Call TmpTable

tmp = " ALTER TABLE MST_COMMODITY ADD  " & C & " NCDEXCODE " & TextType & "(100)"
Call TmpTable

tmp = " ALTER TABLE MST_GODOWN ADD  " & C & " NCDEXCODE " & TextType & "(500)"
Call TmpTable


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_UpdateCommtrackNCDEXCODE.frm','Update Commtrack NCDEX Code','Frm_UpdateCommtrackNCDEXCODE','Y','A',0,232,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CommTrackData.frm','Comm Track Data','FrmRpt_CommTrackData','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("191")
MsgBox "Version 191 Installed successfully"

End Function

Public Function version190()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Mst_StorageContract Add WtQty Decimal (18,3) "
Call TmpTable

tmp = "Update Mst_StorageContract Set WtQty = 0 "
Call TmpTable


tmp = "Alter table Mst_StorageContract Add PeakFrom " & TextType & "(2),MasterStorageContractID  decimal(18,0)"
Call TmpTable


tmp = "Create Table Mst_SContractBillingDetail"
tmp = tmp & " (SCBDID  decimal(18,0) not null,"
tmp = tmp & " ContractID decimal(18,0),"
tmp = tmp & " FromDate datetime,"
tmp = tmp & " ToDate   datetime,"
tmp = tmp & " BillingCycleID decimal(18,0),"
tmp = tmp & " BillingQty decimal(18,3),"
tmp = tmp & " BillingRate decimal(18,2),"
tmp = tmp & " Remarks  " & TextType & "(255),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " CreatedBy   " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30))"

Call TmpTable

If Gen_DBType <> "MDB" Then

   tmp = "ALTER TABLE Mst_SContractBillingDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_SContractBillingDetail PRIMARY KEY CLUSTERED (SCBDID),"
   tmp = tmp & " CONSTRAINT DF_Mst_SContractBillingDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_SContractBillingDetail_Mst_StorageContract FOREIGN KEY (ContractID) REFERENCES Mst_StorageContract (ContractID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SContractBillingDetail_Mst_BillingCycle FOREIGN KEY (BillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID)"
   
   Call TmpTable

End If


UpdateConfig ("190")
MsgBox "Version 190 Installed successfully"
End Function


Public Function version189()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_UpdateInvoiceFAGAdjAmt.frm','Update Invoice FAG Adjustment Amount','Frm_UpdateInvoiceFAGAdjAmt','Y','T',0,228,1"
Call TmpTable

UpdateConfig ("189")
MsgBox "Version 189 Installed successfully"
End Function

Public Function version188()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_SecurityAttendance.frm','Security Attendance Report','FrmRpt_SecurityAttendance','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NewGodownMapping.frm','New Godown Mapping','FrmRpt_NewGodownMapping','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("188")
MsgBox "Version 188 Installed successfully"
End Function


Public Function version187()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Mst_GSL Add MadeupFlag " & TextType & " (1),ConnectedGSLID decimal(18,0) "
Call TmpTable

tmp = " Alter Table Txn_spillage Add BagsType " & TextType & " (1),NewGSLID decimal(18,0) "
Call TmpTable

tmp = "Update Mst_GSL Set MadeupFlag = 0 "
Call TmpTable

tmp = "Update Txn_Spillage Set BagsType = 'S'"
Call TmpTable

UpdateConfig ("187")
MsgBox "Version 187 Installed successfully"

End Function

Public Function Version186()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NNNRMismatchInvoice.frm','NNNR Mismatch Invoice','FrmRpt_NNNRMismatchInvoice','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("186")
MsgBox "Version 186 Installed successfully"
End Function


Public Function Version185()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Create table Mst_FF_RateCard_Godown_Fix ("
tmp = tmp & " FF_RateCard_Godown_FixID Decimal(18,0) Not Null,"
tmp = tmp & " FF_RateCard_GodownID Decimal(18,0),"
tmp = tmp & " FixAmount Decimal(18,2) ,"
tmp = tmp & " Remarks " & TextType & "(255),"
tmp = tmp & " StartDate Datetime,"
tmp = tmp & " EndDate Datetime,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency Datetime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate Datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate Datetime"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (FF_RateCard_Godown_FixID)"
End If
tmp = tmp & " )"


Call TmpTable

If Gen_DBType = "MDB" Then

   tmp = " ALTER TABLE Mst_FF_RateCard_Godown_Fix WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_FF_RateCard_Godown_Fix PRIMARY KEY CLUSTERED (FF_RateCard_Godown_FixID),"
   tmp = tmp & " CONSTRAINT DF_Mst_FF_RateCard_Godown_Fix_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Godown_Fix_Mst_FF_RateCard_Godown FOREIGN KEY (FF_RateCard_GodownID) REFERENCES Mst_FF_RateCard_Godown (FF_RateCard_GodownID))"
   
End If

tmp = "Alter Table Txn_InvMonGSLDetail  Add FranchProcessdBy " & TextType & "(30),FranchProcessedDate datetime"

Call TmpTable

UpdateConfig ("185")
MsgBox "Version 185 Installed successfully"

End Function


Public Function Version184()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " ALTER TABLE Mst_Client ADD " & C & " Status " & TextType & "(1)"
Call TmpTable

tmp = " Update Mst_Client  set Status = 'A'"
Call TmpTable

UpdateConfig ("184")
MsgBox "Version 184 Installed successfully"
End Function

Public Function version183()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_MenuAccess.frm','Menu Access','FrmRpt_MenuAccess','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("183")
MsgBox "Version 183 Installed successfully"
End Function

Public Function version182()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_MonthWiseSecurityDetailBilledUnbilled.frm','Month Wise Security Detail Billed Unbilled','FrmRpt_MonthWiseSecurityDetailBilledUnbilled','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("182")
MsgBox "Version 182 Installed successfully"
End Function


Public Function version181()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Log_Mst_CLientDepository"
tmp = tmp & " ("
tmp = tmp & " ClientDepositoryID Decimal(18,0),"
tmp = tmp & " ClientDepositorFlag " & TextType & "(1),"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " NAVCode " & TextType & "(25),"
tmp = tmp & " Remarks " & TextType & "(5000),"
tmp = tmp & " Concurrency DateTime"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> MDB Then
   tmp = " ALTER TABLE Log_Mst_CLientDepository WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Log_Mst_CLientDepository_Concurrency DEFAULT (getdate()) FOR Concurrency"
   Call TmpTable
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CashReceiptnotdepositedinBank.frm','Cash Receipt not deposited in Bank','FrmRpt_CashReceiptnotdepositedinBank','Y','R',0,255,1"
Call TmpTable


UpdateConfig ("181")
MsgBox "Version 181 Installed successfully"

End Function


Public Function version180()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Alter Table Mst_GSL Add Remarks " & TextType & " (8000)"
Call TmpTable

tmp = " Create Table Log_Mst_GSL"
tmp = tmp & " ("
tmp = tmp & " GSLID  Decimal(18,0),"
tmp = tmp & " SM_Prefix " & TextType & "(5),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " WHRBags Decimal(18,0),"
tmp = tmp & " WHRWeight Decimal(18,3),"
tmp = tmp & " CDTFBags Decimal(18,0),"
tmp = tmp & " CDTFWeight Decimal(18,3),"
tmp = tmp & " BalanceBags Decimal(18,0),"
tmp = tmp & " BalanceWeight Decimal(18, 3),"
tmp = tmp & " Remarks " & TextType & "(8000),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedOn DateTime"
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> MDB Then
   tmp = " ALTER TABLE Log_Mst_GSL WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Log_Mst_GSL_Concurrency DEFAULT (getdate()) FOR Concurrency"
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_GSL ON Mst_GSL"
   tmp = tmp & " AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mGSLID decimal(18, 0),"
   tmp = tmp & " @mSM_Prefix " & TextType & "(5),"
   tmp = tmp & " @mFlag " & TextType & "(1),"
   tmp = tmp & " @mWHRBags Decimal(18,0),"
   tmp = tmp & " @mWHRWeight Decimal(18,3),"
   tmp = tmp & " @mCDTFBags Decimal(18,0),"
   tmp = tmp & " @mCDTFWeight Decimal(18,3),"
   tmp = tmp & " @mBalanceBags Decimal(18,0),"
   tmp = tmp & " @mBalanceWeight Decimal(18,3),"
   tmp = tmp & " @mRemarks " & TextType & " (8000),"
   tmp = tmp & " @mUpdatedBy  " & TextType & " (30),"
   tmp = tmp & " @mUpdatedOn DateTime"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mGSLID = d.GSLID,"
   tmp = tmp & " @mSM_Prefix = d.SM_Prefix,"
   tmp = tmp & " @mFlag = d.Flag,"
   tmp = tmp & " @mWHRBags = d.WHRBags,"
   tmp = tmp & " @mWHRWeight = d.WHRWeight,"
   tmp = tmp & " @mCDTFBags = d.CDTFBags,"
   tmp = tmp & " @mCDTFWeight = d.CDTFWeight,"
   tmp = tmp & " @mBalanceBags = d.BalanceBags,"
   tmp = tmp & " @mBalanceWeight = d.BalanceWeight,"
   tmp = tmp & " @mRemarks = d.Remarks,"
   tmp = tmp & " @mUpdatedBy    = d.UpdatedBy,"
   tmp = tmp & " @mUpdatedOn    = d.UpdatedDate"
   tmp = tmp & " FROM inserted i"
   tmp = tmp & " INNER JOIN deleted d ON i.GSLID = d.GSLID and i.SM_Prefix = d.SM_Prefix"
   tmp = tmp & " insert into log_Mst_GSL (GSLID, SM_Prefix, Flag,  WHRBags, WHRWeight, CDTFBags, CDTFWeight, BalanceBags, BalanceWeight, Remarks, UpdatedBy, UpdatedOn)"
   tmp = tmp & " values (@mGSLID, @mSM_Prefix, @mFlag, @mWHRBags, @mWHRWeight, @mCDTFBags, @mCDTFWeight, @mBalanceBags, @mBalanceWeight, @mRemarks, @mUpdatedBy, @mUpdatedOn)"
   Call TmpTable
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_GSLRectification.frm','GSL Rectification','Frm_GSLRectification','Y','A',0,231,1"
Call TmpTable

UpdateConfig ("180")
MsgBox "Version 180 Installed successfully"
End Function


Public Function version179()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CDTFvsWHRMonitoringMIS.frm','CDTF vs WHR Monitoring MIS','FrmRpt_CDTFvsWHRMonitoringMIS','Y','R',0,229,1"
Call TmpTable

UpdateConfig ("179")
MsgBox "Version 179 Installed successfully"
End Function

Public Function Version178()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Txn_CDTFBOE Add TxnType " & TextType & "(2)"

Call TmpTable

UpdateConfig ("178")
MsgBox "Version 178 Installed successfully"

End Function
Public Function Version177()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create Table Log_EmpCMPMapping("
tmp = tmp & " LocationID Decimal(18,0),"
tmp = tmp & " CMPID Decimal(18,0),"
tmp = tmp & " EmployeeID Decimal(18,0),"
tmp = tmp & " Concurrency Datetime)"

Call TmpTable

UpdateConfig ("177")
MsgBox "Version 177 Installed successfully"

End Function
Public Function Version176()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Txn_FumigationApproval Add MonthYear Datetime "
Call TmpTable

UpdateConfig ("176")
MsgBox "Version 176 Installed successfully"

End Function

Public Function Version175()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_CxTF_Details Add OutstandingAmt decimal(18,2),CommodityValue decimal(18,2),BalQty decimal(18,3)"
Call TmpTable

UpdateConfig ("175")
MsgBox "Version 175 Installed successfully"

End Function


Public Function Version174()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_MonthwiseGodownRent.frm','Month wise Godown Rent MIS','FrmRpt_MonthwiseGodownRent','Y','R',0,229,1"

Call TmpTable

UpdateConfig ("174")
MsgBox "Version 174 Installed successfully"

End Function


Public Function Version173()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Alter Table Txn_LessorAgreement Add AgreenotReceieveFlag " & TextType & "(1)"
Call TmpTable

tmp = " Update Txn_LessorAgreement set AgreenotReceieveFlag = 0 "
Call TmpTable

tmp = "Alter Table Txn_InvMonGSLDetail Add " & C & " MonG_UID " & TextType & " (36) "
Call TmpTable

UpdateConfig ("173")
MsgBox "Version 173 Installed successfully"

End Function
Public Function Version172()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_MonthwiseSecAttendanceProcess.frm','Month wise Security Attendance Process','Frm_MonthwiseSecAttendanceProcess','Y','A',0,232,1"

Call TmpTable


UpdateConfig ("172")
MsgBox "Version 172 Installed successfully"

End Function
Public Function Version171()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Txn_TiDetails Add CollectionRemark " & TextType & "(2000),CRemarkUpdatedBy " & TextType & "(30),CRemarkUpdatedDate Datetime"

Call TmpTable

tmp = "Create Table Log_TICollectionRemark("
tmp = tmp & " TIID decimal(18,0),"
tmp = tmp & " OldCollectionRemark " & TextType & "(2000),"
tmp = tmp & " NewCollectionRemark " & TextType & "(2000),"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate Datetime)"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_UpdateTIRemark.frm','Collection Remark against Tax Invoices','FrmTxn_UpdateTIRemark','Y','T',0,228,1"
Call TmpTable

UpdateConfig ("171")
MsgBox "Version 171 Installed successfully"

End Function

Public Function Version170()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Mst_Users Add Flag " & TextType & "(2)"

Call TmpTable

tmp = "Create table Log_EmployeeAllocationDetails("
tmp = tmp & " TxnDate datetime,"
tmp = tmp & " EmployeeID decimal (18,0),"
tmp = tmp & " AllocationType " & TextType & "(2),"
tmp = tmp & " AllocationID decimal(18,0),"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " Concurrency datetime)"

Call TmpTable

UpdateConfig ("170")
MsgBox "Version 170 Installed successfully"

End Function

Public Function Version169()
Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Mst_SecurityRequest"
tmp = tmp & " ("
tmp = tmp & " SecurityRequestID Decimal(18,0) Not Null,"
tmp = tmp & " RequestDate DateTime,"
tmp = tmp & " LocationID Decimal(18,0),"
tmp = tmp & " RequestToID Decimal(18,0),"
tmp = tmp & " SecurityAgencyID Decimal(18,0),"
tmp = tmp & " NoofOpenGodowns Decimal(18,0),"
tmp = tmp & " NoofGuards Decimal(18,0),"
tmp = tmp & " NoofGodownsMapped Decimal(18,0),"
tmp = tmp & " TotalCost Decimal(18,2),"
tmp = tmp & " TotalAUM Decimal(18,2),"
tmp = tmp & " AgreeStartDate DateTime,"
tmp = tmp & " AgreeEndDate DateTime,"
tmp = tmp & " SalaryPayableUpto DateTime,"
tmp = tmp & " SecurityTypeID Decimal(18,0),"
tmp = tmp & " SecurityDesignationID Decimal(18,0),"
tmp = tmp & " RatePerMonth Decimal(18,2),"
tmp = tmp & " Status " & TextType & "(1),"
tmp = tmp & " Remarks " & TextType & "(2000),"
tmp = tmp & " ApprovedRemarks " & TextType & "(2000),"
tmp = tmp & " ApprovedDate DateTime,"
tmp = tmp & " ApprovedBy " & TextType & "(30),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (SecurityRequestID)"
End If
tmp = tmp & " )"
Call TmpTable


tmp = " Create Table Mst_SecurityRequestGodownMapping"
tmp = tmp & " ("
tmp = tmp & " SecurityRequestGodownMappingID Decimal(18,0) Not Null,"
tmp = tmp & " SecurityRequestID Decimal(18,0),"
tmp = tmp & " CMPID Decimal(18,0),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " BalanceBags Decimal(18,0),"
tmp = tmp & " BalanceQty Decimal(18,3),"
tmp = tmp & " AUM Decimal(18,2),"
tmp = tmp & " Remarks " & TextType & "(2000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (SecurityRequestID)"
End If
tmp = tmp & " )"
Call TmpTable


tmp = " Create Table Mst_SecLocGodMappingDetail"
tmp = tmp & " ("
tmp = tmp & " GodMapDID Decimal(18,0) Not Null,"
tmp = tmp & " TxnID Decimal(18,0),"
tmp = tmp & " TxnDetailID Decimal(18,0),"
tmp = tmp & " CMPID Decimal(18,0),"
tmp = tmp & " GodownID Decimal(18,0),"
tmp = tmp & " BalanceBags Decimal(18,0),"
tmp = tmp & " BalanceQty Decimal(18,3),"
tmp = tmp & " AUM Decimal(18,2),"
tmp = tmp & " Remarks " & TextType & "(2000),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DateTime"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (SecurityRequestID)"
End If
tmp = tmp & " )"
Call TmpTable


tmp = " Create Table Txn_SecurityAttendance"
tmp = tmp & " ("
tmp = tmp & " AttendanceID Decimal(18,0) Not Null,"
tmp = tmp & " TxnDetailID Decimal(18,0),"
tmp = tmp & " AttendanceDate DateTime,"
tmp = tmp & " GuardName " & TextType & "(200),"
tmp = tmp & " Status " & TextType & "(1),"
tmp = tmp & " Remarks " & TextType & "(2000),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,PRIMARY KEY (MGWPID)"
End If
tmp = tmp & " )"
Call TmpTable


tmp = " Create Table Log_Txn_SecurityAttendance"
tmp = tmp & " ("
tmp = tmp & " AttendanceID    decimal (18, 0),"
tmp = tmp & " TxnDetailID decimal (18, 0),"
tmp = tmp & " AttendanceDate  datetime,"
tmp = tmp & " GuardName   " & TextType & " (200),"
tmp = tmp & " UpdatedOn   datetime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30),"
tmp = tmp & " Remarks " & TextType & " (2000),"
tmp = tmp & " Status VarChar(1)"
tmp = tmp & " )"
Call TmpTable


tmp = " Create Table Log_Mst_SecLocMappingDetail"
tmp = tmp & " ("
tmp = tmp & " TxnDetailID decimal (18,0),"
tmp = tmp & " TxnID decimal (18,0),"
tmp = tmp & " SecurityTypeID decimal (18,0),"
tmp = tmp & " SecurityDesignationID decimal (18,0),"
tmp = tmp & " StartDate datetime,"
tmp = tmp & " EndDate datetime ,"
tmp = tmp & " NoofGodown decimal (18,0),"
tmp = tmp & " RatePerMonth decimal (18,2),"
tmp = tmp & " Remarks varchar (2000),"
tmp = tmp & " UpdatedBy varchar (30),"
tmp = tmp & " UpdatedOn datetime,"
tmp = tmp & " SalaryPaidUpto datetime,"
tmp = tmp & " RequestID decimal(18, 0)"
tmp = tmp & " )"
Call TmpTable

tmp = " Alter Table Mst_SecLocMapping"
tmp = tmp & " Add " & C & " AgreementStartDate DateTime, " & C & " AgreementEndDate DateTime"
Call TmpTable

tmp = " Alter Table Mst_SecLocMappingDetail Add " & C & " RequestID Decimal(18,0)"
Call TmpTable




If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Mst_SecurityRequest WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_SecurityRequest PRIMARY KEY CLUSTERED (SecurityRequestID),"
   tmp = tmp & " CONSTRAINT DF_Mst_SecurityRequest_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequest_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequest_Mst_Employee FOREIGN KEY (RequestToID) REFERENCES Mst_Employee (EmployeeID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequest_Mst_SecurityShifts FOREIGN KEY (SecurityTypeID) REFERENCES Mst_SecurityShifts (SecurityShiftsID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequest_Mst_SecurityAgency FOREIGN KEY (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequest_Mst_SecurityDesignation FOREIGN KEY (SecurityDesignationID) REFERENCES Mst_SecurityDesignation (SecurityDesignationID)"
   
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_SecurityRequestGodownMapping WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_SecurityRequestGodownMapping PRIMARY KEY CLUSTERED (SecurityRequestGodownMappingID),"
   tmp = tmp & " CONSTRAINT DF_Mst_SecurityRequestGodownMapping_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequestGodownMapping_Mst_SecurityRequest FOREIGN KEY (SecurityRequestID) REFERENCES Mst_SecurityRequest(SecurityRequestID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequestGodownMapping_Mst_CMP FOREIGN KEY (CMPID) REFERENCES Mst_CMP(CMPID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecurityRequestGodownMapping_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID)"
   
   Call TmpTable
   
   tmp = " ALTER TABLE Mst_SecLocMappingDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Mst_SecLocMappingDetail_Mst_SecurityRequest FOREIGN KEY (RequestID) REFERENCES Mst_SecurityRequest(SecurityRequestID)"
   
   Call TmpTable

   tmp = " ALTER TABLE Mst_SecLocGodMappingDetail WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_SecLocGodMappingDetail PRIMARY KEY CLUSTERED (GodMapDID),"
   tmp = tmp & " CONSTRAINT DF_Mst_SecLocGodMappingDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_SecLocGodMappingDetail_Mst_SecLocMapping FOREIGN KEY (TxnID) REFERENCES Mst_SecLocMapping(TxnID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecLocGodMappingDetail_Mst_SecLocMappingDetail FOREIGN KEY (TxnDetailID) REFERENCES Mst_SecLocMappingDetail(TxnDetailID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecLocGodMappingDetail_Mst_CMP FOREIGN KEY (CMPID) REFERENCES Mst_CMP(CMPID),"
   tmp = tmp & " CONSTRAINT FK_Mst_SecLocGodMappingDetail_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID)"
   
   Call TmpTable

   tmp = " ALTER TABLE Txn_SecurityAttendance WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_SecurityAttendance PRIMARY KEY CLUSTERED (AttendanceID),"
   tmp = tmp & " CONSTRAINT DF_Txn_SecurityAttendance_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_SecurityAttendance_Mst_SecLocMappingDetail FOREIGN KEY (TxnDetailID) REFERENCES Mst_SecLocMappingDetail(TxnDetailID)"
   
   Call TmpTable

   tmp = " Create Trigger tr_Txn_SecurityAttendance"
   tmp = tmp & " ON Txn_SecurityAttendance AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mAttendanceID  decimal (18, 0),"
   tmp = tmp & " @mTxnDetailID   decimal (18, 0),"
   tmp = tmp & " @mAttendanceDate    datetime,"
   tmp = tmp & " @mGuardName " & TextType & " (200),"
   tmp = tmp & " @mUpdatedOn datetime,"
   tmp = tmp & " @mUpdatedBy " & TextType & "(30),"
   tmp = tmp & " @mRemarks   " & TextType & " (2000),"
   tmp = tmp & " @mStatus " & TextType & "(1)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mAttendanceID  = d.AttendanceID,"
   tmp = tmp & " @mTxnDetailID   = d.TxnDetailID,"
   tmp = tmp & " @mAttendanceDate = d.AttendanceDate,"
   tmp = tmp & " @mGuardName = d.GuardName,"
   tmp = tmp & " @mUpdatedOn = i.UpdatedDate,"
   tmp = tmp & " @mUpdatedBy = i.UpdatedBy,"
   tmp = tmp & " @mRemarks   = d.Remarks,"
   tmp = tmp & " @mStatus = d.Status"
   tmp = tmp & " FROM    inserted i INNER JOIN deleted d ON i.AttendanceID = d.AttendanceID"
   tmp = tmp & " insert into Log_Txn_SecurityAttendance"
   tmp = tmp & " (AttendanceID, TxnDetailID, AttendanceDate, GuardName, Remarks, Status, UpdatedOn, UpdatedBy)"
   tmp = tmp & " Values"
   tmp = tmp & " (@mAttendanceID, @mTxnDetailID, @mAttendanceDate, @mGuardName, @mRemarks, @mStatus, @mUpdatedOn, @mUpdatedBy)"
   
   Call TmpTable
   
   tmp = " Create Trigger tr_Mst_SecLocMappingDetail"
   tmp = tmp & " ON Mst_SecLocMappingDetail AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mTxnDetailID decimal (18,0),"
   tmp = tmp & " @mTxnID decimal (18,0),"
   tmp = tmp & " @mSecurityTypeID decimal (18,0),"
   tmp = tmp & " @mSecurityDesignationID decimal (18,0),"
   tmp = tmp & " @mStartDate datetime,"
   tmp = tmp & " @mEndDate datetime ,"
   tmp = tmp & " @mNoofGodown decimal (18,0),"
   tmp = tmp & " @mRatePerMonth decimal (18,2),"
   tmp = tmp & " @mRemarks varchar (2000),"
   tmp = tmp & " @mUpdatedBy varchar (30),"
   tmp = tmp & " @mUpdatedOn datetime,"
   tmp = tmp & " @mSalaryPaidUpto datetime,"
   tmp = tmp & " @mRequestID decimal (18,0)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mTxnDetailID = d.TxnDetailID,"
   tmp = tmp & " @mTxnID  = d.TxnID,"
   tmp = tmp & " @mSecurityTypeID  = d.SecurityTypeID,"
   tmp = tmp & " @mSecurityDesignationID = d.SecurityDesignationID,"
   tmp = tmp & " @mStartDate = d.StartDate,"
   tmp = tmp & " @mEndDate = d.EndDate,"
   tmp = tmp & " @mNoofGodown = d.NoofGodown,"
   tmp = tmp & " @mRatePerMonth = d.RatePerMonth,"
   tmp = tmp & " @mRemarks = d.Remarks,"
   tmp = tmp & " @mUpdatedOn = i.CreatedDate,"
   tmp = tmp & " @mUpdatedBy = i.CreatedBy,"
   tmp = tmp & " @mSalaryPaidUpto = d.SalaryPaidUpto,"
   tmp = tmp & " @mRequestID = d.RequestID"
   tmp = tmp & " FROM    inserted i INNER JOIN deleted d ON i.TxnDetailID = d.TxnDetailID"
   tmp = tmp & " insert into Log_Mst_SecLocMappingDetail"
   tmp = tmp & " (TxnDetailID , TxnID , SecurityTypeID , SecurityDesignationID , StartDate , EndDate , NoofGodown , RatePerMonth ,Remarks ,"
   tmp = tmp & " UpdatedBy , UpdatedOn ,  SalaryPaidUpto , RequestID)"
   tmp = tmp & " Values"
   tmp = tmp & " (@mTxnDetailID , @mTxnID , @mSecurityTypeID , @mSecurityDesignationID , @mStartDate , @mEndDate , @mNoofGodown , @mRatePerMonth ,@mRemarks ,"
   tmp = tmp & " @mUpdatedBy , @mUpdatedOn ,  @mSalaryPaidUpto , @mRequestID)"
   
   Call TmpTable

   tmp = " Alter Table Txn_CXTF_details drop constraint FK_Txn_CxTF_Details_Txn_TIDetails "
   
   Call TmpTable

End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_SecurityRequest.frm','Security Request','FrmMst_SecurityRequest','Y','M',0,227,1"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_SecurityAttendance.frm','Security Daily Attendance','FrmTxn_SecurityAttendance','Y','T',0,228,1"

Call TmpTable


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1, 'FrmRpt_SecurityRequestGodownMapping.frm', 'Security Request Godown Mapping', 'FrmRpt_SecurityRequestGodownMapping', 'Y','R',0,229,1"

Call TmpTable


UpdateConfig ("169")
MsgBox "Version 169 Installed successfully"

End Function



Public Function Version168()
Dim TextType, C As Variant

If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "CREATE TABLE Txn_FumigationApproval("
tmp = tmp & " MappingID decimal(18,0) NOT NULL ,"
tmp = tmp & " MappingDate datetime ,"
tmp = tmp & " VendorID decimal(18,0),"
tmp = tmp & " LocationID decimal(18,0),"
tmp = tmp & " ActualQty decimal(18,3),"
tmp = tmp & " PayableAmount decimal(18,2),"
tmp = tmp & " ActualAmount decimal(18,2),"
tmp = tmp & " BillNo " & TextType & "(1000),"
tmp = tmp & " BillDate DateTime,"
tmp = tmp & " STAmount Decimal(18,2),"
tmp = tmp & " ServiceTaxID Decimal(18,0),"
tmp = tmp & " Remark " & TextType & "(2000),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " BatchID Decimal(18,0),"
tmp = tmp & " ForwardTo Decimal(18,0),"
tmp = tmp & " ApprovalDate datetime,"
tmp = tmp & " ApprovalRemark " & TextType & "(250),"
tmp = tmp & " Flag " & TextType & "(1),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedDate datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,Primary Key(MappingID)"
End If
tmp = tmp & " )"

Call TmpTable


tmp = "Alter Table Txn_Fumigation Add " & C & " MappingID Decimal(18,0)," & C & " VendorQuantity Decimal(18,3)," & C & "VendorRate decimal(18,2)," & C & " BillFlag " & TextType & "(1)"
Call TmpTable
tmp = "Alter Table Txn_Spraying Add " & C & " MappingID Decimal(18,0)," & C & " VendorQuantity Decimal(18,3)," & C & " VendorRate decimal(18,2)," & C & " BillFlag " & TextType & "(1)"
Call TmpTable

tmp = "Create table Log_FumigationApproval("
tmp = tmp & " MappingID decimal(18,0),"
tmp = tmp & " OldForwardToID decimal(18,0),"
tmp = tmp & " OldActualAmt decimal(18,2),"
tmp = tmp & " NewActualAmt decimal(18,2),"
tmp = tmp & " NewForwardToID decimal(18,0),"
tmp = tmp & " UpdatedBy " & TextType & "(36),"
tmp = tmp & " Concurrency datetime)"

Call TmpTable

tmp = "Create table Log_ISINCommodity("
tmp = tmp & " ISINID " & TextType & "(25),"
tmp = tmp & " OldCommodityID decimal(18,0),"
tmp = tmp & " NewCommodityID decimal(18,0),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate datetime)"

Call TmpTable

If Gen_DBType <> "MDB" Then
    
   tmp = "ALTER TABLE Log_FumigationApproval WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Log_FumigationApproval_Concurrency DEFAULT (getdate()) FOR Concurrency"
   
   Call TmpTable
   
   tmp = " ALTER TABLE Txn_FumigationApproval WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Txn_FumigationApproval PRIMARY KEY CLUSTERED (MappingID),"
   tmp = tmp & " CONSTRAINT DF_Txn_FumigationApproval_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Txn_FumigationApproval_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
   tmp = tmp & " CONSTRAINT FK_Txn_FumigationApproval_Mst_Vendor FOREIGN KEY (VendorID) REFERENCES Mst_Vendor (VendorID)"
   
   Call TmpTable
    
   tmp = "CREATE TRIGGER tr_Txn_FumigationApproval"
   tmp = tmp & " ON Txn_FumigationApproval AFTER UPDATE AS"
   tmp = tmp & " DECLARE"
   tmp = tmp & " @mMappingID decimal(18,0),"
   tmp = tmp & " @mOldForwardToID decimal(18,0),"
   tmp = tmp & " @mOldActualAmt decimal(18,2),"
   tmp = tmp & " @mNewActualAmt decimal(18,2),"
   tmp = tmp & " @mNewForwardToID decimal(18,0),"
   tmp = tmp & " @mUpdatedBy varchar(36)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mMappingID= d.MappingID,"
   tmp = tmp & " @mOldForwardToID = d.ForwardTo,"
   tmp = tmp & " @mOldActualAmt = d.ActualAmount,"
   tmp = tmp & " @mNewActualAmt = i.ActualAmount,"
   tmp = tmp & " @mNewForwardToID = i.ForwardTo,"
   tmp = tmp & " @mUpdatedBy = i.UpdatedBy"
   tmp = tmp & " FROM inserted i INNER JOIN deleted d"
   tmp = tmp & " ON i.MappingID = d.MappingID"
   tmp = tmp & " insert into Log_FumigationApproval"
   tmp = tmp & " (MappingID, OldForwardToID, OldActualAmt, NewActualAmt, NewForwardToID, UpdatedBy)"
   tmp = tmp & " Values"
   tmp = tmp & " (@mMappingID, @mOldForwardToID, @mOldActualAmt, @mNewActualAmt, @mNewForwardToID, @mUpdatedBy)"
   
   Call TmpTable
   
   tmp = "ALTER TABLE Log_ISINCommodity WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Log_ISINCommodity_Concurrency DEFAULT (getdate()) FOR Concurrency"
   
   Call TmpTable
   
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_FumigationApproval.frm','Fumigation Approval','FrmTxn_FumigationApproval','Y','T',0,228,1"
Call TmpTable

UpdateConfig ("168")
MsgBox "Version 168 Installed successfully"

End Function

Public Function Version167()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_InvMonGSLDetail Add " & C & " ProcessedBy " & TextType & "(30)"
Call TmpTable

UpdateConfig ("167")
MsgBox "Version 167 Installed successfully"

End Function
Public Function Version166()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_CxTF_Details Add " & C & " TIID Decimal(18,0)"
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Txn_CxTF_Details WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Txn_CxTF_Details_Txn_TIDetails FOREIGN KEY (TIID) REFERENCES Txn_TIDetails (TIID)"

   Call TmpTable
End If

UpdateConfig ("166")
MsgBox "Version 166 Installed successfully"

End Function

Public Function Version165()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

If Gen_DBType <> "MDB" Then
   tmp = "Alter Table Txn_InvMonGSLDetail Drop Constraint FK_Txn_InvMonGSLDetail_Mst_GSL "
   Call TmpTable
End If

UpdateConfig ("165")
MsgBox "Version 165 Installed successfully"

End Function
Public Function Version164()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter table Txn_CxTF_Details Add ApprovePrintFlag " & TextType & "(1),ApprovePrintBy " & TextType & "(25),ApprovePrintRemark " & TextType & "(255),ApprovePrintDate Datetime "
Call TmpTable

tmp = "Alter table Txn_CxTF_GSL Add TruckNo " & TextType & "(50),WeighBridgeSlipNo " & TextType & "(50) "

Call TmpTable

tmp = "Alter table Mst_ExchangeType add MultOfAll decimal(18,0),MultOfAM decimal(18,0) "

Call TmpTable

tmp = "Update Mst_ExchangeType set MultOfAll=10,MultOfAM=5 Where ExchangeTypeID=1 "
Call TmpTable

tmp = " Alter Table Mst_Location add " & C & " HRLocationID decimal(18,0)"
Call TmpTable


UpdateConfig ("164")
MsgBox "Version 164 Installed successfully"

End Function

Public Function Version163()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownWiseBalanceStockandCommodityValue.frm','Godown Wise Balance Stock and Commodity Value','FrmRpt_GodownWiseBalanceStockandCommodityValue','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("163")
MsgBox "Version 163 Installed successfully"

End Function

Public Function Version162()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Alter Table Txn_ProInvoice Add " & C & " RejectionRemarks " & TextType & " (2000)"
Call TmpTable

UpdateConfig ("162")
MsgBox "Version 162 Installed successfully"

End Function
Public Function Version161()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = "Create Table Log_Client2ClientTransfer("
tmp = tmp & " TxnDate Datetime,"
tmp = tmp & " OldNAVCode " & TextType & "(25),"
tmp = tmp & " NewNAVCode " & TextType & "(25),"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " ClientIDRows " & TextType & "(1000),"
tmp = tmp & " Concurrency datetime)"

Call TmpTable


If Gen_DBType <> "MDB" Then
    
   tmp = " ALTER TABLE Log_Client2ClientTransfer ADD"
   tmp = tmp & " CONSTRAINT DF_Log_Client2ClientTransfer_Concurrency DEFAULT (getdate()) FOR Concurrency"
   Call TmpTable

End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_Cust2CustTransfer.frm','Blocking of Customer Account','FrmTxn_Cust2CustTransfer','Y','T',0,228,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportReverseDPInvoice.frm','Export Reverse DP Invoice','Frm_ExportReverseDPInvoice','Y','C',0,230,1"
Call TmpTable

UpdateConfig ("161")
MsgBox "Version 161 Installed successfully"

End Function

Public Function Version160()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Alter Table Mst_State add " & C & " RegionID Decimal(18,0)"
Call TmpTable

tmp = " Create Table Mst_EmployeeAllocationDetails"
tmp = tmp & " ("
tmp = tmp & " EmpAllDetailID Decimal(18,0) Not Null,"
tmp = tmp & " EmployeeID Decimal(18,0),"
tmp = tmp & " AllocationType " & TextType & "(2),"
tmp = tmp & " AllocationID Decimal(18,0),"
tmp = tmp & " G_UID   " & TextType & "(36),"
tmp = tmp & " Concurrency datetime,"
tmp = tmp & " CreatedBy " & TextType & "(36),"
tmp = tmp & " CreatedDate datetime,"
tmp = tmp & " UpdatedBy " & TextType & "(36),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
    tmp = tmp & " ,Primary Key(EmpAllDetailID)"
End If
tmp = tmp & " )"

Call TmpTable


If Gen_DBType <> "MDB" Then
   
    
   tmp = " ALTER TABLE Mst_State WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT FK_Mst_State_Mst_Region FOREIGN KEY (RegionID) REFERENCES Mst_Region (RegionID)"
   Call TmpTable

   tmp = " ALTER TABLE Mst_EmployeeAllocationDetails WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_EmployeeAllocationDetails PRIMARY KEY CLUSTERED (EmpAllDetailID),"
   tmp = tmp & " CONSTRAINT DF_Mst_EmployeeAllocationDetails_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_EmployeeAllocationDetails_Mst_Employee FOREIGN KEY (EmployeeID) REFERENCES Mst_Employee (EmployeeID)"
   Call TmpTable

End If


UpdateConfig ("160")
MsgBox "Version 160 Installed successfully"

End Function

Public Function Version159()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create  Table Mst_LocationCommodityMapping"
tmp = tmp & " ("
tmp = tmp & " LocationCommodityMappingID Decimal(18,0) not null,"
tmp = tmp & " LocationID Decimal(18, 0),"
tmp = tmp & " CommodityID Decimal(18, 0),"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime"
If Gen_DBType = "MDB" Then
   tmp = " ,PRIMARY KEY (LocationCommodityMappingID)"
End If
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then
   
   tmp = " ALTER TABLE Mst_LocationCommodityMapping WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT PK_Mst_LocationCommodityMapping PRIMARY KEY CLUSTERED (LocationCommodityMappingID),"
   tmp = tmp & " CONSTRAINT DF_Mst_LocationCommodityMapping_Concurrency DEFAULT (getdate()) FOR Concurrency,"
   tmp = tmp & " CONSTRAINT FK_Mst_LocationCommodityMapping_Mst_Commodity FOREIGN KEY (CommodityID) REFERENCES Mst_Commodity (CommodityID),"
   tmp = tmp & " CONSTRAINT FK_Mst_LocationCommodityMapping_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID)"
   
   Call TmpTable

End If

tmp = " Alter Table Mst_Godown Add " & C & " LeasedUptoDate DateTime"
Call TmpTable

tmp = " Update Mst_Godown Set LeasedUptoDate = '" & Date & "' Where UnitTypeID = 6"
Call TmpTable

tmp = " Alter Table Txn_Fumigation Add " & C & " Status " & TextType & "(1), " & C & " ApprovalDate DateTime, " & C & " ApprovalRemark " & TextType & "(1000), " & C & " BatchID Decimal(18,0), " & C & " ForwardTo Decimal(18, 0)"
Call TmpTable


tmp = " Create Table Mst_ChequeReturnReason"
tmp = tmp & " ( "
tmp = tmp & " CRRID Decimal(18,0) Not null,"
tmp = tmp & " Reason " & TextType & "(1000),"
tmp = tmp & " GenerateAdHoc " & TextType & "(1),"
tmp = tmp & " ReturnCharges decimal(18,2) , "
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30),"
tmp = tmp & " UpdatedDate DateTime,"
tmp = tmp & " UpdatedBy " & TextType & "(30)"

If Gen_DBType = "MDB" Then
    tmp = " ,PRIMARY KEY (CRRID)"
End If

tmp = tmp & " )"

Call TmpTable

tmp = "Create Table Txn_ChequeDishonour"
tmp = tmp & " ( "
tmp = tmp & " ChequeDishonourID Decimal(18,0) Not null,"
tmp = tmp & " CRID Decimal(18,0),"
tmp = tmp & " CRRID Decimal(18, 0),"
tmp = tmp & " SM_Prefix " & TextType & "(5),"
tmp = tmp & " CRDate DateTime,"
tmp = tmp & " LocationID Decimal(18,0),"
tmp = tmp & " InstrumentType " & TextType & "(50),"
tmp = tmp & " DrawnOn " & TextType & "(150),"
tmp = tmp & " BranchName " & TextType & "(150),"
tmp = tmp & " InstrumentNo " & TextType & "(25),"
tmp = tmp & " InstrumentDate DateTime,"
tmp = tmp & " InstrumentAmount Decimal(18,2),"
tmp = tmp & " ReceivedFrom " & TextType & "(100),"
tmp = tmp & " Remark " & TextType & "(255),"
tmp = tmp & " ReturnRemark " & TextType & "(100),"
tmp = tmp & " ChequeReturnChargesAmount Decimal(18, 2),"
tmp = tmp & " GenerateAdHocInvoice " & TextType & "(1),"
tmp = tmp & " PINo Decimal(18,0),"
tmp = tmp & " Concurrency DateTime,"
tmp = tmp & " G_UID " & TextType & "(36),"
tmp = tmp & " CreatedDate DateTime,"
tmp = tmp & " CreatedBy " & TextType & "(30)"
If Gen_DBType = "MDB" Then
   tmp = " ,PRIMARY KEY (ChequeDishonourID)"
End If
tmp = tmp & " )"
Call TmpTable

If Gen_DBType <> "MDB" Then

    tmp = " ALTER TABLE Mst_ChequeReturnReason WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Mst_ChequeReturnReason PRIMARY KEY CLUSTERED (CRRID),"
    tmp = tmp & " CONSTRAINT DF_Mst_ChequeReturnReason_Concurrency DEFAULT (getdate()) FOR Concurrency"
    Call TmpTable
    
    
    tmp = " ALTER TABLE Txn_ChequeDishonour WITH NOCHECK ADD"
    tmp = tmp & " CONSTRAINT PK_Txn_ChequeDishonour PRIMARY KEY CLUSTERED (ChequeDishonourID),"
    tmp = tmp & " CONSTRAINT DF_Txn_ChequeDishonour_Concurrency DEFAULT (getdate()) FOR Concurrency,"
    tmp = tmp & " CONSTRAINT FK_Txn_ChequeDishonour_Txn_NewCashReceipt FOREIGN KEY (CRID, SM_Prefix) REFERENCES Txn_NewCashReceipt (CRID, SM_Prefix),"
    tmp = tmp & " CONSTRAINT FK_Txn_ChequeDishonour_Mst_ChequeReturnReason FOREIGN KEY (CRRID) REFERENCES Mst_ChequeReturnReason(CRRID),"
    tmp = tmp & " CONSTRAINT FK_Txn_ChequeDishonour_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID)"

    Call TmpTable

   tmp = " ALTER TABLE Txn_Fumigation WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Txn_Fumigation_BatchID DEFAULT (0) FOR BatchID,"
   tmp = tmp & " CONSTRAINT FK_Txn_Fumigation_Mst_Employee FOREIGN KEY (ForwardTo) REFERENCES Mst_Employee (EmployeeID)"
      
   Call TmpTable


End If


tmp = " Alter Table Mst_DPInvoiceUpload Add " & C & " ReverseRemark " & TextType & "(2000), " & C & " ReverseFlag " & TextType & "(1)," & C & " ReverseDate datetime "
Call TmpTable

tmp = " Update Mst_DPInvoiceUpload Set ReverseFlag = 'N' "

Call TmpTable


tmp = " Alter Table Mst_DPInvoiceUpload Add  RBatchID Decimal (18,0) "
Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " ALTER TABLE Mst_DPInvoiceUpload WITH NOCHECK ADD CONSTRAINT DF_Mst_DPInvoiceUpload_RBatchID DEFAULT (0) FOR RBatchID "
   TmpTable
End If

tmp = " Update Mst_DPInvoiceUpload Set RBatchID = 0"
TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportFumigationExpense.frm','Export Fumigation Expense','Frm_ExportFumigationExpense','Y','C',0,230,1"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_DPInvoice.frm','DP Invoices','FrmTxn_DPInvoice','Y','T',0,228,1"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_ChequeReturnReason.frm','Cheque Dishounor Reason','FrmMst_ChequeReturnReason','Y','M',0,227,1"

Call TmpTable


UpdateConfig ("159")
MsgBox "Version 159 Installed successfully"

End Function

Public Function Version158()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NCDEXRateCard.frm','NCDEX Rate Card','FrmRpt_NCDEXRateCard','Y','R',0,229,1"
Call TmpTable


UpdateConfig ("158")
MsgBox "Version 158 Installed successfully"

End Function

Public Function Version157()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If

tmp = " Create table Log_FranchiseeFeeHistory"
tmp = tmp & " ("
tmp = tmp & " ClaimID   decimal(18,0),"
tmp = tmp & " OldForwardToID    decimal(18,0),"
tmp = tmp & " OldActualFee decimal(18,2),"
tmp = tmp & " NewActualFee decimal(18,2),"
tmp = tmp & " NewForwardToID    decimal(18,0),"
tmp = tmp & " UpdatedBy varchar(36),"
tmp = tmp & " Concurrency   datetime"
tmp = tmp & " )"

Call TmpTable

tmp = " Create Table Log_FranchiseeFeeRejectHistory"
tmp = tmp & " ("
tmp = tmp & " ClaimID decimal(18,0),"
tmp = tmp & " CMPName " & TextType & "(100),"
tmp = tmp & " GSLID Decimal(18,0),"
tmp = tmp & " GodownNo " & TextType & "(25), "
tmp = tmp & " StackNo " & TextType & "(10),"
tmp = tmp & " LotNo " & TextType & "(10),"
tmp = tmp & " Commodity " & TextType & "(75),"
tmp = tmp & " NoOfBags Decimal(18,0),"
tmp = tmp & " Quantity Decimal(18,3),"
tmp = tmp & " TransactionFlag " & TextType & "(2),"
tmp = tmp & " CxTFNo Varchar(25),"
tmp = tmp & " BillingCycle " & TextType & "(50),"
tmp = tmp & " BillingUnit " & TextType & "(10),"
tmp = tmp & " NoofCycle Decimal(18,0),"
tmp = tmp & " RateperCycle Decimal(18,2),"
tmp = tmp & " Amount Decimal(18,2),"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
    
   tmp = " CREATE TRIGGER tr_Txn_FranchiseeRent "
   tmp = tmp & " ON Txn_FranchiseeRent  AFTER UPDATE AS "
   tmp = tmp & " DECLARE "
   tmp = tmp & " @mClaimID decimal(18,0),"
   tmp = tmp & " @mOldForwardToID decimal(18,0),"
   tmp = tmp & " @mOldActualFee decimal(18,2),"
   tmp = tmp & " @mNewActualFee decimal(18,2),"
   tmp = tmp & " @mNewForwardToID decimal(18,0),"
   tmp = tmp & " @mUpdatedBy   varchar(36)"
   tmp = tmp & " SELECT  "
   tmp = tmp & " @mClaimID= d.ClaimID, "
   tmp = tmp & " @mOldForwardToID = d.ForwardToID,     "
   tmp = tmp & " @mOldActualFee = d.ActualFee,"
   tmp = tmp & " @mNewActualFee = i.ActualFee,"
   tmp = tmp & " @mNewForwardToID = i.ForwardToID,"
   tmp = tmp & " @mUpdatedBy = i.UpdatedBy  "
   tmp = tmp & " FROM    inserted i INNER JOIN deleted d "
   tmp = tmp & " ON i.ClaimID = d.ClaimID  "
   tmp = tmp & " insert into Log_FranchiseeFeeHistory "
   tmp = tmp & " (ClaimID, OldForwardToID, OldActualFee, NewActualFee, NewForwardToID, UpdatedBy)"
   tmp = tmp & " values"
   tmp = tmp & " (@mClaimID, @mOldForwardToID, @mOldActualFee, @mNewActualFee, @mNewForwardToID, @mUpdatedBy)"
    
   Call TmpTable

   tmp = " ALTER TABLE Log_FranchiseeFeeHistory WITH NOCHECK ADD"
   tmp = tmp & " CONSTRAINT DF_Log_FranchiseeFeeHistory_Concurrency DEFAULT (getdate()) FOR Concurrency "
   
   Call TmpTable

End If

tmp = " Create Table Log_Mst_Employee"
tmp = tmp & " ("
tmp = tmp & " EmployeeID        Decimal(18, 0),"
tmp = tmp & " EmployeeNo        " & TextType & "(25),"
tmp = tmp & " EmployeeName      " & TextType & "(100),"
tmp = tmp & " DesignationID     Decimal(18, 0),"
tmp = tmp & " SelfReporting     " & TextType & "(1),"
tmp = tmp & " DirectReportingTo Decimal(18, 0),"
tmp = tmp & " ReportingToDRC    Decimal(18, 0),"
tmp = tmp & " ConnectedToRAG    Decimal(18, 0),"
tmp = tmp & " EmployeeEmailID   " & TextType & "(100),"
tmp = tmp & " MobileNo          " & TextType & "(50),"
tmp = tmp & " Address           " & TextType & "(250),"
tmp = tmp & " Flag              " & TextType & "(1),"
tmp = tmp & " NAVCode           " & TextType & "(25),"
tmp = tmp & " AllocateExpense   " & TextType & "(1),"
tmp = tmp & " CMPAllocation     " & TextType & "(1),"
tmp = tmp & " BaseLocationID    Decimal(18,0),"
tmp = tmp & " FAGUpdated        " & TextType & "(30),"
tmp = tmp & " FAGUpdatedDate    DateTime,"
tmp = tmp & " UpdatedOn         DateTime,"
tmp = tmp & " UpdatedBy         " & TextType & "(30)"
tmp = tmp & " )"

Call TmpTable

tmp = " Create Table Log_Mst_EMPCMPMapping"
tmp = tmp & " ("
tmp = tmp & " EmpCMPMapID       Decimal(18, 0),"
tmp = tmp & " LocationID        Decimal(18, 0),"
tmp = tmp & " CMPID             Decimal(18, 0),"
tmp = tmp & " EmployeeID        Decimal(18, 0),"
tmp = tmp & " DesignationID     Decimal(18, 0),"
tmp = tmp & " SM_Prefix         " & TextType & "(5),"
tmp = tmp & " Flag              " & TextType & "(1),"
tmp = tmp & " UpdatedOn         DateTime,"
tmp = tmp & " UpdatedBy         " & TextType & "(30)"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
   tmp = " CREATE TRIGGER tr_Mst_Employee ON Mst_Employee"
   tmp = tmp & " For Update, Delete"
   tmp = tmp & " AS DECLARE"
   tmp = tmp & " @mEmployeeID        Decimal(18, 0),"
   tmp = tmp & " @mEmployeeNo        " & TextType & "(25),"
   tmp = tmp & " @mEmployeeName      " & TextType & "(100),"
   tmp = tmp & " @mDesignationID     Decimal(18, 0),"
   tmp = tmp & " @mSelfReporting     " & TextType & "(1),"
   tmp = tmp & " @mDirectReportingTo Decimal(18, 0),"
   tmp = tmp & " @mReportingToDRC    Decimal(18, 0),"
   tmp = tmp & " @mConnectedToRAG    Decimal(18, 0),"
   tmp = tmp & " @mEmployeeEmailID   " & TextType & "(100),"
   tmp = tmp & " @mMobileNo          " & TextType & "(50),"
   tmp = tmp & " @mAddress           " & TextType & "(250),"
   tmp = tmp & " @mFlag              " & TextType & "(1),"
   tmp = tmp & " @mNAVCode           " & TextType & "(25),"
   tmp = tmp & " @mAllocateExpense   " & TextType & "(1),"
   tmp = tmp & " @mCMPAllocation     " & TextType & "(1),"
   tmp = tmp & " @mBaseLocationID    Decimal(18,0),"
   tmp = tmp & " @mFAGUpdated        " & TextType & "(30),"
   tmp = tmp & " @mFAGUpdatedDate    DateTime,"
   tmp = tmp & " @mUpdatedOn         DateTime,"
   tmp = tmp & " @mUpdatedBy         " & TextType & "(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mEmployeeID        = d.EmployeeID,"
   tmp = tmp & " @mEmployeeNo        = d.EmployeeNo,"
   tmp = tmp & " @mEmployeeName      = d.EmployeeName,"
   tmp = tmp & " @mDesignationID     = d.DesignationID,"
   tmp = tmp & " @mSelfReporting     = d.SelfReporting,"
   tmp = tmp & " @mDirectReportingTo = d.DirectReportingTo,"
   tmp = tmp & " @mReportingToDRC    = d.ReportingToDRC,"
   tmp = tmp & " @mConnectedToRAG    = d.ConnectedToRAG,"
   tmp = tmp & " @mEmployeeEmailID   = d.EmployeeEmailID,"
   tmp = tmp & " @mMobileNo          = d.MobileNo,"
   tmp = tmp & " @mAddress           = d.Address,"
   tmp = tmp & " @mFlag              = d.Flag,"
   tmp = tmp & " @mNAVCode           = d.NAVCode,"
   tmp = tmp & " @mAllocateExpense   = d.AllocateExpense,"
   tmp = tmp & " @mCMPAllocation     = d.CMPAllocation,"
   tmp = tmp & " @mBaseLocationID    = d.BaseLocationID,"
   tmp = tmp & " @mFAGUpdated        = d.FAGUpdated,"
   tmp = tmp & " @mFAGUpdatedDate    = d.FAGUpdatedDate,"
   tmp = tmp & " @mUpdatedOn         = GetDate(),"
   tmp = tmp & " @mUpdatedBy         = isnull(d.UpdatedBy,d.CreatedBy)"
   tmp = tmp & " FROM  deleted d"
   tmp = tmp & " insert into Log_Mst_Employee"
   tmp = tmp & "a (EmployeeID, EmployeeNo, EmployeeName, DesignationID, SelfReporting, DirectReportingTo, ReportingToDRC, ConnectedToRAG, EmployeeEmailID, MobileNo, Address, Flag, NAVCode, AllocateExpense, CMPAllocation, BaseLocationID, FAGUpdated, FAGUpdatedDate, UpdatedOn, UpdatedBy)"
   tmp = tmp & " values (@mEmployeeID, @mEmployeeNo, @mEmployeeName, @mDesignationID, @mSelfReporting, @mDirectReportingTo, @mReportingToDRC, @mConnectedToRAG, @mEmployeeEmailID, @mMobileNo, @mAddress, @mFlag, @mNAVCode, @mAllocateExpense, @mCMPAllocation, @mBaseLocationID, @mFAGUpdated, @mFAGUpdatedDate, @mUpdatedOn, @mUpdatedBy)"
   
   Call TmpTable
   
   tmp = " CREATE TRIGGER tr_Mst_EMPCMPMapping ON Mst_EMPCMPMapping"
   tmp = tmp & " For Update, Delete"
   tmp = tmp & " AS DECLARE"
   tmp = tmp & " @mEmpCMPMapID       Decimal(18, 0),"
   tmp = tmp & " @mLocationID        Decimal(18, 0),"
   tmp = tmp & " @mCMPID             Decimal(18, 0),"
   tmp = tmp & " @mEmployeeID        Decimal(18, 0),"
   tmp = tmp & " @mDesignationID     Decimal(18, 0),"
   tmp = tmp & " @mSM_Prefix         Varchar(5),"
   tmp = tmp & " @mFlag              Varchar(1),"
   tmp = tmp & " @mUpdatedOn         DateTime,"
   tmp = tmp & " @mUpdatedBy         Varchar(30)"
   tmp = tmp & " SELECT"
   tmp = tmp & " @mEmpCMPMapID     = d.EmpCMPMapID,"
   tmp = tmp & " @mLocationID      = d.LocationID,"
   tmp = tmp & " @mCMPID           = d.CMPID,"
   tmp = tmp & " @mEmployeeID      = d.EmployeeID,"
   tmp = tmp & " @mDesignationID   = d.DesignationID,"
   tmp = tmp & " @mSM_Prefix   = d.SM_Prefix,"
   tmp = tmp & " @mFlag            = d.Flag,"
   tmp = tmp & " @mUpdatedOn       = GetDate(),"
   tmp = tmp & " @mUpdatedBy       = isnull(d.UpdatedBy,d.CreatedBy)"
   tmp = tmp & " FROM  deleted d"
   tmp = tmp & " insert into Log_Mst_EMPCMPMapping"
   tmp = tmp & " (EmpCMPMapID, LocationID, CMPID, EmployeeID, DesignationID, SM_Prefix, Flag, UpdatedOn, UpdatedBy)"
   tmp = tmp & " values (@mEmpCMPMapID, @mLocationID, @mCMPID, @mEmployeeID, @mDesignationID, @mSM_Prefix, @mFlag, @mUpdatedOn, @mUpdatedBy)"
   
   Call TmpTable

End If

tmp = " Alter table Txn_FranchiseeRent Add LocationID decimal(18,0),FranchInvoiceNo varchar(25),FranchInvoiceDate Datetime,ServiceTaxID decimal(18,0)"
Call TmpTable

tmp = " Alter table Txn_InvMonGSLDetail Add OldFranchRate decimal(18,2),OldFranchBillingCycleID decimal(18,0),OldFranchBillingUnit varchar(1)"
Call TmpTable

tmp = " Alter Table Txn_NewCashReceipt Add " & C & " ChequeReturnChargesAmount Decimal(18,2)"
Call TmpTable

tmp = " Alter Table Mst_Client Add " & C & " ChequeReturn " & TextType & "(1)"
Call TmpTable

tmp = " Alter Table Mst_Depository Add " & C & " ChequeReturn " & TextType & "(1)"
Call TmpTable

tmp = " Update Mst_Client set  ChequeReturn = 'N'"
Call TmpTable

tmp = " Update Mst_Depository set  ChequeReturn = 'N'"
Call TmpTable

tmp = " Alter Table Mst_Agent Alter Column Address " & TextType & "(2000)"
Call TmpTable

tmp = " Alter Table Mst_Agent Alter Column ContactNo " & TextType & "(25)"
Call TmpTable

tmp = " Alter Table Mst_Agent Alter Column MobileNo " & TextType & "(25)"
Call TmpTable


tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ISInWiseDPCharges.frm','ISIN wise DP Charges','FrmRpt_ISInWiseDPCharges','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_FranchiseeRateCard.frm','Franchisee Rate Card','FrmRpt_FranchiseeRateCard','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_LocationWiseDPIncome.frm','Location Wise DP Income','FrmRpt_LocationWiseDPIncome','Y','R',0,229,1"
Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_UpdateClientChequeReturn.frm','Update Client Cheque Return','FrmTxn_UpdateClientChequeReturn','Y','T',0,228,1"
Call TmpTable

UpdateConfig ("157")
MsgBox "Version 157 Installed successfully"

End Function

Public Function Version156()

Dim TextType, C As Variant
If Gen_DBType = "MDB" Then
   TextType = "Text"
   C = " COLUMN "
Else
   TextType = "Varchar"
   C = " "
End If


tmp = " Create Table Log_ContractGSL"
tmp = tmp & " ("
tmp = tmp & " GSLID       Decimal(18,0),"
tmp = tmp & " SM_Prefix   " & TextType & "(5),"
tmp = tmp & " ContractID  Decimal(18,0),"
tmp = tmp & " UpdatedOn   DateTime,"
tmp = tmp & " UpdatedBy   " & TextType & "(30)"
tmp = tmp & " )"

Call TmpTable

tmp = " Create Table Log_Txn_InvMonGSLDetail"
tmp = tmp & " ("
tmp = tmp & " GSLID         Decimal(18,0),"
tmp = tmp & " SM_Prefix     " & TextType & "(5),"
tmp = tmp & " YearMonth     DateTime,"
tmp = tmp & " PINo          Decimal(18,0),"
tmp = tmp & " InvoiceDate   DateTime,"
tmp = tmp & " UpdatedBy     " & TextType & "(30),"
tmp = tmp & " UpdatedOn     DateTime"
tmp = tmp & " )"

Call TmpTable

tmp = " Create Table Log_Txn_ProInvoice"
tmp = tmp & " ("
tmp = tmp & " InvoiceNo         Decimal(18,0),"
tmp = tmp & " SM_Prefix         " & TextType & "(5),"
tmp = tmp & " LocationID        Decimal(18,0),"
tmp = tmp & " ExchangeTypeID    Decimal(18,0),"
tmp = tmp & " ClientIDRow       Decimal(18,0),"
tmp = tmp & " NoOfBags          Decimal(18,2),"
tmp = tmp & " Weight            Decimal(18,2),"
tmp = tmp & " Rate              Decimal(18,2),"
tmp = tmp & " Amount            Decimal(18,2),"
tmp = tmp & " BatchID           Decimal(18,0),"
tmp = tmp & " ServiceTaxNAVCode " & TextType & "(30),"
tmp = tmp & " FAGPostingDate    DateTime,"
tmp = tmp & " G_UID             " & TextType & "(36),"
tmp = tmp & " TFlag             " & TextType & "(1),"
tmp = tmp & " InvoiceDate       DateTime,"
tmp = tmp & " BillingMonth      DateTime,"
tmp = tmp & " Discount          Decimal(18,2),"
tmp = tmp & " Premium           Decimal(18,2),"
tmp = tmp & " Remarks           " & TextType & "(2000),"
tmp = tmp & " Flag              " & TextType & "(10),"
tmp = tmp & " ReverseDate       DateTime,"
tmp = tmp & " ReverseFlag       " & TextType & "(2),"
tmp = tmp & " ReverseRemark     " & TextType & "(255),"
tmp = tmp & " RBatchID          Decimal(18,0),"
tmp = tmp & " UpdatedBy         " & TextType & "(30),"
tmp = tmp & " UpdatedOn         DateTime"
tmp = tmp & " )"

Call TmpTable

If Gen_DBType <> "MDB" Then
    tmp = " CREATE TRIGGER tr_Txn_ProInvoice ON Txn_ProInvoice"
    tmp = tmp & " For Delete"
    tmp = tmp & " AS DECLARE"
    tmp = tmp & " @mInvoiceNo           decimal(18,0),"
    tmp = tmp & " @mSM_Prefix           " & TextType & "(5),"
    tmp = tmp & " @mLocationID          decimal(18,0),"
    tmp = tmp & " @mExchangeTypeID      decimal(18,0),"
    tmp = tmp & " @mClientIDRow         decimal(18,0),"
    tmp = tmp & " @mNoOfBags            decimal(18,2),"
    tmp = tmp & " @mWeight              decimal(18,2),"
    tmp = tmp & " @mRate                decimal(18,2),"
    tmp = tmp & " @mAmount              decimal(18,2),"
    tmp = tmp & " @mBatchID             decimal(18,0),"
    tmp = tmp & " @mServiceTaxNAVCode   " & TextType & "(30),"
    tmp = tmp & " @mFAGPostingDate      datetime,"
    tmp = tmp & " @mG_UID               " & TextType & "(36),"
    tmp = tmp & " @mTFlag               " & TextType & "(1),"
    tmp = tmp & " @mInvoiceDate         datetime,"
    tmp = tmp & " @mBillingMonth        datetime,"
    tmp = tmp & " @mDiscount            decimal(18,2),"
    tmp = tmp & " @mPremium             decimal(18,2),"
    tmp = tmp & " @mRemarks             " & TextType & "(2000),"
    tmp = tmp & " @mFlag                " & TextType & "(10),"
    tmp = tmp & " @mReverseDate         datetime,"
    tmp = tmp & " @mReverseFlag         " & TextType & "(2),"
    tmp = tmp & " @mReverseRemark       " & TextType & "(255),"
    tmp = tmp & " @mRBatchID            decimal(18,0),"
    tmp = tmp & " @mUpdatedOn           DateTime,"
    tmp = tmp & " @mUpdatedBy           " & TextType & "(30)"
    tmp = tmp & " SELECT"
    tmp = tmp & " @mInvoiceNo           = d.InvoiceNo,"
    tmp = tmp & " @mSM_Prefix           = d.SM_Prefix,"
    tmp = tmp & " @mLocationID          = d.LocationID,"
    tmp = tmp & " @mExchangeTypeID      = d.ExchangeTypeID,"
    tmp = tmp & " @mClientIDRow         = d.ClientIDRow,"
    tmp = tmp & " @mNoOfBags            = d.NoOfBags,"
    tmp = tmp & " @mWeight              = d.Weight,"
    tmp = tmp & " @mRate                = d.Rate,"
    tmp = tmp & " @mAmount              = d.Amount,"
    tmp = tmp & " @mBatchID             = d.BatchID,"
    tmp = tmp & " @mServiceTaxNAVCode   = d.ServiceTaxNAVCode,"
    tmp = tmp & " @mFAGPostingDate      = d.FAGPostingDate,"
    tmp = tmp & " @mG_UID               = d.G_UID,"
    tmp = tmp & " @mTFlag               = d.TFlag,"
    tmp = tmp & " @mInvoiceDate         = d.InvoiceDate,"
    tmp = tmp & " @mBillingMonth        = d.BillingMonth,"
    tmp = tmp & " @mDiscount            = d.Discount,"
    tmp = tmp & " @mPremium             = d.Premium,"
    tmp = tmp & " @mRemarks             = d.Remarks,"
    tmp = tmp & " @mFlag                = d.Flag,"
    tmp = tmp & " @mReverseDate         = d.ReverseDate,"
    tmp = tmp & " @mReverseFlag         = d.ReverseFlag,"
    tmp = tmp & " @mReverseRemark       = d.ReverseRemark,"
    tmp = tmp & " @mRBatchID            = d.RBatchID,"
    tmp = tmp & " @mUpdatedOn           = GetDate(),"
    tmp = tmp & " @mUpdatedBy           = d.UpdatedBy"
    tmp = tmp & " FROM  deleted d"
    tmp = tmp & " insert into Log_Txn_ProInvoice"
    tmp = tmp & " (InvoiceNo, SM_Prefix, LocationID, ExchangeTypeID, ClientIDRow, NoOfBags, Weight, Rate, Amount, BatchID, ServiceTaxNAVCode, FAGPostingDate, G_UID, TFlag, InvoiceDate, BillingMonth, Discount, Premium, Remarks, Flag, ReverseDate, ReverseFlag, ReverseRemark, RBatchID, UpdatedOn, UpdatedBy)"
    tmp = tmp & " Values"
    tmp = tmp & " (@mInvoiceNo, @mSM_Prefix, @mLocationID, @mExchangeTypeID, @mClientIDRow, @mNoOfBags, @mWeight, @mRate, @mAmount, @mBatchID, @mServiceTaxNAVCode, @mFAGPostingDate, @mG_UID, @mTFlag, @mInvoiceDate, @mBillingMonth, @mDiscount, @mPremium, @mRemarks, @mFlag, @mReverseDate, @mReverseFlag, @mReverseRemark, @mRBatchID, @mUpdatedOn, @mUpdatedBy)"
    
    Call TmpTable
    
    tmp = " CREATE TRIGGER tr_TxN_InvMonGSLDetail ON TxN_InvMonGSLDetail"
    tmp = tmp & " For Update"
    tmp = tmp & " AS DECLARE"
    tmp = tmp & " @mGSLID       Decimal(18,0),"
    tmp = tmp & " @mSM_Prefix   Varchar(5),"
    tmp = tmp & " @mYearMonth   DateTime,"
    tmp = tmp & " @mPINo        Decimal(18,0),"
    tmp = tmp & " @mInvoiceDate DateTime,"
    tmp = tmp & " @mUpdatedOn   DateTime,"
    tmp = tmp & " @mUpdatedBy   Varchar(30)"
    tmp = tmp & " SELECT"
    tmp = tmp & " @mGSLID       = d.GSLID,"
    tmp = tmp & " @mSM_Prefix   = d.SM_Prefix,"
    tmp = tmp & " @mYearMonth   = d.YearMonth,"
    tmp = tmp & " @mPINo        = d.PINo,"
    tmp = tmp & " @mInvoiceDate = d.InvoiceDate,"
    tmp = tmp & " @mUpdatedOn   = GetDate(),"
    tmp = tmp & " @mUpdatedBy   = d.UpdatedBy"
    tmp = tmp & " FROM  deleted d"
    tmp = tmp & " insert into Log_TxN_InvMonGSLDetail"
    tmp = tmp & " (GSLID, SM_Prefix, YearMonth, PINo, InvoiceDate, UpdatedOn, UpdatedBy)"
    tmp = tmp & " Values"
    tmp = tmp & " (@mGSLID, @mSM_Prefix, @mYearMonth, @mPINo, @mInvoiceDate, @mUpdatedOn, @mUpdatedBy)"
    
    Call TmpTable

End If



tmp = " Alter Table Mst_Godown Add " & C & " BalanceBags Decimal(18,0),  " & C & " BalanceQty Decimal(18,3), "
tmp = tmp & " " & C & " NoOfTransactions Decimal(18,0), " & C & " NoOfBags Decimal(18,0), " & C & " NoOfQty Decimal(18,3), "
tmp = tmp & " " & C & " YearMonth DateTime"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_UpdateGodownBalanceStock.frm','Update Godown Balance Stock','FrmTxn_UpdateGodownBalanceStock','Y','T',0,228,1"

Call TmpTable

tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportReturnPledgeCashReceipt.frm','Export Return Pledge CashReceipt','Frm_ExportReturnPledgeCashReceipt','Y','C',0,230,1"

Call TmpTable


UpdateConfig ("156")
MsgBox "Version 156 Installed successfully"

End Function

'Public Function Version155()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Create Table Mst_EarthquakeZone("
'tmp = tmp & " EQZoneID Decimal(18,0) Not null,"
'tmp = tmp & " EQZone " & TextType & "(250),"
'tmp = tmp & " Flag " & TextType & "(2),"
'tmp = tmp & " G_UID   " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
''tmp = tmp & " UpdatedBy   " & TextType & "(30))"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy   " & TextType & "(30),"
'   tmp = " PRIMARY KEY (EQZoneID))"
'Else
'   tmp = tmp & " UpdatedBy   " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'tmp = "Alter Table Mst_Location Add EQZoneID decimal(18,0)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Mst_EarthquakeZone WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_EarthquakeZone PRIMARY KEY  CLUSTERED (EQZoneID),"
'   tmp = tmp & " CONSTRAINT DF_Mst_EarthquakeZone_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_Location WITH NOCHECK ADD"
'   tmp = tmp & " Constraint FK_Mst_Location_Mst_EarthquakeZone Foreign KEY (EQZoneID) references Mst_EarthquakeZone(EQZoneID)"
'
'   Call TmpTable
'
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_EarthquakeZone.frm','Earthquake Zone','FrmMst_EarthquakeZone','Y','M',0,227,1"
'
'Call TmpTable
'
'UpdateConfig ("155")
'MsgBox "Version 155 Installed successfully"
'
'End Function
'
'Public Function Version154()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_InvMonGSLDetail Add " & c & " OldBalanceWeight decimal(18,3) "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CRInvoices.frm','Cash Receipt Invoices','FrmRpt_CRInvoices','Y','R',0,229,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportPledgeCashReceipt.frm','Export Pledge Cash Receipt','Frm_ExportPledgeCashReceipt','Y','C',0,230,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_LocationOrGodownwiseStockSummary.frm','Location Or Godown Wise Stock Summary','FrmRpt_LocationOrGodownwiseStockSummary','Y','R',0,229,1"
'
'Call TmpTable
'
'UpdateConfig ("154")
'MsgBox "Version 154 Installed successfully"
'
'End Function
'
'Public Function Version153()
'Dim TextType, c As Variant
'
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Mst_NCDEXAccreditedCapacity"
'tmp = tmp & " ("
'tmp = tmp & " AccreditedID          decimal(18,0) not null,"
'tmp = tmp & " AccreditedDate        Datetime,"
'tmp = tmp & " StartDate             Datetime,"
'tmp = tmp & " CloseDate             Datetime,"
'tmp = tmp & " AccreditedDescription " & TextType & "(250),"
'tmp = tmp & " LocationID            Decimal(18,0),"
'tmp = tmp & " Capacity              Decimal(18,3),"
'tmp = tmp & " Status                " & TextType & "(1),"
'tmp = tmp & " Remarks               " & TextType & "(2000),"
'tmp = tmp & " G_UID                 " & TextType & "(36),"
'tmp = tmp & " Concurrency           datetime,"
'tmp = tmp & " CreatedBy             " & TextType & "(30),"
'tmp = tmp & " CreatedDate           datetime,"
'tmp = tmp & " UpdatedBy             " & TextType & "(30),"
'tmp = tmp & " UpdatedDate           DateTime"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " PRIMARY KEY (AccreditedID)"
'End If
'
'tmp = tmp & " )"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_NCDEXAccreditedCapacity WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_NCDEXAccreditedCapacity PRIMARY KEY  CLUSTERED (AccreditedID),"
'   tmp = tmp & " CONSTRAINT DF_Mst_NCDEXAccreditedCapacity_Concurrency DEFAULT (getdate()) FOR Concurrency ,"
'   tmp = tmp & " Constraint FK_Mst_NCDEXAccreditedCapacity_Mst_Location Foreign KEY (LocationID) references Mst_Location(LocationID)"
'
'   Call TmpTable
'End If
'
'
'tmp = " ALTER TABLE Mst_Godown Add " & c & " AccreditedID decimal(18,0)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_Godown WITH NOCHECK ADD"
'   tmp = tmp & " Constraint FK_Mst_Godown_Mst_NCDEXAccreditedCapacity Foreign KEY (AccreditedID) references Mst_NCDEXAccreditedCapacity(AccreditedID)"
'
'   Call TmpTable
'End If
'
'tmp = " Create Table Log_Mst_NCDEXAccreditedCapacity"
'tmp = tmp & " ("
'tmp = tmp & " AccreditedID Decimal(18,0),"
'tmp = tmp & " GodownID Decimal(18,0),"
'tmp = tmp & " Concurrency Datetime,"
'tmp = tmp & " CreatedDate Datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " UpdatedDate Datetime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30)"
'tmp = tmp & " )"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Log_Mst_NCDEXAccreditedCapacity WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Log_Mst_NCDEXAccreditedCapacity_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmMst_NCDEXAccreditedCapacity.frm','NCDEX Accredited Capacity','FrmMst_NCDEXAccreditedCapacity','Y','M',0,227,1"
'
'Call TmpTable
'
'UpdateConfig ("153")
'MsgBox "Version 153 Installed successfully"
'
'End Function
'
'
'Public Function Version152()
'
'Dim TextType, c As Variant
'
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Txn_FranchiseeRent"
'tmp = tmp & " (ClaimID decimal(18,0) Not Null,"
'tmp = tmp & " ClaimDate DateTime,"
'tmp = tmp & " ExpenseFrom DateTime,"
'tmp = tmp & " ExpenseTo DateTime,"
'tmp = tmp & " VendorID decimal(18,0),"
'tmp = tmp & " Amount decimal(18,2),"
'tmp = tmp & " ActualFee decimal(18,2),"
'tmp = tmp & " Remarks " & TextType & "(200),"
'tmp = tmp & " ApprovalRemarks " & TextType & "(200),"
'tmp = tmp & " ApprovalDate Datetime,"
'tmp = tmp & " BatchID decimal(18,0),"
'tmp = tmp & " ForwardToID decimal(18,0),"
'tmp = tmp & " Flag " & TextType & "(2),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " Concurrency Datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate Datetime,"
'tmp = tmp & " UpdatedDate Datetime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " UpdatedByFAG " & TextType & "(30),"
'tmp = tmp & " UpdatedDateFAG Datetime,"
'tmp = tmp & " BankName " & TextType & "(100),"
'tmp = tmp & " ChequeNo " & TextType & "(100),"
'tmp = tmp & " PaidAmount Decimal(18,0),"
'tmp = tmp & " TDS Decimal(18,0),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " PaidDate Datetime,"
'   tmp = " PRIMARY KEY (ClaimID))"
'Else
'   tmp = tmp & " PaidDate Datetime)"
'End If
'
'
'Call TmpTable
'
'tmp = " Alter table Txn_InvMonGSLDetail Add"
'tmp = tmp & " ISINID " & TextType & "(25),"
'tmp = tmp & " CRDate DateTime,"
'tmp = tmp & " FranchBillingCycleID Decimal(18,0),"
'tmp = tmp & " FranchBillingUnit " & TextType & "(1),"
'tmp = tmp & " FranchNoofDays Decimal(18,0),"
'tmp = tmp & " FranchInvoiceAmount Decimal(18,2),"
'tmp = tmp & " FF_RateCardID Decimal(18,0),"
'tmp = tmp & " FranchClaimID decimal(18,0),"
'tmp = tmp & " FranchProcessFlag " & TextType & "(2)"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Txn_FranchiseeRent WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_FranchiseeRent PRIMARY KEY  CLUSTERED (ClaimID)  ON [PRIMARY],"
'   tmp = tmp & " CONSTRAINT DF_Txn_FranchiseeRent_Concurrency DEFAULT (getdate()) FOR [Concurrency],"
'   tmp = tmp & " CONSTRAINT FK_Txn_PledgeCashReceipt_Mst_Lessor FOREIGN KEY (VendorID) REFERENCES Mst_Lessor (LessorID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_PledgeCashReceipt_Mst_Employee FOREIGN KEY (ForwardToID) REFERENCES Mst_Employee (EmployeeID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_InvMonGSLDetail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_InvMonGSLDetail_Mst_FBillingCycle FOREIGN KEY (FranchBillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_InvMonGSLDetail_Mst_FF_RateCard FOREIGN KEY (FF_RateCardID) REFERENCES Mst_FF_RateCard (FF_RateCardID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_InvMonGSLDetail_Txn_FranchiseeRent FOREIGN KEY (FranchClaimID) REFERENCES Txn_FranchiseeRent (ClaimID)"
'
'   Call TmpTable
'
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_DPMonthlyGSLProcess.frm','Franchisee Fees Stock Op Balance','FrmTxn_DPMonthlyGSLProcess','Y','T',0,228,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_FranchiseeRent.frm','Franchisee Fees Process','FrmTxn_FranchiseeRent','Y','T',0,228,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_FranchiseeRentApproval.frm','Franchisee Fees Approval','FrmTxn_FranchiseeRentApproval','Y','T',0,228,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportFranchiseeFees.frm','Export Franchisee Fees','Frm_ExportFranchiseeFees','Y','C',0,230,1"
'
'Call TmpTable
'
'UpdateConfig ("152")
'MsgBox "Version 152 Installed successfully"
'
'End Function
'
'
'Public Function Version151()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Create Table Txn_PledgeCashReceipt("
'tmp = tmp & " PCRID decimal(18,0) Not Null,"
'tmp = tmp & " SM_Prefix " & TextType & "(5) Not Null ,"
'tmp = tmp & " PCRDate datetime,"
'tmp = tmp & " LocationID decimal (18,0),"
'tmp = tmp & " InstrumentType  " & TextType & "(50),"
'tmp = tmp & " DrawnOn " & TextType & "(150),"
'tmp = tmp & " BranchName  " & TextType & "(150),"
'tmp = tmp & " InstrumentNo    " & TextType & "(25),"
'tmp = tmp & " InstrumentDate  datetime,"
'tmp = tmp & " InstrumentAmount    decimal(18,0),"
'tmp = tmp & " ReceivedFrom    " & TextType & "(100),"
'tmp = tmp & " ExcessAmount Decimal(18,2),"
'tmp = tmp & " Remark  " & TextType & "(255),"
'tmp = tmp & " AdvanceFlag " & TextType & "(2),"
'tmp = tmp & " Flag    " & TextType & "(2),"
'tmp = tmp & " HologramNo  " & TextType & "(25),"
'tmp = tmp & " G_UID   " & TextType & "(36),"
'tmp = tmp & " SightedDate datetime,"
'tmp = tmp & " ReturnFlag  " & TextType & "(2),"
'tmp = tmp & " BatchID decimal(18,0),"
'tmp = tmp & " DocNo " & TextType & "(30),"
'tmp = tmp & " EntryNo Decimal(18,0),"
'tmp = tmp & " ReturnDate  datetime,"
'tmp = tmp & " RBatchID    decimal(18,0),"
'tmp = tmp & " ReturnRemark    " & TextType & "(100),"
'tmp = tmp & " FAGBankID   decimal(18,0),"
'tmp = tmp & " FAGUpdatedBy    " & TextType & "(30),"
'tmp = tmp & " FAGUpdatedDate  datetime,"
'tmp = tmp & " DepositDate datetime,"
'tmp = tmp & " RejectFlag  " & TextType & "(1),"
'tmp = tmp & " RejectDate  datetime,"
'tmp = tmp & " RejectRemark    " & TextType & "(100),"
'tmp = tmp & " RejectedBy  " & TextType & "(30),"
'tmp = tmp & " FAGRemarks  " & TextType & "(255),"
'tmp = tmp & " FAGPostingDate  datetime,"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy   " & TextType & "(30),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate Datetime,"
'   tmp = " PRIMARY KEY (PCRID,SM_Prefix))"
'Else
'   tmp = tmp & " UpdatedDate Datetime)"
'End If
'
'Call TmpTable
'
'tmp = " Create Table Txn_PledgeCashReceiptDetail("
'tmp = tmp & " PCRDID  decimal(18,0) Not Null,"
'tmp = tmp & " PCRID   decimal(18,0),"
'tmp = tmp & " SM_Prefix   " & TextType & "(5),"
'tmp = tmp & " PledgeID    decimal(18,0),"
'tmp = tmp & " RecAmount   decimal(18,2),"
'tmp = tmp & " Flag    " & TextType & "(2),"
'tmp = tmp & " G_UID   " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy   " & TextType & "(30),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate Datetime,"
'   tmp = " PRIMARY KEY (PCRDID))"
'Else
'   tmp = tmp & " UpdatedDate Datetime)"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'    tmp = "ALTER TABLE Txn_PledgeCashReceipt WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_PledgeCashReceipt PRIMARY KEY  CLUSTERED (PCRID,SM_Prefix)  ON [PRIMARY],"
'    tmp = tmp & " CONSTRAINT DF_Txn_PledgeCashReceipt_Concurrency DEFAULT (getdate()) FOR [Concurrency],"
'    tmp = tmp & " CONSTRAINT FK_Txn_PledgeCashReceipt_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
'    tmp = tmp & " CONSTRAINT FK_Txn_PledgeCashReceipt_Mst_FAGBank FOREIGN KEY (FAGBankId) REFERENCES Mst_FAGBank (FAGBankId)"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_PledgeCashReceiptDetail WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_PledgeCashReceiptDetail PRIMARY KEY  CLUSTERED (PCRDID)  ON [PRIMARY],"
'    tmp = tmp & " CONSTRAINT DF_Txn_PledgeCashReceiptDetail_Concurrency DEFAULT (getdate()) FOR [Concurrency],"
'    tmp = tmp & " CONSTRAINT FK_Txn_PledgeCashReceiptDetail_Txn_PledgeCreation FOREIGN KEY (PledgeID,SM_Prefix) REFERENCES Txn_PledgeCreation (PledgeID,SM_Prefix)"
'
'    Call TmpTable
'
'End If
'
'tmp = "Alter table Txn_PledgeCreation Add " & c & " Discount Decimal(18,2)," & c & " RecFees Decimal(18,2)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_PledgeCashReceipt.frm','Pledge Cash Receipt','FrmTxn_PledgeCashReceipt','Y','T',0,228,1"
'
'Call TmpTable
'
'UpdateConfig ("151")
'MsgBox "Version 151 Installed successfully"
'
'End Function
'
'Public Function Version150()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Log_FormAccessDetail"
'tmp = tmp & " (LogID Decimal(18,0),"
'tmp = tmp & " UserID Decimal(10,0),"
'tmp = tmp & " MenuID Decimal(18,0),"
'tmp = tmp & " Concurrency DateTime)"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Log_FormAccessDetail WITH NOCHECK ADD"
'   tmp = tmp & "  CONSTRAINT DF_Log_FormAccessDetail_Concurrency DEFAULT (getdate()) FOR Concurrency "
'   Call TmpTable
'
'   tmp = " ALTER TABLE Log_FormAccessDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Log_FormAccessDetail_Mst_Users FOREIGN KEY (UserID) REFERENCES Mst_Users(UsersID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Log_FormAccessDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Log_FormAccessDetail_Mst_MenuDetail FOREIGN KEY (MenuID) REFERENCES Mst_MenuDetail(MenuID)"
'
'   Call TmpTable
'
'End If
'
'
'UpdateConfig ("150")
'MsgBox "Version 150 Installed successfully"
'
'End Function
'
'Public Function Version149()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Txn_NAVInvoiceOutStanding( "
'tmp = tmp & " CustomerNo  " & TextType & " (10), "
'tmp = tmp & " CustomerName    " & TextType & " (500), "
'tmp = tmp & " EntryNo decimal(18,0), "
'tmp = tmp & " EHOPInvoiceNo   " & TextType & "(50), "
'tmp = tmp & " DocumentType    " & TextType & "(50), "
'tmp = tmp & " DocumentNo  " & TextType & "(50), "
'tmp = tmp & " VoucherNo   " & TextType & "(50), "
'tmp = tmp & " PostingDate datetime, "
'tmp = tmp & " InvoiceAmount   decimal (18,2), "
'tmp = tmp & " BaseAmount  decimal (18,2), "
'tmp = tmp & " TaxAmount   decimal (18,2), "
'tmp = tmp & " RemainingAmount decimal (18,2), "
'tmp = tmp & " Narration   " & TextType & " (2000), "
'tmp = tmp & " ActualDatePosted datetime, "
'tmp = tmp & " ActualTimePosted " & TextType & "(20), "
'tmp = tmp & " RegionCode  " & TextType & "(30), "
'tmp = tmp & " DepartmentCode  " & TextType & "(30), "
'tmp = tmp & " CustomerPostingGroup " & TextType & " (20), "
'tmp = tmp & " CreatedBy   " & TextType & " (30), "
'tmp = tmp & " CreatedDate datetime, "
'tmp = tmp & " Concurrency datetime, "
'tmp = tmp & " G_UID   " & TextType & "(36)) "
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ImportNAVInvoiceOutstand.frm','Import NAV Outstanding Invoices','Frm_ImportNAVInvoiceOutstand','Y','C',0,230,1"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NAVOutstandInvoices.frm','NAV Outstanding Invoices','FrmRpt_NAVOutstandInvoices','Y','R',0,229,1"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_WHRvsActualStock.frm','WHR vs Actual Stock','FrmRpt_WHRvsActualStock','Y','R',0,229,1"
'Call TmpTable
'
'UpdateConfig ("149")
'MsgBox "Version 149 Installed successfully"
'
'End Function
'
'
'Public Function Version148()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Create Table Txn_CDTFBOE("
'tmp = tmp & " BOEID   decimal(18,0) not Null,"
'tmp = tmp & " LocationID  decimal(18,0),"
'tmp = tmp & " SM_Prefix " & TextType & "(5),"
'tmp = tmp & " CDTFNo  decimal(18,0),"
'tmp = tmp & " CDTFDate    datetime,"
'tmp = tmp & " ClientIDRow decimal(18,0),"
'tmp = tmp & " InvoiceDate datetime,"
'tmp = tmp & " BillingMonth    datetime,"
'tmp = tmp & " Remarks " & TextType & "(1000),"
'tmp = tmp & " Flag " & TextType & "(2),"
'tmp = tmp & " Discount    decimal(18,2),"
'tmp = tmp & " Premium decimal(18,2),"
'tmp = tmp & " Amount  decimal(18,2),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " BatchID decimal(18,0),"
'tmp = tmp & " FAGPostingDate  datetime,"
'tmp = tmp & " FAGRemarks  " & TextType & "(2000),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy  " & TextType & "(30),"
'tmp = tmp & " ReverseDate DateTime,"
'tmp = tmp & " ReverseFlag " & TextType & "(2),"
'tmp = tmp & " ReverseRemark " & TextType & "(255),"
'tmp = tmp & " RBatchID decimal(18, 0),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate Datetime,"
'   tmp = " PRIMARY KEY (BOEID))"
'Else
'   tmp = tmp & " UpdatedDate Datetime)"
'End If
'
'Call TmpTable
'
'tmp = "Create table Txn_CDTFBOE_GSL("
'tmp = tmp & " BOEGSLID decimal(18,0) not null,"
'tmp = tmp & " BOEID   decimal(18,0),"
'tmp = tmp & " SM_Prefix   " & TextType & "(5),"
'tmp = tmp & " GSLID   decimal(18,0),"
'tmp = tmp & " CDTFBags    decimal(18,0),"
'tmp = tmp & " CDTFQty decimal(18,3),"
'tmp = tmp & " NoOfDays    decimal(18,0),"
'tmp = tmp & " Rate    decimal (18,2),"
'tmp = tmp & " BillingUnit " & TextType & " (2),"
'tmp = tmp & " BillingCycleID  decimal (18,0),"
'tmp = tmp & " Amount  decimal(18,2),"
'tmp = tmp & " G_UID   " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy   " & TextType & "(30),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate Datetime,"
'   tmp = " PRIMARY KEY (BOEGSLID))"
'Else
'   tmp = tmp & " UpdatedDate Datetime)"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_CDTFBOE WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_CDTFBOE PRIMARY KEY  CLUSTERED (BOEID)  ON [PRIMARY], "
'   tmp = tmp & " CONSTRAINT DF_Txn_CDTFBOE_Concurrency DEFAULT (getdate()) FOR [Concurrency]"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_CDTFBOE_GSL WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_CDTFBOE_GSL PRIMARY KEY  CLUSTERED (BOEGSLID)  ON [PRIMARY], "
'   tmp = tmp & " CONSTRAINT DF_Txn_CDTFBOE_GSL_Concurrency DEFAULT (getdate()) FOR [Concurrency],"
'   tmp = tmp & " CONSTRAINT FK_Txn_CDTFBOE_GSL_Txn_CDTFBOE FOREIGN KEY "
'   tmp = tmp & " (BOEID) REFERENCES Txn_CDTFBOE (BOEID) "
'
'   Call TmpTable
'
'End If
'
'tmp = " Alter table Mst_MenuDetail  Add FormDesc " & TextType & "(2000),DetailDesc " & TextType & "(5000) "
'
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_7DaysInvoice.frm','7 Days Invoice','FrmTxn_7DaysInvoice','Y','T',0,228,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_Export7DaysInvoice.frm','Export 7 Days Invoice','Frm_Export7DaysInvoice','Y','C',0,230,1"
'
'Call TmpTable
'
'UpdateConfig ("148")
'MsgBox "Version 148 Installed successfully"
'
'End Function
'
'Public Function Version147()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Alter Table Txn_TIDetails Add AdjAmountFAG " & c & " Decimal(18,2)," & c & " Remarks " & TextType & "(2000) "
'Call TmpTable
'
'tmp = " Update Txn_TIDetails Set AdjAmountFAG = 0 Where  AdjAmountFAG Is Null "
'Call TmpTable
'
'
'
'tmp = " Create Table Log_TIDetails ("
'tmp = tmp & " TIID Decimal(18,0),"
'tmp = tmp & " AdjAmountFAG Decimal(18,2),"
'tmp = tmp & " Remarks " & TextType & "(2000),"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " UpdatedDate DateTime)"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " CREATE TRIGGER tr_Log_TIDetails_U ON Txn_TIDetails AFTER UPDATE"
'   tmp = tmp & " AS"
'   tmp = tmp & " DECLARE @mTIID  decimal(18, 0),"
'   tmp = tmp & " @mAdjAmountFAG decimal(18, 2),"
'   tmp = tmp & " @mRemarks " & TextType & "(2000),"
'   tmp = tmp & " @mUpdatedBy  " & TextType & " (30),"
'   tmp = tmp & " @mUpdatedDate  DateTime"
'   tmp = tmp & " SELECT  @mTIID= d.TIID,"
'   tmp = tmp & " @mAdjAmountFAG = d.AdjAmountFAG,"
'   tmp = tmp & " @mRemarks = i.Remarks,"
'   tmp = tmp & " @mUpdatedBy= i.UpdatedBy,"
'   tmp = tmp & " @mUpdatedDate= i.UpdatedDate"
'   tmp = tmp & " FROM    inserted i"
'   tmp = tmp & " INNER JOIN deleted d ON i.TIID = d.TIID"
'   tmp = tmp & " insert into Log_TIDetails (TIID ,AdjAmountFAG ,Remarks ,UpdatedBy,UpdatedDate)"
'   tmp = tmp & " values(@mTIID ,@mAdjAmountFAG ,@mRemarks ,@mUpdatedBy,@mUpdatedDate)"
'
'   Call TmpTable
'
'
'   tmp = " ALTER TABLE Txn_TIDetails WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_TIDetails_AdjAmountFAG DEFAULT (0) FOR AdjAmountFAG"
'
'   Call TmpTable
'
'End If
'
'UpdateConfig ("147")
'MsgBox "Version 147 Installed successfully"
'
'End Function
'
'Public Function Version146()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
''tmp = " Alter Table Txn_PledgeCreation  add " & c & " Discount decimal(18,2)," & c & " DiscountRemarks  " & TextType & "(2000) "
''Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_NNNRInvConfirmation.frm','NNNR Invoice Confirmation','FrmTxn_NNNRInvConfirmation','Y','T',0,228,1"
'Call TmpTable
'
'UpdateConfig ("146")
'MsgBox "Version 146 Installed successfully"
'
'End Function
'
'
'Public Function Version145()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_NewCashReceipt Add " & c & " RejectFlag " & TextType & " (1), " & c & " RejectDate datetime, " & c & " RejectRemark " & TextType & "(100), " & c & " RejectedBy " & TextType & "(30)"
'Call TmpTable
'
'tmp = " Create Table Log_Txn_NewCashReceiptTIDetail"
'tmp = tmp & " (CRTIDID decimal(18, 0) ,"
'tmp = tmp & " CRID decimal(18, 0) ,"
'tmp = tmp & " SM_Prefix " & TextType & "(5) ,"
'tmp = tmp & " TIID decimal(18, 0) ,"
'tmp = tmp & " RecAmount decimal(18, 2) ,"
'tmp = tmp & " RecTDS decimal(18, 2),"
'tmp = tmp & " RecOtherDed decimal(18, 2) ,"
'tmp = tmp & " Flag " & TextType & "(2),"
'tmp = tmp & " ClientIDRow decimal(18, 0) ,"
'tmp = tmp & " DPID " & TextType & "(50),"
'tmp = tmp & " DocNo " & TextType & "(50),"
'tmp = tmp & " EntryNo decimal(18, 0) ,"
'tmp = tmp & " CRAAID decimal(18, 0) ,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " Concurrency datetime )"
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Log_Txn_NewCashReceiptTIDetail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Log_Txn_NewCashReceiptTIDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " Create TRIGGER tr_Txn_NewCashReceiptTIDetail ON Txn_NewCashReceiptTIDetail"
'   tmp = tmp & " FOR DELETE AS"
'   tmp = tmp & " insert into Log_Txn_NewCashReceiptTIDetail (CRTIDID ,CRID ,SM_Prefix, TIID, RecAmount, RecTDS, RecOtherDed,"
'   tmp = tmp & " Flag, ClientIDRow, DocNo, DPID, EntryNo, CRAAID, UpdatedBy)"
'   tmp = tmp & " Select CRTIDID ,CRID ,SM_Prefix, TIID, RecAmount, RecTDS, RecOtherDed,"
'   tmp = tmp & " Flag, ClientIDRow, DocNo, DPID, EntryNo, CRAAID, UpdatedBy  from deleted"
'
'   Call TmpTable
'
'End If
'
'UpdateConfig ("145")
'MsgBox "Version 145 Installed successfully"
'
'End Function
'Public Function Version144()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Create Table Log_WHRPrint"
'tmp = tmp & " (WHRNo decimal(18,0),"
'tmp = tmp & " PrintDate Datetime,"
'tmp = tmp & " Version " & TextType & " (25),"
'tmp = tmp & " HologramNo " & TextType & " (25),"
'tmp = tmp & " PledgeID decimal (18,0),"
'tmp = tmp & " SM_prefix " & TextType & " (5),"
'tmp = tmp & " UpdatedBy " & TextType & " (30),"
'tmp = tmp & " UpdatedDate Datetime)"
'
'Call TmpTable
'
'tmp = "Alter table Txn_WHR_Detail Add " & c & " CARNo Decimal(18,0)"
'
'Call TmpTable
'
'tmp = " Create Table Log_Txn_WHR_Detail("
'tmp = tmp & " CARNo Decimal(18,0),"
'tmp = tmp & " WHRNo Decimal(18,0),"
'tmp = tmp & " Concurrency Datetime,"
'tmp = tmp & " UpdatedOn DateTime,"
'tmp = tmp & " UpdatedBy Varchar(30))"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Log_Txn_WHR_Detail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Log_Txn_WHR_Detail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'End If
'
'
'tmp = "Alter Table Txn_PledgeCreation Add " & c & " ActualFees Decimal (18,2) "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_MonthwiseInvoices.frm','Income Category wise Monthly Income MIS','FrmRpt_MonthwiseInvoices','Y','R',0,229,1"
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_FAGMIS.frm','FAG MIS','FrmRpt_FAGMIS','Y','R',0,229,2"
'Call TmpTable
'
'
'UpdateConfig ("144")
'MsgBox "Version 144 Installed successfully"
'
'End Function
'
'
'Public Function Version143()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_Depository Add " & c & " NAVDepositoryName " & TextType & "(200)"
'Call TmpTable
'
'UpdateConfig ("143")
'MsgBox "Version 143 Installed successfully"
'
'End Function
'Public Function Version142()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_NewCashReceipt Add " & c & " InvAllocFlag " & TextType & " (1)"
'Call TmpTable
'
'tmp = "Alter Table Txn_CxTF_TruckDetail Add " & c & " NoOfBags decimal(18,0)"
'
'Call TmpTable
'
'tmp = " Alter Table Mst_Bank Add " & c & "  CMAgreement " & TextType & "(1)"
'Call TmpTable
'
'tmp = " Alter Table Mst_Location Add " & c & " FeesForCM Decimal(18,0), " & c & " FeesForWithoutCM Decimal(18,0)"
'Call TmpTable
'
'tmp = " Update Mst_Location Set FeesForCM = 500,FeesForWithoutCM = 1000 "
'Call TmpTable
'
'tmp = " Update Mst_Bank Set CMAgreement = 0 "
'Call TmpTable
'
'tmp = "Alter Table Mst_Client Add " & c & " NAVClientName " & TextType & "(200)"
'Call TmpTable
'
'UpdateConfig ("142")
'MsgBox "Version 142 Installed successfully"
'
'End Function
'
'Public Function Version141()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Mst_CommodityPrice Add " & c & " YearMonth " & TextType & "(50) "
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ClientLocationOutStandingInvoice.frm','Client Location OutStanding Invoice','FrmRpt_ClientLocationOutStandingInvoice','Y','R',0,229,1"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_MonthWiseOutStandingAmount.frm','MonthWise OutStanding Amount','FrmRpt_MonthWiseOutStandingAmount','Y','R',0,229,1"
'Call TmpTable
'
'UpdateConfig ("141")
'MsgBox "Version 141 Installed successfully"
'
'End Function
'
'Public Function Version140()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_PendingAdvanceCashReceipt.frm','Pending Advance Cash Receipt','FrmRpt_PendingAdvanceCashReceipt','Y','R',0,229,1"
'Call TmpTable
'
'
'UpdateConfig ("140")
'MsgBox "Version 140 Installed successfully"
'
'End Function
'
'Public Function Version139()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Alter table Mst_Commodity Add " & c & " MaxRate decimal(18,3)"
'Call TmpTable
'
'tmp = " Alter table Mst_Commodity Add " & c & " MinRate decimal(18,3)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_ReservationInvoice WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_ReservationInvoice_RBatchID DEFAULT (0) FOR RBatchID"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_TempInvoice WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_TempInvoice_RBatchID DEFAULT (0) FOR RBatchID"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_ProInvoice_RBatchID DEFAULT (0) FOR RBatchID"
'
'   Call TmpTable
'
'End If
'
'tmp = " Alter table Txn_CRTDSAdjustmentDetail add TDSEntryNo " & c & " Decimal(18,0), " & c & " OtherDedEntryNo Decimal(18,0),"
'tmp = tmp & " " & c & " TDSDocNo  " & TextType & "(50), " & c & " OtherDedDocNo " & TextType & " (50)"
'
'Call TmpTable
'
'
'tmp = " Alter Table Txn_CRTDSAdjustmentDetail add " & c & " STAmount decimal(18,2) "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportCRTDSAdj.frm','Export CR TDS Adjustment','Frm_ExportCRTDSAdj','Y','C',0,230,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_GodownwiseNCDEXStock.frm','Godown wise NCDEX Stock','FrmRpt_GodownwiseNCDEXStock','Y','R',0,229,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportCRAdvInvAlloc.frm','Export CR Advance Allocation','Frm_ExportCRAdvInvAlloc','Y','C',0,230,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_UpdateTDSAdjEntryNo.frm','Update TDS Adjustment EntryNo','Frm_UpdateTDSAdjEntryNo','Y','C',0,230,1"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportTDSAdjInvoice.frm','Export TDSAdjustment Invoices','Frm_ExportTDSAdjInvoice','Y','C',0,230,1"
'Call TmpTable
'
'
'
'UpdateConfig ("139")
'MsgBox "Version 139 Installed successfully"
'
'End Function
'
'
'Public Function Version138()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_NNNRBOEMonitoringMIS.frm','NNNR BOE Monitoring MIS','FrmRpt_NNNRBOEMonitoringMIS','Y','R',0,229,1"
'Call TmpTable
'
'UpdateConfig ("138")
'MsgBox "Version 138 Installed successfully"
'
'End Function
'
'Public Function Version137()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CITFInvoiceMIS.frm','CITF Invoice MIS','FrmRpt_CITFInvoiceMIS','Y','R',0,229,1"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_SecurityBillingDetail.frm','Security Billing Detail Report','FrmRpt_SecurityBillingDetail','Y','R',0,229,2"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_UnmappedCITF.frm','CITF unmapped to Cash Receipt','FrmRpt_UnmappedCITF','Y','R',0,229,3"
'Call TmpTable
'
'UpdateConfig ("137")
'MsgBox "Version 137 Installed successfully"
'
'End Function
'
'Public Function Version136()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_CxTF_Details Add " & c & " SCMFlag " & TextType & "(1) "
'Call TmpTable
'
'tmp = " Alter Table Mst_GSL Add  " & c & "  SCMFlag  " & TextType & "(1) "
'Call TmpTable
'
'
'tmp = " Update Txn_CxTF_Details Set SCMFlag = 0 Where SCMFlag is null "
'Call TmpTable
'
'tmp = " Update Mst_GSL Set SCMFlag = 0 Where SCMFlag is null "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_CxTFCreatedDiscripency.frm','CxTF Created Discripency','FrmRpt_CxTFCreatedDiscripency','Y','R',0,229,1"
'Call TmpTable
'
'
'UpdateConfig ("136")
'MsgBox "Version 136 Installed successfully"
'
'End Function
'
'Public Function Version135()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_SpotPending.frm','Spot Pending Report','FrmRpt_SpotPending','Y','R',0,229,1"
'Call TmpTable
'
'
'UpdateConfig ("135")
'MsgBox "Version 135 Installed successfully"
'
'End Function
'Public Function Version134()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter table Mst_ExpenseType Add " & c & " ReverseFlag " & TextType & "(1)"
'Call TmpTable
'
'UpdateConfig ("134")
'MsgBox "Version 134 Installed successfully"
'
'End Function
'
'Public Function Version133()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_FF_RateCard Add VendorID decimal(18,0) "
'Call TmpTable
'
'tmp = " Alter Table Mst_FF_RateCard Alter Column LessorID decimal(18,0) "
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Mst_FF_RateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Mst_Vendor FOREIGN KEY (VendorID) REFERENCES Mst_Vendor(VendorID)"
'
'   Call TmpTable
'
'End If
'
'UpdateConfig ("133")
'MsgBox "Version 133 Installed successfully"
'
'
'End Function
'
'
'
'Public Function Version132()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Log_CxTF_Details ("
'tmp = tmp & " CxTFNo decimal(18,0) Not Null,"
'tmp = tmp & " TxnType   " & TextType & "(1) Not Null,"
'tmp = tmp & " SM_Prefix " & TextType & "(5) Not Null,"
'tmp = tmp & " CxTFDate datetime,"
'tmp = tmp & " CMPID decimal(18,0),"
'tmp = tmp & " CommodityID decimal(18,0),"
'tmp = tmp & " ExchangeTypeID decimal(18,0),"
'tmp = tmp & " BookedSpaceFlag   " & TextType & "(1),"
'tmp = tmp & " Flag " & TextType & "(1),"
'tmp = tmp & " Agent " & TextType & "(100),"
'tmp = tmp & " Remark " & TextType & "(255),"
'tmp = tmp & " ContractID decimal(18,0),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " HologramNo " & TextType & "(25),"
'tmp = tmp & " AgentID   decimal(18,0),"
'tmp = tmp & " ManualCDTFNo " & TextType & "(18),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " DeletedBy " & TextType & "(30),"
'   tmp = " PRIMARY KEY (CxTFNo,TxnType,SM_Prefix))"
'Else
'   tmp = tmp & " DeletedBy " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'tmp = " Create Table Log_WHR_Detail (       "
'tmp = tmp & " WHRNo decimal(18,0) Not Null,"
'tmp = tmp & " SM_Prefix " & TextType & "(5) Not Null,"
'tmp = tmp & " WHRDate   datetime,"
'tmp = tmp & " ConditionID decimal(18,0),"
'tmp = tmp & " StorageFrom datetime,"
'tmp = tmp & " StorageTo datetime,   "
'tmp = tmp & " Grade " & TextType & "(30),"
'tmp = tmp & " Weigher   " & TextType & "(200),"
'tmp = tmp & " Assayer   " & TextType & "(200),"
'tmp = tmp & " GradeByAssayer    " & TextType & "(50),"
'tmp = tmp & " GradeValidTill    Datetime,"
'tmp = tmp & " StorageCharges    decimal(18,2),"
'tmp = tmp & " StorageAmount decimal(18,2),"
'tmp = tmp & " ClientIDRow   decimal(18,0),"
'tmp = tmp & " CommodityID   decimal(18,0),"
'tmp = tmp & " CMPID decimal (18,0),"
'tmp = tmp & " LicenseID decimal(18,0),"
'tmp = tmp & " WHRNoOfBags   decimal(18,0),"
'tmp = tmp & " WHRQuantity   decimal(18,3),"
'tmp = tmp & " PrintFlag " & TextType & "(1),"
'tmp = tmp & " Remark    " & TextType & "(200),"
'tmp = tmp & " Flag  " & TextType & "(1),"
'tmp = tmp & " PledgeID  decimal(18,0),"
'tmp = tmp & " OldLicenseID  decimal(18,0),"
'tmp = tmp & " GodownID  decimal(18,0),"
'tmp = tmp & " DeliveredQuantity decimal (18,3),"
'tmp = tmp & " DeliveredNoOfBags decimal (18,0),"
'tmp = tmp & " Rate  decimal(18,2),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " HologramNo    " & TextType & "(25),"
'tmp = tmp & " PolicyNo  " & TextType & "(255),"
'tmp = tmp & " CoverNoteNo   " & TextType & "(255),"
'tmp = tmp & " SaidToContain " & TextType & "(1),"
'tmp = tmp & " EstimatedWeight   " & TextType & "(1),"
'tmp = tmp & " Concurrency   datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate   datetime,"
'tmp = tmp & " UpdatedDate   datetime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " SRAsPerBag    decimal(18,0),"
''tmp = tmp & " DeletedBy       " & TextType & "(30))"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " DeletedBy " & TextType & "(30),"
'   tmp = " PRIMARY KEY (WHRNo,SM_Prefix))"
'Else
'   tmp = tmp & " DeletedBy " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE [dbo].[Log_CxTF_Details] WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT [PK_Log_CxTF_Details] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ([CxTFNo],[TxnType],[SM_Prefix])  ON [PRIMARY] "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE [dbo].[Log_CxTF_Details] WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT [DF_Log_CxTF_Details_Concurrency] DEFAULT (getdate()) FOR [Concurrency]"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE [dbo].[Log_CxTF_Details] ADD "
'   tmp = tmp & " CONSTRAINT [FK_Log_CxTF_Details_Mst_ExchangeType] FOREIGN KEY "
'   tmp = tmp & " ([ExchangeTypeID]) REFERENCES [dbo].[Mst_ExchangeType] ([ExchangeTypeID]) ON UPDATE CASCADE "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE [dbo].[Log_WHR_Detail] WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT [PK_Log_WHR_Detail] PRIMARY KEY  CLUSTERED ([WHRNo],[SM_Prefix])  ON [PRIMARY] "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE [dbo].[Log_WHR_Detail] WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_StorageCharges] DEFAULT (0) FOR [StorageCharges],CONSTRAINT [DF_Log_WHR_Detail_StorageAmount] DEFAULT (0) FOR [StorageAmount],"
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_WHRNoOfBags] DEFAULT (0) FOR [WHRNoOfBags],"
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_WHRQuantity] DEFAULT (0) FOR [WHRQuantity],"
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_OldLicenseID] DEFAULT (0) FOR [OldLicenseID],"
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_DeliveredQuantity] DEFAULT (0) FOR [DeliveredQuantity],"
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_DeliveredNoOfBags] DEFAULT (0) FOR [DeliveredNoOfBags],"
'   tmp = tmp & " CONSTRAINT [DF_Log_WHR_Detail_Concurrency] DEFAULT (getdate()) FOR [Concurrency]"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE [dbo].[Log_WHR_Detail] ADD "
'   tmp = tmp & " CONSTRAINT [FK_Log_WHR_Detail_Mst_License] FOREIGN KEY ([LicenseID]) REFERENCES [dbo].[Mst_License] ([LicenseID]) ON UPDATE CASCADE "
'
'   Call TmpTable
'
'End If
'
'tmp = " Alter Table Txn_TIDetails Add " & c & " LocationUpdatedBy " & TextType & "(30)," & c & " LocationUpdatedDate DateTime "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ChequewiseClientLedger.frm','Cheque wise Client Ledger','FrmRpt_ChequewiseClientLedger','Y','R',0,229,1"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_LocationInvoiceMapping.frm','Location Tax Invoice Mapping','FrmTxn_LocationInvoiceMapping','Y','T',0,228,1"
'Call TmpTable
'
'UpdateConfig ("132")
'MsgBox "Version 132 Installed successfully"
'
'
'End Function
'
'
'Public Function Version131()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ClientWiseCxTFDetails.frm','Client Wise CxTF Details','FrmRpt_ClientWiseCxTFDetails','Y','R',0,229,1"
'
'Call TmpTable
'
'UpdateConfig ("131")
'MsgBox "Version 131 Installed successfully"
'
'
'End Function
'
'Public Function Version130()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter table Txn_TIDetails Add " & c & " ReverseFlag " & TextType & "(2)"
'Call TmpTable
'
'tmp = "Alter Table Txn_TempInvoice Add " & c & " DPID " & TextType & " (50)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportReverseTempInv.frm','Export Reverse Ad-Hoc Invoice','Frm_ExportReverseTempInv','Y','C',0,230,1"
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportReverseNNNRInvoice.frm','Export Reverse NNNR Invoice','Frm_ExportReverseNNNRInvoice','Y','C',0,230,1"
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'Frm_ExportReverseReservationInvoice.frm','Export Reverse Reservation Invoice','Frm_ExportReverseReservationInvoice','Y','C',0,230,1"
'Call TmpTable
'
'UpdateConfig ("130")
'MsgBox "Version 130 Installed successfully"
'
'End Function
'
'
'Public Function Version129()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Create table Txn_LockSealHologramDetail"
'tmp = tmp & " (LSHDID decimal(18,0) Not Null,"
'tmp = tmp & " LockSealHologramID Decimal(18,0),"
'tmp = tmp & " IssueDate Datetime,"
'tmp = tmp & " LocationID decimal(18,0),"
'tmp = tmp & " FromLSHologramNo Decimal(18,0),"
'tmp = tmp & " ToLSHologramNo Decimal(18,0),"
'tmp = tmp & " EmployeeID decimal(18,0),"
'tmp = tmp & " Address " & TextType & " (2000),"
'tmp = tmp & " Remark " & TextType & " (250)  NULL ,"
'tmp = tmp & " G_UID  " & TextType & "(36)  NULL ,"
'tmp = tmp & " Flag  " & TextType & "(1)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & "(30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & "(30)  NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime, "
'   tmp = " PRIMARY KEY (LSHDID))"
'Else
'   tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'tmp = " Create TABLE Log_LockSealHologramHistory ("
'tmp = tmp & " LockSealHologramID decimal(18, 0),"
'tmp = tmp & " UpdatedOn datetime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30))"
'
'Call TmpTable
'
'tmp = " Alter Table Txn_LockSealHologram add " & c & " EmployeeID Decimal(18,0), " & c & " Address " & TextType & "(2000) "
'
'Call TmpTable
'
'tmp = "Alter Table Txn_TIDetails Add " & c & " ServiceTaxID decimal (18,0) ," & c & " EntryFlag " & TextType & "(15) "
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " Alter table Txn_LockSealHologram WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_LockSealHologram_Mst_Employee FOREIGN KEY  (EmployeeID) REFERENCES Mst_Employee (EmployeeID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_LockSealHologramDetail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_LockSealHologramDetail  PRIMARY KEY  CLUSTERED(LSHDID),"
'   tmp = tmp & " CONSTRAINT DF_Txn_LockSealHologramDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT FK_Txn_LockSealHologramDetail_Txn_LockSealHologram FOREIGN KEY  (LockSealHologramID) REFERENCES Txn_LockSealHologram (LockSealHologramID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_LockSealHologramDetail_Mst_Location FOREIGN KEY  (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_LockSealHologramDetail_Mst_Employee FOREIGN KEY  (EmployeeID) REFERENCES Mst_Employee (EmployeeID)"
'
'   Call TmpTable
'
'   tmp = " CREATE TRIGGER tr_Txn_LockSealHologram_U ON Txn_LockSealHologram"
'   tmp = tmp & " AFTER UPDATE AS DECLARE"
'   tmp = tmp & " @mLockSealHologramID  decimal(18, 0),"
'   tmp = tmp & " @mUpdatedOn DateTime,"
'   tmp = tmp & " @mUpdatedBy Varchar(30)"
'   tmp = tmp & " SELECT"
'   tmp = tmp & " @mLockSealHologramID= d.LockSealHologramID,"
'   tmp = tmp & " @mUpdatedOn = GetDate(),"
'   tmp = tmp & " @mUpdatedBy = i.UpdatedBy"
'   tmp = tmp & " FROM  inserted i"
'   tmp = tmp & " INNER JOIN deleted d ON i.LockSealHologramID = d.LockSealHologramID"
'   tmp = tmp & " insert into Log_LockSealHologramHistory (LockSealHologramID ,UpdatedOn ,UpdatedBy)"
'   tmp = tmp & " values(@mLockSealHologramID ,@mUpdatedOn ,@mUpdatedBy)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_TIDetails WITH NOCHECK ADD"
'   tmp = tmp & " Constraint FK_Txn_TIDetails_Mst_ServiceTax Foreign KEY (ServiceTaxID) references Mst_ServiceTax(ServiceTaxID)"
'
'   Call TmpTable
'
'
'End If
'
'tmp = " Alter table Txn_TempInvoice add " & c & " ReverseDate Datetime, " & c & " ReverseFlag " & TextType & " (2), " & c & " ReverseRemark " & TextType & " (255), " & c & " RBatchID Decimal(18,0) "
'Call TmpTable
'
'tmp = " Update Txn_TempInvoice Set RBatchID=0 "
'Call TmpTable
'
'tmp = " Alter table Txn_ReservationInvoice add " & c & " ReverseDate Datetime, " & c & " ReverseFlag " & TextType & " (2), " & c & " ReverseRemark " & TextType & " (255), " & c & " RBatchID Decimal(18,0)"
'Call TmpTable
'
'tmp = " Update Txn_ReservationInvoice Set RBatchID=0"
'Call TmpTable
'
'tmp = " Alter table Txn_ProInvoice add " & c & " ReverseDate Datetime, " & c & " ReverseFlag " & TextType & " (2), " & c & " ReverseRemark " & TextType & " (255), " & c & " RBatchID Decimal(18,0)"
'Call TmpTable
'
'tmp = " Update Txn_ProInvoice Set RBatchID=0"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_TIDetails.frm','Manual Entry of Tax Invoice','FrmTxn_TIDetails','Y','T',0,228,1"
'
'Call TmpTable
'
'
'UpdateConfig ("129")
'MsgBox "Version 129 Installed successfully"
'
'End Function
'
'Public Function Version128()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Txn_LockSealHologram ("
'tmp = tmp & " LockSealHologramID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LockSealHologramDate datetime NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " FromLockSealHologramNo decimal(18, 0) NULL ,"
'tmp = tmp & " ToLockSealHologramNo decimal(18, 0) NULL ,"
'tmp = tmp & " TotalQty decimal(18, 0) NULL ,"
'tmp = tmp & " Remark " & TextType & " (1000)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime, "
'   tmp = " PRIMARY KEY (LockSealHologramID ))"
'Else
'   tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_LockSealHologram WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_LockSealHologram  PRIMARY KEY  CLUSTERED(LockSealHologramID),"
'   tmp = tmp & " CONSTRAINT DF_Txn_LockSealHologram_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT FK_Txn_LockSealHologram_Mst_Location FOREIGN KEY  (LocationID) REFERENCES Mst_Location (LocationID)"
'
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_LockSealHologram.frm','Lock Seal Hologram','FrmTxn_LockSealHologram','Y','T',0,228,1"
'
'Call TmpTable
'
'UpdateConfig ("128")
'MsgBox "Version 128 Installed successfully"
'
'End Function
'
'Public Function Version127()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Create Table Txn_CRAdvanceAllocation("
'tmp = tmp & " CRAAID  decimal(18,0) Not Null,"
'tmp = tmp & " CRID  decimal(18,0),"
'tmp = tmp & " SM_Prefix " & TextType & "(5),"
'tmp = tmp & " AllocDate  Datetime,"
'tmp = tmp & " LocationID  decimal(18,0),"
'tmp = tmp & " Remark  " & TextType & "(255),"
'tmp = tmp & " Flag  " & TextType & "(2),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " BatchID Decimal(18,0),"
'tmp = tmp & " FAGRemarks  " & TextType & "(255),"
'tmp = tmp & " FAGPostingDate  datetime,"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy  " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = " PRIMARY KEY (CRAAID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'tmp = " Create Table Txn_CRTDSAdjustment("
'tmp = tmp & " TDSAID  decimal(18,0) Not Null,"
'tmp = tmp & " SM_Prefix " & TextType & "(5)Not Null,"
'tmp = tmp & " AdjDate  Datetime,"
'tmp = tmp & " LocationID  decimal(18,0),"
'tmp = tmp & " ExchangeTypeID  decimal(18,0),"
'tmp = tmp & " ShowDP decimal(18,0),"
'tmp = tmp & " ClientIDRow decimal(18,0),"
'tmp = tmp & " DPID  " & TextType & "(50),"
'tmp = tmp & " Remark  " & TextType & "(255),"
'tmp = tmp & " Flag  " & TextType & "(2),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " BatchID Decimal(18,0),"
'tmp = tmp & " FAGRemarks  " & TextType & "(255),"
'tmp = tmp & " FAGPostingDate  datetime,"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy  " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = " PRIMARY KEY (TDSAID,SM_Prefix))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'tmp = " Create Table Txn_CRTDSAdjustmentDetail("
'tmp = tmp & " TDSADID decimal(18,0)Not Null,"
'tmp = tmp & " TDSAID  decimal(18,0),"
'tmp = tmp & " SM_Prefix " & TextType & "(5),"
'tmp = tmp & " TIID  decimal(18,0),"
'tmp = tmp & " RecTDS  decimal(18,2),"
'tmp = tmp & " RecOtherDed decimal(18,2),"
'tmp = tmp & " ExpenseAccID decimal(18,0),"
'tmp = tmp & " Flag  " & TextType & "(2),"
'tmp = tmp & " G_UID  " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy  " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = " PRIMARY KEY (TDSADID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'tmp = " Create Table Txn_ClientLedger( "
'tmp = tmp & " CustomerNo  " & TextType & " (10), "
'tmp = tmp & " CustomerName    " & TextType & " (500), "
'tmp = tmp & " EntryNo decimal(18,0), "
'tmp = tmp & " EHOPInvoiceNo   " & TextType & "(50), "
'tmp = tmp & " DocumentType    " & TextType & "(50), "
'tmp = tmp & " DocumentNo  " & TextType & "(50), "
'tmp = tmp & " VoucherNo   " & TextType & "(50), "
'tmp = tmp & " PostingDate datetime, "
'tmp = tmp & " InvoiceAmount   decimal (18,2), "
'tmp = tmp & " BaseAmount  decimal (18,2), "
'tmp = tmp & " TaxAmount   decimal (18,2), "
'tmp = tmp & " RemainingAmount decimal (18,2), "
'tmp = tmp & " Narration   " & TextType & " (2000), "
'tmp = tmp & " ActualDatePosted datetime, "
'tmp = tmp & " ActualTimePosted " & TextType & "(20), "
'tmp = tmp & " RegionCode  " & TextType & "(30), "
'tmp = tmp & " DepartmentCode  " & TextType & "(30), "
'tmp = tmp & " CustomerPostingGroup " & TextType & " (20), "
'tmp = tmp & " CreatedBy   " & TextType & " (30), "
'tmp = tmp & " CreatedDate datetime, "
'tmp = tmp & " Concurrency datetime, "
'tmp = tmp & " G_UID   " & TextType & "(36)) "
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_CRAdvanceAllocation WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_CRAdvanceAllocation PRIMARY KEY  CLUSTERED (CRAAID),"
'   tmp = tmp & " CONSTRAINT DF_Txn_CRAdvanceAllocation_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " Constraint FK_Txn_CRAdvanceAllocation_Mst_Location Foreign KEY (LocationID) references Mst_Location(LocationID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_CRTDSAdjustment WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_CRTDSAdjustment PRIMARY KEY  CLUSTERED (TDSAID,SM_Prefix),"
'   tmp = tmp & " CONSTRAINT DF_Txn_CRTDSAdjustment_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " Constraint FK_Txn_CRTDSAdjustment_Mst_Location Foreign KEY (LocationID) references Mst_Location(LocationID),"
'   tmp = tmp & " Constraint FK_Txn_CRTDSAdjustment_Mst_ExchangeType Foreign KEY (ExchangeTypeID) references Mst_ExchangeType(ExchangeTypeID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_CRTDSAdjustmentDetail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_CRTDSAdjustmentDetail PRIMARY KEY  CLUSTERED (TDSADID),"
'   tmp = tmp & " CONSTRAINT DF_Txn_CRTDSAdjustmentDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " Constraint FK_Txn_CRTDSAdjustmentDetail_Txn_CRTDSAdjustment Foreign KEY (TDSAID,SM_Prefix) references Txn_CRTDSAdjustment(TDSAID,SM_Prefix),"
'   tmp = tmp & " Constraint FK_Txn_CRTDSAdjustmentDetail_Txn_TIDetails Foreign KEY (TIID) references Txn_TIDetails(TIID)"
'
'   Call TmpTable
'
'End If
'
'tmp = " Alter table Txn_NewCashReceiptTIDetail Add " & c & " CRAAID decimal(18,0)"
'Call TmpTable
'
'tmp = " Alter Table Txn_NewCashReceiptClientDetail Add  " & c & " AdvNAVCode " & TextType & "(25)"
'
'Call TmpTable
'
'tmp = " Alter Table Mst_ExpenseType Add " & c & " OtherDeductionFlag " & TextType & "(1)"
'Call TmpTable
'
'tmp = " Update Mst_ExpenseType Set OtherDeductionFlag = 'N' "
'Call TmpTable
'
'tmp = " Alter Table Mst_MenuDetail Add " & c & " SelfMenu decimal (18,0), " & c & " MasterMenuID decimal(18,0), " & c & " SequenceNo decimal(18,0)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Values (227,'Master','Master','Master','Y','',1,Null,1)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Values(228,'Transaction','Transaction','Transaction','Y','',1,Null,2)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Values(229,'Report','Report','Report','Y','',1,Null,3)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Values(230,'CMG Process','CMG Process','CMG Process','Y','',1,Null,4)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Values(231,'Database Managment','Database Managment','Database Managment','Y','',1,Null,5)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Values(232,'Admin','Admin','Admin','Y','',1,Null,6)"
'
'Call TmpTable
'
'tmp = " Update Mst_MenuDetail Set SelfMenu = 0,MasterMenuID = 227,SequenceNo = 1 Where HMenuFlag = 'M' "
'Call TmpTable
'
'tmp = " Update Mst_MenuDetail Set SelfMenu = 0,MasterMenuID = 228,SequenceNo = 1 Where HMenuFlag = 'T'"
'Call TmpTable
'
'tmp = " Update Mst_MenuDetail Set SelfMenu = 0,MasterMenuID = 229,SequenceNo = 1 Where HMenuFlag = 'R'"
'Call TmpTable
'
'tmp = " Update Mst_MenuDetail Set SelfMenu = 0,MasterMenuID = 230,SequenceNo = 1 Where HMenuFlag = 'C'"
'Call TmpTable
'
'tmp = " Update Mst_MenuDetail Set SelfMenu = 0,MasterMenuID = 232,SequenceNo = 1 Where HMenuFlag = 'A'"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_CRAdvanceAllocation.frm','CR Advance Allocation','FrmTxn_CRAdvanceAllocation','Y','T',0,228,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_CRTDSAdjustment.frm','CR TDS Adjustment','FrmTxn_CRTDSAdjustment','Y','T',0,228,1"
'
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmTxn_ClientLedgerUpload.frm','Client Ledger Upload','FrmTxn_ClientLedgerUpload','Y','C',0,230,1"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, SelfMenu,MasterMenuID,SequenceNo)"
'tmp = tmp & " Select(Select Max(MenuID) from Mst_MenuDetail)+1,'FrmRpt_ClientLedger.frm','Client Ledger','FrmRpt_ClientLedger','Y','R',0,229,1 "
'
'Call TmpTable
'
'UpdateConfig ("127")
'MsgBox "Version 127 Installed successfully"
'
'
'End Function
'
'
'
'
'Public Function Version126()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Create Table Log_BODUpdate("
'tmp = tmp & " CMPID Decimal (18,0),"
'tmp = tmp & " NewBODDate DateTime,"
'tmp = tmp & " OldBODDate DateTime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " UpdatedDate DateTime)"
'
'Call TmpTable
'
'
'tmp = " Alter Table  Txn_CxTF_ClientDetail Add  SplRRNQty Decimal(18,3),SplRRNNoofBags Decimal(18,0) "
'Call TmpTable
'
'tmp = " Alter Table Txn_CxTF_GSL Add WHRNo Decimal(18,0) "
'Call TmpTable
'
'
'tmp = "Alter Table Mst_RRN Add " & c & " OutSystemUpdatedBy " & TextType & " (30), " & c & " OutSystemUpdatedDate DateTime "
'Call TmpTable
'
'
'tmp = "Alter Table Txn_MonthEmp_History Add FAGUpdated " & TextType & "(30),FAGUpdatedDate DateTime "
'Call TmpTable
'
'UpdateConfig ("126")
'MsgBox "Version 126 Installed successfully"
'
'End Function
'
'Public Function Version125()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter table Txn_NewCashReceiptTIDetail Add " & c & " DocNo " & TextType & "(50)," & c & " EntryNo Decimal(18,0)"
'Call TmpTable
'
'tmp = "Alter table Txn_NewCashReceiptClientDetail Add " & c & " DocNo " & TextType & "(50)," & c & " EntryNo Decimal(18,0)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (224,'Frm_ExportCRInvoice.frm','Export CR Invoices Allocation','Frm_ExportCRInvoice','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (225,'Frm_UpdateCREntryNo.frm','Update CR EntryNo','Frm_UpdateCREntryNo','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (226,'FrmChangePassword.frm','Change Password','FrmChangePassword','Y','A',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'UpdateConfig ("125")
'MsgBox "Version 125 Installed successfully"
'
'End Function
'Public Function Version124()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_NewCashReceipt Add DepositDate DateTime "
'Call TmpTable
'
'UpdateConfig ("124")
'MsgBox "Version 124 Installed successfully"
'
'End Function
'
'Public Function Version123()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_NewCashReceipt Add FAGBankID Decimal(18,0),FAGUpdatedBy " & TextType & "(30),FAGUpdatedDate DateTime"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_NewCashReceipt WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_NewCashReceipt_Mst_FAGBank FOREIGN KEY"
'   tmp = tmp & " (FAGBankID) REFERENCES Mst_FAGBank(FAGBankID)"
'
'   Call TmpTable
'
'End If
'UpdateConfig ("123")
'MsgBox "Version 123 Installed successfully"
'
'End Function
'
'Public Function Version122()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_NewCashReceiptTIDetail Add " & c & " ClientIDRow Decimal(18,0), " & c & " DPID " & TextType & "(50)"
'Call TmpTable
'
'UpdateConfig ("122")
'MsgBox "Version 122 Installed successfully"
'
'End Function
'Public Function Version121()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Mst_StateIndia("
'tmp = tmp & " StateIndiaID Decimal(18,0) Not Null,"
'tmp = tmp & " StateIndiaName " & TextType & "(100) Not Null,"
'tmp = tmp & " Flag " & TextType & "(1),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " Concurrency Datetime Not Null,"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " UpdatedDate Datetime,"
'tmp = tmp & " Remarks " & TextType & "(1000),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (StateIndiaID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = " Alter table Mst_State Add " & c & "  StateIndiaID Decimal(18,0)"
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Mst_StateIndia WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_StateIndia PRIMARY KEY CLUSTERED (StateIndiaID)"
'   Call TmpTable
'
'
'   tmp = " ALTER TABLE Mst_StateIndia ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_StateIndia_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_State WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_State_Mst_StateIndia FOREIGN KEY"
'   tmp = tmp & " (StateIndiaID) REFERENCES Mst_StateIndia(StateIndiaID)"
'   Call TmpTable
'
'End If
'
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " Values(1,'Andhra Pradesh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(2,'Arunachal Pradesh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(3,'Assam','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(4,'Chandigarh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(5,'Chattisgarh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(6,'Delhi','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(7,'Goa','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(8,'Gujarat','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(9,'Bihar','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(10,'Haryana','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(11,'Himachal Pradesh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(12,'Jammu and Kashmir','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(13,'Jharkhand','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(14,'Karnataka','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(15,'Kerala','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(16,'Madhya Pradesh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(17,'Maharashtra','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(18,'Manipur','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(19,'Nagaland','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(20,'Orissa','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(21,'Punjab','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(22,'Rajasthan','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(23,'Tamil Nadu','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(24,'Tripura','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(25,'Uttar Pradesh','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(26,'Uttarakhand','A',Newid())"
'Call TmpTable
'
'tmp = " Insert Into Mst_StateIndia (StateIndiaID,StateIndiaName ,Flag,G_UID)"
'tmp = tmp & " values(27,'West Bengal','A',Newid())"
'Call TmpTable
'
'
'tmp = " Update Mst_State set StateIndiaID = StateId,Flag = 'A' "
'tmp = tmp & " Where StateIndiaId is null"
'
'Call TmpTable
'
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(223, 'FrmMst_StateIndia.frm', 'State India', 'FrmMst_StateIndia', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'UpdateConfig ("121")
'MsgBox "Version 121 Installed successfully"
'
'
'
'End Function
'Public Function Version120()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter Table Txn_QualityInformation Add " & c & " Remarks " & TextType & " (2000)"
'Call TmpTable
'
'
'tmp = "Update Txn_NewCashReceipt Set ReturnFlag='N',RBatchID=0 Where ReturnFlag is null "
'Call TmpTable
'
'
'tmp = " Alter Table Txn_TIDetails Add " & c & " SMID " & TextType & " (50)"
'Call TmpTable
'
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (222,' Frm_ExportSecurityBilling.frm','Export Security Billing','Frm_ExportSecurityBilling','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'
'
'UpdateConfig ("120")
'MsgBox "Version 120 Installed successfully"
'
'
'
'End Function
'
'Public Function Version119()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_SecurityAgency Add " & c & " TDSID Decimal(18,0), " & c & " NAVCode " & TextType & "(25), " & c & " FAGUpdated " & TextType & "(30), " & c & "  FAGUpdatedDate DateTime "
'Call TmpTable
'
'tmp = "Alter table Txn_NewCashReceipt Add " & c & " EntryNo " & TextType & "(25)," & c & "ReturnDate datetime," & c & "RBatchID Decimal(18,0)," & c & "ReturnRemark " & TextType & "(100)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_SecurityAgency WITH NOCHECK ADD"
'   tmp = tmp & " Constraint FK_Mst_SecurityAgency_Mst_TDS Foreign KEY (TDSID) references Mst_TDS(TDSID)"
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(220,'FrmRpt_ModifiedNNNRInvoice.frm','Modified NNNR Invoice','FrmRpt_ModifiedNNNRInvoice', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (221,'Frm_ExportReturnCashReceipt.frm','Export Return Cash Receipt','Frm_ExportReturnCashReceipt','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'UpdateConfig ("119")
'MsgBox "Version 119 Installed successfully"
'
'
'
'End Function
'
'Public Function Version118()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (219,' FrmRpt_InvoiceDetails.frm','Invoice Details',' FrmRpt_InvoiceDetails','Y','R',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'
'UpdateConfig ("118")
'MsgBox "Version 118 Installed successfully"
'
'
'
'End Function
'
'Public Function Version117()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter Table Mst_Location Add " & c & " Collect7Days " & TextType & "(1)"
'Call TmpTable
'
'UpdateConfig ("117")
'MsgBox "Version 117 Installed successfully"
'
'
'
'End Function
'
'Public Function Version116()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Create Table Txn_NewCashReceipt("
'tmp = tmp & " CRID  decimal(18,0) Not Null,"
'tmp = tmp & " SM_Prefix  " & TextType & "(5)Not Null,"
'tmp = tmp & " CRDate  Datetime,"
'tmp = tmp & " LocationID  decimal(18,0),"
'tmp = tmp & " InstrumentType   " & TextType & "(50),"
'tmp = tmp & " DrawnOn  " & TextType & "(150),"
'tmp = tmp & " BranchName  " & TextType & "(150),"
'tmp = tmp & " InstrumentNo   " & TextType & "(25),"
'tmp = tmp & " InstrumentDate Datetime,"
'tmp = tmp & " InstrumentAmount  decimal(18,2),"
'tmp = tmp & " ReceivedFrom   " & TextType & "(100),"
'tmp = tmp & " Remark   " & TextType & "(255),"
'tmp = tmp & " AdvanceFlag  " & TextType & "(2),"
'tmp = tmp & " Flag   " & TextType & "(2),"
'tmp = tmp & " G_UID  " & TextType & "(36),"
'tmp = tmp & " HologramNo   " & TextType & "(25),"
'tmp = tmp & " CRDepositID decimal(18,0),"
'tmp = tmp & " SightedDate datetime,"
'tmp = tmp & " ReturnFlag  " & TextType & "(2),"
'tmp = tmp & " BatchID Decimal(18,0),"
'tmp = tmp & " FAGRemarks   " & TextType & "(255),"
'tmp = tmp & " FAGPostingDate DateTime,"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy   " & TextType & "(30),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (CRID,SM_Prefix))"
'Else
'     tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'tmp = " Create table Txn_NewCashReceiptClientDetail("
'tmp = tmp & " CRCDID  decimal(18,0)Not Null,"
'tmp = tmp & " CRID    decimal(18,0),"
'tmp = tmp & " SM_Prefix  " & TextType & "(5),"
'tmp = tmp & " ExchangeTypeID  decimal(18,0),"
'tmp = tmp & " ClientIDRow decimal(18,0),"
'tmp = tmp & " Remark   " & TextType & "(255),"
'tmp = tmp & " Amount  decimal(18,2),"
'tmp = tmp & " DPID   " & TextType & "(50),"
'tmp = tmp & " ShowDP decimal(18,0),"
'tmp = tmp & " Flag   " & TextType & "(2),"
'tmp = tmp & " G_UID  " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy  " & TextType & "(30),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (CRCDID))"
'Else
'     tmp = tmp & " UpdatedDate datetime) "
'End If
'Call TmpTable
'
'tmp = " Create Table Txn_NewCashReceiptTIDetail("
'tmp = tmp & " CRTIDID decimal(18,0)Not Null,"
'tmp = tmp & " CRID  decimal(18,0),"
'tmp = tmp & " SM_Prefix  " & TextType & "(5),"
'tmp = tmp & " TIID  decimal(18,0),"
'tmp = tmp & " RecAmount  decimal(18,2),"
'tmp = tmp & " RecTDS  decimal(18,2),"
'tmp = tmp & " RecOtherDed decimal(18,2),"
'tmp = tmp & " Flag   " & TextType & "(2),"
'tmp = tmp & " G_UID   " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy   " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy  " & TextType & "(30),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (CRTIDID))"
'Else
'     tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_NewCashReceiptDepositDetails ("
'tmp = tmp & " CRDepositID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FAGBankID decimal(18, 0) NULL ,"
'tmp = tmp & " DepositDate datetime NULL ,"
'tmp = tmp & " Remark  " & TextType & " (250)  NULL ,"
'tmp = tmp & " G_UID  " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy  " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy  " & TextType & " (30),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (CRDepositID))"
'Else
'     tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'    tmp = " ALTER TABLE Txn_NewCashReceipt WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_NewCashReceipt PRIMARY KEY  CLUSTERED (CRID,SM_Prefix),"
'    tmp = tmp & " CONSTRAINT DF_Txn_NewCashReceipt_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceipt_Mst_Location Foreign KEY (LocationID) references Mst_Location(LocationID)"
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_NewCashReceiptClientDetail WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_NewCashReceiptClientDetail PRIMARY KEY  CLUSTERED (CRCDID),"
'    tmp = tmp & " CONSTRAINT DF_Txn_NewCashReceiptClientDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceiptClientDetail_Txn_NewCashReceipt Foreign KEY (CRID,SM_Prefix) references Txn_NewCashReceipt(CRID,SM_Prefix),"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceiptClientDetail_Mst_ExchangeType Foreign KEY (ExchangeTypeID) references Mst_ExchangeType(ExchangeTypeID)"
'
'    Call TmpTable
'
'    tmp = " AlTER TABLE Txn_NewCashReceiptTIDetail WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_NewCashReceiptTIDetail PRIMARY KEY  CLUSTERED (CRTIDID),"
'    tmp = tmp & " CONSTRAINT DF_Txn_NewCashReceiptTIDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceiptTIDetail_Txn_NewCashReceipt Foreign KEY (CRID,SM_Prefix) references Txn_NewCashReceipt(CRID,SM_Prefix),"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceiptTIDetail_Txn_TIDetails Foreign KEY (TIID) references Txn_TIDetails(TIID)"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_NewCashReceiptDepositDetails WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_NewCashReceiptDepositDetails PRIMARY KEY  CLUSTERED (CRDepositID),"
'    tmp = tmp & " CONSTRAINT DF_Txn_NewCashReceiptDepositDetails_Concurrency DEFAULT (getdate()) FOR Concurrency ,"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceiptDepositDetails_Mst_Location Foreign KEY (LocationID) references Mst_Location(LocationID),"
'    tmp = tmp & " Constraint FK_Txn_NewCashReceiptDepositDetails_Mst_FAGBank Foreign KEY (FAGBankID) references Mst_FAGBank(FAGBankID)"
'
'    Call TmpTable
'
'End If
'
'
'tmp = "Alter Table Txn_TIDetails add " & c & " PostInvoiceNo " & TextType & "(25), " & c & " EntryNo decimal(18,0), " & c & " NAVCode " & TextType & "(25)"
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (212,'Frm_ImportReservationInvoice.frm','Import Reservation Invoice','Frm_ImportReservationInvoice','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (213,'Frm_UpdateNNNRInvoiceNo.frm','Import NNNR Invoice','Frm_UpdateNNNRInvoiceNo','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (214,'Frm_UpdateTempInvoiceNo.frm','Import Ad-Hoc Invoice','Frm_UpdateTempInvoiceNo','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (215,'FrmTxn_NewCashReceipt.frm','New Cash Receipt','FrmTxn_NewCashReceipt','Y','T',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (216,'Frm_ExportCashReceipt.frm','Export Cash Receipt','Frm_ExportCashReceipt','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (217,'FrmTxn_NewCRDepositeDetails.frm','New Cash Receipt Deposit','FrmTxn_NewCRDepositeDetails','Y','T',NULL,NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (218,'FrmTxn_CRSightedDetails.frm','New Cash Receipt Sighted','FrmTxn_CRSightedDetails','Y','T',NULL,NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'UpdateConfig ("116")
'MsgBox "Version 116 Installed successfully"
'
'End Function
'
'Public Function Version115()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Create Table Txn_TempInvoice( "
'tmp = tmp & " PINo    decimal(18,0) Not Null,"
'tmp = tmp & " PIDate  datetime,"
'tmp = tmp & " InvoiceTypeID   decimal(18,0),"
'tmp = tmp & " ExchangeTypeID  decimal(18,0),"
'tmp = tmp & " LocationID  decimal(18,0),"
'tmp = tmp & " CommodityID decimal(18,0),"
'tmp = tmp & " ClientIDRow decimal(18,0),"
'tmp = tmp & " InvoiceDate datetime,"
'tmp = tmp & " Amount  decimal(18,2),"
'tmp = tmp & " Flag     " & TextType & "(2),"
'tmp = tmp & " BatchID decimal(18,0),"
'tmp = tmp & " ServiceTaxID    decimal(18,0),"
'tmp = tmp & " Quantity decimal(18, 3),"
'tmp = tmp & " Remarks  " & TextType & "(255),"
'tmp = tmp & " FAGRemarks   " & TextType & "(255),"
'tmp = tmp & " FAGPostingDate  datetime,"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy    " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy    " & TextType & "(30),"
'tmp = tmp & " G_UID    " & TextType & "(36),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (PINo ))"
'Else
'     tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'    tmp = " ALTER TABLE Txn_TempInvoice WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_TempInvoice PRIMARY KEY  CLUSTERED (PINo),"
'    tmp = tmp & " CONSTRAINT DF_Txn_TempInvoice_Concurrency DEFAULT (getdate()) FOR Concurrency ,"
'    tmp = tmp & " Constraint FK_Txn_TempInvoice_Mst_Location Foreign KEY (LocationID) references Mst_Location(LocationID),"
'    tmp = tmp & " Constraint FK_Txn_TempInvoice_Mst_ExchangeType Foreign KEY (ExchangeTypeID) references Mst_ExchangeType(ExchangeTypeID),"
'    tmp = tmp & " Constraint FK_Txn_TempInvoice_Mst_Commodity Foreign KEY (CommodityID) references Mst_Commodity(CommodityID),"
'    tmp = tmp & " Constraint FK_Txn_TempInvoice_Mst_IncomeExpenseType Foreign KEY (InvoiceTypeID) references Mst_IncomeExpenseType(IETypeID)"
'
'    Call TmpTable
'
'End If
'
'tmp = "Alter table Mst_Client Add " & c & " RNAVCode " & TextType & " (25) "
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (210,'FrmTxn_TempInvoice.frm','Ad-Hoc Invoice','FrmTxn_TempInvoice','Y','T',NULL,NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (211,'Frm_ExportTempInvoice.frm','Export Adhoc Invoice','Frm_ExportTempInvoice','Y','C',NULL,NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'UpdateConfig ("115")
'MsgBox "Version 115 Installed successfully"
'
'End Function
'
'
'Public Function Version114()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (209,'Frm_ExportReservationInvoice.frm','Export Reservation Invoice','Frm_ExportReservationInvoice','Y','C',NULL,NULL,NULL,NULL,NULL)"
'
'Call TmpTable
'
'UpdateConfig ("114")
'MsgBox "Version 114 Installed successfully"
'
'End Function
'
'Public Function Version113()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Create Table Txn_ReservationInvoice ("
'tmp = tmp & " InvoiceNo Decimal(18,0) Not Null, "
'tmp = tmp & " ContractID Decimal (18,0) Not Null,"
'tmp = tmp & " LocationID Decimal(18,0) Not Null,"
'tmp = tmp & " YearMonth DateTime Not Null,"
'tmp = tmp & " BillingDate DateTime Not Null,"
'tmp = tmp & " BillingAmount Decimal(18,2),"
'tmp = tmp & " ActualAmount Decimal(18,2),"
'tmp = tmp & " TINo " & TextType & "(50),"
'tmp = tmp & " TIDate DateTime,"
'tmp = tmp & " TIAmount Decimal(18,2),"
'tmp = tmp & " BatchID Decimal(18,0),"
'tmp = tmp & " Remarks " & TextType & " (2000),"
'tmp = tmp & " Concurrency DateTime Not Null,"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate DateTime,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " UpdatedDate DateTime ,"
'tmp = tmp & " Flag " & TextType & "(1),"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " ServiceTaxNAVCode " & TextType & " (25),"
'   tmp = " PRIMARY KEY (InvoiceNo))"
'Else
'   tmp = tmp & " ServiceTaxNAVCode " & TextType & " (25) )"
'End If
'
'Call TmpTable
'
'
'tmp = " Create Table Txn_TIDetails("
'tmp = tmp & " TIID    decimal(18,0)Not Null,"
'tmp = tmp & " TINo  " & TextType & "(50),"
'tmp = tmp & " TIDate  datetime,"
'tmp = tmp & " TIAmount    decimal(18,2),"
'tmp = tmp & " TISTAmount  decimal(18,2),"
'tmp = tmp & " PIID    decimal(18,0),"
'tmp = tmp & " PIFlag   " & TextType & "(25),"
'tmp = tmp & " LocationID  decimal(18,0),"
'tmp = tmp & " ExchangeTypeID  decimal(18,0),"
'tmp = tmp & " ClientIDRow decimal(18,0),"
'tmp = tmp & " DepositoryID     " & TextType & "(50),"
'tmp = tmp & " RecAmount   decimal(18,2),"
'tmp = tmp & " RecTDSAmount    decimal(18,2),"
'tmp = tmp & " RecOtherDed decimal(18,2),"
'tmp = tmp & " G_UID    " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy    " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy    " & TextType & "(30),"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (TIID ))"
'Else
'     tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Txn_ReservationInvoice ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_ReservationInvoice_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT DF_Txn_ReservationInvoice_BatchID DEFAULT (0) FOR BatchID,"
'   tmp = tmp & " CONSTRAINT PK_Txn_ReservationInvoice PRIMARY KEY  CLUSTERED (InvoiceNo)"
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_ReservationInvoice ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ReservationInvoice_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ReservationInvoice ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ReservationInvoice_Mst_StorageContract FOREIGN KEY (ContractID) REFERENCES Mst_StorageContract (ContractID)"
'
'   Call TmpTable
'
'
'    tmp = " ALTER TABLE Txn_TIDetails WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_TIDetails PRIMARY KEY  CLUSTERED (TIID),"
'    tmp = tmp & " CONSTRAINT DF_Txn_TIDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'End If
'
'tmp = " Alter Table Mst_IncomeExpenseType Add " & c & " ExchangeTypeID decimal(18,0)"
'Call TmpTable
'
'tmp = " Alter Table Mst_FAGBank Add " & c & "  NAVCode Varchar(25)"
'
'Call TmpTable
'
'tmp = " Create table Log_ReservationInvoice ("
'tmp = tmp & " InvoiceNo decimal(18,0), "
'tmp = tmp & " UpdatedDate datetime,"
'tmp = tmp & " UpdatedBy decimal(18,0),"
'tmp = tmp & " Remarks " & TextType & "(2000))"
'Call TmpTable
'
'tmp = " CREATE TRIGGER TR_Txn_ReservationInvoice ON Txn_ReservationInvoice"
'tmp = tmp & "  after Update"
'tmp = tmp & " AS"
'tmp = tmp & " DECLARE @mInvoiceNo decimal(18,0),"
'tmp = tmp & " @mUpdatedDate datetime,"
'tmp = tmp & " @mUpdatedBy   " & TextType & "  (30),"
'tmp = tmp & " @mRemarks   " & TextType & " (2000),"
'tmp = tmp & " @mOldRemark   " & TextType & " (2000)"
'tmp = tmp & " SELECT"
'tmp = tmp & " @mInvoiceNo=i.InvoiceNo,"
'tmp = tmp & " @mUpdatedDate = i.UpdatedDate,"
'tmp = tmp & " @mUpdatedBy = i.UpdatedBy,"
'tmp = tmp & " @mRemarks= i.Remarks,"
'tmp = tmp & " @mOldRemark=d.Remarks"
'tmp = tmp & " FROM inserted i"
'tmp = tmp & " INNER JOIN deleted d ON i.InvoiceNo = d.InvoiceNo"
'
'tmp = tmp & " if @mRemarks = @mOldRemark"
'tmp = tmp & " begin"
'tmp = tmp & " set @mRemarks =null;"
'tmp = tmp & " End"
'tmp = tmp & " insert into Log_ReservationInvoice(InvoiceNo,UpdatedDate ,UpdatedBy,Remarks)"
'tmp = tmp & " values(@mInvoiceNo,@mUpdatedDate,@mUpdatedBy,@mRemarks)"
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (205,'FrmTxn_GenerateReservationInvoice.frm','Generate Reservation Invoice','FrmTxn_GenerateReservationInvoice','Y','T',NULL,NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " Values (206,'FrmTxn_ReservationInvoice.frm','Reservation Invoice','FrmTxn_ReservationInvoice','Y','T',NULL,NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'
'UpdateConfig ("113")
'MsgBox "Version 113 Installed successfully"
'
'End Function
'
'
'Public Function Version112()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Mst_State Add Remarks " & TextType & " (1000) "
'Call TmpTable
'
'tmp = " Create Table Mst_StateDRCMapping ( "
'tmp = tmp & " MappingID Decimal(18,0) Not Null,"
'tmp = tmp & " EmployeeID Decimal(18,0),"
'tmp = tmp & " StateID Decimal(18,0),"
'tmp = tmp & " EmailID Varchar(100),"
'tmp = tmp & " Flag Varchar (1),"
'tmp = tmp & " Concurrency DateTime,"
'tmp = tmp & " G_UID Varchar(36),"
'tmp = tmp & " CreatedBy Varchar (30),"
'tmp = tmp & " CreatedDate DateTime,"
'tmp = tmp & " UpdatedDate DateTime,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30),"
'   tmp = " PRIMARY KEY (MappingID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30) )"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Mst_StateDRCMapping ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_StateDRCMapping_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Mst_StateDRCMapping PRIMARY KEY  CLUSTERED (MappingID)"
'
'   Call TmpTable
'
'End If
'
'
'tmp = "Alter Table Txn_GSL_Transfer Add " & c & " TransferedBags Decimal(18,0), " & c & "  TransferedWeight Decimal(18,3) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(204, 'FrmMst_StateDRCMapping.frm', 'State DRC Mapping', 'FrmMst_StateDRCMapping', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("112")
'MsgBox "Version 112 Installed successfully"
'
'End Function
'
'Public Function Version111()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Mst_Region ("
'tmp = tmp & " RegionID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Region  " & TextType & "(100),"
'tmp = tmp & " Flag  " & TextType & "(2),"
'tmp = tmp & " G_UID  " & TextType & "(36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy  " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30),"
'   tmp = " PRIMARY KEY (RegionID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30) )"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_Region ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_Region_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Mst_Region PRIMARY KEY  CLUSTERED (RegionID)"
'
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(203, 'FrmMst_Region.frm', 'Region Master', 'FrmMst_Region', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("111")
'MsgBox "Version 111 Installed successfully"
'
'End Function
'
'Public Function Version110()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(202, 'FrmRpt_ShortageMIS.frm', 'Shortage MIS', 'FrmRpt_ShortageMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("110")
'MsgBox "Version 110 Installed successfully"
'
'End Function
'
'Public Function Version109()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " CREATE TABLE Mst_CoExpense ("
'tmp = tmp & " CoExpenseID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FinancialYr " & TextType & " (50),"
'tmp = tmp & " ExpenseAmt decimal(18, 2),"
'tmp = tmp & " FromDate datetime,"
'tmp = tmp & " ToDate datetime,"
'tmp = tmp & " Flag " & TextType & " (1) ,"
'tmp = tmp & " G_UID " & TextType & " (36) ,"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30),"
'   tmp = " PRIMARY KEY (CoExpenseID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30) )"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_CoExpense ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_CoExpense_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Mst_CoExpense PRIMARY KEY  CLUSTERED (CoExpenseID)"
'
'   Call TmpTable
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(201, 'FrmMst_CoExpense.frm', 'CO Expense', 'FrmMst_CoExpense', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("109")
'MsgBox "Version 109 Installed successfully"
'
'End Function
'Public Function Version108()
'
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_Client Add FAGUpdated " & TextType & " (30) , FAGUpdatedDate datetime "
'Call TmpTable
'
'tmp = "Alter Table Mst_Vendor Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime "
'Call TmpTable
'
'tmp = "Alter Table Mst_ITVendor Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime "
'Call TmpTable
'
'tmp = "Alter Table Mst_Employee Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime "
'Call TmpTable
'
'tmp = "Alter Table Mst_Depository Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime "
'Call TmpTable
'
'
'tmp = "Alter Table Mst_Location Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime"
'Call TmpTable
'
'tmp = "Alter Table Mst_ExpenseType Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime"
'Call TmpTable
'
'tmp = "Alter Table Mst_Agent Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime"
'Call TmpTable
'
'tmp = "Alter Table Mst_CommodityGroup Add FAGUpdated " & TextType & "(30) , FAGUpdatedDate datetime"
'Call TmpTable
'
'
'
'UpdateConfig ("108")
'MsgBox "Version 108 Installed successfully"
'
'
'End Function
'
'
'
'Public Function Version107()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Mst_StorageRateCardDepoDetail ("
'tmp = tmp & " DepoSRCDID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FromBagSize decimal(18, 0) ,"
'tmp = tmp & " ToBagSize decimal(18, 0),"
'tmp = tmp & " RatePerBag decimal(18, 4), "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Concurrency DateTime ,"
'   tmp = tmp & " PRIMARY KEY (DepoSRCDID)) "
'Else
'   tmp = tmp & " Concurrency DateTime ) "
'End If
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Mst_SpecialRateCardDepoDetail ("
'tmp = tmp & " SPRCDDID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SPRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SRCDID decimal(18, 0) ,"
'tmp = tmp & " FromBagSize decimal(18, 0) ,"
'tmp = tmp & " ToBagSize decimal(18, 0),"
'tmp = tmp & " RatePerBag decimal(18, 4), "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Concurrency DateTime ,"
'   tmp = tmp & " PRIMARY KEY (SPRCDID)) "
'Else
'   tmp = tmp & " Concurrency DateTime ) "
'End If
'
'Call TmpTable
'
'
'tmp = " Alter Table Mst_StorageRateCard Add " & c & " DepoBillingCycleID Decimal(18,0) "
'Call TmpTable
'
'
'tmp = " Alter Table Mst_SpecialRateCard Add " & c & " DepoBillingCycleID Decimal(18,0) "
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_StorageRateCardDepoDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_StorageRateCardDepoDetail PRIMARY KEY  CLUSTERED (DepoSRCDID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StorageRateCardDepoDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_StorageRateCardDepoDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StorageRateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_StorageRateCard_Mst_BillingCycleDepo FOREIGN KEY (DepoBillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCard_Mst_BillingCycleDepo FOREIGN KEY (DepoBillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID)"
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDepoDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SpecialRateCardDepoDetail PRIMARY KEY  CLUSTERED (SPRCDDID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDepoDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SpecialRateCardDepoDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDepoDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCardDepoDetail_Mst_SpecialRateCard FOREIGN KEY (SPRCID) REFERENCES Mst_SpecialRateCard    (SPRCID)"
'   Call TmpTable
'
'End If
'
'tmp = " Update Mst_SpecialRateCard Set Flag = 'A'"
'Call TmpTable
'
'tmp = " Update Mst_StorageRateCard Set DepoBillingCycleID = BillingCycleID "
'Call TmpTable
'
'tmp = "Insert Into Mst_StorageRateCardDepoDetail Select * from Mst_StorageRateCardDetail "
'Call TmpTable
'
'tmp = " Update Mst_StorageRateCardDepoDetail Set Concurrency = '" & Now & "'"
'Call TmpTable
'
'
'tmp = " Update Mst_SpecialRateCard Set DepoBillingCycleID = BillingCycleID "
'Call TmpTable
'
'tmp = "Insert Into Mst_SpecialRateCardDepoDetail Select * from Mst_SpecialRateCardDetail "
'Call TmpTable
'
'tmp = " Update Mst_SpecialRateCardDepoDetail Set Concurrency = '" & Now & "'"
'Call TmpTable
'
''tmp = "Alter Table Txn_InvMonGSLDetail Add " & c & " DateOfWithdrawal datetime"
''Call TmpTable
'
'UpdateConfig ("107")
'MsgBox "Version 107 Installed successfully"
'
'End Function
'
'
'Public Function Version106()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "DROP TABLE Txn_MonthwiseSecLocDetail"
'Call TmpTable
'
'tmp = " DROP TABLE Txn_MonthwiseSecLoc"
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_MonthwiseSecLoc ("
'tmp = tmp & " MonthwiseTxnID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthwiseTxnDate datetime NULL ,"
'tmp = tmp & " MonthYear datetime NULL ,"
'tmp = tmp & " SecurityAgencyID decimal(18, 0) NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " G_UID   " & TextType & "   (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy   " & TextType & "   (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL,"
'   tmp = " PRIMARY KEY (MonthwiseTxnID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL)"
'End If
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_MonthwiseSecLocDetail ("
'tmp = tmp & " MonthwiseTxnDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthwiseTxnID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SecLocMapDetailID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityTypeID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityDesignationID decimal(18, 0) NULL ,"
'tmp = tmp & " StartDate datetime NULL ,"
'tmp = tmp & " EndDate datetime NULL ,"
'tmp = tmp & " SalarypaidUpto datetime NULL ,"
'tmp = tmp & " NoofGodown decimal(18, 0) NULL ,"
'tmp = tmp & " RatePerMonth decimal(18, 0) NULL ,"
'tmp = tmp & " CMPMID decimal(18, 0) NULL ,"
'tmp = tmp & " AttendanceSheetID   " & TextType & "   (50)  NULL ,"
'tmp = tmp & " NoofDays decimal(18, 0) NULL ,"
'tmp = tmp & " EstimatedAmt decimal(18, 2) NULL ,"
'tmp = tmp & " EstimatedSTax decimal(18, 2) NULL ,"
'tmp = tmp & " TotEstimatedAmt decimal(18, 2) NULL ,"
'tmp = tmp & " BillingNoOfDays decimal(18, 0) NULL ,"
'tmp = tmp & " BillingAmount decimal(18, 2) NULL ,"
'tmp = tmp & " BillingSTax decimal(18, 2) NULL ,"
'tmp = tmp & " BillingTotalAmount decimal(18, 2) NULL ,"
'tmp = tmp & " BillingID decimal(18, 0) NULL ,"
'tmp = tmp & " Remarks   " & TextType & "   (1000)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy   " & TextType & "   (30)  NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " CreatedDate datetime NULL,"
'   tmp = " PRIMARY KEY (MonthwiseTxnDetailID))"
'Else
'   tmp = tmp & " CreatedDate datetime NULL)"
'End If
'Call TmpTable
'
''SecurityBilling....!
'tmp = "CREATE TABLE Txn_SecurityBilling ("
'tmp = tmp & " BillTxnID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BillTxnDate datetime NULL,"
'tmp = tmp & " MonthYear datetime NULL ,"
'tmp = tmp & " SecurityAgencyID decimal(18, 0) NULL,"
'tmp = tmp & " STaxID decimal(18, 0) NULL ,"
'tmp = tmp & " BillNo " & TextType & " (100) NULL ,"
'tmp = tmp & " BillDate datetime NULL ,"
'tmp = tmp & " BillAmt decimal(18, 2) NULL ,"
'tmp = tmp & " BillSTax decimal(18, 2) NULL ,"
'tmp = tmp & " TotalAmt decimal(18, 2) NULL ,"
'tmp = tmp & " ForwardTo decimal(18, 0) NULL ,"
'tmp = tmp & " Remarks   " & TextType & " (1000) NULL ,"
'tmp = tmp & " FAGRemarks   " & TextType & " (1000) NULL ,"
'tmp = tmp & " ApprovalDate   DateTime NULL ,"
'tmp = tmp & " Status   " & TextType & "   (1)  NULL ,"
'tmp = tmp & " G_UID   " & TextType & "   (36)  NULL ,"
'tmp = tmp & " BatchID decimal(18, 0) NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy   " & TextType & "   (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL,"
'   tmp = " PRIMARY KEY (BillTxnID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL)"
'End If
'Call TmpTable
'
'tmp = "CREATE TABLE Txn_SecurityBillingDetail ("
'tmp = tmp & " BillTxnDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BillTxnID decimal(18, 0) NULL ,"
'tmp = tmp & " MonthwiseTxnDetailID decimal(18, 0) NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityTypeID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityDesignationID decimal(18, 0) NULL ,"
'tmp = tmp & " NoofGodown decimal(18, 0) NULL ,"
'tmp = tmp & " RatePerMonth decimal(18, 0) NULL ,"
'tmp = tmp & " CMPMID decimal(18, 0) NULL ,"
'tmp = tmp & " AttendanceSheetID   " & TextType & "   (50)  NULL ,"
'tmp = tmp & " NoofDays decimal(18, 0) NULL ,"
'tmp = tmp & " EstimatedAmt decimal(18, 2) NULL ,"
'tmp = tmp & " EstimatedSTax decimal(18, 2) NULL ,"
'tmp = tmp & " TotEstimatedAmt decimal(18, 2) NULL ,"
'tmp = tmp & " BillDetailAmt decimal(18, 2) NULL ,"
'tmp = tmp & " BillDetailSTax decimal(18, 2) NULL ,"
'tmp = tmp & " BillNoOfDays decimal(18, 0) NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy   " & TextType & "   (30)  NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " CreatedDate datetime NULL,"
'   tmp = " PRIMARY KEY (BillTxnDetailID))"
'Else
'   tmp = tmp & " CreatedDate datetime NULL)"
'End If
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Txn_MonthwiseSecLoc ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_MonthwiseSecLoc_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_MonthwiseSecLoc PRIMARY KEY  CLUSTERED (MonthwiseTxnID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_MonthwiseSecLocDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_MonthwiseSecLocDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_MonthwiseSecLocDetail PRIMARY KEY  CLUSTERED (MonthwiseTxnDetailID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_MonthwiseSecLoc ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLoc_Mst_SecurityAgency FOREIGN KEY"
'   tmp = tmp & " (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLoc_Mst_Location FOREIGN KEY"
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_MonthwiseSecLocDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (CMPMID) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_SecLocMappingDetail1 FOREIGN KEY"
'   tmp = tmp & " (SecLocMapDetailID) REFERENCES Mst_SecLocMappingDetail (TxnDetailID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_SecurityDesignation FOREIGN KEY"
'   tmp = tmp & " (SecurityDesignationID) REFERENCES Mst_SecurityDesignation (SecurityDesignationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_SecurityShifts FOREIGN KEY"
'   tmp = tmp & " (SecurityTypeID) REFERENCES Mst_SecurityShifts (SecurityShiftsID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Txn_MonthwiseSecLoc FOREIGN KEY"
'   tmp = tmp & " (MonthwiseTxnID) REFERENCES Txn_MonthwiseSecLoc (MonthwiseTxnID)"
'   Call TmpTable
'
''SecurityBilling....!
'   tmp = "ALTER TABLE Txn_SecurityBilling ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_SecurityBilling_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_SecurityBilling PRIMARY KEY  CLUSTERED (BillTxnID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_SecurityBillingDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_SecurityBillingDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_SecurityBillingDetail PRIMARY KEY  CLUSTERED (BillTxnDetailID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_SecurityBilling ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBilling_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (ForwardTo) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBilling_Mst_SecurityAgency FOREIGN KEY"
'   tmp = tmp & " (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBilling_Mst_ServiceTax FOREIGN KEY"
'   tmp = tmp & " (STaxID) REFERENCES Mst_ServiceTax (ServiceTaxID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_SecurityBillingDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBillingDetail_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (CMPMID) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBillingDetail_Mst_Location FOREIGN KEY"
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBillingDetail_Mst_SecurityDesignation FOREIGN KEY"
'   tmp = tmp & " (SecurityDesignationID) REFERENCES Mst_SecurityDesignation (SecurityDesignationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBillingDetail_Mst_SecurityShifts FOREIGN KEY"
'   tmp = tmp & " (SecurityTypeID) REFERENCES Mst_SecurityShifts (SecurityShiftsID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBillingDetail_Txn_MonthwiseSecLocDetail FOREIGN KEY"
'   tmp = tmp & " (MonthwiseTxnDetailID) REFERENCES Txn_MonthwiseSecLocDetail (MonthwiseTxnDetailID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SecurityBillingDetail_Txn_SecurityBilling FOREIGN KEY"
'   tmp = tmp & " (BillTxnID) REFERENCES Txn_SecurityBilling (BillTxnID)"
'   Call TmpTable
'End If
'
'
'tmp = "ALTER TABLE Mst_SecLocMappingDetail ADD " & c & "  SalaryPaidUpto datetime "
'Call TmpTable
'
'
'tmp = "Update Mst_SecLocMappingDetail Set SalaryPaidUpto = '2011/05/31'  Where SalaryPaidUpto is null "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy)"
'tmp = tmp & " VALUES(200, 'FrmTxn_SecurityBilling.frm ', 'Security Billing', 'FrmTxn_SecurityBilling', 'Y', 'T',NULL,NULL,NULL,NULL)"
'Call TmpTable
'
'UpdateConfig ("106")
'MsgBox "Version 106 Installed successfully"
'
'End Function
'
'
'Public Function Version105()
'
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_Designation Add " & c & " SecReportingDesignationID Decimal(18,0)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'
'   tmp = "ALTER TABLE Mst_Designation ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_Designation_Mst_DesignationForSecReporting FOREIGN KEY (SecReportingDesignationID) REFERENCES Mst_Designation (DesignationID)"
'   Call TmpTable
'
'End If
'
'
'UpdateConfig ("105")
'MsgBox "Version 105 Installed successfully"
'
'
'End Function
'
'Public Function Version104()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " CREATE TABLE Txn_MonthwiseSecLoc ("
'tmp = tmp & " MonthwiseTxnID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthwiseTxnDate datetime NOT NULL ,"
'tmp = tmp & " MonthYear datetime NOT NULL ,"
'tmp = tmp & " SecurityAgencyID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " STaxID decimal(18, 0) NULL ,"
'tmp = tmp & " BillNo decimal(18, 0) NULL ,"
'tmp = tmp & " BillDate datetime NULL ,"
'tmp = tmp & " BillAmt decimal(18, 2) NULL ,"
'tmp = tmp & " BillSTax decimal(18, 2) NULL ,"
'tmp = tmp & " TotalAmt decimal(18, 2) NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime NULL ,"
'   tmp = " PRIMARY KEY (MonthwiseTxnID))"
'Else
'   tmp = tmp & " UpdatedDate datetime NULL)"
'End If
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_MonthwiseSecLocDetail ("
'tmp = tmp & " MonthwiseTxnDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthwiseTxnID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SecLocMapDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityTypeID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityDesignationID decimal(18, 0) NULL ,"
'tmp = tmp & " StartDate datetime NULL ,"
'tmp = tmp & " EndDate datetime NULL ,"
'tmp = tmp & " NoofGodown decimal(18, 0) NULL ,"
'tmp = tmp & " RatePerMonth decimal(18, 0) NULL ,"
'tmp = tmp & " CMPMID decimal(18, 0) NULL ,"
'tmp = tmp & " AttendanceSheetID " & TextType & " (50)  NULL ,"
'tmp = tmp & " NoofDays decimal(18, 0) NULL ,"
'tmp = tmp & " EstimatedAmt decimal(18, 2) NULL ,"
'tmp = tmp & " BillAmount decimal(18, 2) NULL ,"
'tmp = tmp & " BillSTax decimal(18, 2) NULL ,"
'tmp = tmp & " Remarks " & TextType & " (1000)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " CreatedDate datetime NULL,"
'   tmp = " PRIMARY KEY (MonthwiseTxnDetailID))"
'Else
'   tmp = tmp & " CreatedDate datetime NULL)"
'End If
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "Alter table Mst_SecLocMappingDetail alter column txndetailid decimal(18,0) not null"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecLocMappingDetail ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_SecLocMappingDetail PRIMARY KEY  CLUSTERED (TxnDetailID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthwiseSecLoc ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_MonthwiseSecLoc_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_MonthwiseSecLoc PRIMARY KEY  CLUSTERED (MonthwiseTxnID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthwiseSecLocDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_MonthwiseSecLocDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_MonthwiseSecLocDetail PRIMARY KEY  CLUSTERED (MonthwiseTxnDetailID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthwiseSecLoc ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLoc_Mst_SecurityAgency FOREIGN KEY"
'   tmp = tmp & " (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLoc_Mst_ServiceTax FOREIGN KEY"
'   tmp = tmp & " (STaxID) REFERENCES Mst_ServiceTax (ServiceTaxID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthwiseSecLocDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Txn_MonthwiseSecLoc FOREIGN KEY"
'   tmp = tmp & " (MonthwiseTxnID) REFERENCES Txn_MonthwiseSecLoc (MonthwiseTxnID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (CMPMID) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_Location FOREIGN KEY"
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_SecLocMappingDetail1 FOREIGN KEY"
'   tmp = tmp & " (SecLocMapDetailID) REFERENCES Mst_SecLocMappingDetail (TxnDetailID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_SecurityDesignation FOREIGN KEY"
'   tmp = tmp & " (SecurityDesignationID) REFERENCES Mst_SecurityDesignation (SecurityDesignationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthwiseSecLocDetail_Mst_SecurityShifts FOREIGN KEY"
'   tmp = tmp & " (SecurityTypeID) REFERENCES Mst_SecurityShifts (SecurityShiftsID)"
'   Call TmpTable
'End If
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(199, 'FrmTxn_MonthwiseSecLoc.frm', 'Monthwise Security Location', 'FrmTxn_MonthwiseSecLoc', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("104")
'MsgBox "Version 104 Installed successfully"
'
'End Function
'
'
'Public Function Version103()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Mst_SecurityDesignation ("
'tmp = tmp & " SecurityDesignationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SecurityDesignation " & TextType & " (100) NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (10),"
'tmp = tmp & " Remarks " & TextType & " (250),"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL,"
'   tmp = " PRIMARY KEY (SecurityDesignationID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL)"
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_SecLocMapping ("
'tmp = tmp & " TxnID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " TxnDate datetime NULL ,"
'tmp = tmp & " SecurityAgencyID decimal(18, 0) NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36) NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30) NULL,"
'   tmp = " PRIMARY KEY (TxnID))"
'Else
'    tmp = tmp & " UpdatedBy " & TextType & " (30) NULL)"
'End If
'Call TmpTable
'
'tmp = " CREATE TABLE Mst_SecLocMappingDetail ("
'tmp = tmp & " TxnDetailID decimal(18, 0) NULL ,"
'tmp = tmp & " TxnID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityTypeID decimal(18, 0) NULL ,"
'tmp = tmp & " SecurityDesignationID Decimal(18,0) ,"
'tmp = tmp & " StartDate datetime NULL ,"
'tmp = tmp & " EndDate datetime NULL ,"
'tmp = tmp & " NoofGodown decimal(18, 0) NULL ,"
'tmp = tmp & " RatePerMonth decimal(18, 2) NULL ,"
'tmp = tmp & " Remarks " & TextType & " (2000),"
'tmp = tmp & " Concurrency datetime NOT NULL,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL,"
'tmp = tmp & " CreatedDate datetime NULL)"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'    tmp = "ALTER TABLE Mst_SecLocMapping ADD"
'    tmp = tmp & " CONSTRAINT DF_Mst_SecLocMapping_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'    tmp = tmp & " CONSTRAINT PK_Mst_SecLocMapping PRIMARY KEY  CLUSTERED (TxnID)"
'    Call TmpTable
'
'    tmp = "ALTER TABLE Mst_SecLocMappingDetail ADD"
'    tmp = tmp & " CONSTRAINT DF_Mst_SecLocMappingDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'    Call TmpTable
'
'
'    tmp = "ALTER TABLE Mst_SecurityDesignation ADD"
'    tmp = tmp & " CONSTRAINT DF_Mst_SecurityDesignation_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'    tmp = tmp & " CONSTRAINT PK_Mst_SecurityDesignation PRIMARY KEY  CLUSTERED (SecurityDesignationID)"
'    Call TmpTable
'
'    tmp = "ALTER TABLE Mst_SecLocMapping ADD"
'    tmp = tmp & " CONSTRAINT FK_Mst_SecLocMapping_Mst_Location FOREIGN KEY"
'    tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
'    tmp = tmp & " CONSTRAINT FK_Mst_SecLocMapping_Mst_SecurityAgency FOREIGN KEY"
'    tmp = tmp & " (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID)"
'    Call TmpTable
'
'    tmp = "ALTER TABLE Mst_SecLocMappingDetail ADD"
'    tmp = tmp & " CONSTRAINT FK_Mst_SecLocMappingDetail_Mst_SecurityDesignation FOREIGN KEY"
'    tmp = tmp & " (SecurityDesignationID) REFERENCES Mst_SecurityDesignation (SecurityDesignationID)"
'    Call TmpTable
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(197, 'FrmMst_SecurityDesignation.frm', 'Security Designation', 'FrmMst_SecurityDesignation', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(198, 'FrmMst_SecurityLocationMapping.frm', 'Security Location Mapping', 'FrmMst_SecurityLocationMapping', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'
'UpdateConfig ("103")
'MsgBox "Version 103 Installed successfully"
'
'End Function
'
'Public Function Version102()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter table Mst_Designation add " & c & "  MISFlag " & TextType & " (1)"
'
'Call TmpTable
'
'UpdateConfig ("102")
'MsgBox "Version 102 Installed successfully"
'
'End Function
'Public Function Version101()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_ProInvoice Add " & c & " Remarks " & TextType & "(2000), " & c & " Flag " & TextType & "(10) "
'Call TmpTable
'
'tmp = "Update Txn_ProInvoice Set Flag = 'C' Where BatchID <> 0 "
'Call TmpTable
'
'tmp = "Update Txn_ProInvoice Set Flag = 'P' Where BatchID = 0 "
'Call TmpTable
'
'
'UpdateConfig ("101")
'MsgBox "Version 101 Installed successfully"
'
'End Function
'Public Function Version100()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " CREATE TABLE Txn_MonthEmp_History ("
'tmp = tmp & " MonthEmpID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " TxnDate datetime ,"
'tmp = tmp & " EmployeeID decimal(18, 0),"
'tmp = tmp & " EmployeeNo " & TextType & "(25),"
'tmp = tmp & " EmployeeName " & TextType & "(100),"
'tmp = tmp & " DesignationID decimal(18, 0),"
'tmp = tmp & " SelfReporting " & TextType & "(1) ,"
'tmp = tmp & " DirectReportingTo decimal(18, 0),"
'tmp = tmp & " ReportingToDRC decimal(18, 0),"
'tmp = tmp & " ConnectedToRAG decimal(18, 0),"
'tmp = tmp & " EmployeeEmailID " & TextType & "(100),"
'tmp = tmp & " MobileNo " & TextType & "(50),"
'tmp = tmp & " Address " & TextType & "(250),"
'tmp = tmp & " Flag " & TextType & "(1),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime ,"
'tmp = tmp & " UpdatedBy " & TextType & "(30),"
'tmp = tmp & " NAVCode " & TextType & "(25),"
'tmp = tmp & " AllocateExpense " & TextType & "(1),"
'tmp = tmp & " CMPAllocation " & TextType & "(1),"
'tmp = tmp & " BaseLocationID decimal(18, 0),"
'tmp = tmp & " CTC decimal(18, 2),"
'tmp = tmp & " GrossPay decimal(18, 2),"
'tmp = tmp & " DOL datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " DOJ datetime,"
'   tmp = " PRIMARY KEY (MonthEmpID))"
'Else
'   tmp = tmp & " DOJ datetime)"
'End If
'Call TmpTable
'
'
'tmp = "Create table Txn_MonthEmp_HistoryDetail("
'tmp = tmp & " MonthEmpDetailID decimal(18,0) Not Null,"
'tmp = tmp & " MonthEmpID decimal(18,0),"
'tmp = tmp & " LocationID decimal(18,0),"
'tmp = tmp & " AllocatedCTC decimal(18,2),"
'tmp = tmp & " AllocatedGrossPay decimal(18,2),"
'tmp = tmp & " Flag " & TextType & "(1),"
'tmp = tmp & " Remark " & TextType & "(255),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = " PRIMARY KEY (MonthEmpDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & "(30))"
'End If
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Mst_IncomeExpenseType ("
'tmp = tmp & " IETypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " IEType " & TextType & "(250),"
'tmp = tmp & " IEFlag " & TextType & " (50),"
'tmp = tmp & " Remark " & TextType & " (1000),"
'tmp = tmp & " Flag " & TextType & " (50),"
'tmp = tmp & " G_UID " & TextType & " (36),"
'tmp = tmp & " Concurrency datetime ,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30),"
'   tmp = " PRIMARY KEY (IETypeID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30))"
'End If
'Call TmpTable
'
'tmp = "Alter Table Txn_CxTF_BOE Add " & c & " ExchangeTypeID decimal(18,0)"
'Call TmpTable
'
'tmp = "Update Txn_CxTF_BOE Set ExchangeTypeID=1 Where ExchangeTypeID is Null"
'Call TmpTable
'
'tmp = " Alter table Txn_OtherIncome ADD " & c & " IETypeID decimal(18,0)"
'Call TmpTable
'
'tmp = " Alter table Txn_OtherIncome alter Column Amount decimal(18,2)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_MonthEmp_History ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_MonthEmp_History_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_MonthEmp_History PRIMARY KEY  CLUSTERED (MonthEmpID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthEmp_History ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthEmp_History_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (EmployeeID) REFERENCES Mst_Employee (EmployeeID)"
''   tmp = tmp & " CONSTRAINT FK_Txn_MonthEmp_History_Mst_Location FOREIGN KEY"
''   tmp = tmp & " (BaseLocationID) REFERENCES Mst_Location (LocationID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthEmp_HistoryDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_MonthEmp_HistoryDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_MonthEmp_HistoryDetail PRIMARY KEY  CLUSTERED (MonthEmpDetailID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_MonthEmp_HistoryDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthEmp_HistoryDetail_Mst_Location FOREIGN KEY"
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_MonthEmp_HistoryDetail_Txn_MonthEmp_HistoryDetail FOREIGN KEY"
'   tmp = tmp & " (MonthEmpID) REFERENCES Txn_MonthEmp_History (MonthEmpID)"
'
'   Call TmpTable
'
'   tmp = " Alter Table Txn_CxTF_BOE Add Constraint FK_Txn_CxTF_BOE_Mst_ExchangeType FOREIGN KEY"
'   tmp = tmp & " (ExchangeTypeID) REFERENCES Mst_ExchangeType (ExchangeTypeID)"
'
'   Call TmpTable
'
'
'   tmp = " ALTER TABLE Mst_IncomeExpenseType ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_IncomeExpenseType_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Mst_IncomeExpenseType PRIMARY KEY  CLUSTERED (IETypeID)"
'   Call TmpTable
'
'   tmp = " Alter table Txn_OtherIncome Add CONSTRAINT FK_Txn_OtherIncome_Mst_IncomeExpenseType FOREIGN KEY"
'   tmp = tmp & " (IETypeID) REFERENCES Mst_IncomeExpenseType(IETypeID)"
'   Call TmpTable
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(194, 'FrmTxn_EmpHistory.frm', 'Monthwise Employee History', 'FrmTxn_EmpHistory', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(195, 'Frm_MonthEmpProces.frm', 'Monthwise Employee Process', 'Frm_MonthEmpProces', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(196, 'FrmMst_IncomeExpenseType.frm', 'Income Expense Type', 'FrmMst_IncomeExpenseType', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("100")
'MsgBox "Version 100 Installed successfully"
'
'End Function
'
'Public Function Version99()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " CREATE TABLE Txn_OtherIncome ("
'tmp = tmp & " OtherIncomeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " TxnDate datetime NOT NULL ,"
'tmp = tmp & " MonthYear " & TextType & " (50)  NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " ClientName " & TextType & " (200)  NULL ,"
'tmp = tmp & " Amount decimal(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (50)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'If Gen_DBType = "MDB" Then
'tmp = tmp & " CreatedDate datetime NULL"
'   tmp = " PRIMARY KEY (OtherIncomeID))"
'Else
'tmp = tmp & " CreatedDate datetime NULL)"
'End If
'
'Call TmpTable
'
'tmp = "ALTER Table Mst_ExpenseType ADD MISFlag " & TextType & "(1), MISExpenseGroup " & TextType & "(250)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_OtherIncome WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_OtherIncome PRIMARY KEY  CLUSTERED (OtherIncomeID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_OtherIncome ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_OtherIncome_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'End If
'
'tmp = "Update mst_ExpenseType Set MISFlag='Y' "
'Call TmpTable
'
'tmp = "Update Mst_ExpenseType Set MISExpenseGroup='Other' "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(191, 'Frm_ImportOtherIncome.frm', 'Import Other Income', 'Frm_ImportOtherIncome', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(192, 'FrmRpt_NewIncomeMIS.frm', 'AM wise New Income MIS', 'FrmRpt_NewIncomeMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(193, 'FrmRpt_CompFumgation.frm', 'Comparative Fumigation', 'FrmRpt_CompFumgation', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("99")
'MsgBox "Version 99 Installed successfully"
'
'End Function
'
'
'Public Function Version98()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " ALTER TABLE Txn_Audit ADD " & c & " Remarks " & TextType & "(1000)"
'Call TmpTable
'
'
'tmp = " ALTER TABLE Mst_Location ADD " & c & " AM_EmployeeID Decimal(18,0)"
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_Location ADD"
'   tmp = tmp & " CONSTRAINT FK_Mst_Location_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (AM_EmployeeID) REFERENCES Mst_Employee (EmployeeID)"
'   Call TmpTable
'End If
'
'
'UpdateConfig ("98")
'MsgBox "Version 98 Installed successfully"
'
'End Function
'Public Function Version97()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Txn_CITFCR ("
'tmp = tmp & " CITFCRID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " TxnDate datetime NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL,"
'   tmp = " PRIMARY KEY (CITFCRID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL)"
'End If
'Call TmpTable
'
'
'tmp = " CREATE TABLE Txn_CITFCRDetail ("
'tmp = tmp & " CITFCRDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CITFCRID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (50)  NULL ,"
'tmp = tmp & " CITFNo decimal(18, 0) NULL ,"
'tmp = tmp & " CRID decimal(18, 0) NULL ,"
'tmp = tmp & " Remarks " & TextType & " (1000)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL,"
'   tmp = " PRIMARY KEY (CITFCRDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL)"
'End If
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_CITFCR ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_CITF_CR_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_CITF_CR PRIMARY KEY  CLUSTERED (CITFCRID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_CITFCR ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_CITFCR_Mst_Location FOREIGN KEY"
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_CITFCRDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_CITFCRDetail_Concurrency DEFAULT (getdate()) FOR Concurrency,"
'   tmp = tmp & " CONSTRAINT PK_Txn_CITFCRDetail PRIMARY KEY  CLUSTERED (CITFCRDetailID)"
'   Call TmpTable
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(190, 'FrmTxn_CITF_CR.frm', 'CITF Cash Receipt Mapping', 'FrmTxn_CITF_CR', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("97")
'MsgBox "Version 97 Installed successfully"
'
'End Function
'
'
'Public Function Version96()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_Spillage Alter Column Remarks " & TextType & "(1000)"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(189, 'FrmRpt_UnpaidGodownRent.frm', 'Unpaid Godown Rent', 'FrmRpt_UnpaidGodownRent', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("96")
'MsgBox "Version 96 Installed successfully"
'
'End Function
'
'Public Function Version95()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Txn_Hologram ("
'tmp = tmp & " HologramID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " HologramDate datetime NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " FromHologramNo decimal(18, 0) NULL ,"
'tmp = tmp & " ToHologramNo decimal(18, 0) NULL ,"
'tmp = tmp & " TotalQty decimal(18, 0) NULL ,"
'tmp = tmp & " Remark " & TextType & " (1000)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime, "
'   tmp = " PRIMARY KEY (HologramID ))"
'Else
'   tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Txn_Hologram WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_Hologram  PRIMARY KEY  CLUSTERED(HologramID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_Hologram ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_Hologram_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_Hologram  ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_Hologram_Mst_Location FOREIGN KEY  (LocationID) REFERENCES Mst_Location (LocationID)"
'
'   Call TmpTable
'
'   tmp = "Alter Table Mst_hologram drop constraint PK_Mst_Hologram"
'   Call TmpTable
'
'End If
'
'
'tmp = "Alter Table Mst_hologram alter Column HologramNo " & TextType & "(25)"
'Call TmpTable
'
'
'tmp = "Alter Table Mst_Hologram Add " & c & " NewHologramNo decimal(18,0)"
'Call TmpTable
'
'tmp = "Select * from Mst_Hologram"
'
'Call TmpTable
'If TmpRs.RecordCount > 0 Then
'
'   tmp = "Update Mst_Hologram Set NewHologramNo = 0 Where Isnumeric(HologramNo)= 0 And NewHologramNo is Null"
'   Call Tmp1Table
'
'   tmp = "Select * from  Mst_Hologram Where HologramNo like '-%'  And Isnumeric(HologramNo)= 1 And NewHologramNo is Null"
'   Call Tmp1Table
'
'   For r = 0 To TmpRs1.RecordCount - 1
'       TmpRs1!NewHologramNo = -Val(TmpRs1!HologramNo)
'       TmpRs1.Update
'       TmpRs1.MoveNext
'   Next
'
'   tmp = "Update Mst_Hologram Set NewHologramNo = 0  Where HologramNo like '%.%' And NewHologramNo is Null"
'   Call Tmp1Table
'
'   tmp = "Select * from  Mst_Hologram Where (HologramNo like '0%' or HologramNo like ' %')  And NewHologramNo is Null And Isnumeric(HologramNo)= 1"
'   Call Tmp1Table
'   For r = 0 To TmpRs1.RecordCount - 1
'       TmpRs1!NewHologramNo = Val(Trim(TmpRs1!HologramNo))
'       TmpRs1.Update
'       TmpRs1.MoveNext
'   Next
'   tmp = "Select * from  Mst_Hologram Where NewHologramNo is Null And Isnumeric(HologramNo)= 1"
'   Call Tmp1Table
'   For r = 0 To TmpRs1.RecordCount - 1
'       TmpRs1!NewHologramNo = Val(Trim(TmpRs1!HologramNo))
'       TmpRs1.Update
'       TmpRs1.MoveNext
'   Next
'End If
'
'
'tmp = "ALTER Table Txn_VendorClaimDetail ADD " & c & " ServiceTax decimal(18,2), " & c & " TotalAmount decimal(18,2) "
'Call TmpTable
'
'tmp = "UPDATE Txn_VendorClaimDetail SET TotalAmount = Amount, ServiceTax = 0 "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(188, 'FrmTxn_Hologram.frm', 'Hologram', 'FrmTxn_Hologram', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("95")
'MsgBox "Version 95 Installed successfully"
'
'End Function
'
'
'
'
'Public Function Version94()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_VendorClaim Add  InvoiceNo " & TextType & "(300)"
'Call TmpTable
'
'
'UpdateConfig ("94")
'MsgBox "Version 94 Installed successfully"
'
'End Function
'
'Public Function Version93()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(186, 'FrmRpt_NRejectionStorageAmount.frm', 'NCDEX Rejection Storage Amount', 'FrmRpt_NRejectionStorageAmount', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(187, 'FrmRpt_LocationWiseClient.frm', 'Location Wise Client List', 'FrmRpt_LocationWiseClient', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("93")
'MsgBox "Version 93 Installed successfully"
'
'End Function
'
'Public Function Version92()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " CREATE TABLE Txn_CxTF_BOE ("
'tmp = tmp & " BOEID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5) NULL ,"
'tmp = tmp & " CxTFNo decimal(18, 0) NULL ,"
'tmp = tmp & " CxTFDate datetime NULL ,"
'tmp = tmp & " ClientIDRow decimal(18, 0) NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NULL ,"
'tmp = tmp & " WithdrawalBags decimal(18, 0) NULL ,"
'tmp = tmp & " WithdrawalQty decimal(18, 3) NULL ,"
'tmp = tmp & " BOEAmount decimal(18, 2) NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36) NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime, "
'   tmp = " PRIMARY KEY (BOEID))"
'Else
'   tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'    tmp = " ALTER TABLE Txn_CxTF_BOE WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_CxTF_BOE PRIMARY KEY  CLUSTERED (BOEID)"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_CxTF_BOE ADD"
'    tmp = tmp & " CONSTRAINT DF_Txn_CxTF_BOE_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_CxTF_BOE ADD"
'    tmp = tmp & " CONSTRAINT FK_Txn_CxTF_BOE_Mst_Commodity FOREIGN KEY"
'    tmp = tmp & " (CommodityID) REFERENCES Mst_Commodity (CommodityID),"
'    tmp = tmp & " CONSTRAINT FK_Txn_CxTF_BOE_Mst_Location FOREIGN KEY"
'    tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID)"
'
'    Call TmpTable
'
'End If
'
'tmp = "Alter Table Mst_Designation Add " & c & " DirectLink " & TextType & "(1) "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(185, 'FrmRpt_CITFBOE.frm', 'CITF BOE', 'FrmRpt_CITFBOE', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("92")
'MsgBox "Version 92 Installed successfully"
'
'End Function
'
'
'
'Public Function Version91()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter Table Txn_CashReceipt Add " & c & " ShowDP Decimal(18,0) "
'Call TmpTable
'
'tmp = "Alter Table Txn_CashReceiptInvoiceDetails Add " & c & " DPID " & TextType & "(50), " & c & "  NCCharges Decimal(18,2) "
'Call TmpTable
'
'UpdateConfig ("91")
'MsgBox "Version 91 Installed successfully"
'
'End Function
'
'Public Function Version90()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_CommodityQualityParameter ("
'tmp = tmp & " CommodityQualityParameterID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " QualityParameterID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MinValue decimal(18, 0) NULL ,"
'tmp = tmp & " MaxValue decimal(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (36), "
'   tmp = " PRIMARY KEY (CommodityQualityParameterID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (36)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_QualityParameter ("
'tmp = tmp & " QualityParameterID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " QualityParameter " & TextType & " (500)  NULL ,"
'tmp = tmp & " FumigationApplicable " & TextType & " (1)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (36), "
'   tmp = " PRIMARY KEY (QualityParameterID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (36)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_TestingAgency ("
'tmp = tmp & " AgencyID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " AgencyName " & TextType & " (500)  NULL ,"
'tmp = tmp & " Status " & TextType & " (1)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (36), "
'   tmp = " PRIMARY KEY (AgencyID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (36)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_QualityInformation ("
'tmp = tmp & " QualityInformationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Date datetime NULL ,"
'tmp = tmp & " EmployeeID decimal(18, 0) NULL ,"
'tmp = tmp & " GodownID decimal(18, 0) NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NULL ,"
'tmp = tmp & " AgencyID decimal(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (36), "
'   tmp = " PRIMARY KEY (QualityInformationID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (36)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_QualityInformationDetail ("
'tmp = tmp & " QualityInformationDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " QualityInformationID decimal(18, 0) NULL ,"
'tmp = tmp & " GSLID decimal(18, 0) NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5) Not NULL ,"
'tmp = tmp & " QualityParameterID decimal(18, 0) NULL ,"
'tmp = tmp & " TestValue decimal(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (36), "
'   tmp = " PRIMARY KEY (QualityInformationDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (36)) "
'End If
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_CommodityQualityParameter WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_CommodityQualityParameter PRIMARY KEY  CLUSTERED"
'   tmp = tmp & " (CommodityQualityParameterID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_QualityParameter WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_QualityParameter PRIMARY KEY  CLUSTERED"
'   tmp = tmp & " (QualityParameterID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_TestingAgency WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_TestingAgency PRIMARY KEY  CLUSTERED (AgencyID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_QualityInformation WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_QualityInformation PRIMARY KEY  CLUSTERED (QualityInformationID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_QualityInformationDetail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_QualityInformationDetail PRIMARY KEY  CLUSTERED"
'   tmp = tmp & " (QualityInformationDetailID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_CommodityQualityParameter ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_CommodityQualityParameter_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_QualityParameter ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_QualityParameter_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_TestingAgency ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_TestingAgency_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_QualityInformation ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_QualityInformation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_QualityInformationDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_QualityInformationDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_CommodityQualityParameter ADD"
'   tmp = tmp & " CONSTRAINT FK_Mst_CommodityQualityParameter_Mst_Commodity FOREIGN KEY"
'   tmp = tmp & " (CommodityID) REFERENCES Mst_Commodity (CommodityID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_CommodityQualityParameter_Mst_QualityParameter FOREIGN KEY"
'   tmp = tmp & " (QualityParameterID) REFERENCES Mst_QualityParameter (QualityParameterID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_QualityInformation ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformation_Mst_Commodity FOREIGN KEY"
'   tmp = tmp & " (CommodityID) REFERENCES Mst_Commodity (CommodityID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformation_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (EmployeeID) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformation_Mst_Godown FOREIGN KEY"
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformation_Mst_TestingAgency FOREIGN KEY"
'   tmp = tmp & " (AgencyID) REFERENCES Mst_TestingAgency (AgencyID)"
'   Call TmpTable
'
'
'   tmp = " ALTER TABLE Txn_QualityInformationDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformationDetail_Mst_QualityParameter FOREIGN KEY"
'   tmp = tmp & " (QualityParameterID) REFERENCES Mst_QualityParameter (QualityParameterID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformationDetail_Txn_QualityInformation FOREIGN KEY"
'   tmp = tmp & " (QualityInformationID) REFERENCES Txn_QualityInformation (QualityInformationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_QualityInformationDetail_Mst_GSL FOREIGN KEY"
'   tmp = tmp & " (GSLID,SM_Prefix) REFERENCES Mst_GSL (GSLID,SM_Prefix)"
'
'   Call TmpTable
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(181, 'FrmMst_QualityParameter.frm', 'Quality Parameter', 'FrmMst_QualityParameter', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(182, 'FrmMst_CommodityQualityParameter.frm', 'Commodity Quality Parameter', 'FrmMst_CommodityQualityParameter', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(183, 'FrmMst_TestingAgency.frm', 'Testing Agency', 'FrmMst_TestingAgency', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(184, 'FrmTxn_QualityInformation.frm', 'Quality Information', 'FrmTxn_QualityInformation', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("90")
'MsgBox "Version 90 Installed successfully"
'
'End Function
'
'Public Function Version89()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Txn_RAGMemberTransfer ("
'tmp = tmp & " TransferID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " TransferDate datetime NOT NULL ,"
'tmp = tmp & " FromEmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ToEmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Remark " & TextType & " (2000)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedBy " & TextType & " (36)  ,"
''tmp = tmp & " UpdatedDate datetime)"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime, "
'   tmp = " PRIMARY KEY (TransferID))"
'Else
'   tmp = tmp & " UpdatedDate datetime) "
'End If
'
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Txn_RAGMemberTransferDetail ("
'tmp = tmp & " RAGMemberTransferDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " TransferID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " EmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  ,"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedBy " & TextType & " (36)  ,"
''tmp = tmp & " UpdatedDate datetime)"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime, "
'   tmp = " PRIMARY KEY (RAGMemberTransferDetailID))"
'Else
'   tmp = tmp & " UpdatedDate datetime) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Txn_RAGMemberTransfer WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_TransferToRAG PRIMARY KEY  CLUSTERED  (TransferID)  "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_RAGMemberTransfer WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_TransferToRAG_ConCurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_RAGMemberTransferDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_RAGMemberTransferDetail PRIMARY KEY  CLUSTERED (RAGMemberTransferDetailID)  "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_RAGMemberTransferDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_RAGMemberTransferDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_RAGMemberTransferDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_RAGMemberTransferDetail_Txn_RAGMemberTransfer FOREIGN KEY  "
'   tmp = tmp & " (TransferID) REFERENCES Txn_RAGMemberTransfer (TransferID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_RAGMemberTransferDetail_Mst_Employee FOREIGN KEY  "
'   tmp = tmp & " (EmployeeID) REFERENCES Mst_Employee (EmployeeID)"
'
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(180, 'FrmTxn_RAGMemberTransfer.frm', 'RAG Member Transfer', 'FrmTxn_RAGMemberTransfer', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("89")
'MsgBox "Version 89 Installed successfully"
'
'End Function
'
'Public Function Version88()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Update Txn_ProInvoice Set Discount = 0, Premium = 0 Where Discount is null Or Premium is null "
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD CONSTRAINT DF_Txn_ProInvoice_Discount DEFAULT (0) FOR "
'   tmp = tmp & " Discount "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD CONSTRAINT DF_Txn_ProInvoice_Premium DEFAULT (0) FOR "
'   tmp = tmp & " Premium "
'   Call TmpTable
'
'End If
'
'tmp = "Alter table Txn_InvMonGSLDetail add " & c & " OldRate Decimal(18,4)," & c & " OldBillingCycleID Decimal(18,0), " & c & "  OldBillingUnit " & TextType & " (1)"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(179, 'FrmRpt_RAGGodownListing.frm', 'RAG Godown Listing', 'FrmRpt_RAGGodownListing', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("88")
'MsgBox "Version 88 Installed successfully"
'
'End Function
'
'
'
'Public Function Version87()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Update Mst_Commodity set FumigationFlag = 'Y',EffectiveDays = 45 "
'Call TmpTable
'
'UpdateConfig ("87")
'MsgBox "Version 87 Installed successfully"
'
'End Function
'
'Public Function Version86()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'tmp = "CREATE TABLE Txn_Fumigation ("
'tmp = tmp & " FumigationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FumigationDate datetime ,"
'tmp = tmp & " NextFumigationDate datetime,"
'tmp = tmp & " VendorID decimal(18, 0),"
'tmp = tmp & " GodownID decimal(18, 0),"
'tmp = tmp & " Remark " & TextType & " (255),"
'tmp = tmp & " Rate decimal(18, 2),"
'tmp = tmp & " TotalQty decimal(18, 3),"
'tmp = tmp & " BillableQty decimal(18, 3),"
'tmp = tmp & " Amount decimal(18, 2),"
'tmp = tmp & " BillNo " & TextType & " (50),"
'tmp = tmp & " G_UID " & TextType & " (36),"
'tmp = tmp & " Flag " & TextType & " (1),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (FumigationID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_FumigationDetail ("
'tmp = tmp & " FumgDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FumigationID decimal(18, 0),"
'tmp = tmp & " StackNo " & TextType & " (10),"
'tmp = tmp & " CommodityID decimal(18, 0),"
'tmp = tmp & " NoOfBags decimal(18, 0),"
'tmp = tmp & " BalanceQty decimal(18, 3),"
'tmp = tmp & " ActualNoOfBags decimal(18, 0),"
'tmp = tmp & " ActualQty decimal(18, 3),"
'tmp = tmp & " BillableFlag " & TextType & " (1),"
'tmp = tmp & " LastFumgDate datetime,"
'tmp = tmp & " Remark " & TextType & " (100),"
'tmp = tmp & " G_UID " & TextType & " (36),"
'tmp = tmp & " Concurrency datetime ,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (FumgDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
'tmp = "Alter Table Mst_Commodity Add " & c & " FumigationFlag varchar(1)," & c & " EffectiveDays decimal(18,0)"
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_Spraying ("
'tmp = tmp & " SprayingID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SprayingDate datetime NULL ,"
'tmp = tmp & " VendorID decimal(18, 0) NULL ,"
'tmp = tmp & " GodownID decimal(18, 0) NULL ,"
'tmp = tmp & " Remark varchar (255)  NULL ,"
'tmp = tmp & " Rate decimal(18, 2) NULL ,"
'tmp = tmp & " NoOfBags decimal(18, 0) NULL ,"
'tmp = tmp & " BalanceQty decimal(18, 3) NULL ,"
'tmp = tmp & " ActualNoOfBags decimal(18, 0) NULL ,"
'tmp = tmp & " ActualQty decimal(18, 3) NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NULL ,"
'tmp = tmp & " G_UID varchar (36)  NULL ,"
'tmp = tmp & " Flag varchar (1)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy varchar (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy varchar (30)  NULL )"
'
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'    tmp = "ALTER TABLE Txn_Fumigation WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_Fumigation PRIMARY KEY  CLUSTERED"
'    tmp = tmp & " (FumigationID)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_Fumigation ADD"
'    tmp = tmp & " CONSTRAINT DF_Txn_Fumigation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_Fumigation ADD"
'    tmp = tmp & " CONSTRAINT FK_Txn_Fumigation_Mst_Godown FOREIGN KEY"
'    tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
'    tmp = tmp & " CONSTRAINT FK_Txn_Fumigation_Mst_Vendor FOREIGN KEY"
'    tmp = tmp & " (VendorID) REFERENCES Mst_Vendor (VendorID)"
'
'    Call TmpTable
'
'
'    tmp = "ALTER TABLE Txn_FumigationDetail WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_FumigationDetail PRIMARY KEY  CLUSTERED (FumgDetailID)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_FumigationDetail ADD"
'    tmp = tmp & " CONSTRAINT DF_Txn_FumigationDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_FumigationDetail ADD"
'    tmp = tmp & " CONSTRAINT FK_Txn_FumigationDetail_Mst_Commodity FOREIGN KEY"
'    tmp = tmp & " (CommodityID) REFERENCES Mst_Commodity (CommodityID),"
'    tmp = tmp & " CONSTRAINT FK_Txn_FumigationDetail_Txn_Fumigation FOREIGN KEY"
'    tmp = tmp & " (FumigationID) REFERENCES Txn_Fumigation (FumigationID)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_Spraying WITH NOCHECK ADD "
'    tmp = tmp & " CONSTRAINT PK_Txn_Spraying PRIMARY KEY  CLUSTERED "
'    tmp = tmp & " (SprayingID)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_Spraying ADD "
'    tmp = tmp & " CONSTRAINT DF_Txn_Spraying_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_Spraying ADD "
'    tmp = tmp & " CONSTRAINT FK_Txn_Spraying_Mst_Godown FOREIGN KEY "
'    tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
'    tmp = tmp & " CONSTRAINT FK_Txn_Spraying_Mst_Vendor FOREIGN KEY "
'    tmp = tmp & " (VendorID) REFERENCES Mst_Vendor (VendorID)"
'
'    Call TmpTable
'
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(177, 'FrmTxn_Fumigation.frm', 'Fumigation', 'FrmTxn_Fumigation', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(178, 'FrmTxn_Spraying.frm', 'Spraying', 'FrmTxn_Spraying', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("86")
'MsgBox "Version 86 Installed successfully"
'
'End Function
'
'
'Public Function Version85()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "ALTER TABLE Mst_SpecialRateCard ALTER COLUMN Amount Decimal(18,4)"
'Call TmpTable
'
'
'UpdateConfig ("85")
'MsgBox "Version 85 Installed successfully"
'
'End Function
'
'
'Public Function Version84()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Txn_DayWiseGSLDetail ("
'tmp = tmp & " GSLID numeric(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL ,"
'tmp = tmp & " CMPID numeric(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityID numeric(18, 0) NULL ,"
'tmp = tmp & " GodownID numeric(18, 0) NULL ,"
'tmp = tmp & " StackNo " & TextType & " (10)  NOT NULL ,"
'tmp = tmp & " LotNo " & TextType & " (10)  NOT NULL ,"
'tmp = tmp & " BalanceBags numeric(18, 0) NULL ,"
'tmp = tmp & " BalanceWeight numeric(18, 3) NULL ,"
'tmp = tmp & " ExchangeTypeID numeric(18, 0) NULL ,"
'tmp = tmp & " ClientIDRow numeric(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " Flag1 " & TextType & " (1)  NULL ,"
'tmp = tmp & " GrossWt numeric(18, 3) NULL ,"
'tmp = tmp & " CxTFNO numeric(18, 0) NULL ,"
'tmp = tmp & " TFlag " & TextType & " (1)  NULL ,"
'tmp = tmp & " CDTFNo numeric(18, 0) NULL , "
'tmp = tmp & " CDTFDate datetime NULL , "
'tmp = tmp & " CxTFDate datetime NULL ) "
'
'Call TmpTable
'
'UpdateConfig ("84")
'MsgBox "Version 84 Installed successfully"
'
'End Function
'
'Public Function Version83()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter table Mst_Designation add " & c & " SelfReporting " & TextType & "(1), " & c & " ReportingDesignationID Decimal(18,0) "
'Call TmpTable
'
'
'tmp = "Alter table Mst_InsuranceDetails add " & c & " Flag " & TextType & "(1)"
'Call TmpTable
'
'
'tmp = "Update Mst_InsuranceDetails Set Flag='A'"
'Call TmpTable
'
'tmp = "Update Mst_InsuranceCoMaster Set Flag='A'"
'Call TmpTable
'
'
'tmp = "Update Mst_MenuDetail Set MenuName = 'Base Storage Rate Card' Where MenuID = 62 "
'Call TmpTable
'
'tmp = "Update Mst_MenuDetail Set MenuName = 'Special Storage Rate Card' Where MenuID = 151 "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(174, 'Frm_ITVendorNAVExport.frm', 'Export ITVendor Expense Claim for Navision', 'Frm_ITVendorNAVExport', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(175, 'FrmRpt_EmpAllocation.frm', 'Employee Allocation', 'FrmRpt_EmpAllocation', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(176, 'FrmRpt_Fumigation.frm', 'Fumigation Report', 'FrmRpt_Fumigation', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("83")
'MsgBox "Version 83 Installed successfully"
'
'End Function
'Public Function Version82()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE Mst_NCDEXStorageRateCard("
'tmp = tmp & " NSRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) ,"
'tmp = tmp & " CommodityGroupID decimal(18, 0) ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (NSRCID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_NCDEXStorageRateCardDetail("
'tmp = tmp & " NSRCDID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " NSRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " GodownID decimal(18, 0) ,"
'tmp = tmp & " Rate decimal(18, 4), "
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (NSRCDID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_FF_RateCard ("
'tmp = tmp & " FF_RateCardID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID Decimal(18, 0)  NOT NULL ,"
'tmp = tmp & " CommodityGroupID Decimal(18, 0)  NOT NULL ,"
'tmp = tmp & " LessorID Decimal(18, 0)  NOT NULL  ,"
'tmp = tmp & " ExchangeTypeID Decimal(18, 0)  NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (36)  NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime NULL,"
'   tmp = tmp & " PRIMARY KEY (FF_RateCardID))"
'Else
'   tmp = tmp & " UpdatedDate datetime NULL)"
'End If
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_FF_RateCard_Godown ("
'tmp = tmp & " FF_RateCard_GodownID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FF_RateCardID Decimal(18, 0) NULL ,"
'tmp = tmp & " CMPID Decimal(18, 0) NULL ,"
'tmp = tmp & " GodownID Decimal(18, 0) NULL ,"
'tmp = tmp & " BillingCycleID Decimal(18, 0) NULL ,"
'tmp = tmp & " BillingUnit " & TextType & " (1)  NULL ,"
'tmp = tmp & " RateOnQty Decimal(18, 4) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (36)  NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime NULL,"
'   tmp = tmp & " PRIMARY KEY (FF_RateCard_GodownID))"
'Else
'   tmp = tmp & " UpdatedDate datetime NULL)"
'End If
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_FF_RateCard_Godown_Bag ("
'tmp = tmp & " FF_RateCard_Godown_BagID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FF_RateCard_GodownID Decimal(18, 0) NULL ,"
'tmp = tmp & " BagSizeFrom Decimal(18, 0) NULL ,"
'tmp = tmp & " BagSizeTo Decimal(18, 0) NULL ,"
'tmp = tmp & " RatePerBag Decimal(18, 4) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (36)  NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime NULL,"
'   tmp = tmp & " PRIMARY KEY (FF_RateCard_Godown_BagID))"
'Else
'   tmp = tmp & " UpdatedDate datetime NULL)"
'End If
'Call TmpTable
'
'
'tmp = "Alter table Txn_VendorClaimDetail add " & c & " Quantity Decimal(18,0)"
'Call TmpTable
'
'tmp = "Alter table Txn_ClaimSubmissionDetail  add  " & c & " Quantity Decimal(18,0)"
'Call TmpTable
'
'tmp = "Alter table Txn_ITVendorClaimDetail  add " & c & " Quantity Decimal(18,0), " & c & " InvoiceNo Varchar(100)"
'Call TmpTable
'
'tmp = "Alter Table Txn_InvMonGSLDetail Add FranchiseeRate " & c & "  decimal(18,4)"
'Call TmpTable
'
'tmp = "ALTER TABLE Txn_InvMonGSLDetail ALTER COLUMN  Rate Decimal(18,4)"
'Call TmpTable
'
'tmp = "Alter Table Mst_Employee Add " & c & " BaseLocationID Decimal(18,0) "
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_SpecialRateCard ALTER COLUMN  RateOnQty Decimal(18,4)"
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_SpecialRateCardDetail ALTER COLUMN RatePerBag Decimal(18,4)"
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_StorageRateCard ALTER COLUMN  RateOnQty Decimal(18,4)"
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_StorageRateCardDetail ALTER COLUMN RatePerBag Decimal(18,4)"
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Mst_NCDEXStorageRateCard WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_NCDEXStorageRateCard PRIMARY KEY  CLUSTERED (NSRCID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_NCDEXStorageRateCard ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_NCDEXStorageRateCard_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_NCDEXStorageRateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_NCDEXStorageRateCard_Mst_Location FOREIGN KEY "
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_NCDEXStorageRateCard_Mst_CommodityGroupID FOREIGN KEY "
'   tmp = tmp & " (CommodityGroupID) REFERENCES Mst_CommodityGroup (CommodityGroupID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_NCDEXStorageRateCardDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_NCDEXStorageRateCardDetail PRIMARY KEY  CLUSTERED (NSRCDID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_NCDEXStorageRateCardDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_NCDEXStorageRateCardDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_NCDEXStorageRateCardDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_NCDEXStorageRateCardDetail_Mst_Godown FOREIGN KEY "
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_NCDEXStorageRateCardDetail_Mst_NCDEXStorageRateCard FOREIGN KEY "
'   tmp = tmp & " (NSRCID) REFERENCES Mst_NCDEXStorageRateCard (NSRCID)"
'
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_FF_RateCard WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_FF_RateCard PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (FF_RateCardID)  "
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_FF_RateCard_Godown WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_FF_RateCard_Godown PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (FF_RateCard_GodownID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard_Godown_Bag WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_FF_RateCard_Godown_Bag PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (FF_RateCard_Godown_BagID)  "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_FF_RateCard_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard_Godown ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_FF_RateCard_Godown_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard_Godown_Bag ADD "
'   tmp = tmp & " CONSTRAINT DF_FF_RateCard_Godown_Bag_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Mst_CommodityGroup FOREIGN KEY "
'   tmp = tmp & " (CommodityGroupID) REFERENCES Mst_CommodityGroup (CommodityGroupID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Mst_Lessor FOREIGN KEY "
'   tmp = tmp & " (LessorID) REFERENCES Mst_Lessor (LessorID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Mst_Location FOREIGN KEY "
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID) ,"
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Mst_ExchangeType FOREIGN KEY "
'   tmp = tmp & " (ExchangeTypeID) REFERENCES Mst_ExchangeType (ExchangeTypeID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard_Godown ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Godown_Mst_BillingCycle FOREIGN KEY "
'   tmp = tmp & " (BillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Godown_Mst_CMP FOREIGN KEY "
'   tmp = tmp & " (CMPID) REFERENCES Mst_CMP (CMPID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Godown_Mst_FF_RateCard FOREIGN KEY "
'   tmp = tmp & " (FF_RateCardID) REFERENCES Mst_FF_RateCard (FF_RateCardID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Godown_Mst_Godown FOREIGN KEY "
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_FF_RateCard_Godown_Bag ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_FF_RateCard_Godown_Bag_Mst_FF_RateCard_Godown FOREIGN KEY "
'   tmp = tmp & " (FF_RateCard_GodownID) REFERENCES Mst_FF_RateCard_Godown (FF_RateCard_GodownID)"
'   Call TmpTable
'
'
'   tmp = " ALTER TABLE Mst_Employee ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_Employee_Mst_Location FOREIGN KEY "
'   tmp = tmp & " (BaseLocationID) REFERENCES Mst_Location (LocationID) "
'   Call TmpTable
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(169, 'FrmRpt_AMwiseIncomeMIS.frm', 'AM wise Income MIS', 'FrmRpt_AMwiseIncomeMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(170, 'FrmMst_NCDEXStorageRateCard.frm', 'NCDEX Storage Rate Card', 'FrmMst_NCDEXStorageRateCard', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(171, 'FrmMst_FranchiseeGodownRateCard.frm', 'Franchisee Godown Rate Card', 'FrmMst_FranchiseeGodownRateCard', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(172, 'FrmRpt_DuplicateGodownRent.frm', 'Duplicate Godown Rent', 'FrmRpt_DuplicateGodownRent ', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(173, 'FrmTxn_GodownTransfer.frm', 'Godown Transfer', 'FrmTxn_GodownTransfer', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("82")
'MsgBox "Version 82 Installed successfully"
'
'End Function
'
'Public Function Version81()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(168, 'Frm_GodownRentCoverLetter.frm', 'Godown Rent CoverLetter', 'Frm_GodownRentCoverLetter', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("81")
'MsgBox "Version 81 Installed successfully"
'
'End Function
'
'Public Function Version80()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_GodownRentDetail Add " & c & " Arrears Decimal (18,2) "
'Call TmpTable
'
'tmp = " Update Txn_GodownRentDetail Set Arrears = 0 "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(167,'FrmRpt_RateCard.frm','Rate Card','FrmRpt_RateCard', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("80")
'MsgBox "Version 80 Installed successfully"
'
'End Function
'
'
'
'Public Function Version79()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Insert into Mst_GLAccType(GLAccTypeID,GLAccType,NAVCode,G_UID) Values (3,'N3C','XXXXXXXXXXXXX','0E895CA0759F4C9C8D7C7D97172AAAAA')"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(165, 'Frm_ExportNonNCDEXNonCorporateInvoice.frm', 'Export Non NCDEX/Non Corporate Invoice', 'Frm_ExportNonNCDEXNonCorporateInvoice', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(166,'FrmRpt_CITFDiscrepancy.frm','CITF Discrepancy','FrmRpt_CITFDiscrepancy', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("79")
'MsgBox "Version 79 Installed successfully"
'
'End Function
'
'Public Function Version78()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(164, 'FrmRpt_StockSummaryReport.frm', 'Stock Summary Report', 'FrmRpt_StockSummaryReport', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'UpdateConfig ("78")
'MsgBox "Version 78 Installed successfully"
'
'End Function
'
'Public Function Version77()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Alter Table Txn_ITVendorClaimDetail Add " & c & " AssetsID Decimal (18,0) "
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_ITVendorClaimDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ITVendorClaimDetail_Mst_AssetType FOREIGN KEY "
'   tmp = tmp & " (AssetsID) REFERENCES Mst_AssetsType (AssetsID)"
'
'   Call TmpTable
'
'End If
'
'UpdateConfig ("77")
'MsgBox "Version 77 Installed successfully"
'
'End Function
'
'Public Function Version76()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_ClientGroup ("
'tmp = tmp & " ClientGroupID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClientGroupName " & TextType & " (100) NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (ClientGroupID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'
'End If
'
'Call TmpTable
'
'
'tmp = " Alter Table Mst_Client Add  " & c & " ClientGroupID Decimal(18,0)"
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_ITVendor ("
'tmp = tmp & " VendorID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " VendorName " & TextType & " (100) NOT NULL ,"
'tmp = tmp & " VendorAddress " & TextType & " (255)  NULL ,"
'tmp = tmp & " LocationID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ContactNo " & TextType & "(100)  NULL ,"
'tmp = tmp & " MobileNo " & TextType & "(100) NULL ,"
'tmp = tmp & " Fax " & TextType & "(100)  NULL ,"
'tmp = tmp & " EmailID " & TextType & " (50)  NULL ,"
'tmp = tmp & " TDSID Decimal(18, 0) Not NULL ,"
'tmp = tmp & " NavCode " & TextType & " (25)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (VendorID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'
'End If
'
'Call TmpTable
'
'
'
'tmp = " CREATE TABLE Txn_ITVendorClaim ( "
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClaimDate datetime NOT NULL ,"
'tmp = tmp & " ExpenseFrom datetime NULL ,"
'tmp = tmp & " ExpenseTo datetime NULL ,"
'tmp = tmp & " VendorID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NULL ,"
'tmp = tmp & " Remarks " & TextType & " (200)  NOT NULL ,"
'tmp = tmp & " ApprovalRemarks " & TextType & " (200)  NULL ,"
'tmp = tmp & " ApprovalDate datetime NULL ,"
'tmp = tmp & " BatchID decimal(18, 0) NULL ,"
'tmp = tmp & " ForwardToID decimal(18, 0) NULL,"
'tmp = tmp & " Flag " & TextType & " (2)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NOT NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " UpdatedByFAG " & TextType & " (30)  NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDateFAG " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (ClaimID))"
'Else
'   tmp = tmp & " UpdatedDateFAG " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_ITVendorClaimDetail ("
'tmp = tmp & " ClaimDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExpenseTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Particulars " & TextType & " (250)  NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NOT NULL ,"
'tmp = tmp & " InvoiceDate datetime ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (ClaimDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_ClientGroup WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_ClientGroup PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (ClientGroupID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_ClientGroup ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_ClientGroup_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_Client ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_Client_Mst_ClientGroup FOREIGN KEY "
'   tmp = tmp & " (ClientGroupID) REFERENCES Mst_ClientGroup (ClientGroupID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_ITVendor WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_ITVendor PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (VendorID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_ITVendor ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_ITVendor_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = " ALTER TABLE Txn_ITVendorClaim WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_ITVendorClaim PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (ClaimID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ITVendorClaimDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_ITVendorClaimDetail PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (ClaimDetailID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ITVendorClaim WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_ITVendorClaim_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_ITVendorClaimDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_ITVendorClaimDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ITVendorClaim ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ITVendorClaim_Mst_ITVendor FOREIGN KEY "
'   tmp = tmp & " (VendorID) REFERENCES Mst_ITVendor (VendorID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_ITVendorClaim_Mst_Employee FOREIGN KEY "
'   tmp = tmp & " (ForwardToID) REFERENCES Mst_Employee (EmployeeID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ITVendorClaimDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ITVendorClaimDetail_Mst_Location FOREIGN KEY "
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_ITVendorClaimDetail_Mst_ExpenseType FOREIGN KEY "
'   tmp = tmp & " (ExpenseTypeID) REFERENCES Mst_ExpenseType (ExpenseTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_ITVendorClaimDetail_Txn_ITVendorClaim FOREIGN KEY "
'   tmp = tmp & " (ClaimID) REFERENCES Txn_ITVendorClaim (ClaimID)"
'   Call TmpTable
'
'End If
'
'tmp = "CREATE TABLE Log_ITVendorClaimFwdToHistory ("
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " OldForwardToID decimal(18, 0) NULL ,"
'tmp = tmp & " NewForwardToID decimal(18, 0) NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) NULL, "
'tmp = tmp & " Concurrency datetime NOT NULL CONSTRAINT DF_Log_ITVendorClaimFwdToHistory DEFAULT (getdate()))"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "SET QUOTED_IDENTIFIER ON "
'   Call TmpTable
'
'   tmp = "SET ANSI_NULLS ON "
'   Call TmpTable
'
'   tmp = "CREATE TRIGGER tr_Txn_ITVendorClaim_U ON Txn_ITVendorClaim AFTER "
'   tmp = tmp & " UPDATE AS DECLARE @mClaimID  decimal(18, 0),  @mOldForwardToID decimal(18, 0), @mNewForwardToID decimal(18, 0), @mUpdatedBy " & TextType & "(30) "
'   tmp = tmp & " SELECT  @mClaimID= d.ClaimID, @mOldForwardToID = d.ForwardToID, @mNewForwardToID = i.ForwardToID, @mUpdatedBy= i.UpdatedBy  FROM   "
'   tmp = tmp & " inserted i INNER JOIN deleted d ON i.ClaimID = d.ClaimID  "
'   tmp = tmp & " insert into Log_ITVendorClaimFwdToHistory (ClaimID ,OldForwardToID ,NewForwardToID ,UpdatedBy) values(@mClaimID ,@mOldForwardToID ,@mNewForwardToID ,@mUpdatedBy)"
'   Call TmpTable
'
'   tmp = "SET QUOTED_IDENTIFIER OFF "
'   Call TmpTable
'
'   tmp = "SET ANSI_NULLS ON "
'   Call TmpTable
'End If
'
'
'tmp = " Alter Table Txn_ProInvoice Add  " & c & " InvoiceDate DateTime,  " & c & " BillingMonth DateTime,  " & c & "  Discount Decimal(18,2)," & c & "  Premium Decimal(18,2)  "
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID]) "
'tmp = tmp & " VALUES(160,'FrmTxn_ProInvoice.frm','Proforma Invoice','FrmTxn_ProInvoice', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(161, 'FrmMst_ITVendor.frm', 'IT Vendor', 'FrmMst_ITVendor', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(162, 'FrmTxn_ITVendorClaim.frm', 'IT Vendor Claim', 'FrmTxn_ITVendorClaim', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(163, 'FrmMst_ClientGroup.frm', 'Client Group', 'FrmMst_ClientGroup', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'
'UpdateConfig ("76")
'MsgBox "Version 76 Installed successfully"
'
'
'End Function
'
'
'
'Public Function Version75()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Mst_SpecialRateCard Add  " & c & " Remarks " & TextType & "(500)"
'Call TmpTable
'
'UpdateConfig ("75")
'MsgBox "Version 75 Installed successfully"
'
'
'End Function
'
'Public Function Version74()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Txn_GodownRent Add  " & c & " BankName " & TextType & "(100)," & c & " ChequeNo " & TextType & "(100)," & c & " PaidAmount decimal(18,2)," & c & " TDS decimal(18,2)," & c & " PaidDate Datetime"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(159,'Frm_ImportGodClaimPayDetails.frm','Payment Detail Upload','Frm_ImportGodClaimPayDetails', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("74")
'MsgBox "Version 74 Installed successfully"
'
'
'End Function
'Public Function Version73()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Mst_FAGBank ( "
'tmp = tmp & " FAGBankId decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BankName " & TextType & "  (150) Not Null ,"
'tmp = tmp & " AccountNo " & TextType & "  (100) Not Null ,"
'tmp = tmp & " ContactPerson " & TextType & "  (150)  ,"
'tmp = tmp & " Address " & TextType & "  (200)  ,"
'tmp = tmp & " Flag " & TextType & "  (1)  ,"
'tmp = tmp & " Remarks " & TextType & "  (200)  ,"
'tmp = tmp & " G_UID " & TextType & "  (36)  ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & "  (30)  ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (FAGBankId))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'
'tmp = " CREATE TABLE Txn_CashReceiptDepositDetails ("
'tmp = tmp & " CRDepositID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FAGBankID decimal(18, 0) NULL ,"
'tmp = tmp & " DepositDate datetime NULL ,"
'tmp = tmp & " Remark " & TextType & " (250)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (CRDepositID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_ProInvoice("
'tmp = tmp & " InvoiceNo decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL  ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExchangeTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClientIDRow decimal(18, 0) NOT NULL ,"
'tmp = tmp & " NoOfBags decimal (18,2) Not Null ,"
'tmp = tmp & " Weight decimal(18, 2) NOT NULL ,"
'tmp = tmp & " Rate decimal (18,2) Not Null ,"
'tmp = tmp & " Amount decimal(18, 2) NOT NULL ,"
'tmp = tmp & " BatchID decimal(18, 0),"
'tmp = tmp & " ServiceTaxNAVCode " & TextType & " (30)  ,"
'tmp = tmp & " FAGPostingDate datetime ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " TFlag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (InvoiceNo,SM_Prefix))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = " Alter Table Txn_CashReceipt Add " & c & " CRDepositID Decimal(18,0) "
'
'Call TmpTable
'
'tmp = " Alter Table Mst_StorageContract Add " & c & " ExtendedPeriod Datetime"
'Call TmpTable
'
'tmp = " Alter Table  Txn_CxTF_TruckDetail Add " & c & " WBSlipDate DateTime"
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'    tmp = " ALTER TABLE Mst_FAGBank WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Mst_FAGBank PRIMARY KEY  CLUSTERED"
'    tmp = tmp & " (FAGBankID)"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Mst_FAGBank ADD"
'    tmp = tmp & " CONSTRAINT DF_Mst_FAGBank_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_CashReceiptDepositDetails WITH NOCHECK ADD"
'    tmp = tmp & " CONSTRAINT PK_Txn_CashReceiptDepositDetails PRIMARY KEY  CLUSTERED"
'    tmp = tmp & " (CRDepositID)"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_CashReceiptDepositDetails ADD"
'    tmp = tmp & " CONSTRAINT DF_Txn_CashReceiptDepositDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'    Call TmpTable
'
'    tmp = " ALTER TABLE Txn_CashReceiptDepositDetails ADD"
'    tmp = tmp & " CONSTRAINT FK_Txn_CashReceiptDepositDetails_Mst_FAGBank FOREIGN KEY (FAGBankID) REFERENCES Mst_FAGBank (FAGBankId)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_CashReceiptDepositDetails ADD CONSTRAINT FK_Txn_CashReceiptDepositDetails_Mst_Location FOREIGN KEY"
'    tmp = tmp & "(LocationID) REFERENCES Mst_Location (LocationID)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE Txn_CashReceipt ADD"
'    tmp = tmp & " CONSTRAINT FK_Txn_CashReceipt_Txn_CashReceiptDepositDetails FOREIGN KEY"
'    tmp = tmp & " (CRDepositID) REFERENCES Txn_CashReceiptDepositDetails (CRDepositID)"
'
'    Call TmpTable
'
'    tmp = "Alter table Txn_CashReceipt nocheck constraint FK_Txn_CashReceipt_Txn_CashReceiptDepositDetails"
'    Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD CONSTRAINT PK_Txn_ProInvoice PRIMARY KEY  CLUSTERED (InvoiceNo,SM_Prefix) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD CONSTRAINT DF_Txn_ProInvoice_Concurrency DEFAULT (getdate()) FOR "
'   tmp = tmp & " Concurrency "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ProInvoice WITH NOCHECK ADD CONSTRAINT DF_Txn_ProInvoice_BatchID DEFAULT (0) FOR "
'   tmp = tmp & " BatchID "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ProInvoice ADD CONSTRAINT FK_Txn_ProInvoice_Mst_Location FOREIGN KEY"
'   tmp = tmp & "(LocationID) REFERENCES Mst_Location (LocationID)"
'
'   Call TmpTable
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(154,'FrmMst_FAGBank.frm','FAG Bank','FrmMst_FAGBank', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(155,'FrmTxn_CRDepositeDetails.frm','CR Deposit Details','FrmTxn_CRDepositeDetails', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(156,'FrmTxn_StockOpBal.frm','Stock Opening Balance','FrmTxn_StockOpBal', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(157,'FrmTxn_InvoiceGenProcess.frm','Monthly Transaction Details','FrmTxn_InvoiceGenProcess', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(158,'FrmTxn_GenerateInvoice.frm','Generate Invoice','FrmTxn_GenerateInvoice', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("73")
'MsgBox "Version 73 Installed successfully"
'End Function
'
'Public Function Version72()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_ClaimSubmission Add " & c & " DesignationID Decimal(18,0) "
'Call TmpTable
'
'
'tmp = " Update Txn_ClaimSubmission Set Txn_ClaimSubmission.DesignationID= ME.DesignationID From Mst_Employee ME "
'tmp = tmp & " Inner Join Txn_ClaimSubmission On ME.EmployeeID=Txn_ClaimSubmission.EmployeeID "
'Call TmpTable
'
'
'tmp = "Alter Table Mst_Commodity Add " & c & "  ServiceTaxID Decimal(18,0) "
'Call TmpTable
'
'tmp = " Update Mst_ExpenseType Set Flag ='A' "
'Call TmpTable
'
'tmp = "Alter Table Mst_Employee Add " & c & "  CMPAllocation " & TextType & " (1) "
'Call TmpTable
'
'tmp = "Update Mst_Employee Set CMPAllocation = 1 "
'Call TmpTable
'
'tmp = "Alter Table Txn_InvMonGSLDetail Add  " & c & "   CDTFNo Decimal(18,0), " & c & "  CDTFDate DateTime, " & c & "  CxTFDate DateTime, " & c & "  PINo Decimal(18,0)"
'Call TmpTable
'
'tmp = "ALTER TABLE Txn_ClaimSubmission ADD CONSTRAINT FK_Txn_ClaimSubmission_Mst_Designation FOREIGN KEY"
'tmp = tmp & "(DesignationID) REFERENCES Mst_Designation (DesignationID)"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(153, 'FrmRpt_AuditorGodownMIS.frm', 'Auditor Godown MIS', 'FrmRpt_AuditorGodownMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("72")
'MsgBox "Version 72 Installed successfully"
'
'
'End Function
'
'
'Public Function Version71()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(152, 'FrmRpt_CMPBOMDetails.frm', 'Pending Month Begin Process', 'FrmRpt_CMPBOMDetails', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("71")
'MsgBox "Version 71 Installed successfully"
'
'
'End Function
'Public Function Version70()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " Alter Table Mst_GSL Add " & c & " NContractID Decimal(18,0) "
'Call TmpTable
'
'tmp = " Update Mst_GSL Set  NContractID = ContractID where ContractID is not null "
'Call TmpTable
'
''If Gen_DBType <> "MDB" Then
''
''   tmp = " ALTER TABLE Mst_GSL ADD "
''   tmp = tmp & " CONSTRAINT FK_Mst_GSL_Mst_StorageContractNContract FOREIGN KEY "
''   tmp = tmp & " (NContractID) REFERENCES Mst_StorageContract (ContractID)"
''   Call TmpTable
''
''End If
'
'UpdateConfig ("70")
'MsgBox "Version 70 Installed successfully"
'
'
'End Function
'Public Function Version69()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE Mst_SpecialRateCard("
'tmp = tmp & " SPRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SPRCDate datetime  ,"
'tmp = tmp & " SRCID decimal(18, 0) , "
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExchangeTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityGroupID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClientIDRow decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BillingCycleID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BillingUnit " & TextType & " (2)   NOT NULL ,"
'tmp = tmp & " RateOnQty decimal (18,2)  ,"
'tmp = tmp & " BaseRate " & TextType & " (2)   NOT NULL  ,"
'tmp = tmp & " AbsoluteDiscount " & TextType & " (35)   NOT NULL  ,"
'tmp = tmp & " Amount decimal(18, 2) NOT NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SRCID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'Call TmpTable
'
'
'tmp = " CREATE TABLE Mst_SpecialRateCardDetail ("
'tmp = tmp & " SPRCDID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SPRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SRCDID decimal(18, 0) ,"
'tmp = tmp & " FromBagSize decimal(18, 0) ,"
'tmp = tmp & " ToBagSize decimal(18, 0),"
'tmp = tmp & " RatePerBag decimal(18, 2), "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Concurrency DateTime ,"
'   tmp = tmp & " PRIMARY KEY (SPRCDID)) "
'Else
'   tmp = tmp & " Concurrency DateTime ) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_InvMonGSLDetail ("
'tmp = tmp & " GSLID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL ,"
'tmp = tmp & " CMPID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityID Decimal(18, 0) NULL ,"
'tmp = tmp & " GodownID Decimal(18, 0) NULL ,"
'tmp = tmp & " StackNo " & TextType & " (10)  NOT NULL ,"
'tmp = tmp & " LotNo " & TextType & " (10)  NOT NULL ,"
'tmp = tmp & " BalanceBags Decimal(18, 0) NULL ,"
'tmp = tmp & " BalanceWeight Decimal(18, 3) NULL ,"
'tmp = tmp & " ExchangeTypeID Decimal(18, 0) NULL ,"
'tmp = tmp & " ClientIDRow Decimal(18, 0) NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " Flag1 " & TextType & " (1)  NULL ,"
'tmp = tmp & " GrossWt Decimal(18, 3) NULL ,"
'tmp = tmp & " KRNo Decimal(18, 0) NULL ,"
'tmp = tmp & " KRNoBatchID Decimal(18, 0) NULL ,"
'tmp = tmp & " KRDate datetime NULL ,"
'tmp = tmp & " ContractID Decimal(18, 0) NULL ,"
'tmp = tmp & " BagSize Decimal(18, 0) NULL ,"
'tmp = tmp & " DRNo Decimal(18, 0) NULL ,"
'tmp = tmp & " DRDate datetime NULL ,"
'tmp = tmp & " YearMonth datetime NULL ,"
'tmp = tmp & " CxTFNO Decimal(18, 0) NULL ,"
'tmp = tmp & " BillingCycleID Decimal(18, 0) NULL ,"
'tmp = tmp & " BillingUnit " & TextType & " (1)  NULL ,"
'tmp = tmp & " InvoiceNo " & TextType & " (50)  NULL ,"
'tmp = tmp & " InvoiceDate datetime NULL ,"
'tmp = tmp & " InvoiceAmount Decimal(18, 2) NULL ,"
'tmp = tmp & " Rate Decimal(18, 2) NULL ,"
'tmp = tmp & " NoofDays Decimal(18, 0) NULL ,"
'tmp = tmp & " TFlag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL)"
'
'Call TmpTable
'
'
'
'
'
'tmp = "Alter Table Txn_GodownRent Add " & c & " CMPMID Decimal(18,0)"
'Call TmpTable
'
'tmp = "Alter Table Txn_GodownLessorAgreeDetail Add " & c & "  RentPaidUpto DateTime"
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_GodownRent ADD CONSTRAINT FK_Txn_GodownRent_Mst_Employee1 FOREIGN KEY"
'   tmp = tmp & "(CMPMID) REFERENCES Mst_Employee (EmployeeID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCard WITH NOCHECK ADD CONSTRAINT PK_Mst_SpecialRateCard PRIMARY KEY  CLUSTERED (SPRCID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCard WITH NOCHECK ADD CONSTRAINT DF_Mst_SpecialRateCard_Concurrency DEFAULT (getdate()) FOR "
'   tmp = tmp & " Concurrency "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCard_Mst_CommodityGroup FOREIGN KEY (CommodityGroupID) REFERENCES Mst_CommodityGroup    (CommodityGroupID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCard_Mst_ExchangeType FOREIGN KEY (ExchangeTypeID) REFERENCES Mst_ExchangeType(ExchangeTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCard_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCard_Mst_BillingCycle FOREIGN KEY (BillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SpecialRateCardDetail PRIMARY KEY  CLUSTERED (SPRCDID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SpecialRateCardDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCardDetail_Mst_SpecialRateCard FOREIGN KEY (SPRCID) REFERENCES Mst_SpecialRateCard    (SPRCID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_SpecialRateCard ADD"
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCard_Mst_StorageRateCard FOREIGN KEY    (SRCID) REFERENCES    Mst_StorageRateCard(SRCID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SpecialRateCardDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Mst_SpecialRateCardDetail_Mst_StorageRateCardDetail    FOREIGN KEY (SRCDID)    REFERENCES Mst_StorageRateCardDetail(SRCDID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_InvMonGSLDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_InvMonGSLDetail_BalanceBags DEFAULT (0) FOR BalanceBags,"
'   tmp = tmp & " CONSTRAINT DF_Txn_InvMonGSLDetail_BalanceWeight DEFAULT (0) FOR BalanceWeight,"
'   tmp = tmp & " CONSTRAINT DF_Txn_InvMonGSLDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_InvMonGSLDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_InvMonGSLDetail_Mst_BillingCycle FOREIGN KEY"
'   tmp = tmp & " (BillingCycleID) REFERENCES Mst_BillingCycle "
'   tmp = tmp & " (BillingCycleID),CONSTRAINT FK_Txn_InvMonGSLDetail_Mst_Godown FOREIGN KEY"
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID) ON UPDATE CASCADE ,"
'   tmp = tmp & " CONSTRAINT FK_Txn_InvMonGSLDetail_Mst_GSL FOREIGN KEY"
'   tmp = tmp & " (GSLID,SM_Prefix) REFERENCES Mst_GSL (GSLID,SM_Prefix)"
'   Call TmpTable
'
'   tmp = "Alter Table Txn_InvMonGSLDetail nocheck constraint FK_Txn_InvMonGSLDetail_Mst_BillingCycle"
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(150,'FrmTxn_CMPMonthEndProcess.frm','CMP Month End Process','FrmTxn_CMPMonthEndProcess', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(151,'FrmMst_SpecialRateCard.frm','Special Rate Card','FrmMst_SpecialRateCard', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'tmp = " Alter Table Txn_CxTF_GSL Add " & c & " BillFlag " & TextType & " (2) "
'Call TmpTable
'
'tmp = "Update Mst_CMP Set CMPLockedDate='2010-05-01 00:00:00' Where CMPLockedDate is null "
'Call TmpTable
'
'UpdateConfig ("69")
'MsgBox "Version 69 Installed successfully"
'
'End Function
'
'
'Public Function Version68()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_NewISINGSLRevalidation Add " & c & " DRNo " & TextType & "(50)," & c & " DRDate datetime, " & c & " OldQty decimal(18,3), " & c & " NewQty decimal(18,3)"
'Call TmpTable
'
'tmp = "Alter Table Mst_CMP Add " & c & " CMPLockedDate Datetime"
'Call TmpTable
'
'tmp = "Update Mst_CMP Set CMPLockedDate='2010-05-01 00:00:00'"
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(148, 'FrmRpt_ISINRevalidation.frm', 'Pending ISIN for Revalidation', 'FrmRpt_ISINRevalidation', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(149, 'FrmRpt_GodownRentClaim.frm', 'Godown Rent Claim Report', 'FrmRpt_GodownRentClaim', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'UpdateConfig ("68")
'MsgBox "Version 68 Installed successfully"
'
'End Function
'Public Function Version67()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_GLAccType Add " & c & " Flag " & TextType & "(2)"
'Call TmpTable
'
'
'tmp = " Create View Vw_ReservationContractStock as("
'tmp = tmp & " Select Mst_GSL.ContractID,Sum(BalanceBags)BalanceBags ,Sum(BalanceWeight)BalanceQty  From Mst_GSL"
'tmp = tmp & " inner join Mst_StorageContract on Mst_GSL.ContractID=Mst_StorageContract.ContractID"
'tmp = tmp & " Where (Mst_GSL.ContractID is not Null  AND Mst_GSL.ContractID <> '0') And (Mst_GSL.Flag='O' or Mst_GSL.Flag='J')"
'tmp = tmp & " Group By Mst_GSL.ContractID)"
'
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_NewISINRevalidation ("
'tmp = tmp & " NIRID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & "(5)  NOT NULL ,"
'tmp = tmp & " ARNo " & TextType & "(50)  NULL ,"
'tmp = tmp & " ARDate DateTime  NULL ,"
'tmp = tmp & " RevalDate datetime NULL ,"
'tmp = tmp & " Grade " & TextType & "(50) NULL ,"
'tmp = tmp & " OldISIN " & TextType & "(25) NULL ,"
'tmp = tmp & " ClientIDRow decimal(18, 0) NULL ,"
'tmp = tmp & " G_UID " & TextType & "(36) NULL ,"
'tmp = tmp & " Flag " & TextType & "(2) NULL ,"
'tmp = tmp & " Remark " & TextType & "(250) NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & "(30) NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (NIRID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'
'tmp = " CREATE TABLE Txn_NewISINGSLRevalidation ("
'tmp = tmp & " NIGRID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & "(5)  NOT NULL ,"
'tmp = tmp & " NIRID decimal(18, 0) NULL ,"
'tmp = tmp & " GSLID decimal(18, 0) NULL ,"
'tmp = tmp & " NewISIN " & TextType & "(25)  NULL ,"
'tmp = tmp & " BalanceBags decimal(18, 0) NULL ,"
'tmp = tmp & " BalanceWeight decimal(18, 3) NULL ,"
'tmp = tmp & " G_UID " & TextType & "(36)  NULL ,"
'tmp = tmp & " Flag " & TextType & "(2) NULL ,"
'tmp = tmp & " Remark " & TextType & "(250) NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & "(30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (NIGRID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Mst_StorageRateCardDetail ("
'tmp = tmp & " SRCDID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FromBagSize decimal(18, 0) ,"
'tmp = tmp & " ToBagSize decimal(18, 0),"
'tmp = tmp & " RatePerBag decimal(18, 2), "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Concurrency DateTime ,"
'   tmp = tmp & " PRIMARY KEY (SRCDID)) "
'Else
'   tmp = tmp & " Concurrency DateTime ) "
'End If
'
'
'Call TmpTable
'
'
'
'tmp = "ALTER TABLE Mst_StorageRateCard Add " & c & " BillingCycleID decimal(18,0), " & c & " RateOnQty decimal(18,2)," & c & " DefaultRate decimal(18,0) "
'Call TmpTable
'
'tmp = " Alter Table Mst_StorageRateCard alter column FromBagSize Decimal(18,0) "
'Call TmpTable
'
'tmp = " Alter Table Mst_StorageRateCard alter column ToBagSize Decimal(18,0) "
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_NewISINRevalidation WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_NewISINRevalidation PRIMARY KEY  CLUSTERED (NIRID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_NewISINRevalidation ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_NewISINRevalidation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_NewISINRevalidation ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_NewISINRevalidation_Mst_ISIN FOREIGN KEY "
'   tmp = tmp & " (OldISIN) REFERENCES Mst_ISIN (ISINID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_NewISINGSLRevalidation WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_NewISINGSLRev PRIMARY KEY  CLUSTERED (NIGRID)"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_NewISINGSLRevalidation ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_NewISINGSLRev_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_NewISINGSLRevalidation ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_NewISINGSLRevalidation_Mst_GSL FOREIGN KEY (GSLID,SM_Prefix) REFERENCES Mst_GSL (GSLID,SM_Prefix),"
'   tmp = tmp & " CONSTRAINT FK_Txn_NewISINGSLRevalidation_Txn_NewISINRevalidation FOREIGN KEY (NIRID) REFERENCES Txn_NewISINRevalidation (NIRID)"
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_StorageRateCardDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_StorageRateCardDetail PRIMARY KEY  CLUSTERED (SRCDID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StorageRateCardDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_StorageRateCardDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StorageRateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_StorageRateCard_Mst_BillingCycle FOREIGN KEY (BillingCycleID) REFERENCES Mst_BillingCycle (BillingCycleID)"
'
'   Call TmpTable
'
'End If
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(145, 'FrmRpt_ReservationContractStock.frm', 'Reservation Contract Stock', 'FrmRpt_ReservationContractStock', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(146, 'Frm_NewISINRevalidation.frm', 'New ISIN Revalidation', 'Frm_NewISINRevalidation', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(147, 'FrmRpt_VendorClaimDetails.frm', 'Vendor Claim Details', 'FrmRpt_VendorClaimDetails', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("67")
'MsgBox "Version 67 Installed successfully"
'
'End Function
'
'Public Function Version66()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter Table Mst_DPInvoiceUpload Add " & c & " InvoiceNo " & TextType & "(25)," & c & " InvoiceDate datetime," & c & " InvSTAmt decimal(18,2)," & c & " InvoiceAmt decimal(18,2) "
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(141, 'Frm_DPInvoiceCoverLetter.frm', 'DPInvoice Cover Letter', 'Frm_DPInvoiceCoverLetter', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(142, 'Frm_UpdateDPInvoiceNo.frm', 'Update DP Invoice No', 'Frm_UpdateDPInvoiceNo', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(143, 'FrmRpt_GodownWithoutAgreement.frm', 'Godown Without Agreement', 'FrmRpt_GodownWithoutAgreement', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(144, 'FrmRpt_NCDEXUtilization.frm', 'NCDEX Utilization', 'FrmRpt_NCDEXUtilization', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'
'UpdateConfig ("66")
'MsgBox "Version 66 Installed successfully"
'
'End Function
'
'
'Public Function Version65()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_DPInvoiceUpload ("
'tmp = tmp & " DPInvoiceID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthYear datetime ,"
'tmp = tmp & " LocationID decimal(18, 0) ,"
'tmp = tmp & " DPNavCode " & TextType & " (25) ,"
'tmp = tmp & " Amt decimal(18, 2),"
'tmp = tmp & " STAmt decimal(18, 2) ,"
'tmp = tmp & " TotalAmt decimal(18, 2) ,"
'tmp = tmp & " PINo " & TextType & " (50)  ,"
'tmp = tmp & " BatchID decimal(18,0) ,"
'tmp = tmp & " Flag " & TextType & " (1) ,"
'tmp = tmp & " Concurrency datetime ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (DPInvoiceID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_GLAccType ("
'tmp = tmp & " GLAccTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " GLAccType " & TextType & " (100),"
'tmp = tmp & " NAVCode " & TextType & "(25),"
'tmp = tmp & " G_UID " & TextType & " (36),"
'tmp = tmp & " Concurrency datetime,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime,"
'tmp = tmp & " UpdatedDate datetime,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (GLAccTypeID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_DPInvoiceUpload WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_DPInvoiceUpload PRIMARY KEY  CLUSTERED (DPInvoiceID)   "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_DPInvoiceUpload ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_DPInvoiceUpload_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_GLAccType WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_GLAccType PRIMARY KEY  CLUSTERED (GLAccTypeID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_GLAccType ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_GLAccType_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_GodownRent ADD"
'   tmp = tmp & "   CONSTRAINT DF_Txn_GodownRent_BatchID DEFAULT (0) FOR BatchID"
'   Call TmpTable
'
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(138, 'Frm_UploadDPInvoice.frm', 'Import DP Invoice', 'Frm_UploadDPInvoice', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(139, 'Frm_ExportDPInvoice.frm', 'Export DP Invoice', 'Frm_ExportDPInvoice', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(140, 'FrmMst_GLAccType.frm', 'G/L Account Type', 'FrmMst_GLAccType', 'Y', 'A',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "Insert into Mst_GLAccType(GLAccTypeID,GLAccType,NAVCode,G_UID) Values (1,'DP','5020401000000','0E895CA0759F4C9C8D7C7D97172SAGAR')"
'Call TmpTable
'
'tmp = "Insert into Mst_GLAccType(GLAccTypeID,GLAccType,NAVCode,G_UID) Values (2,'DPSTCode','STWS','0E895CA0759F4C9C8D7C7D97172KISHOR')"
'Call TmpTable
'
'
'
'UpdateConfig ("65")
'MsgBox "Version 65 Installed successfully"
'
'End Function
'
'
'Public Function Version64()
'
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " alter table Txn_GodownRentDetail add LAID decimal(18,0)"
'Call TmpTable
'
'tmp = "ALTER TABLE Txn_GodownRentDetail ADD"
'tmp = tmp & "  CONSTRAINT FK_Txn_GodownRentDetail_Txn_LessorAgreement FOREIGN KEY"
'tmp = tmp & " (LAID) REFERENCES Txn_LessorAgreement (LAID)"
'
'Call TmpTable
'
'UpdateConfig ("64")
'MsgBox "Version 64 Installed successfully"
'
'End Function
'
'Public Function Version63()
'
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " Alter table Txn_LessorAgreement add " & c & " ServiceTaxID Decimal(18,0)"
'Call TmpTable
'
'tmp = "ALTER TABLE Txn_LessorAgreement ADD"
'tmp = tmp & " CONSTRAINT FK_Txn_LessorAgreement_Mst_ServiceTax FOREIGN KEY"
'tmp = tmp & " (ServiceTaxID) REFERENCES Mst_ServiceTax (ServiceTaxID)"
'Call TmpTable
'
'tmp = "Alter table Txn_LessorAgreement nocheck constraint FK_Txn_LessorAgreement_Mst_ServiceTax"
'Call TmpTable
'
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(137, 'FrmRpt_WHRFundingMIS.frm', 'WHR Pledge Detail', 'FrmRpt_WHRFundingMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("63")
'MsgBox "Version 63 Installed successfully"
'
'End Function
'
'Public Function Version62()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Txn_GodownRent ("
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClaimDate datetime NOT NULL ,"
'tmp = tmp & " ExpenseFrom datetime ,"
'tmp = tmp & " ExpenseTo datetime ,"
'tmp = tmp & " LessorID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Amount decimal(18, 2) ,"
'tmp = tmp & " Remarks " & TextType & " (200)  ,"
'tmp = tmp & " ApprovalRemarks " & TextType & " (200)  ,"
'tmp = tmp & " ApprovalDate datetime  ,"
'tmp = tmp & " BatchID decimal(18, 0)  ,"
'tmp = tmp & " ForwardToID decimal(18, 0)  ,"
'tmp = tmp & " Flag " & TextType & " (2)  ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NOT NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)   ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  ,"
'tmp = tmp & " UpdatedByFAG " & TextType & " (30)  ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDateFAG " & TextType & " (30) , "
'   tmp = tmp & " PRIMARY KEY (ClaimID ))"
'Else
'     tmp = tmp & " UpdatedDateFAG " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Txn_GodownRentDetail ("
'tmp = tmp & " ClaimDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " GodownID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExpenseTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Particulars " & TextType & " (250) ,"
'tmp = tmp & " Amount decimal(18, 2) NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30) ,"
'   tmp = tmp & " PRIMARY KEY (ClaimDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30))"
'End If
'
'Call TmpTable
'
'
'
'tmp = "CREATE TABLE Log_GodownRentHistory ("
'tmp = tmp & " ClaimID decimal(18, 0) ,"
'tmp = tmp & " OldForwardToID decimal(18, 0)  ,"
'tmp = tmp & " NewForwardToID decimal(18, 0)  ,"
'tmp = tmp & " UpdatedBy " & TextType & " (36) ,"
'tmp = tmp & " Concurrency datetime )"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_GodownRent WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_GodownRent PRIMARY KEY  CLUSTERED (ClaimID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_GodownRent ADD"
'   tmp = tmp & "   CONSTRAINT DF_Txn_GodownRent_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_GodownRent ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownRent_Mst_Employee FOREIGN KEY"
'   tmp = tmp & " (ForwardToID) REFERENCES Mst_Employee (   EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownRent_Mst_Lessor FOREIGN KEY"
'   tmp = tmp & " (LessorID) REFERENCES Mst_Lessor (LessorID)"
'   Call TmpTable
'
'
'
'   tmp = "ALTER TABLE Txn_GodownRentDetail WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Txn_GodownRentDetail PRIMARY KEY  CLUSTERED (ClaimDetailID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GodownRentDetail ADD"
'   tmp = tmp & " CONSTRAINT DF_Txn_GodownRentDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'
'   tmp = "ALTER TABLE Txn_GodownRentDetail ADD"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownRentDetail_Mst_ExpenseType FOREIGN KEY"
'   tmp = tmp & " (ExpenseTypeID) REFERENCES Mst_ExpenseType (ExpenseTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownRentDetail_Mst_Godown FOREIGN KEY"
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownRentDetail_Txn_GodownRent FOREIGN KEY"
'   tmp = tmp & " (ClaimID) REFERENCES Txn_GodownRent (ClaimID)"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Log_GodownRentHistory ADD"
'   tmp = tmp & " CONSTRAINT DF_Log_GodownRentHistory_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "CREATE TRIGGER tr_Txn_GodownRent "
'   tmp = tmp & " ON Txn_GodownRent  AFTER UPDATE AS "
'   tmp = tmp & " DECLARE @mClaimID  decimal(18, 0), "
'   tmp = tmp & " @mOldForwardToID decimal(18, 0), "
'   tmp = tmp & " @mNewForwardToID decimal(18, 0), "
'   tmp = tmp & " @mUpdatedBy " & TextType & "(30) "
'   tmp = tmp & " SELECT  "
'   tmp = tmp & " @mClaimID= d.ClaimID, "
'   tmp = tmp & " @mOldForwardToID = d.ForwardToID, "
'   tmp = tmp & " @mNewForwardToID = i.ForwardToID, "
'   tmp = tmp & " @mUpdatedBy= i.UpdatedBy  "
'   tmp = tmp & " FROM    inserted i INNER JOIN deleted d "
'   tmp = tmp & " ON i.ClaimID = d.ClaimID  "
'   tmp = tmp & " insert into Log_GodownRentHistory (ClaimID ,OldForwardToID ,NewForwardToID ,UpdatedBy) "
'   tmp = tmp & " values(@mClaimID ,@mOldForwardToID ,@mNewForwardToID ,@mUpdatedBy)"
'
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(134,'FrmTxn_GodownRent.frm','Process','FrmTxn_GodownRent', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(135,'FrmTxn_GodownApproval.frm','Rent Approval','FrmTxn_GodownApproval', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_Lessor ALTER COLUMN TDSCode Decimal(18,0)"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(136,'Frm_GodonRentNAVExport.frm','Export Godown-Rent Claim NAV','Frm_GodonRentNAVExport', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "Alter table Txn_LessorAgreement add  Flag Char(1)"
'Call TmpTable
'
'
'
'UpdateConfig ("62")
'MsgBox "Version 62 Installed successfully"
'
'End Function
'
'Public Function Version61()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Log_UpdateCxTFDetails ("
'tmp = tmp & " CxTFUpdateID decimal(18, 0) NOT NULL ,"
'   tmp = tmp & " ModuleType " & TextType & " (10) ,"
'   tmp = tmp & " ModuleNo decimal(18, 0) ,"
'   tmp = tmp & " PrevCXTFDate datetime ,"
'   tmp = tmp & " NewCXTFDate datetime ,"
'   tmp = tmp & " PrevAgentID decimal(18, 0) ,"
'   tmp = tmp & " NewAgentID decimal(18, 0) ,"
'   tmp = tmp & " PrevManualCDTFNo " & TextType & "(18),"
'   tmp = tmp & " NewManualCDTFNo " & TextType & "(18),"
'   tmp = tmp & " Remark " & TextType & " (255) ,"
'   tmp = tmp & " UpdatedBy " & TextType & " (30) ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Concurrency DateTime ,"
'   tmp = tmp & " PRIMARY KEY (CxTFUpdateID)) "
'Else
'   tmp = tmp & " Concurrency DateTime ) "
'End If
'
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Log_UpdateCxTFDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Log_UpdateCxTFDetails PRIMARY KEY  CLUSTERED (CxTFUpdateID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Log_UpdateCxTFDetails ADD "
'   tmp = tmp & " CONSTRAINT DF_Log_UpdateCxTFDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(133, 'Frm_UpdateCxtfDetails.frm', 'Update CxTF Details', 'Frm_UpdateCxtfDetails', 'Y', 'A',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("61")
'MsgBox "Version 61 Installed successfully"
'
'End Function
'
'
'Public Function Version60()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Mst_EmpLedger ( "
'tmp = tmp & " VendorCode " & TextType & " (25) , "
'tmp = tmp & " VendorName " & TextType & " (100)  ,"
'tmp = tmp & " DocumentType " & TextType & " (50)  ,"
'tmp = tmp & " DocumentNo " & TextType & " (20)  ,"
'tmp = tmp & " TransactionDate DateTime  ,"
'tmp = tmp & " Narration " & TextType & " (500)  ,"
'tmp = tmp & " TransactionAmount Decimal(18,2) ,"
'tmp = tmp & " RemainingAmount Decimal(18,2)  ,"
'tmp = tmp & " AcutalPostingDate DateTime  ,"
'tmp = tmp & " AcutalPostingTime " & TextType & " (30)  ,"
'tmp = tmp & " EmployeeID Decimal(18,0),"
'tmp = tmp & " EntryNo " & TextType & " (20) Not Null ,"
'tmp = tmp & " CreatedBy " & TextType & " (20) ,"
'tmp = tmp & " CreatedOn DateTime ,"
''tmp = tmp & " Concurrency DateTime )"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " Concurrency DateTime,"
'   tmp = tmp & " PRIMARY KEY (EntryNo)) "
'Else
'   tmp = tmp & " Concurrency DateTime ) "
'End If
'
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_Lessor ("
'tmp = tmp & " LessorID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LessorName " & TextType & "(200)  NOT NULL ,"
'tmp = tmp & " Address " & TextType & "(255),"
'tmp = tmp & " ContactNo " & TextType & "(15),"
'tmp = tmp & " MobileNo " & TextType & "(15),"
'tmp = tmp & " EmailID " & TextType & "(100),"
'tmp = tmp & " Fax " & TextType & "(15),"
'tmp = tmp & " PANNo " & TextType & "(10),"
'tmp = tmp & " BankID decimal(18, 0),"
'tmp = tmp & " BranchID decimal(18, 0),"
'tmp = tmp & " AccountNo " & TextType & "(50),"
'tmp = tmp & " RTGSNo " & TextType & "(50),"
'tmp = tmp & " G_UID " & TextType & "(36),"
'tmp = tmp & " NAVCode " & TextType & "(25),"
'tmp = tmp & " TDSCode " & TextType & "(25),"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & "(30),"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (LessorID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_LessorAgreement ("
'tmp = tmp & " LAID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " EntryDate datetime ,"
'tmp = tmp & " StartDate datetime ,"
'tmp = tmp & " ExpiryDate datetime ,"
'tmp = tmp & " AgreementRecvAs " & TextType & " (50) ,"
'tmp = tmp & " AdvanceGiven decimal(18, 2) ,"
'tmp = tmp & " AdvanceAdjusted decimal(18, 2) ,"
'tmp = tmp & " AdvanceReturn decimal(18, 2) ,"
'tmp = tmp & " AdvanceBalance decimal(18, 2) ,"
'tmp = tmp & " Remark " & TextType & " (250) ,"
'tmp = tmp & " G_UID " & TextType & " (36) ,"
'tmp = tmp & " Concurrency datetime NOT Null ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (LAID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_LessorAgreementDetail ("
'tmp = tmp & " LADetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LAID decimal(18, 0) NOT NULL  ,"
'tmp = tmp & " LessorID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LessorAmt decimal(18, 2) ,"
'tmp = tmp & " Remark " & TextType & " (250)  ,"
'tmp = tmp & " G_UID " & TextType & " (36) ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (LADetailID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'tmp = "CREATE TABLE Txn_GodownLessorAgreeDetail ("
'tmp = tmp & " GLADID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LAID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " GodownID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5),"
'tmp = tmp & " LessorID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " RentAmount decimal(18, 2),"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & "(30),"
'   tmp = tmp & " PRIMARY KEY (LADetailID)) "
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
'tmp = "CREATE TABLE Log_LessorAgreementHistory ("
'tmp = tmp & " LAID decimal(18, 0),"
'tmp = tmp & " UpdatedOn datetime,"
'tmp = tmp & " UpdatedBy " & TextType & "  (30))"
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_EmpLedger WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_EmpLedger PRIMARY KEY  CLUSTERED (EntryNo)  "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpLedger ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_EmpLedger_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpLedger ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_EmpLedger_Mst_Employee FOREIGN KEY (EmployeeID) REFERENCES Mst_Employee (EmployeeId)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Lessor WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_Lessor PRIMARY KEY  CLUSTERED (LessorID)  "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Lessor ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_Lessor_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Lessor ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_Lessor_Mst_Bank FOREIGN KEY (  BankID  ) REFERENCES Mst_Bank (BankId   ),"
'   tmp = tmp & " CONSTRAINT FK_Mst_Lessor_Mst_Branch FOREIGN KEY (BranchID) REFERENCES Mst_Branch (BranchId)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_LessorAgreement WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_LessorAgreement PRIMARY KEY  CLUSTERED (LAID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_LessorAgreement ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_LessorAgreement_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_LessorAgreementDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_LessorAgreementDetail PRIMARY KEY  CLUSTERED (LADetailID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_LessorAgreementDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_LessorAgreementDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_LessorAgreementDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_LessorAgreementDetail_Mst_Lessor FOREIGN KEY (LessorID) REFERENCES Mst_Lessor (LessorID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_LessorAgreementDetail_Txn_LessorAgreement FOREIGN KEY (LAID) REFERENCES Txn_LessorAgreement (  LAID)"
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_GodownLessorAgreeDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_GodownLessorAgreeDetail PRIMARY KEY  CLUSTERED (GLADID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GodownLessorAgreeDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_GodownLessorAgreeDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GodownLessorAgreeDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownLessorAgreeDetail_Mst_Godown FOREIGN KEY (GodownID) REFERENCES Mst_Godown (GodownID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownLessorAgreeDetail_Mst_Lessor FOREIGN KEY (LessorID) REFERENCES Mst_Lessor (LessorID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_GodownLessorAgreeDetail_Txn_LessorAgreement FOREIGN KEY (LAID) REFERENCES Txn_LessorAgreement (LAID)"
'
'   Call TmpTable
'
'
'   tmp = "CREATE TRIGGER tr_Txn_LessorAgreement_U ON Txn_LessorAgreement"
'   tmp = tmp & " AFTER UPDATE AS DECLARE"
'   tmp = tmp & " @mLAID  decimal(18, 0),"
'   tmp = tmp & " @mUpdatedOn DateTime,"
'   tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
'   tmp = tmp & " SELECT"
'   tmp = tmp & " @mLAID= d.LAID,"
'   tmp = tmp & " @mUpdatedOn = GetDate(),"
'   tmp = tmp & " @mUpdatedBy = i.UpdatedBy"
'   tmp = tmp & " FROM  inserted i"
'   tmp = tmp & " INNER JOIN deleted d ON i.LAID = d.LAID"
'   tmp = tmp & " insert into Log_LessorAgreementHistory (LAID ,UpdatedOn ,UpdatedBy)"
'   tmp = tmp & " values(@mLAID ,@mUpdatedOn ,@mUpdatedBy)"
'
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(128, 'Frm_EmpLedger.frm', 'Employee Ledger Upload', 'Frm_EmpLedger', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(129, 'FrmMst_Lessor.frm', 'Lessor Master', 'FrmMst_Lessor', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(130, 'FrmTxn_LessorAgreement.frm', 'Lessor Agreement', 'FrmTxn_LessorAgreement', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(131, 'FrmRpt_EmployeeLedger.frm', 'Employee Ledger', 'FrmRpt_EmployeeLedger', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(132, 'FrmRpt_ShortageInPledgedQty.frm', 'Shortage Pledge Quantity', 'FrmRpt_ShortageInPledgedQty', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'UpdateConfig ("60")
'MsgBox "Version 60 Installed successfully"
'
'End Function
'
'Public Function Version59()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
'
'tmp = " CREATE TABLE Log_UnmarkPrint_History ( "
'tmp = tmp & "   SM_Prefix " & TextType & "(5)  , "
'tmp = tmp & "   DocType " & TextType & "(10)  , "
'tmp = tmp & "   DocNo decimal(18, 0)  , "
'tmp = tmp & "   UserId " & TextType & " (15)  , "
'tmp = tmp & "   Remarks " & TextType & " (500)  , "
'tmp = tmp & "   HologramNo " & TextType & " (25)  , "
'tmp = tmp & "   Concurrency datetime Not Null ) "
'
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Log_UnmarkPrint_History WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Log_UnmarkPrint_History_Concurrency DEFAULT (getdate()) FOR           Concurrency"
'   Call TmpTable
'
'End If
'
'
'tmp = "Alter Table Mst_RRN Add " & c & " Remarks " & TextType & "(500)"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(127, 'FrmRpt_NCDEXAdvanceSC.frm', 'NCDEX Advance Storage Charges', 'FrmRpt_NCDEXAdvanceSC', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'UpdateConfig ("59")
'MsgBox "Version 59 Installed successfully"
'End Function
'
'Public Function Version58()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE Mst_WHRUOM ("
'tmp = tmp & " WHRUnitID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Description " & TextType & " (50) ,"
'tmp = tmp & " ApplicableOn " & TextType & " (30) ,"
'tmp = tmp & " G_UID " & TextType & " (36) ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime ,"
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (WHRUnitID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_WHRUOM WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_WHRUOM PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( WHRUnitID )"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_WHRUOM ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_WHRUOM_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(124,'FrmRpt_PendingWHR.frm','Pending WHRs','FrmRpt_PendingWHR', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(125,'FrmRpt_GSLISINLocationMisMatch.frm','Mismatch GSL ISIN Location','FrmRpt_GSLISINLocationMisMatch', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(126, 'FrmMst_WHRUOM.frm', 'WHR Unit Of Measurement', 'FrmMst_WHRUOM', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("58")
'MsgBox "Version 58 Installed successfully"
'
'End Function
'Public Function Version57()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_BillingCycle ("
'tmp = tmp & " BillingCycleID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BillingCycle " & TextType & " (50) Not NULL ,"
'tmp = tmp & " NoofDays decimal(18, 0) ,"
'tmp = tmp & " Flag " & TextType & " (1) ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (36) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedBy " & TextType & " (36) ,"
''tmp = tmp & " UpdatedDate datetime ) "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDate datetime , "
'   tmp = tmp & " PRIMARY KEY (BillingCycleID))"
'Else
'   tmp = tmp & " UpdatedDate datetime ) "
'End If
'
'Call TmpTable
'
'tmp = "Alter Table Mst_StorageRateCard Add " & c & " SM_Prefix " & TextType & " (5)"
'Call TmpTable
'
'
'tmp = "ALTER TABLE Mst_StorageRateCard ALTER COLUMN RatePerBag numeric(18,2)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_BillingCycle WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_BillingCycle PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (BillingCycleID ) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_BillingCycle ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_BillingCycle_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(123,'FrmMst_BillingCycle.frm','Billing Cycle','FrmMst_BillingCycle', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "Update Mst_MenuDetail Set FormFileName = 'FrmMst_StorageRateCard.frm',FormName = 'FrmMst_StorageRateCard', "
'tmp = tmp & " HMenuFlag = 'T' Where MenuID = 62 "
'Call TmpTable
'
'UpdateConfig ("57")
'MsgBox "Version 57 Installed successfully"
'
'End Function
'
'Public Function Version56()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter table Mst_GSL Add " & c & " CRAmount Decimal(18,2), " & c & " CRDate Datetime "
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(122, 'FrmRpt_CMPPV.frm', 'CMP PV', 'FrmRpt_CMPPV', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("56")
'MsgBox "Version 56 Installed successfully"
'
'End Function
'Public Function Version55()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter table Mst_Depository add " & c & " NavCode " & TextType & " (25) "
'
'Call TmpTable
'
'UpdateConfig ("55")
'MsgBox "Version 55 Installed successfully"
'
'End Function
'
'Public Function Version54()
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(121,'FrmRpt_LocationRentMIS.frm','Location Rent MIS','FrmRpt_LocationRentMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("54")
'MsgBox "Version 54 Installed successfully"
'
'End Function
'Public Function Version53()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = " CREATE TABLE Txn_CashReceiptCDTFDetails ("
'tmp = tmp & " CRCDTFID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CRID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL ,"
'tmp = tmp & " CxTFNo decimal(18, 0) NOT NULL  ,"
'tmp = tmp & " CxTFDate datetime NULL  ,"
'tmp = tmp & " TxnType " & TextType & " (1) Not NULL,"
'tmp = tmp & " FromDate datetime NULL ,"
'tmp = tmp & " DueDate datetime NULL ,"
'tmp = tmp & " NoofDays decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClientIDRow decimal(18, 0) NULL ,"
'tmp = tmp & " ChequeAmount decimal(18, 2) NULL ,"
'tmp = tmp & " TDSAmount decimal(18, 2) NULL ,"
'tmp = tmp & " NetAmount decimal(18, 2) NULL ,"
'tmp = tmp & " CDTFRemarks " & TextType & " (255)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime  NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime  NULL ,"
'tmp = tmp & " UpdatedDate datetime  NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (CRCDTFID,SM_Prefix))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "Alter Table Txn_CashReceipt Add " & c & " Predemat decimal(18,0)"
'Call TmpTable
'
'tmp = "Alter Table Txn_CashReceiptInvoiceDetails Add " & c & " TDSAmount decimal(18,2)," & c & " ChequeAmount decimal(18,2)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Txn_CashReceiptCDTFDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_CashReceiptCDTFDetails PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (CRCDTFID,SM_Prefix)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CashReceiptCDTFDetails WITH NOCHECK ADD CONSTRAINT DF_Txn_CashReceiptCDTFDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_CashReceiptCDTFDetails ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_CashReceiptCDTFDetails_Txn_CxTF_Details FOREIGN KEY "
'   tmp = tmp & " (CxTFNo,TxnType,SM_Prefix) REFERENCES Txn_CxTF_Details (CxTFNo,TxnType,SM_Prefix),"
'   tmp = tmp & " CONSTRAINT FK_Txn_CashReceiptCDTFDetails_Txn_CashReceipt FOREIGN KEY "
'   tmp = tmp & " (CRID,SM_Prefix) REFERENCES Txn_CashReceipt (CRID,SM_Prefix)"
'
'   Call TmpTable
'
'End If
'
'tmp = "Update Txn_CashReceipt Set Predemat = 0 "
'Call TmpTable
'
'tmp = "Update Txn_CashReceiptInvoiceDetails Set TDSAmount = 0 "
'Call TmpTable
'
'tmp = "Update Txn_CashReceiptInvoiceDetails Set ChequeAmount = Amount "
'Call TmpTable
'
'tmp = "drop View Vw_CxTFDetails "
'Call TmpTable
'
'tmp = "Create View Vw_CxTFDetails As"
'tmp = tmp & " select a.SM_Prefix,b.ClientIDRow,b.RRNID,sum(b.AdjRRNQty)'TxnQty',sum(b.AdjRRNNoofBags)'TxnBags'"
'tmp = tmp & " From Txn_CxTF_Details a"
'tmp = tmp & " Inner join Txn_CxTF_ClientDetail  b on a.CxTFNO=b.CxTFNO and a.SM_Prefix=b.SM_Prefix and a.txnType=b.txnType"
'tmp = tmp & " Where ExchangeTypeID = 1 and a.txntype = 'W' and (b.AdjRRNQty<>0 Or b.AdjRRNNoofBags<>0)"
'tmp = tmp & " Group by a.SM_Prefix,b.ClientIDRow,b.RRNID"
'Call TmpTable
'
'
'UpdateConfig ("53")
'MsgBox "Version 53 Installed successfully"
'
'End Function
'
'
'Public Function Version52()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(120, 'FrmRpt_MissingEmpID.frm', 'Missing Employees', 'FrmRpt_MissingEmpID', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("52")
'MsgBox "Version 52 Installed successfully"
'
'End Function
'Public Function Version51()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter Table Txn_WHR_Detail Add " & c & " SRAsPerBag Decimal(18,0)"
'Call TmpTable
'
'tmp = "Update Txn_WHR_Detail Set SRAsPerBag = 1"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(118, 'FrmRpt_AvgMonthStockMIS.frm', 'Insurance MIS', 'FrmRpt_AvgMonthStockMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(119, 'FrmRpt_RRNPending.frm', 'RRN Pending', 'FrmRpt_RRNPending', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("51")
'MsgBox "Version 51 Installed successfully"
'
'End Function
'
'Public Function Version50()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Log_LoginDetail ("
'tmp = tmp & " Login_Date datetime NOT NULL ,"
'tmp = tmp & " EmployeeID " & TextType & " (15) NULL )"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Log_LoginDetail ADD "
'   tmp = tmp & " CONSTRAINT DF_Log_LoginDetail_Login_Date DEFAULT (getdate()) FOR Login_Date"
'
'   Call TmpTable
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(116, 'FrmRpt_EmployeeExpenseMIS.frm', 'Employee Expense MIS', 'FrmRpt_EmployeeExpenseMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(117, 'FrmRpt_LoginDetails.frm', 'Login Details', 'FrmRpt_LoginDetails', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("50")
'MsgBox "Version 50 Installed successfully"
'
'End Function
'
'
'Public Function Version49()
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(114,'Frm_BenposImport.frm','BENPOS Upload','Frm_BenposImport', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = " INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(115, 'FrmRpt_GSLBenPosMIS.frm', 'GSL Benpos MIS', 'FrmRpt_GSLBenPosMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("49")
'MsgBox "Version 49 Installed successfully"
'
'
'End Function
'
'Public Function Version48()
'
'Dim TextType, c As Variant
'
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Mst_RRN Add " & c & " OutOfSystemQty decimal(18,3)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_RRN WITH NOCHECK ADD CONSTRAINT DF_Mst_RRN_OutOfSystemQty DEFAULT (0) FOR OutOfSystemQty"
'   Call TmpTable
'End If
'
'tmp = "Update Mst_RRN set OutOfSystemQty = 0"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(113,'FrmRpt_CMPBOD.frm','CMP BOD','FrmRpt_CMPBOD', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("48")
'MsgBox "Version 48 Installed successfully"
'
'
'End Function
'
'Public Function Version47()
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(112,'FrmRpt_GodownUtilizationMIS.frm','Godown Capacity MIS','FrmRpt_GodownUtilizationMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("47")
'MsgBox "Version 47 Installed successfully"
'
'End Function
'
'Public Function Version46()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE Mst_CA (RowID decimal(18, 0) NOT NULL ,CANo decimal(18, 0) NOT NULL , "
'tmp = tmp & " CADate datetime NOT NULL ,RRNID decimal(18, 0) NOT NULL ,Quantity decimal(18, 3) NOT NULL , "
'tmp = tmp & " PartialAcceptance decimal(18, 3) NOT NULL ,Flag " & TextType & " (2)  ,G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,CreatedBy " & TextType & " (35)  ,CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (35), "
'   tmp = tmp & " PRIMARY KEY (RowID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (35)) "
'End If
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Log_KRDelete ("
'tmp = tmp & " GSLID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL ,"
'tmp = tmp & " KRNo decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ISINID " & TextType & " (25) ,"
'tmp = tmp & " Flag " & TextType & " (2)  ,"
'tmp = tmp & " Remarks " & TextType & " (500)  ,"
'tmp = tmp & " DeletedBy " & TextType & " (35),"
'tmp = tmp & " DeletedDate datetime ,"
'tmp = tmp & " G_UID " & TextType & " (36),"
'tmp = tmp & " Concurrency datetime)"
'
'Call TmpTable
'
'
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_CA WITH NOCHECK ADD CONSTRAINT PK_Mst_CA PRIMARY KEY  CLUSTERED (RowID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_CA WITH NOCHECK ADD CONSTRAINT DF_Mst_CA_Concurrency "
'   tmp = tmp & " DEFAULT (getdate()) FOR Concurrency "
'   Call TmpTable
'
'   tmp = "Alter Table Mst_RRN Add " & c & " CA_Qty decimal(18,3)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Mst_RRN WITH NOCHECK ADD CONSTRAINT DF_Mst_RRN_CA_Qty DEFAULT (0) FOR CA_Qty"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Log_KRDelete WITH NOCHECK ADD CONSTRAINT DF_Log_KRDelete_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'End If
'
'tmp = "Update Mst_RRN set CA_Qty = 0"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(109,'FrmRpt_RRNConfirmation.frm','RRN Confirmation','FrmRpt_RRNConfirmation', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(110,'Frm_CAUpload.frm','CA for RRN Upload','Frm_CAUpload', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(111,'Frm_KRDelete.frm','KR-DR Delete','Frm_KRDelete', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("46")
'MsgBox "Version 46 Installed successfully"
'
'End Function
'
'
'
'Public Function Version45()
'
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
''---Gross Wt
'
'tmp = "Create View Vw_CxTFGSL  As"
'tmp = tmp & " SELECT  d.CxTFNO, d.SM_Prefix, d.txnType, SUM(g.GrossWeight) 'GSL-Gross-Wt'"
'tmp = tmp & " FROM  Txn_CxTF_Details d"
'tmp = tmp & " Inner JOIN   Txn_CxTF_GSL g ON d .CxTFNO = g.CxTFNO AND d .SM_Prefix = g.SM_Prefix AND d .txnType = g.txnType"
'tmp = tmp & " WHERE  d .txntype = 'D'"
'tmp = tmp & " GROUP BY d.SM_Prefix, d.CxTFNO, d.txnType"
'Call TmpTable
'
'tmp = "Create View Vw_TruckDetail   As"
'tmp = tmp & " SELECT d.CxTFNO, d.SM_Prefix, d.txnType,SUM(t .GrossWeight) 'TTD-Gross-Wt'"
'tmp = tmp & " FROM   Txn_CxTF_Details d"
'tmp = tmp & " INNER JOIN Txn_CxTF_TruckDetail t ON d .CxTFNO = t .CxTFNO AND d .SM_Prefix = t .SM_Prefix AND d .txnType = t .txnType"
'tmp = tmp & " WHERE  d.txntype = 'D'"
'tmp = tmp & " GROUP BY  d.SM_Prefix, d.CxTFNO, d.txnType"
'Call TmpTable
'
''---Delivered Qty
'
'tmp = "Create View Vw_CxTFDetails As"
'tmp = tmp & " select a.SM_Prefix,b.ClientIDRow,b.RRNID,sum(b.AdjRRNQty)'TxnQty',sum(b.AdjRRNNoofBags)'TxnBags'"
'tmp = tmp & " From Txn_CxTF_Details a"
'tmp = tmp & " Inner join Txn_CxTF_ClientDetail  b on a.CxTFNO=b.CxTFNO and a.SM_Prefix=b.SM_Prefix and a.txnType=b.txnType"
'tmp = tmp & " Where ExchangeTypeID = 1 and a.txntype = 'W' and b.AdjRRNQty<>0 and b.AdjRRNNoofBags<>0"
'tmp = tmp & " Group by a.SM_Prefix,b.ClientIDRow,b.RRNID"
'Call TmpTable
'
'
'tmp = "Create View Vw_Whr_detail As"
'tmp = tmp & " Select SM_Prefix,ClientIDRow,WHRNo,DeliveredQuantity,DeliveredNoOfBags"
'tmp = tmp & " From Txn_Whr_detail"
'tmp = tmp & " Where DeliveredQuantity <> 0 And DeliveredNoOfBags <> 0"
'Call TmpTable
'
''---RRN
'
'tmp = "Create View Vw_RRN As"
'tmp = tmp & " Select r.ClientIDRow,r.RRNID,Sum(r.DeliveredQty)'Del Qty'"
'tmp = tmp & " From MST_RRN r"
'tmp = tmp & " Where r.DeliveredQty <> 0"
'tmp = tmp & " Group by r.ClientIDRow,r.rrnid"
'Call TmpTable
'
'tmp = "Create View Vw_ClientDetail As"
'tmp = tmp & " Select b.SM_Prefix, b.ClientIDRow,b.RRNID,sum(b.AdjRRNQty)'TxnQty'"
'tmp = tmp & " from Txn_CxTF_ClientDetail  b"
'tmp = tmp & " Inner Join Mst_Client c on c.ClientIDRow=b.ClientIDRow"
'tmp = tmp & " where  b.Txntype = 'W' and c.ExchangeTypeID=2 "
'tmp = tmp & " Group By  b.SM_Prefix, b.ClientIDRow,b.RRNID "
'Call TmpTable
'
''---Gsl Detail
'
'tmp = "Create View Vw_GSLDetail As"
'tmp = tmp & " Select SM_Prefix,ClientIDRow,CDCNO,Sum(DeliveredQty)'Del Qty'"
'tmp = tmp & " From Txn_CDC_GSL_Detail"
'tmp = tmp & " Group by SM_Prefix,ClientIDRow,CDCNO"
'Call TmpTable
'
'tmp = "Create View Vw_ClientDetailforGSLDetail As"
'tmp = tmp & " Select b.SM_Prefix, b.ClientIDRow,b.RRNID,sum(b.AdjRRNQty)'TxnQty'"
'tmp = tmp & " from Txn_CxTF_ClientDetail  b"
'tmp = tmp & " Inner Join Mst_Client c on c.ClientIDRow=b.ClientIDRow"
'tmp = tmp & " where  b.Txntype = 'W' and c.ExchangeTypeID=3"
'tmp = tmp & " Group By  b.SM_Prefix, b.ClientIDRow,b.RRNID"
'Call TmpTable
'
'
'tmp = "Alter Table Txn_AdjTransaction Add " & c & " Remarks " & TextType & " (1000) "
'Call TmpTable
'
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(107, 'FrmRpt_NewDailyMIS.frm', 'New Daily MIS', 'FrmRpt_NewDailyMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(108, 'FrmRpt_ControlReport.frm', 'Control Report', 'FrmRpt_ControlReport', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("45")
'MsgBox "Version 45 Installed successfully"
'
'
'End Function
'Public Function Version44()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE Mst_BagEquivalent ("
'tmp = tmp & " BagEquivalentID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " BagFromSize numeric(18, 0) NOT NULL ,"
'tmp = tmp & " BagToSize numeric(18, 0) NOT NULL ,"
'tmp = tmp & " Equivalent numeric(18, 3) NOT NULL ,"
'tmp = tmp & " G_UID varchar (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (35)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (35), "
'   tmp = tmp & " PRIMARY KEY (BagEquivalentID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (35)) "
'End If
'
'Call TmpTable
'
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_BagEquivalent WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_BagEquivalent PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( BagEquivalentID )  "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_BagEquivalent ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_BagEquivalent_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(106,'FrmMst_BagEquivalent.frm','Bag Equivalent','FrmMst_BagEquivalent', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
''---
'tmp = "Alter Table Mst_GSL Add BagSize Decimal(18,3),BagEquivalent Decimal(18,3),WtQty Decimal(18,3)"
'Call TmpTable
'
'tmp = "update Mst_GSL set BagSize=BalanceWeight/BalanceBags,BagEquivalent =1,WtQty = BalanceWeight where flag <> 'B' and BalanceBags<>0"
'Call TmpTable
'
'tmp = "Alter Table Txn_GSLAdjTransaction Alter COLUMN Quantity Decimal(18,2) "
'Call TmpTable
'
'
'
'UpdateConfig ("44")
'MsgBox "Version 44 Installed successfully"
'
'End Function
'
'
'Public Function Version43()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(104, 'FrmTxn_StorageContractGslMapping.frm', 'S-Contract GSL Mapping', 'FrmTxn_StorageContractGslMapping', 'Y', 'T',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(105, 'FrmAccessRightMenu.frm', 'Access Right-Menu', 'FrmAccessRightMenu', 'N', 'A',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'
'UpdateConfig ("43")
'MsgBox "Version 43 Installed successfully"
'
'End Function
'
'Public Function Version42()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "Alter Table Txn_CashReceipt add " & c & " ExchangeTypeID decimal(18,0)"
'Call TmpTable
'
'tmp = "Select * from Txn_CashReceipt "
'Call TmpTable
'
'If TmpRs.RecordCount > 0 Then
'   tmp = "Update Txn_CashReceipt Set ExchangeTypeID = 1 "
'   Call Tmp4Table
'End If
'
'UpdateConfig ("42")
'MsgBox "Version 42 Installed successfully"
'
'End Function
'Public Function Version41()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(103, 'FrmRpt_WHRMonitor.frm', 'WHR Monitor', 'FrmRpt_WHRMonitor', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'
'Call TmpTable
'
'UpdateConfig ("41")
'MsgBox "Version 41 Installed successfully"
'
'End Function
'Public Function Version40()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "ALTER VIEW ExchageStock"
'tmp = tmp & " AS SELECT SUM(s1.BalanceWeight) AS Stock, l1.LocationID, Mst_Godown.UnitTypeID, s1.ExchangeTypeID, s1.ContractID"
'tmp = tmp & " FROM Mst_GSL s1 INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID "
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
'tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ExchangeTypeID, s1.ContractID, Mst_Godown.CloseDate"
'tmp = tmp & " HAVING (Mst_Godown.CloseDate IS NULL)"
'
'Call TmpTable
'
'UpdateConfig ("40")
'MsgBox "Version 40 Installed successfully"
'
'End Function
'Public Function Version39()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Txn_SRGodownMapping (SRGodownMappingID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SRGodownMappingDate datetime NOT NULL ,ContractID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (2),G_UID " & TextType & " (36),"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (SRGodownMappingID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_SRGodownMapDetails (SRGodownMapDetailsID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SRGodownMappingID decimal(18, 0) NOT NULL ,GodownID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ReservationQty decimal(18, 2) NOT NULL ,G_UID " & TextType & " (36),"
'tmp = tmp & " Concurrency datetime NOT NULL ,CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime NULL ,UpdatedDate datetime NULL , "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (SRGodownMappingID,GodownID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE Txn_SRGodownMapping WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_SRGodownMapping PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SRGodownMappingID)  "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_SRGodownMapDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_SRGodownMapDetails PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SRGodownMappingID,GodownID)  "
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_SRGodownMapping WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_SRGodownMapping_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_SRGodownMapDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_SRGodownMapDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_SRGodownMapping ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_SRGodownMapping_Mst_StorageContract FOREIGN KEY "
'   tmp = tmp & " (ContractID) REFERENCES Mst_StorageContract (ContractID)"
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_SRGodownMapDetails ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_SRGodownMapDetails_Mst_Godown FOREIGN KEY "
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_SRGodownMapDetails_Txn_SRGodownMapping FOREIGN KEY "
'   tmp = tmp & " (SRGodownMappingID) REFERENCES Txn_SRGodownMapping (SRGodownMappingID)"
'
'   Call TmpTable
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(101, 'FrmTxn_SRGodownMapping.frm', 'SR Godown Mapping', 'FrmTxn_SRGodownMapping', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'UpdateConfig ("39")
'MsgBox "Version 39 Installed successfully"
'
'End Function
'Public Function Version38()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Log_CMPReopenHistory ("
'tmp = tmp & " CMPID decimal(18, 0) NOT NULL ,OldCloseDate datetime NULL ,ReOpenedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) NULL ,"
'tmp = tmp & " Concurrency datetime NULL) "
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Log_CMPReopenHistory ADD"
'   tmp = tmp & " CONSTRAINT DF_Log_CMPReopenHistory_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'End If
'
'
'
'tmp = " CREATE TABLE Mst_CommodityPrice ("
'tmp = tmp & " StateID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthYear " & TextType & " (50)  NOT NULL ,"
'tmp = tmp & " Day1 decimal(18, 2) NULL ,"
'tmp = tmp & " Day2 decimal(18, 2) NULL ,"
'tmp = tmp & " Day3 decimal(18, 2) NULL ,"
'tmp = tmp & " Day4 decimal(18, 2) NULL ,"
'tmp = tmp & " Day5 decimal(18, 2) NULL ,"
'tmp = tmp & " Day6 decimal(18, 2) NULL ,"
'tmp = tmp & " Day7 decimal(18, 2) NULL ,"
'tmp = tmp & " Day8 decimal(18, 2) NULL ,"
'tmp = tmp & " Day9 decimal(18, 2) NULL ,"
'tmp = tmp & " Day10 decimal(18,2) NULL ,"
'tmp = tmp & " Day11 decimal(18,2) NULL ,"
'tmp = tmp & " Day12 decimal(18, 2) NULL ,"
'tmp = tmp & " Day13 decimal(18, 2) NULL ,"
'tmp = tmp & " Day14 decimal(18, 2) NULL ,"
'tmp = tmp & " Day15 decimal(18, 2) NULL ,"
'tmp = tmp & " Day16 decimal(18, 2) NULL ,"
'tmp = tmp & " Day17 decimal(18, 2) NULL ,"
'tmp = tmp & " Day18 decimal(18, 2) NULL ,"
'tmp = tmp & " Day19 decimal(18, 2) NULL ,"
'tmp = tmp & " Day20 decimal(18, 2) NULL ,"
'tmp = tmp & " Day21 decimal(18, 2) NULL ,"
'tmp = tmp & " Day22 decimal(18, 2) NULL ,"
'tmp = tmp & " Day23 decimal(18, 2) NULL ,"
'tmp = tmp & " Day24 decimal(18, 2) NULL ,"
'tmp = tmp & " Day25 decimal(18, 2) NULL ,"
'tmp = tmp & " Day26 decimal(18, 2) NULL ,"
'tmp = tmp & " Day27 decimal(18, 2) NULL ,"
'tmp = tmp & " Day28 decimal(18, 2) NULL ,"
'tmp = tmp & " Day29 decimal(18, 2) NULL ,"
'tmp = tmp & " Day30 decimal(18, 2) NULL ,"
'tmp = tmp & " Day31 decimal(18, 2) NULL ,"
'tmp = tmp & " Concurrency datetime NULL )"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE dbo.Mst_CommodityPrice WITH NOCHECK ADD "
'   tmp = tmp & "  CONSTRAINT PK_Mst_CommodityPrice PRIMARY KEY  CLUSTERED "
'   tmp = tmp & "  (StateID,CommodityID,MonthYear)  "
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE dbo.Mst_CommodityPrice ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_CommodityPrice_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = " ALTER TABLE dbo.Mst_CommodityPrice ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_CommodityPrice_Mst_Commodity FOREIGN KEY "
'   tmp = tmp & " (CommodityID ) REFERENCES dbo.Mst_Commodity (CommodityID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_CommodityPrice_Mst_State FOREIGN KEY "
'   tmp = tmp & " (StateID) REFERENCES dbo.Mst_State (StateID)"
'
'   Call TmpTable
'
'End If
'
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(98, 'Frm_ReOpenCMP.frm', 'Re-Open CMP', 'Frm_ReOpenCMP', 'Y', 'A',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(99, 'FrmMst_CommodityPrice.frm', 'Commodity Price Upload', 'FrmMst_CommodityPrice', 'Y', 'A',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) "
'tmp = tmp & " VALUES(100, 'FrmRpt_TxnListing.frm', 'Transaction Detail Report', 'FrmRpt_TxnListing', 'Y', 'R',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'
'UpdateConfig ("38")
'MsgBox "Version 38 Installed successfully"
'
'End Function
'
'
'Public Function Version37()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "ALTER  VIEW LeaseFranchStock"
'tmp = tmp & " AS SELECT  SUM(s1.BalanceWeight) AS Stock, l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID"
'tmp = tmp & " FROM Mst_GSL s1 INNER JOIN Mst_Godown ON Mst_Godown.GodownID = s1.GodownID"
'tmp = tmp & " INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
'tmp = tmp & " GROUP BY l1.LocationID, Mst_Godown.UnitTypeID, s1.ContractID, Mst_Godown.CloseDate"
'tmp = tmp & " HAVING (Mst_Godown.CloseDate IS NULL)"
'
'Call TmpTable
'
'tmp = "ALTER  VIEW LeaseFranchCapacity"
'tmp = tmp & " AS SELECT SUM(Mst_Godown.GodownCapacity) AS Capacity, Mst_CMP.LocationID,"
'tmp = tmp & " SUM(ISNULL(Mst_Godown.NCDEXAccreditedQty, 0)) AS AccreQty,Mst_Godown.UnitTypeID"
'tmp = tmp & " FROM Mst_Godown INNER JOIN Mst_CMP ON Mst_CMP.CMPID = Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN Mst_Location l1 ON l1.LocationID = Mst_CMP.LocationID"
'tmp = tmp & " GROUP BY Mst_CMP.LocationID, dbo.Mst_Godown.UnitTypeID, Mst_Godown.CloseDate"
'tmp = tmp & " HAVING (Mst_Godown.CloseDate IS NULL)"
'
'Call TmpTable
'
'UpdateConfig ("37")
'MsgBox "Version 37 Installed successfully"
'
'End Function
'Public Function Version36()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE dbo.Mst_StockStatus ("
'tmp = tmp & " ID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " MonthYear " & TextType & "(10) NULL ,"
'tmp = tmp & " StateID decimal(18, 0) NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " CMPID decimal(18, 0) NULL ,"
'tmp = tmp & " UnitTypeID decimal(18, 0) NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NULL ,"
'tmp = tmp & " ClosingBags decimal(18, 0) NULL ,"
'tmp = tmp & " ClosingWeight decimal(18, 3) NULL )"
'
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_StockStatus WITH NOCHECK ADD "
'tmp = tmp & " CONSTRAINT PK_Mst_StockStatus PRIMARY KEY  CLUSTERED "
'tmp = tmp & " (ID)"
'
'Call TmpTable
'
'tmp = "ALTER TABLE dbo.Mst_StockStatus ADD "
'tmp = tmp & " CONSTRAINT FK_Mst_StockStatus_Mst_CMP FOREIGN KEY "
'tmp = tmp & " ("
'tmp = tmp & "     CMPID"
'tmp = tmp & " ) REFERENCES dbo.Mst_CMP (CMPID),"
'tmp = tmp & " CONSTRAINT FK_Mst_StockStatus_Mst_Commodity FOREIGN KEY "
'tmp = tmp & " ("
'tmp = tmp & "     CommodityID"
'tmp = tmp & " ) REFERENCES dbo.Mst_Commodity (CommodityID),"
'tmp = tmp & " CONSTRAINT FK_Mst_StockStatus_Mst_Location FOREIGN KEY "
'tmp = tmp & " ("
'tmp = tmp & "     LocationID"
'tmp = tmp & " ) REFERENCES dbo.Mst_Location (LocationID),"
'tmp = tmp & " CONSTRAINT FK_Mst_StockStatus_Mst_UnitType FOREIGN KEY "
'tmp = tmp & " (   UnitTypeID"
'tmp = tmp & " ) REFERENCES dbo.Mst_UnitType (UnitTypeID)"
'
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(90, 'FrmMst_StockStatus.frm', 'Stock Status', 'FrmMst_StockStatus', 'Y', 'C',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(91, 'FrmRpt_GSLStockStatus.frm', 'GSL Stock Status', 'FrmRpt_GSLStockStatus', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_ClientLogin ("
'tmp = tmp & " ClientID " & TextType & " (50) NOT NULL ,"
'tmp = tmp & " ClientLoginID " & TextType & " (50) NOT NULL ,"
'tmp = tmp & " ClientLoginName " & TextType & " (200) NOT NULL ,"
'tmp = tmp & " Password " & TextType & " (50) NOT NULL ,"
'tmp = tmp & " Address " & TextType & " (255) NULL ,"
'tmp = tmp & " Organization " & TextType & " (255) NULL ,"
'tmp = tmp & " EmailID " & TextType & " (200) NOT NULL ,"
'tmp = tmp & " TelephoneNo " & TextType & " (50) NULL ,"
'tmp = tmp & " MobileNo " & TextType & " (50) NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36) NULL ,"
'tmp = tmp & " Concurrency datetime NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) NULL ) "
'
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_ClientLogin WITH NOCHECK ADD"
'tmp = tmp & " CONSTRAINT PK_Mst_ClientLogin PRIMARY KEY  CLUSTERED "
'tmp = tmp & " (ClientLoginID) "
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_ClientLoginDetail ("
'tmp = tmp & " ClientDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClientIDRow decimal(18, 0) NULL ,"
'tmp = tmp & " ClientLoginID " & TextType & " (50)  NULL )"
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_ClientLoginDetail WITH NOCHECK ADD "
'tmp = tmp & " CONSTRAINT PK_Mst_ClientLoginDetail PRIMARY KEY  CLUSTERED "
'tmp = tmp & " (ClientDetailID   )  "
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_ClientLoginDetail ADD "
'tmp = tmp & " CONSTRAINT FK_Mst_ClientLoginDetail_Mst_ClientLogin FOREIGN KEY "
'tmp = tmp & " (ClientLoginID) REFERENCES Mst_ClientLogin (ClientLoginID)"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(94, 'FrmMst_ClientLogin.frm', 'Client Login', 'FrmMst_ClientLogin', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID)"
'tmp = tmp & " VALUES(95, 'Frm_RRNUpload.frm', 'RRN Upload', 'Frm_RRNUpload', 'Y', 'A',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(96, 'FrmRpt_DailyStockMIS.frm', 'Daily Stock MIS', 'FrmRpt_DailyStockMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(97, 'FrmRpt_MonitorMIS.frm', 'Employee Monitoring', 'FrmRpt_MonitorMIS', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'
'tmp = "CREATE TABLE Mst_TmpStockStatus ("
'tmp = tmp & " MonthYear " & TextType & " (50)  NULL ,"
'tmp = tmp & " StateID decimal(18, 0) NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " ClosingBags decimal(18, 0) NULL ,"
'tmp = tmp & " ClosingWeight decimal(18, 0) NULL ,"
'tmp = tmp & " CMPID decimal(18, 0) NULL ,"
'tmp = tmp & " UnitTypeID decimal(18, 0) NULL )"
'
'Call TmpTable
'
'tmp = "Alter Table Mst_GSL Add " & c & " ContractID decimal(18,0)"
'Call TmpTable
'
'
'tmp = "CREATE VIEW dbo.LeaseFranchStock"
'tmp = tmp & " AS SELECT SUM(s1.BalanceWeight) AS Stock, l1.LocationID, dbo.Mst_Godown.UnitTypeID,s1.ContractID"
'tmp = tmp & " FROM dbo.Mst_GSL s1 INNER JOIN"
'tmp = tmp & " dbo.Mst_Godown ON dbo.Mst_Godown.GodownID = s1.GodownID INNER JOIN"
'tmp = tmp & " dbo.Mst_CMP ON dbo.Mst_CMP.CMPID = dbo.Mst_Godown.CMPID INNER JOIN"
'tmp = tmp & " dbo.Mst_Location l1 ON l1.LocationID = dbo.Mst_CMP.LocationID"
'tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
'tmp = tmp & " GROUP BY l1.LocationID, dbo.Mst_Godown.UnitTypeID,s1.ContractID"
'
'Call TmpTable
'
'
'tmp = "CREATE VIEW dbo.LeaseFranchCapacity"
'tmp = tmp & " AS SELECT SUM(dbo.Mst_Godown.GodownCapacity) AS Capacity, dbo.Mst_CMP.LocationID,"
'tmp = tmp & " SUM(ISNULL(dbo.Mst_Godown.NCDEXAccreditedQty, 0)) AS AccreQty,dbo.Mst_Godown.UnitTypeID"
'tmp = tmp & " FROM dbo.Mst_Godown INNER JOIN dbo.Mst_CMP ON dbo.Mst_CMP.CMPID = dbo.Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN dbo.Mst_Location l1 ON l1.LocationID = dbo.Mst_CMP.LocationID"
'tmp = tmp & " GROUP BY dbo.Mst_CMP.LocationID, dbo.Mst_Godown.UnitTypeID"
'Call TmpTable
'
'
'tmp = "CREATE VIEW dbo.ExchageStock"
'tmp = tmp & " AS SELECT SUM(s1.BalanceWeight) AS Stock, l1.LocationID, dbo.Mst_Godown.UnitTypeID, s1.ExchangeTypeID, s1.ContractID"
'tmp = tmp & " FROM dbo.Mst_GSL s1 INNER JOIN dbo.Mst_Godown ON dbo.Mst_Godown.GodownID = s1.GodownID "
'tmp = tmp & " INNER JOIN dbo.Mst_CMP ON dbo.Mst_CMP.CMPID = dbo.Mst_Godown.CMPID"
'tmp = tmp & " INNER JOIN dbo.Mst_Location l1 ON l1.LocationID = dbo.Mst_CMP.LocationID"
'tmp = tmp & " WHERE (s1.Flag <> 'E') AND (s1.Flag <> 'Z')"
'tmp = tmp & " GROUP BY l1.LocationID, dbo.Mst_Godown.UnitTypeID, s1.ExchangeTypeID, s1.ContractID"
'Call TmpTable
'
'tmp = "Alter Table Txn_VendorClaimDetail add " & c & " InvoiceDate DateTime"
'Call TmpTable
'
'
'UpdateConfig ("36")
'MsgBox "Version 36 Installed successfully"
'
'End Function
'
'
'Public Function Version35()
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(88, 'FrmRpt_GSLCommodityMap.frm', 'GSL Commodity Map', 'FrmRpt_GSLCommodityMap', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'tmp = " CREATE TABLE Txn_VendorClaim ( "
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClaimDate datetime NOT NULL ,"
'tmp = tmp & " ExpenseFrom datetime NULL ,"
'tmp = tmp & " ExpenseTo datetime NULL ,"
'tmp = tmp & " VendorID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NULL ,"
'tmp = tmp & " Remarks " & TextType & " (200)  NOT NULL ,"
'tmp = tmp & " ApprovalRemarks " & TextType & " (200)  NULL ,"
'tmp = tmp & " ApprovalDate datetime NULL ,"
'tmp = tmp & " BatchID decimal(18, 0) NULL ,"
'tmp = tmp & " ForwardToID decimal(18, 0) NULL,"
'tmp = tmp & " Flag " & TextType & " (2)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NOT NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " UpdatedByFAG " & TextType & " (30)  NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDateFAG " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (ClaimID))"
'Else
'   tmp = tmp & " UpdatedDateFAG " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'tmp = " CREATE TABLE dbo.Txn_VendorClaimDetail ("
'tmp = tmp & " ClaimDetailID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExpenseTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Particulars " & TextType & " (250)  NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NOT NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = tmp & " PRIMARY KEY (ClaimDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = " ALTER TABLE dbo.Txn_VendorClaim WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_VendorClaim PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (ClaimID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE dbo.Txn_VendorClaimDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_VendorClaimDetail PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (ClaimDetailID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE dbo.Txn_VendorClaim WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_VendorClaim_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE dbo.Txn_VendorClaimDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_VendorClaimDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE dbo.Txn_VendorClaim ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_VendorClaim_Mst_Vendor FOREIGN KEY "
'   tmp = tmp & " (VendorID) REFERENCES dbo.Mst_Vendor (VendorID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_VendorClaim_Mst_Employee FOREIGN KEY "
'   tmp = tmp & " (ForwardToID) REFERENCES dbo.Mst_Employee (EmployeeID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE dbo.Txn_VendorClaimDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_VendorClaimDetail_Mst_Location FOREIGN KEY "
'   tmp = tmp & " (LocationID) REFERENCES dbo.Mst_Location (LocationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_VendorClaimDetail_Mst_ExpenseType FOREIGN KEY "
'   tmp = tmp & " (ExpenseTypeID) REFERENCES dbo.Mst_ExpenseType (ExpenseTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_VendorClaimDetail_Txn_VendorClaim FOREIGN KEY "
'   tmp = tmp & " (ClaimID) REFERENCES dbo.Txn_VendorClaim (ClaimID)"
'   Call TmpTable
'
'End If
'
'tmp = "CREATE TABLE Log_VendorClaimFwdToHistory ("
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " OldForwardToID decimal(18, 0) NULL ,"
'tmp = tmp & " NewForwardToID decimal(18, 0) NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) NULL, "
'tmp = tmp & " Concurrency datetime NOT NULL CONSTRAINT DF_Log_VendorClaimFwdToHistory DEFAULT (getdate()))"
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "SET QUOTED_IDENTIFIER ON "
'   Call TmpTable
'
'   tmp = "SET ANSI_NULLS ON "
'   Call TmpTable
'
'   tmp = "CREATE TRIGGER tr_Txn_VendorClaim_U ON Txn_VendorClaim AFTER "
'   tmp = tmp & " UPDATE AS DECLARE @mClaimID  decimal(18, 0),  @mOldForwardToID decimal(18, 0), @mNewForwardToID decimal(18, 0), @mUpdatedBy " & TextType & "(30) "
'   tmp = tmp & " SELECT  @mClaimID= d.ClaimID, @mOldForwardToID = d.ForwardToID, @mNewForwardToID = i.ForwardToID, @mUpdatedBy= i.UpdatedBy  FROM   "
'   tmp = tmp & " inserted i INNER JOIN deleted d ON i.ClaimID = d.ClaimID  "
'   tmp = tmp & " insert into Log_VendorClaimFwdToHistory (ClaimID ,OldForwardToID ,NewForwardToID ,UpdatedBy) values(@mClaimID ,@mOldForwardToID ,@mNewForwardToID ,@mUpdatedBy)"
'   Call TmpTable
'
'   tmp = "SET QUOTED_IDENTIFIER OFF "
'   Call TmpTable
'
'   tmp = "SET ANSI_NULLS ON "
'   Call TmpTable
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(86, 'FrmTxn_VendorClaim.frm', 'Vendor', 'FrmTxn_VendorClaim', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(87, 'Frm_VendorNAVExport.frm', 'Export Vendor Expense Claim for Navision', 'Frm_VendorNAVExport', 'Y', 'C',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'UpdateConfig ("35")
'MsgBox "Version 35 Installed successfully"
'
'End Function
'
'Public Function Version34()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'MsgBox "Please wait!!!"
'
'tmp = "CREATE TABLE Tmp_DPIncomeDetails ("
'tmp = tmp & " FileName " & TextType & " (50)  NULL ,"
'tmp = tmp & " Cycle decimal(18, 0) NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " ISINID " & TextType & " (25)  NULL ,"
'tmp = tmp & " DPID " & TextType & " (50)  NULL ,"
'tmp = tmp & " Qty decimal(18, 0) NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NULL ,"
'tmp = tmp & " Amt_Incl_ST decimal(18, 2) NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NULL ,"
'tmp = tmp & " CMPID decimal(18, 0) NULL  )"
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "Alter Table Mst_EmpAssetMap Add " & c & " Remarks " & TextType & " (1000), AMCExpDate dateTime, AMCAmount decimal(18,2)"
'   Call TmpTable
'
''-----Add New Cloumns
'
'   tmp = "ALTER TABLE Mst_ExchangeType add " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Godown add  " & c & " Concurrency datetime ,  CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_GSL add   " & c & " Concurrency datetime ,CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_InsuranceCoMaster add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_InsuranceDetails add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_ISIN add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_License add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Location add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_RRN add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_State add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StorageContract add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StoreManLocations add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_UnitType add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Agent add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Bank add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Branch add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Category add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Client   add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_CMP add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Commodity add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Depository  add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Assayer add " & c & "  G_UID " & TextType & "(36),Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_AdjTransaction add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CashReceipt add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CDC add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CxTF_Details add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GSL_OpeningBalance add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GSL_Transfer add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GSLAdjTransaction add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ISINGSLRevalidation add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ISINRevalidation add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_PledgeCreation add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_PledgeRelease add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_Spillage add " & c & "  Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_WHR_Detail add  " & c & " Concurrency datetime , CreatedBy " & TextType & " (30), CreatedDate datetime,UpdatedDate datetime,UpdatedBy " & TextType & " (30)"
'   Call TmpTable
'
'
''-------Add Constraints
'
'   tmp = "ALTER TABLE Mst_ExchangeType ADD CONSTRAINT DF_Mst_ExchangeType_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   'tmp = "ALTER TABLE Mst_Godown ADD CONSTRAINT DF_Mst_Godown_Concurrency DEFAULT (getdate()) FOR Concurrency "
'   'Call TmpTable
'
'   tmp = "ALTER TABLE Mst_GSL ADD CONSTRAINT DF_Mst_GSL_Concurrency DEFAULT (getdate()) FOR Concurrency "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_InsuranceCoMaster ADD CONSTRAINT DF_Mst_InsuranceCoMaster_Concurrency DEFAULT (getdate()) FOR Concurrency   "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_InsuranceDetails  ADD CONSTRAINT DF_Mst_InsuranceDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_ISIN ADD CONSTRAINT DF_Mst_ISIN_Concurrency DEFAULT (getdate()) FOR Concurrency  "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_License ADD CONSTRAINT DF_Mst_License_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Location ADD CONSTRAINT DF_Mst_Location_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_RRN ADD CONSTRAINT DF_Mst_RRN_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_State ADD CONSTRAINT DF_Mst_State_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StorageContract ADD CONSTRAINT DF_Mst_StorageContract_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_StoreManLocations ADD CONSTRAINT DF_Mst_StoreManLocations_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_UnitType ADD CONSTRAINT DF_Mst_UnitType_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Agent ADD CONSTRAINT DF_Mst_Agent_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Bank ADD CONSTRAINT DF_Mst_Bank_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Branch ADD CONSTRAINT DF_Mst_Branch_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Category ADD CONSTRAINT DF_Mst_Category_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Client ADD CONSTRAINT DF_Mst_Client_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_CMP ADD CONSTRAINT DF_Mst_CMP_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Commodity ADD CONSTRAINT DF_Mst_Commodity_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Depository ADD CONSTRAINT DF_Mst_Depository_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Assayer ADD CONSTRAINT DF_Mst_Assayer_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_AdjTransaction ADD CONSTRAINT DF_Txn_AdjTransaction_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CashReceipt ADD CONSTRAINT DF_Txn_CashReceipt_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CDC ADD CONSTRAINT DF_Txn_CDC_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_CxTF_Details ADD CONSTRAINT DF_Txn_CxTF_Details_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GSL_OpeningBalance ADD CONSTRAINT DF_Txn_GSL_OpeningBalance_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GSL_Transfer ADD CONSTRAINT DF_Txn_GSL_Transfer_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_GSLAdjTransaction ADD CONSTRAINT DF_Txn_GSLAdjTransaction_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ISINGSLRevalidation ADD CONSTRAINT DF_Txn_ISINGSLRevalidation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ISINRevalidation ADD CONSTRAINT DF_Txn_ISINRevalidation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_PledgeCreation ADD CONSTRAINT DF_Txn_PledgeCreation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_PledgeRelease ADD CONSTRAINT DF_Txn_PledgeRelease_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_Spillage ADD CONSTRAINT DF_Txn_Spillage_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_WHR_Detail ADD CONSTRAINT DF_Txn_WHR_Detail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'End If
'
'UpdateConfig ("34")
'MsgBox "Version 34 Installed Successfully"
'
'End Function
'
'Public Function Version33()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_IncomeDetails Add " & c & " OldPINo " & TextType & "(30), NewPINo decimal(18,0),ExchangeTypeID decimal(18,0)"
'Call TmpTable
'
'tmp = "Alter Table Mst_Config add " & c & " IncomeStartDate DateTime"
'Call TmpTable
'
'tmp = "Update Mst_Config set IncomeStartDate = '10/01/2009'"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(85, 'FrmTxn_IncomeDetails.frm', 'Income Details', 'FrmTxn_IncomeDetails', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'UpdateConfig ("33")
'MsgBox "Version 33 Installed successfully"
'
'End Function
'
'
'
'Public Function Version32()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = "CREATE TABLE Log_ClaimForwardToHistory ("
'tmp = tmp & " ClaimRowID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " OldForwardToID decimal(18, 0) NULL ,"
'tmp = tmp & " NewForwardToID decimal(18, 0) NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, "
'tmp = tmp & " Concurrency datetime NOT NULL CONSTRAINT DF_Log_ClaimForwardToHistory DEFAULT (getdate()))"
'
'Call TmpTable
'
'
'tmp = "CREATE TRIGGER tr_Txn_ClaimSubmission_U ON Txn_ClaimSubmission AFTER UPDATE"
'tmp = tmp & " AS"
'tmp = tmp & " DECLARE @mClaimRowID  decimal(18, 0), "
'tmp = tmp & " @mOldForwardToID decimal(18, 0),"
'tmp = tmp & " @mNewForwardToID decimal(18, 0),"
'tmp = tmp & " @mUpdatedBy " & TextType & "(30)"
'tmp = tmp & " SELECT  @mClaimRowID= d.ClaimRowID,"
'tmp = tmp & " @mOldForwardToID = d.ForwardToID,"
'tmp = tmp & " @mNewForwardToID = i.ForwardToID,"
'tmp = tmp & " @mUpdatedBy= i.UpdatedBy "
'tmp = tmp & " FROM    inserted i"
'tmp = tmp & " INNER JOIN deleted d ON i.ClaimRowID = d.ClaimRowID "
'tmp = tmp & " insert into Log_ClaimForwardToHistory (ClaimRowID ,OldForwardToID ,NewForwardToID ,UpdatedBy)"
'tmp = tmp & " values(@mClaimRowID ,@mOldForwardToID ,@mNewForwardToID ,@mUpdatedBy)"
'
'Call TmpTable
'
'
'
'UpdateConfig ("32")
'MsgBox "Version 32 Installed successfully"
'
'End Function
'Public Function Version31()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_Vendor ("
'tmp = tmp & " VendorID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " VendorName " & TextType & " (100) NOT NULL ,"
'tmp = tmp & " VendorAddress " & TextType & " (255)  NULL ,"
'tmp = tmp & " LocationID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ContactNo " & TextType & "(100)  NULL ,"
'tmp = tmp & " MobileNo " & TextType & "(100) NULL ,"
'tmp = tmp & " Fax " & TextType & "(100)  NULL ,"
'tmp = tmp & " EmailID " & TextType & " (50)  NULL ,"
'tmp = tmp & " TDSID Decimal(18, 0) Not NULL ,"
'tmp = tmp & " NavCode " & TextType & " (25)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL )"
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_ServiceTax ("
'tmp = tmp & " ServiceTaxID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Description " & TextType & " (200)  NOT NULL ,"
'tmp = tmp & " NavCode " & TextType & " (25)  NULL ,"
'tmp = tmp & " Flag " & TextType & " (25)  NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy varchar (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)  NULL )"
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_TDS ("
'tmp = tmp & " TDSID Decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Description " & TextType & " (200) NOT NULL,"
'tmp = tmp & " NavCode " & TextType & " (25) NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36) NULL ,"
'tmp = tmp & " Flag " & TextType & " (25)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) NULL)"
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_Vendor WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT PK_Mst_Vendor PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (VendorID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Vendor ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_Vendor_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_TDS WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_TDS PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (TDSID)   "
'
'   Call TmpTable
'   tmp = "ALTER TABLE Mst_Vendor ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_Vendor_Mst_TDS FOREIGN KEY "
'   tmp = tmp & " (TDSID) REFERENCES dbo.Mst_TDS (TDSID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_TDS ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_TDS_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE dbo.Mst_ServiceTax WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_ServiceTax PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (ServiceTaxID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE dbo.Mst_ServiceTax ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_ServiceTax_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'    tmp = "Alter Table Mst_ExpenseType add " & c & "  ServiceTaxID Decimal(18, 0)," & c & " AllocateTo " & TextType & " (2) "
'    Call TmpTable
'
'    tmp = "Alter Table Txn_ClaimSubmission add " & c & " BatchID Decimal(18,0) ," & c & " ForwardToID Decimal(18,0) "
'    Call TmpTable
'
'
'
'    tmp = " ALTER TABLE Txn_ClaimSubmission ADD  "
'    tmp = tmp & "CONSTRAINT FK_Txn_ClaimSubmission_Mst_Employee1 FOREIGN KEY "
'    tmp = tmp & " (ForwardToID) REFERENCES dbo.Mst_Employee (EmployeeID)"
'
'    Call TmpTable
'
'    tmp = "ALTER TABLE dbo.Mst_ExpenseType ADD "
'    tmp = tmp & " CONSTRAINT FK_Mst_ExpenseType_Mst_ServiceTax FOREIGN KEY "
'    tmp = tmp & " (ServiceTaxID) REFERENCES dbo.Mst_ServiceTax (ServiceTaxID)"
'
'    Call TmpTable
'
'    tmp = "Update Txn_ClaimSubmission Set BatchID = 0 "
'    Call TmpTable
'
'    tmp = "Update Mst_ExpenseType Set AllocateTo = 'B' "
'    Call TmpTable
'
'
'End If
'
'tmp = "Alter Table Txn_ClaimSubmission add " & c & " ClaimRowID Decimal(18,0) "
'Call TmpTable
'
'tmp = "Alter Table Txn_ClaimSubmissionDetail add  " & c & " ClaimRowID Decimal(18,0) "
'Call TmpTable
'
'
'tmp = "Select * from Txn_ClaimSubmission Order By SM_Prefix,ClaimID"
'Call TmpTable
'
'tmp = "Select * from Txn_ClaimSubmissionDetail Order By SM_Prefix,ClaimID"
'Call Tmp1Table
'
'
'For I = 1 To TmpRs.RecordCount
'    TmpRs!ClaimRowID = I
'    TmpRs1.MoveFirst
'    TmpRs1.Filter = "SM_Prefix = '" & TmpRs!SM_Prefix & "' And ClaimID = " & TmpRs!ClaimID & ""
'    For x = 1 To TmpRs1.RecordCount
'        TmpRs1!ClaimRowID = I
'        TmpRs1.MoveNext
'    Next
'    TmpRs1.Filter = ""
'    TmpRs.MoveNext
'Next
'
'
'tmp = "if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_Txn_ClaimSubmissionDetail_Txn_ClaimSubmission]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1) ALTER TABLE [dbo].[Txn_ClaimSubmissionDetail] DROP CONSTRAINT FK_Txn_ClaimSubmissionDetail_Txn_ClaimSubmission "
'Call TmpTable
'
''Query For Droping Pri Key CONSTRAINT
'tmp = "ALTER TABLE [dbo].[Txn_ClaimSubmission] Drop CONSTRAINT PK_Txn_ClaimSubmission "
'Call TmpTable
'
'
'tmp = "Alter Table Txn_ClaimSubmission " & c & " Alter Column ClaimRowID Decimal(18,0) Not Null "
'Call TmpTable
'
'tmp = "Alter Table Txn_ClaimSubmissionDetail " & c & " Alter Column ClaimRowID Decimal(18,0) Not Null "
'Call TmpTable
'
'
'tmp = "ALTER TABLE [dbo].[Txn_ClaimSubmission] WITH NOCHECK ADD  "
'tmp = tmp & " CONSTRAINT [PK_Txn_ClaimSubmission] PRIMARY KEY  CLUSTERED (ClaimRowID)  "
'Call TmpTable
'
'
'tmp = "ALTER TABLE Txn_ClaimSubmissionDetail ADD "
'tmp = tmp & " CONSTRAINT FK_Txn_ClaimSubmissionDetail_Txn_ClaimSubmission FOREIGN KEY "
'tmp = tmp & " (ClaimRowID) REFERENCES Txn_ClaimSubmission(ClaimRowID) "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(80, 'FrmMst_Vendor.frm', 'Vendor', 'FrmMst_Vendor', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & "VALUES(81, 'FrmMst_ServiceTax.frm', 'Service Tax', 'FrmMst_ServiceTax', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & "VALUES(82, 'FrmMst_TDS.frm', 'TDS', 'FrmMst_TDS', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & "VALUES(83, 'FrmNAVExport.frm', 'NAV Export', 'FrmNAVExport', 'Y', 'C',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & "VALUES(84, 'FrmRpt_EmpAssetMap.frm', 'Employee Asset Details', 'FrmRpt_EmpAssetMap', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'UpdateConfig ("31")
'MsgBox "Version 31 Installed successfully"
'
'End Function
'
'
'Public Function Version30()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_ExpenseGroup ("
'tmp = tmp & " ExpenseGroupID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExpenseGroup " & TextType & " (200) NOT NULL ,"
'tmp = tmp & " NavCode " & TextType & " (25)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (ExpenseGroupID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_ExpenseGroup WITH NOCHECK ADD CONSTRAINT PK_Mst_ExpenseGroup PRIMARY KEY  CLUSTERED (ExpenseGroupID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_ExpenseGroup WITH NOCHECK ADD CONSTRAINT DF_Mst_ExpenseGroup_Concurrency DEFAULT (getdate()) FOR "
'   tmp = tmp & " Concurrency "
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(77, 'FrmMst_ExpenseGroup.frm', 'Expense Group', 'FrmMst_ExpenseGroup', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(78, 'FrmRpt_EMPLocation.frm', 'Employee Location Mapping', 'FrmRpt_EMPLocation', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "Alter Table Mst_ExpenseType add " & c & " ExpenseGroupID decimal(18,0) "
'Call TmpTable
'
'tmp = "ALTER TABLE Mst_ExpenseType ADD "
'tmp = tmp & " CONSTRAINT FK_Mst_ExpenseType_Mst_ExpenseGroup FOREIGN KEY "
'tmp = tmp & " (ExpenseGroupID) REFERENCES Mst_ExpenseGroup"
'tmp = tmp & " (ExpenseGroupID)"
'
'Call TmpTable
'
'tmp = "Alter Table Txn_TelephoneDetails add " & c & " Flag " & TextType & " (10)"
'Call TmpTable
'
'
'tmp = "Alter Table Txn_TelephoneTemp add " & c & " Flag " & TextType & " (10)"
'Call TmpTable
'
'
'
'UpdateConfig ("30")
'MsgBox "Version 30 Installed successfully"
'
'End Function
'
'
'Public Function Version29()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Txn_DPIncomeDetails ( "
'tmp = tmp & " FileName " & TextType & " (50)  NULL ,"
'tmp = tmp & " Cycle decimal(18, 0)  NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NULL ,"
'tmp = tmp & " ISINID " & TextType & " (25)  NULL ,"
'tmp = tmp & " DPID " & TextType & " (50)  NULL ,"
'tmp = tmp & " Qty decimal(18, 0) NULL ,"
'tmp = tmp & " Amount decimal(18, 2) NULL ,"
'tmp = tmp & " Amt_Incl_ST decimal(18, 2) NULL ,"
'tmp = tmp & " CommodityID decimal(18, 0) NULL ,"
'tmp = tmp & " CMPID decimal(18, 0) NULL ,"
'tmp = tmp & " G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'
'Call TmpTable
'
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Txn_DPIncomeDetails ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_DPIncomeDetails_Concurrency DEFAULT (getdate()) FOR Concurrency "
'   Call TmpTable
'End If
'
'
'
'tmp = " CREATE TABLE Txn_IncomeDetails ("
'tmp = tmp & " Mode  " & TextType & "  (50)  ,"
'tmp = tmp & "   TxnType  " & TextType & "  (15)  ,"
'tmp = tmp & "   LocationID  Decimal  (18, 0)  ,"
'tmp = tmp & "   SM_Prefix  " & TextType & "  (5)  ,"
'tmp = tmp & "   CMPID decimal(18, 0) NULL ,"
'tmp = tmp & "   BillingMonthStartDate datetime NULL ,"
'tmp = tmp & "   BillingMonthEndDate datetime NULL ,"
'tmp = tmp & "   ClientIDRow decimal(18, 0) NULL ,"
'tmp = tmp & "   CommodityID decimal(18, 0) NULL ,"
'tmp = tmp & "   OpeningNoofBags decimal(18, 0) NULL ,"
'tmp = tmp & "   BagsWithdrawnDuringMonth decimal(18, 0) NULL ,"
'tmp = tmp & "   ClosingNoofBags decimal(18, 0) NULL ,"
'tmp = tmp & "   OpeningNoOfMadeUpBags decimal(18, 0) NULL ,"
'tmp = tmp & "   NewBags_Billing decimal(18, 0) NULL ,"
'tmp = tmp & "   WithdrawnBags_Billing decimal(18, 0) NULL ,"
'tmp = tmp & "   ClosingNoOfMadeUpBags decimal(18, 0) NULL ,"
'tmp = tmp & "   RatePerBagPerMonth_OPBAL decimal(18, 2) NULL ,"
'tmp = tmp & "   StorageCharges decimal(18, 2) NULL ,"
'tmp = tmp & "   ApplyServiceTax  " & TextType & "  (2)  ,"
'tmp = tmp & "   ServiceTax decimal(18, 2) NULL ,"
'tmp = tmp & "   InsuranceCharges decimal(18, 2) NULL ,"
'tmp = tmp & "   AmtInclOfSerTaxnInsurance decimal(18, 2) NULL ,"
'tmp = tmp & "   CRID decimal(18, 0) NULL ,"
'tmp = tmp & "   Remarks  " & TextType & "  (255)  ,"
'tmp = tmp & "   CxTFNO decimal(18, 0) NULL ,"
'tmp = tmp & "   CxTFDate datetime NULL ,"
'tmp = tmp & "   WHRNo decimal(18, 0) NULL ,"
'tmp = tmp & "   NoOfBagsDeposited decimal(18, 0) NULL ,"
'tmp = tmp & "   InvoicedUptoDate datetime NULL ,"
'tmp = tmp & "   Rate decimal(18, 2) NULL ,"
'tmp = tmp & "   NoOfDays decimal(18, 0) NULL ,"
'tmp = tmp & "   CorrCDTFNo decimal(18, 2) NULL ,"
'tmp = tmp & "   CDTFDate datetime NULL ,"
'tmp = tmp & "   NoOfBagsWithdrawn decimal(18, 0) NULL ,"
'tmp = tmp & "   NoOfMadeUpbagsWithdrawn decimal(18, 0) NULL ,"
'tmp = tmp & "   BillingCycle  " & TextType & "  (25)  ,"
'tmp = tmp & "   RatePerBagPerMonth_W decimal(18, 2) NULL ,"
'tmp = tmp & "   Daily decimal(18, 2) NULL ,"
'tmp = tmp & "   Monthly decimal(18, 2) NULL ,"
'tmp = tmp & "   Weekly decimal(18, 2) NULL ,"
'tmp = tmp & "   Fortnightly decimal(18, 2) NULL ,"
'tmp = tmp & "   ApplicableNoOfDays decimal(18, 0) NULL ,"
'tmp = tmp & "   D  " & TextType & "  (25)  ,"
'tmp = tmp & "   W  " & TextType & "  (25)  ,"
'tmp = tmp & "   F  " & TextType & "  (25)  ,"
'tmp = tmp & "   M  " & TextType & "  (25)  ,"
'tmp = tmp & "   RRNID decimal(18, 0) NULL ,"
'tmp = tmp & "   CxTFNo_RejWithdrawal decimal(18, 0) NULL ,"
'tmp = tmp & "   CxTFDate_RejWithdrawal datetime NULL ,"
'tmp = tmp & "   WithdrawnNoOfBags decimal(18, 0) NULL ,"
'tmp = tmp & "   WithdrawnQuantityInMT decimal(18, 0) NULL ,"
'tmp = tmp & "   RatePerMTPerDay decimal(18, 2) NULL ,"
'tmp = tmp & "   NoOfDaysChargesCollected decimal(18, 0) NULL ,"
'tmp = tmp & "   AmtInclOfSerTax decimal(18, 2) NULL ,"
'tmp = tmp & "   TDSAmt decimal(18, 2) NULL ,"
'tmp = tmp & "   DateOfDepositedInBank datetime NULL ,"
'tmp = tmp & "   DepositedBankId  decimal(18, 0) NULL  ,"
'tmp = tmp & "   DepositedAccountNo decimal(18, 0) NULL ,"
'tmp = tmp & "   G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & "   Concurrency datetime NOT NULL ,"
'tmp = tmp & "   CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & "   CreatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'
'Call TmpTable
'
'
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Txn_IncomeDetails ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_IncomeDetails_Concurrency DEFAULT (getdate()) FOR Concurrency "
'   Call TmpTable
'
'End If
'
'
'tmp = " CREATE TABLE Txn_TelephoneTemp ("
'tmp = tmp & " BillMonthYear varchar (10)  NULL ,"
'tmp = tmp & " EmpNo varchar (25)  NULL ,"
'tmp = tmp & " Name varchar (150)  NULL ,"
'tmp = tmp & " Designation varchar (100)  NULL ,"
'tmp = tmp & " MobNo varchar (50)  NULL ,"
'tmp = tmp & " AccountNo varchar (50)  NULL ,"
'tmp = tmp & " Eligibility varchar (50)  NULL ,"
'tmp = tmp & " MonthlyBillAmt decimal(18, 2) NULL ,"
'tmp = tmp & " Excess decimal(18, 2) NULL ,"
'tmp = tmp & " NCMSLBillAmt decimal(18, 2) NULL )"
'
'Call TmpTable
'
'
'tmp = " CREATE TABLE Txn_TelephoneDetails ("
'tmp = tmp & " BillMonthYear varchar (15)  NULL ,"
'tmp = tmp & " EmpNo varchar (25)  NULL ,"
'tmp = tmp & " Name varchar (150)  NULL ,"
'tmp = tmp & " Designation varchar (100)  NULL ,"
'tmp = tmp & " MobNo varchar (50)  NULL ,"
'tmp = tmp & " AccountNo varchar (50)  NULL ,"
'tmp = tmp & " Eligibility varchar (50)  NULL ,"
'tmp = tmp & " MonthlyBillAmt decimal(18, 2) NULL ,"
'tmp = tmp & " Excess decimal(18, 2) NULL ,"
'tmp = tmp & " NCMSLBillAmt decimal(18, 2) NULL ,"
'tmp = tmp & " G_UID varchar (36)  NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy varchar (30)  NULL ,"
'tmp = tmp & " CreatedDate datetime NULL ,"
'tmp = tmp & " UpdatedDate datetime NULL ,"
'tmp = tmp & " UpdatedBy varchar (30)  NULL)"
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Txn_TelephoneDetails ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_TelephoneDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'End If
'
'
'
'tmp = "Alter Table Mst_Employee add " & c & " AllocateExpense " & TextType & " (1)"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) VALUES(72, 'Frm_DPUpload.frm', 'DP Upload', 'Frm_DPUpload', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) VALUES(73, 'Frm_IncomeDetails.frm', 'Income Details Upload', 'Frm_IncomeDetails', 'Y', 'C',NULL,NULL,NULL,NULL,NULL )"
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail (MenuID, FormFileName, MenuName, FormName, Flag, HMenuFlag, CreatedBy, CreatedDate, UpdatedDate, UpdatedBy, G_UID) VALUES(75, 'Frm_TelephoneUpload.frm', 'Telephone Upload', 'Frm_TelephoneUpload', 'Y', 'C',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'UpdateConfig ("29")
'MsgBox "Version 29 Installed successfully"
'
'End Function
'
'
'
'Public Function Version28()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "Alter Table Txn_ClaimSubmission add " & c & " ApprovalRemarks " & TextType & " (200)," & c & " ApprovalDate DateTime "
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_EmpCTC ("
'tmp = tmp & " EmployeeCode varchar (25)  NOT NULL ,"
'tmp = tmp & " Category varchar (50)  NOT NULL ,"
'tmp = tmp & " Department varchar (150)  NOT NULL ,"
'tmp = tmp & " ProfessionalFees_Monthly decimal(12, 2) NOT NULL ,"
'tmp = tmp & " CTC_Yearly decimal(12, 2) NOT NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " CTC_Monthly decimal(12, 2) NOT NULL ), "
'   tmp = " PRIMARY KEY (EmployeeCode))"
'Else
'   tmp = tmp & " CTC_Monthly decimal(12, 2) NOT NULL )"
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE dbo.Txn_EmpCTC WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_EmpCTC PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (EmployeeCode)"
'End If
'
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(71, 'Frm_EmployeeCTC.frm', 'Employee CTC Upload', 'Frm_EmployeeCTC', 'Y', 'A',NULL,NULL,NULL,NULL,NULL ) "
'
'Call TmpTable
'
'
'
'UpdateConfig ("28")
'MsgBox "Version 28 Installed successfully"
'
'End Function
'
'Public Function Version27()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'
'tmp = " CREATE TABLE Mst_AuditPoint ("
'tmp = tmp & " AuditPointID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   AuditPoint " & TextType & " (300)  NOT NULL ,"
'tmp = tmp & "   AuditTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   SegmentTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   RequiredType " & TextType & " (3)  NOT NULL ,"
'tmp = tmp & "   Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & "   G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & "   Concurrency datetime NOT NULL ,"
'tmp = tmp & "   CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & "   CreatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (AuditPointID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_AuditSegment ("
'tmp = tmp & "   SegmentTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   SegmentType " & TextType & " (100)  NOT NULL ,"
'tmp = tmp & "   Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & "   G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & "         Concurrency datetime NOT NULL ,"
'tmp = tmp & "   CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & "         CreatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SegmentTypeID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_AuditType ("
'tmp = tmp & "   AuditTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   AuditType " & TextType & " (200)  NOT NULL ,"
'tmp = tmp & "   Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & "   G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & "   Concurrency datetime NOT NULL ,"
'tmp = tmp & "   CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & "   CreatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (AuditTypeID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_Audit ("
'tmp = tmp & "   AuditID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   AuditDate datetime NOT NULL ,"
'tmp = tmp & "   SegmentTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   GodownID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   DesignationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   EmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   AuditTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & "   G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & "   Concurrency datetime NOT NULL ,"
'tmp = tmp & "   CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & "   CreatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedDate datetime NULL ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (AuditID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Txn_AuditDetails ("
'tmp = tmp & "   AuditDetailsID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   AuditID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   AuditPointID decimal(18, 0) NOT NULL ,"
'tmp = tmp & "   Observation " & TextType & " (7500)  NOT NULL ,"
'tmp = tmp & "   Comment " & TextType & " (7500)  NOT NULL ,"
'tmp = tmp & "   RequiredType " & TextType & " (2)  NOT NULL ,"
'tmp = tmp & "   Required decimal(18, 0) NULL ,"
'tmp = tmp & "   Actual decimal(18, 0) NULL ,"
'tmp = tmp & "   RequiredDate datetime NULL ,"
'tmp = tmp & "   RequiredValue decimal(18, 0) NULL ,"
'tmp = tmp & "   StartDate datetime NOT NULL ,"
'tmp = tmp & "   CloseDate datetime NULL ,"
'tmp = tmp & "   Flag " & TextType & " (1)  NULL ,"
'tmp = tmp & "   G_UID " & TextType & " (36)  NULL ,"
'tmp = tmp & "   Concurrency datetime NOT NULL ,"
'tmp = tmp & "   CreatedBy " & TextType & " (30)  NULL ,"
'tmp = tmp & "   CreatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedDate datetime NULL ,"
'tmp = tmp & "   UpdatedBy " & TextType & " (30)  NULL ) "
'
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_AuditPoint WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_AuditPoint PRIMARY KEY  CLUSTERED (AuditPointID)  "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AuditSegment WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_AuditSegment PRIMARY KEY  CLUSTERED (SegmentTypeID ) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AuditType WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_AuditType PRIMARY KEY  CLUSTERED (AuditTypeID)"
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_Audit WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_Audit PRIMARY KEY  CLUSTERED (AuditID) "
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_AuditPoint WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_AuditPoint_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AuditSegment WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_AuditSegment_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AuditType WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_AuditType_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_Audit WITH NOCHECK ADD  "
'   tmp = tmp & " CONSTRAINT DF_Txn_Audit_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_AuditDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_AuditDetails_Concurrency DEFAULT (getdate()) FOR Concurrency"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AuditPoint ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_AuditPoint_Mst_AuditSegment FOREIGN KEY "
'   tmp = tmp & " (SegmentTypeID) REFERENCES Mst_AuditSegment (SegmentTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_AuditPoint_Mst_AuditType FOREIGN KEY "
'   tmp = tmp & " (AuditTypeID) REFERENCES Mst_AuditType (AuditTypeID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_Audit ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_Audit_Mst_AuditSegment FOREIGN KEY "
'   tmp = tmp & " (SegmentTypeID) REFERENCES Mst_AuditSegment (SegmentTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_Audit_Mst_AuditType FOREIGN KEY "
'   tmp = tmp & " (AuditTypeID) REFERENCES Mst_AuditType (AuditTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_Audit_Mst_Designation FOREIGN KEY "
'   tmp = tmp & " (DesignationID) REFERENCES Mst_Designation (DesignationID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_Audit_Mst_Employee FOREIGN KEY "
'   tmp = tmp & " (EmployeeID) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_Audit_Mst_Godown FOREIGN KEY "
'   tmp = tmp & " (GodownID) REFERENCES Mst_Godown (GodownID)"
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_AuditDetails ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_AuditDetails_Mst_AuditPoint FOREIGN KEY "
'   tmp = tmp & " (AuditPointID) REFERENCES Mst_AuditPoint (AuditPointID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_AuditDetails_Txn_Audit FOREIGN KEY "
'   tmp = tmp & " (AuditID) REFERENCES Txn_Audit (AuditID)"
'
'   Call TmpTable
'
'End If
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(64, 'FrmMst_AuditSegment.frm', 'Audit Segment', 'FrmMst_AuditSegment', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(65, 'FrmMst_AuditType.frm', 'Audit Type', 'FrmMst_AuditType', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(66, 'FrmMst_AuditPoint.frm', 'Audit Point', 'FrmMst_AuditPoint', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(67, 'FrmTxn_AuditDetails.frm', 'Audit Details', 'FrmTxn_AuditDetails', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(68, 'FrmRpt_GodownListing.frm', 'Godown Listing', 'FrmRpt_GodownListing', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(69, 'FrmRpt_ExpenseListing.frm', 'Expense Listing', 'FrmRpt_ExpenseListing', 'Y', 'R',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'UpdateConfig ("27")
'MsgBox "Version 27 Installed successfully"
'
'End Function
'
'
'Public Function Version26()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'tmp = "CREATE TABLE Txn_ClaimSubmission ( "
'tmp = tmp & " ClaimID decimal(18, 0) NOT NULL ,ClaimDate datetime NOT NULL ,ExpenseFrom datetime,ExpenseTo datetime, "
'tmp = tmp & " EmployeeID decimal(18, 0) NOT NULL ,LocationID decimal(18, 0) NOT NULL , SM_Prefix " & TextType & " (5)  NOT NULL ,Amount decimal(18, 2),"
'tmp = tmp & " Remarks " & TextType & " (200)  NOT NULL ,Flag " & TextType & " (2)  ,G_UID " & TextType & " (36)  NOT NULL ,Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,CreatedDate datetime  ,  UpdatedDate datetime  , "
'tmp = tmp & " UpdatedBy " & TextType & " (30) ,UpdatedByFAG " & TextType & " (30), "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedDateFAG datetime, "
'   tmp = " PRIMARY KEY (ClaimID,SM_Prefix))"
'Else
'   tmp = tmp & " UpdatedDateFAG datetime) "
'End If
'
'Call TmpTable
'
'tmp = "CREATE TABLE Txn_ClaimSubmissionDetail (ClaimDetailID decimal(18, 0) NOT NULL , ClaimID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL ,ExpenseTypeID decimal(18, 0) NOT NULL ,Particulars " & TextType & " (250)  ,"
'tmp = tmp & " Amount decimal(18, 2) NOT NULL ,Recoverable " & TextType & " (2) ,ExchangeTypeID decimal(18, 0)  ,RecoverableAmount decimal(18, 2),"
'tmp = tmp & " ClientIDRow decimal(18, 0),Flag " & TextType & " (1) ,Concurrency datetime NOT NULL ,CreatedBy " & TextType & " (30)   ,"
'tmp = tmp & " CreatedDate datetime ,UpdatedDate datetime  , "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (ClaimDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Txn_ClaimSubmission WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_ClaimSubmission PRIMARY KEY  CLUSTERED (ClaimID,SM_Prefix)  "
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_ClaimSubmissionDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Txn_ClaimSubmissionDetail PRIMARY KEY  CLUSTERED (ClaimDetailID)  "
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Txn_ClaimSubmission WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_ClaimSubmission_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ClaimSubmissionDetail WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Txn_ClaimSubmissionDetail_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ClaimSubmission ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ClaimSubmission_Mst_Employee FOREIGN KEY (EmployeeID) REFERENCES Mst_Employee (EmployeeID),"
'   tmp = tmp & " CONSTRAINT FK_Txn_ClaimSubmission_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Txn_ClaimSubmissionDetail ADD "
'   tmp = tmp & " CONSTRAINT FK_Txn_ClaimSubmissionDetail_Mst_ExchangeType FOREIGN KEY (ExchangeTypeID) REFERENCES Mst_ExchangeType"
'   tmp = tmp & " (ExchangeTypeID), CONSTRAINT FK_Txn_ClaimSubmissionDetail_Mst_ExpenseType FOREIGN KEY (ExpenseTypeID) REFERENCES"
'   tmp = tmp & " Mst_ExpenseType (ExpenseTypeID), CONSTRAINT FK_Txn_ClaimSubmissionDetail_Txn_ClaimSubmission FOREIGN KEY "
'   tmp = tmp & " (ClaimID,SM_Prefix) REFERENCES Txn_ClaimSubmission (ClaimID,SM_Prefix)"
'
'   Call TmpTable
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(61, 'FrmTxn_ClaimSubmission.frm', 'Claim Submission', 'FrmTxn_ClaimSubmission', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(62, 'FrmTxn_StorageRateCard.frm', 'Storage Rate Card', 'FrmTxn_StorageRateCard', 'Y', 'T',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_StorageRateCard ("
'tmp = tmp & " SRCID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SRCDate datetime  ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExchangeTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityGroupID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FromBagSize decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ToBagSize decimal(18, 0) NOT NULL ,"
'tmp = tmp & " RatePerBag decimal(18, 0) NOT NULL ,"
'tmp = tmp & " FromDate datetime  ,"
'tmp = tmp & " ToDate datetime  ,"
'tmp = tmp & " TillDate datetime  ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SRCID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_StorageRateCard WITH NOCHECK ADD CONSTRAINT PK_Mst_StorageRateCard PRIMARY KEY  CLUSTERED (SRCID) "
'   Call TmpTable
'   tmp = "ALTER TABLE Mst_StorageRateCard WITH NOCHECK ADD CONSTRAINT DF_Mst_StorageRateCard_Concurrency DEFAULT (getdate()) FOR "
'   tmp = tmp & " Concurrency "
'   Call TmpTable
'   tmp = "ALTER TABLE Mst_StorageRateCard ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_StorageRateCard_Mst_CommodityGroup FOREIGN KEY (CommodityGroupID) REFERENCES Mst_CommodityGroup(CommodityGroupID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_StorageRateCard_Mst_ExchangeType FOREIGN KEY (ExchangeTypeID) REFERENCES Mst_ExchangeType(ExchangeTypeID),"
'   tmp = tmp & " CONSTRAINT FK_Mst_StorageRateCard_Mst_Location FOREIGN KEY (LocationID) REFERENCES Mst_Location (LocationID)"
'   Call TmpTable
'End If
'
'
'UpdateConfig ("26")
'MsgBox "Version 26 Installed successfully"
'
'End Function
'Public Function Version25()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_CommodityGroup ("
'tmp = tmp & " CommodityGroupID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CommodityGroup " & TextType & " (200) NOT NULL ,"
'tmp = tmp & " NavCode " & TextType & " (25)  ,"
'tmp = tmp & " G_UID " & TextType & " (36)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (CommodityGroupID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_CommodityGroup WITH NOCHECK ADD CONSTRAINT PK_Mst_CommodityGroup PRIMARY KEY  CLUSTERED (CommodityGroupID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_CommodityGroup WITH NOCHECK ADD CONSTRAINT DF_Mst_CommodityGroup_Concurrency DEFAULT (getdate()) FOR "
'   tmp = tmp & " Concurrency "
'
'   Call TmpTable
'
'End If
'
'
'tmp = "alter table mst_commodity add " & c & " CommodityGroupID decimal(18, 0) "
'Call TmpTable
'
'tmp = "ALTER TABLE mst_commodity ADD "
'tmp = tmp & " CONSTRAINT FK_mst_commodity_mst_commodityGroup FOREIGN KEY "
'tmp = tmp & " (CommodityGroupID) REFERENCES mst_commodityGroup (CommodityGroupID) ON UPDATE CASCADE "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(59, 'FrmMst_CommodityGroup.frm', 'Commodity Group', 'FrmMst_CommodityGroup', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(60, 'FrmMst_NAVCodeMapping.frm', 'Navision Code Mapping', 'FrmMst_NAVCodeMapping', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'
'tmp = "update Mst_ExchangeType set flag = 'A'"
'Call TmpTable
'
'
'tmp = "Alter Table Mst_Client add " & c & " NAVCode " & TextType & " (25) "
'Call TmpTable
'
'tmp = "Alter Table Mst_Location add " & c & " NAVCode " & TextType & " (25) "
'Call TmpTable
'
'tmp = "Alter Table Mst_Employee add " & c & " NAVCode " & TextType & " (25) "
'Call TmpTable
'
'tmp = "Alter Table Mst_Agent add " & c & " NAVCode " & TextType & " (25) "
'Call TmpTable
'
'UpdateConfig ("25")
'MsgBox "Version 25 Installed successfully"
'End Function
'
'Public Function Version24()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'
'tmp = "CREATE TABLE Mst_ExpenseType ("
'tmp = tmp & " ExpenseTypeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " ExpenseType " & TextType & " (200) NOT NULL ,"
'tmp = tmp & " NavCode " & TextType & " (25)  ,"
'tmp = tmp & " Flag " & TextType & " (1)  ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30)  ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (ExpenseTypeID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_ExpenseType WITH NOCHECK ADD CONSTRAINT PK_Mst_ExpenseType PRIMARY KEY  CLUSTERED (ExpenseTypeID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_ExpenseType WITH NOCHECK ADD CONSTRAINT DF_Mst_ExpenseType_Concurrency DEFAULT (getdate()) FOR "
'   tmp = tmp & " Concurrency "
'
'   Call TmpTable
'
'End If
'
'tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'tmp = tmp & " VALUES(58, 'FrmMst_ExpenseType.frm', 'Expense Type', 'FrmMst_ExpenseType', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'Call TmpTable
'
'
'UpdateConfig ("24")
'MsgBox "Version 24 Installed successfully"
'End Function
'
'Public Function Version23()
''-------Version Script
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = " COLUMN "
'Else
'   TextType = "Varchar"
'   c = " "
'End If
'tmp = "Alter Table Mst_Users add " & c & " EmailID " & TextType & " (100) "
'Call TmpTable
'UpdateConfig ("23")
'MsgBox "Version 23 Installed successfully"
'End Function
'
'Public Function Version22()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
'tmp = "CREATE TABLE Mst_SecurityAgency ( "
'tmp = tmp & " AgencyID decimal(18, 0) NOT NULL , "
'tmp = tmp & " AgencyName " & TextType & " (100)  NOT NULL , "
'tmp = tmp & " ContactPerson " & TextType & " (50), "
'tmp = tmp & " Address " & TextType & " (250) , "
'tmp = tmp & " Mobile " & TextType & " (30)  , "
'tmp = tmp & " Telephone " & TextType & " (30) , "
'tmp = tmp & " Flag char (1)  NOT NULL , "
'tmp = tmp & " Remarks " & TextType & " (250)  , "
'tmp = tmp & " Concurrency datetime NOT NULL , "
'tmp = tmp & " CreatedBy " & TextType & " (30)  , "
'tmp = tmp & " CreatedDate datetime , "
'tmp = tmp & " UpdatedDate datetime  , "
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (AgencyID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
''tmp = tmp & " UpdatedBy " & TextType & " (30))
'
'
'tmp = "CREATE TABLE Mst_SecurityShifts ( "
'tmp = tmp & " SecurityShiftsID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SecurityShifts " & TextType & " (50)  NOT NULL , "
'tmp = tmp & " Concurrency datetime NOT NULL , "
'tmp = tmp & " CreatedBy " & TextType & " (30)  , "
'tmp = tmp & " CreatedDate datetime  , "
'tmp = tmp & " UpdatedDate datetime  , "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SecurityShiftsID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
''tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'
'
'Call TmpTable
'
'
'tmp = "CREATE TABLE Mst_SecurityAssets ( "
'tmp = tmp & " SecurityAssetID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SecurityAsset " & TextType & " (100)  NOT NULL , "
'tmp = tmp & " Concurrency datetime NOT NULL , "
'tmp = tmp & " CreatedBy " & TextType & " (30) , "
'tmp = tmp & " CreatedDate datetime , "
'tmp = tmp & " UpdatedDate datetime , "
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SecurityAssetID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
''tmp = tmp & " UpdatedBy " & TextType & " (30))
'
'
'tmp = "CREATE TABLE Mst_SecurityDetails ( "
'tmp = tmp & " SecurityDetailID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SecurityName " & TextType & " (150)  NOT NULL , "
'tmp = tmp & " Designation " & TextType & " (25) , "
'tmp = tmp & " SuperwiserID decimal(18, 0) , "
'tmp = tmp & " SecurityAgencyID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SecurityTypeID decimal(18, 0), "
'tmp = tmp & " Amount decimal(18, 0) NOT NULL , "
'tmp = tmp & " ShiftTimings " & TextType & " (30)  NOT NULL , "
'tmp = tmp & " Flag " & TextType & " (1)  NOT NULL , "
'tmp = tmp & " Concurrency datetime NOT NULL , "
'tmp = tmp & " CreatedBy " & TextType & " (30) , "
'tmp = tmp & " CreatedDate datetime  , "
'tmp = tmp & " UpdatedDate datetime  , "
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SecurityDetailID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
''tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'
'
'tmp = "CREATE TABLE Mst_SecurityCMPMap ( "
'tmp = tmp & " SecurityCMPMapID decimal(18, 0) NOT NULL , "
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL , "
'tmp = tmp & " CMPID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SecurityDetailID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SM_Prefix " & TextType & " (5) , "
'tmp = tmp & " Flag " & TextType & " (1) , "
'tmp = tmp & " Concurrency datetime NOT NULL , "
'tmp = tmp & " CreatedBy " & TextType & " (30)  , "
'tmp = tmp & " CreatedDate datetime  , "
'tmp = tmp & " UpdatedDate datetime  , "
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SecurityCMPMapID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
''tmp = tmp & " UpdatedBy " & TextType & " (30))  "
'
'Call TmpTable
'
'tmp = "CREATE TABLE Mst_SecurityAssetMap ( "
'tmp = tmp & " SecurityAssetMapID decimal(18, 0) NOT NULL , "
'tmp = tmp & " SecurityAssetID decimal(18, 0) NOT NULL , "
'tmp = tmp & " AssetStatus " & TextType & " (75)  NOT NULL , "
'tmp = tmp & " VerifiedOn datetime NOT NULL , "
'tmp = tmp & " Description " & TextType & " (500)  NOT NULL , "
'tmp = tmp & " SecurityDetailID decimal(18, 0) NOT NULL , "
'tmp = tmp & " Flag " & TextType & " (1) , "
'tmp = tmp & " Concurrency datetime NOT NULL , "
'tmp = tmp & " CreatedBy " & TextType & " (30) , "
'tmp = tmp & " CreatedDate datetime  , "
'tmp = tmp & " UpdatedDate datetime  , "
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (SecurityAssetMapID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_SecurityAgency WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_Agency PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (AgencyID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityAssetMap WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SecurityAssetMap PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SecurityAssetMapID) "
'
'   Call TmpTable
'
'
'   tmp = "ALTER TABLE Mst_SecurityAssets WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SecurityAssets PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SecurityAssetID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SecurityDetails PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SecurityDetailID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityCMPMap WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SecurityCMPMap PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SecurityCMPMapID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityShifts WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_SecurityShifts PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (SecurityShiftsID) "
'
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityAgency WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_Agency_Concurrency DEFAULT (getdate()) FOR Concurrency "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityAssetMap WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SecurityAssetMap_Concurrency DEFAULT (getdate()) FOR Concurrency "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityAssets WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SecurityAssets_Concurrency DEFAULT (getdate()) FOR Concurrency "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityDetails WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SecurityDetails_Concurrency DEFAULT (getdate()) FOR Concurrency "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityCMPMap WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SecurityCMPMap_Concurrency DEFAULT (getdate()) FOR Concurrency "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityShifts WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_SecurityShifts_Concurrency DEFAULT (getdate()) FOR Concurrency "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityAssetMap ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SecurityAssetMap_Mst_SecurityAssets FOREIGN KEY "
'   tmp = tmp & " (SecurityAssetID) REFERENCES Mst_SecurityAssets (SecurityAssetID), "
'   tmp = tmp & " CONSTRAINT FK_Mst_SecurityAssetMap_Mst_SecurityDetails FOREIGN KEY "
'   tmp = tmp & " (SecurityDetailID) REFERENCES Mst_SecurityDetails (SecurityDetailID) "
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityDetails ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SecurityDetails_Mst_SecurityAgency FOREIGN KEY "
'   tmp = tmp & " (SecurityAgencyID) REFERENCES Mst_SecurityAgency (AgencyID) ON UPDATE CASCADE "
'
'
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_SecurityCMPMap ADD "
'   tmp = tmp & " CONSTRAINT FK_Mst_SecurityCMPMap_Mst_CMP FOREIGN KEY "
'   tmp = tmp & " (CMPID) REFERENCES Mst_CMP (CMPID), "
'   tmp = tmp & " CONSTRAINT FK_Mst_SecurityCMPMap_Mst_Location FOREIGN KEY "
'   tmp = tmp & " (LocationID) REFERENCES Mst_Location (LocationID) "
'
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(54, 'FrmMst_SecurityAgency.frm', 'Security Agency', 'FrmMst_SecurityAgency', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(55, 'FrmMst_SecurityDetails.frm', 'Security Details', 'FrmMst_SecurityDetails', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(56, 'FrmMst_SecurityAsset.frm', 'Security Assets', 'FrmMst_SecurityAsset', 'Y', 'M',NULL,NULL,NULL,NULL,NULL ) "
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(57, 'FrmMst_SecurityShifts.frm', 'Security Shifts', 'FrmMst_SecurityShifts', 'Y', 'M',NULL,NULL,NULL,NULL,NULL )"
'   Call TmpTable
'
'End If
'
'UpdateConfig ("22")
'MsgBox "Version 22 Installed successfully"
'End Function
'
'
'
'Public Function Version21()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
''Table Mst_Hologram
'tmp = "alter table Mst_Godown add " & c & " NCDEXAccreditedQty Decimal(18,3) "
'Call TmpTable
'
''NCDEXAccreditedQty
'UpdateConfig ("21")
'MsgBox "Version 21 Installed successfully"
'End Function
'
'Public Function Version20()
'
'If Gen_DBType <> "MDB" Then
'
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(49, 'FrmMst_Designation.frm', 'Designation', 'FrmMst_Designation', 'Y', 'M',NULL,NULL,NULL,NULL,'7543339123864497A7A8D2A8A7B226D8' ) "
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(50, 'FrmMst_AssetType.frm', 'Asset Type', 'FrmMst_AssetType', 'Y', 'M',NULL,NULL,NULL,NULL,'4F94C1542DBC44FF8DFF58EB2688C8FF' )"
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(51, 'FrmMst_AssetStatus.frm', 'Asset Status', 'FrmMst_AssetStatus', 'Y', 'M',NULL,NULL,NULL,NULL,'5DD3D724958F469EA08E9F230070C500' ) "
'   Call TmpTable
'
'   tmp = "INSERT INTO Mst_MenuDetail ([MenuID], [FormFileName], [MenuName], [FormName], [Flag], [HMenuFlag], [CreatedBy], [CreatedDate], [UpdatedDate], [UpdatedBy], [G_UID])"
'   tmp = tmp & " VALUES(52, 'FrmMst_Employee.frm', 'Employee', 'FrmMst_Employee', 'Y', 'M',NULL,NULL,NULL,NULL,'422A2F0704634C589554A1B6FBCC97FE' )"
'   Call TmpTable
'End If
'
'UpdateConfig ("20")
'MsgBox "Version 20 Installed successfully"
'
'End Function
'
'
'Public Function Version19()
'
'Dim TextType, c As Variant
'
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
''Table Mst_AssetStatus
'tmp = "CREATE TABLE Mst_AssetStatus ( "
'tmp = tmp & " AssetsStatusID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " AssetStatus " & TextType & " (100)  NOT NULL , "
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (AssetsStatusID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
''Table Mst_AssetsType
'tmp = "CREATE TABLE Mst_AssetsType ("
'tmp = tmp & " AssetsID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " AssetType " & TextType & " (150)  NOT NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (AssetsID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
''Table Mst_Designation
'
'tmp = "CREATE TABLE Mst_Designation ("
'tmp = tmp & " DesignationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " Designation " & TextType & " (30)  NOT NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (DesignationID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
''Table Mst_Employee
'tmp = " CREATE TABLE Mst_Employee ("
'tmp = tmp & " EmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " EmployeeNo " & TextType & " (25)  NOT NULL ,"
'tmp = tmp & " EmployeeName " & TextType & " (100)  NOT NULL ,"
'tmp = tmp & " DesignationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SelfReporting " & TextType & " (1)  ,"
'tmp = tmp & " DirectReportingTo decimal(18, 0)  ,"
'tmp = tmp & " ReportingToDRC decimal(18, 0)  ,"
'tmp = tmp & " ConnectedToRAG decimal(18, 0)  ,"
'tmp = tmp & " EmployeeEmailID " & TextType & " (100)  ,"
'tmp = tmp & " MobileNo " & TextType & " (50)   ,"
'tmp = tmp & " Address " & TextType & " (250)   ,"
'tmp = tmp & " Flag " & TextType & " (1)   ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (EmployeeID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'Call TmpTable
'
''Table Mst_EmpAssetMap
'tmp = "CREATE TABLE Mst_EmpAssetMap ("
'tmp = tmp & " EmpAssetMapID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " EmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " AssetID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " AssetQty decimal(18, 0) ,"
'tmp = tmp & " Description " & TextType & " (500) ,"
'tmp = tmp & " AssetStatusID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " VerifiedDate datetime NOT NULL ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30),"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " UpdatedBy " & TextType & " (30), "
'   tmp = " PRIMARY KEY (EmpAssetMapID))"
'Else
'   tmp = tmp & " UpdatedBy " & TextType & " (30)) "
'End If
'
'
'Call TmpTable
'
''Table Mst_EmpCMPMapping
'tmp = "CREATE TABLE Mst_EmpCMPMapping ("
'tmp = tmp & " EmpCMPMapID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " LocationID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " CMPID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " EmployeeID decimal(18, 0) NOT NULL ,"
'tmp = tmp & " SM_Prefix " & TextType & " (5)  NOT NULL ,"
'tmp = tmp & " Flag char (1) ,"
'tmp = tmp & " Concurrency datetime NOT NULL ,"
'tmp = tmp & " CreatedBy " & TextType & " (30) ,"
'tmp = tmp & " CreatedDate datetime  ,"
'tmp = tmp & " UpdatedDate datetime  ,"
'tmp = tmp & " UpdatedBy " & TextType & " (30) ,"
'
'
'If Gen_DBType = "MDB" Then
'   tmp = tmp & " DesignationID decimal(18, 0),"
'   tmp = " PRIMARY KEY (EmpCMPMapID))"
'Else
'   tmp = tmp & " DesignationID decimal(18, 0)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'
'   tmp = "ALTER TABLE Mst_AssetStatus WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_AssetStatus PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (AssetsStatusID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AssetsType WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_AssetsType PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (AssetsID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Designation WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT PK_Mst_Designation PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (DesignationID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Employee WITH NOCHECK ADD "
'   tmp = tmp & "CONSTRAINT PK_Mst_Employee PRIMARY KEY  CLUSTERED "
'   tmp = tmp & "(EmployeeID)"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpAssetMap WITH NOCHECK ADD "
'   tmp = tmp & "CONSTRAINT PK_Mst_EmpAssetMap PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " (EmpAssetMapID) "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpCMPMapping WITH NOCHECK ADD "
'   tmp = tmp & "CONSTRAINT PK_Mst_EmpGodownMapping PRIMARY KEY  CLUSTERED "
'   tmp = tmp & "(EmpCMPMapID)"
'   Call TmpTable
'
'End If
'
'
'If Gen_DBType <> "MDB" Then
'   tmp = "ALTER TABLE Mst_AssetStatus WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_AssetStatus_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_AssetsType WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_AssetsType_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Designation WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_Designation_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Employee WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_Employee_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpAssetMap WITH NOCHECK ADD"
'   tmp = tmp & " CONSTRAINT DF_Mst_EmpAssetMap_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpCMPMapping WITH NOCHECK ADD "
'   tmp = tmp & " CONSTRAINT DF_Mst_EmpGodownMapping_Concurrency DEFAULT (getdate()) FOR Concurrency"
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_EmpAssetMap ADD"
'   tmp = tmp & " CONSTRAINT FK_Mst_EmpAssetMap_Mst_AssetStatus FOREIGN KEY "
'   tmp = tmp & " (AssetStatusID) REFERENCES Mst_AssetStatus (AssetsStatusID) ON UPDATE CASCADE ,"
'   tmp = tmp & " CONSTRAINT FK_Mst_EmpAssetMap_Mst_AssetsType FOREIGN KEY (AssetID) REFERENCES Mst_AssetsType ( "
'   tmp = tmp & " AssetsID) ON UPDATE CASCADE ,CONSTRAINT FK_Mst_EmpAssetMap_Mst_Employee FOREIGN KEY (EmployeeID "
'   tmp = tmp & " ) REFERENCES Mst_Employee (EmployeeID) ON UPDATE CASCADE "
'   Call TmpTable
'
'   tmp = "ALTER TABLE Mst_Employee ADD  CONSTRAINT FK_Mst_Employee_Mst_Designation FOREIGN KEY (DesignationID)"
'   tmp = tmp & " REFERENCES Mst_Designation (DesignationID) ON UPDATE CASCADE "
'   Call TmpTable
'
'End If
'
'
'UpdateConfig ("19")
'MsgBox "Version 19 Installed successfully"
'
'End Function
'
'
'
'Public Function Version18()
''-------Version Script
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Date"
'   c = " COLUMN "
'Else
'   TextType = "smalldatetime"
'   c = " "
'End If
'tmp = "Alter Table Txn_CashReceipt add " & c & " AgentID Decimal(18,0) "
'Call TmpTable
'UpdateConfig ("18")
'MsgBox "Version 18 Installed successfully"
'End Function
'
'Public Function Version17()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
''Table Mst_Hologram
'tmp = "alter table Mst_Hologram add " & c & " SM_Prefix " & TextType & "(5) "
'Call TmpTable
'
''NCDEXAccreditedQty
'UpdateConfig ("17")
'MsgBox "Version 17 Installed successfully"
'End Function
'Public Function Version16()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
''Table Txn_CashReceiptInvoiceDetails
'tmp = " CREATE TABLE Help_Detail( "
'tmp = tmp & "  FormName " & TextType & " (75) NOT Null, "
'tmp = tmp & "  FieldName " & TextType & " (75) NOT Null, "
'tmp = tmp & "  DisplayFieldName " & TextType & " (100) NOT Null, "
'If Gen_DBType = "MDB" Then
'    tmp = tmp & "  EngDescription " & TextType & " (255) "
'
'Else
'    tmp = tmp & "  EngDescription " & TextType & " (7000)) "
'End If
'
'Call TmpTable
'
'UpdateConfig ("16")
'MsgBox "Version 16 Installed successfully"
'End Function
'
'
'Public Function Version15()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
'tmp = "alter table Txn_CxTF_Details add " & c & " ManualCDTFNo " & TextType & "(18) "
'Call TmpTable
'
''If Gen_DBType = "MDB" Then
''   TextType = "Date"
''   c = "COLUMN"
''Else
''   TextType = "DateTime"
''   c = ""
''End If
'''tmp = "alter table Txn_GSLAdjTransaction add " & c & " TxnGSLAdjDate " & TextType
'''Call TmpTable
'
'
'UpdateConfig ("15")
'MsgBox "Version 15 Installed successfully!!!"
'End Function
'
'Public Function Version14()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
''Table Txn_CashReceiptInvoiceDetails
'tmp = " CREATE TABLE Txn_CashReceiptInvoiceDetails( "
'tmp = tmp & "  CRIDID decimal (18,0) NOT Null, "
'tmp = tmp & "  CRID decimal (18,0)NOT Null, "
'tmp = tmp & "  SM_Prefix " & TextType & " (5) NOT Null, "
'tmp = tmp & "  PI_No " & TextType & " (50), "
'tmp = tmp & "  PI_Date  datetime , "
'tmp = tmp & "  ClientIDRow decimal (18,0) , "
'tmp = tmp & "  Amount decimal (18,2) , "
'tmp = tmp & "  Flag " & TextType & " (1), "
'
'If Gen_DBType = "MDB" Then
'    tmp = tmp & "  InvoiceDetail " & TextType & " (255) "
'    tmp = " PRIMARY KEY (CRIDID,SM_Prefix))"
'Else
'    tmp = tmp & "  InvoiceDetail " & TextType & " (255)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_CashReceiptInvoiceDetails WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Txn_CashReceiptInvoiceDetails] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [CRIDID],[SM_Prefix] ) ON [PRIMARY]  "
'   Call TmpTable
'End If
'
''Table Txn_AdjTransaction
'tmp = " CREATE TABLE Txn_AdjTransaction( "
'tmp = tmp & "  TxnID decimal (18,0) NOT NULL, "
'tmp = tmp & "  TxnDate  datetime NOT NULL, "
'tmp = tmp & "  CMPID decimal (18,0) NOT NULL, "
'tmp = tmp & "  SM_Prefix " & TextType & " (5) NOT Null, "
'tmp = tmp & "  CommodityID decimal (18,0) NOT NULL , "
'tmp = tmp & "  TxnType " & TextType & " (2), "
'tmp = tmp & "  Amount decimal (18,0) NOT NULL, "
'tmp = tmp & "  Flag " & TextType & " (1), "
'tmp = tmp & "  AgentID decimal (18,0) NOT NULL, "
''tmp = tmp & "   " & TextType & " (1), "
'
''tmp = tmp & "  G_UID " & TextType & " (1), "
'
'If Gen_DBType = "MDB" Then
'    tmp = tmp & "  G_UID " & TextType & " (36), "
'    tmp = " PRIMARY KEY (TxnID))"
'Else
'    tmp = tmp & "  G_UID " & TextType & " (36)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_AdjTransaction WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Txn_AdjTransaction] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [TxnID]) ON [PRIMARY]  "
'   Call TmpTable
'End If
'
''Table Txn_GSLAdjTransaction
'tmp = " CREATE TABLE Txn_GSLAdjTransaction( "
'tmp = tmp & "  TxnGSLAdjID decimal (18,0) NOT NULL, "
'tmp = tmp & "  TxnGSLAdjDat  datetime NOT NULL, "
'tmp = tmp & "  TxnID decimal (18,0) NOT NULL, "
'tmp = tmp & "  GSLID decimal (18,0) , "
'tmp = tmp & "  NoOfBags decimal (18,0) , "
'tmp = tmp & "  Quantity decimal (18,0) , "
'tmp = tmp & "  Flag " & TextType & " (1), "
'tmp = tmp & "  G_UID " & TextType & " (36), "
'
'
'
'If Gen_DBType = "MDB" Then
'    tmp = tmp & "  SpillageFlag " & TextType & " (1), "
'    tmp = " PRIMARY KEY (TxnGSLAdjID))"
'Else
'    tmp = tmp & "  SpillageFlag " & TextType & " (1)) "
'End If
'
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_GSLAdjTransaction WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Txn_GSLAdjTransaction] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [TxnGSLAdjID]) ON [PRIMARY]  "
'   Call TmpTable
'End If
'
'
'UpdateConfig ("14")
'MsgBox "Version 14 Installed successfully"
'End Function
'
'
'Public Function Version13()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
'tmp = "alter table Txn_WHR_Detail add SaidToContain " & TextType & "(1), EstimatedWeight " & TextType & "(1)"
'Call TmpTable
'
'tmp = " CREATE TABLE Txn_BenposDateWise( "
'tmp = tmp & "  ISINID " & TextType & " (25), "
'tmp = tmp & "  ImportDate datetime, "
'tmp = tmp & "  Qty  decimal (18,4)) "
'
'Call TmpTable
'UpdateConfig ("13")
'MsgBox "Version 13 Installed successfully"
'End Function
'Public Function Version12()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'   tmp = " CREATE TABLE Mst_DR( "
'   tmp = tmp & "  DRID decimal (18,0) NOT Null, "
'   tmp = tmp & "  DRNo decimal (18,0), "
'   tmp = tmp & "  DRDate datetime, "
'   tmp = tmp & "  Txn_Flag " & TextType & " (2), "
'   tmp = tmp & "  OldISINID " & TextType & " (25), "
'   tmp = tmp & "  NewISINID " & TextType & " (25), "
'   tmp = tmp & "  ClientIDRow " & TextType & " (36), "
'   tmp = tmp & "  Net_Weight_d  decimal (18,3), "
'   tmp = tmp & "  Net_Weight_e  decimal (18,3), "
'   tmp = tmp & "  Remark " & TextType & " (255) , "
'   tmp = tmp & "  Status " & TextType & " (200) , "
'   tmp = tmp & "  G_UID " & TextType & " (36) ) "
'   Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_DR WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Mst_DR] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [DRID] ) ON [PRIMARY]  "
'   Call TmpTable
'
'End If
'
'
'
'tmp = " CREATE TABLE Txn_GSL_Transfer( "
'tmp = tmp & "  GSLTransferID decimal (18) NOT Null  , "
'tmp = tmp & "  GSLTransferDate datetime  , "
'tmp = tmp & "  NewGSLID decimal(18)  , "
'tmp = tmp & "  OLDGSLID decimal(18)  , "
'tmp = tmp & "  SM_Prefix " & TextType & " (5) , "
'tmp = tmp & "  G_UID " & TextType & " (36) , "
'tmp = tmp & "  Flag " & TextType & " (1) , "
'tmp = tmp & "  Remarks " & TextType & " (200) ) "
'Call TmpTable
'
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Txn_GSL_Transfer WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Txn_GSL_Transfer] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [GSLTransferID] ) ON [PRIMARY]  "
'   Call TmpTable
'
'End If
'
'
'tmp = "alter table Mst_GSL add " & c & " BGSLID decimal (18)," & c & " FGSLID decimal (18)"
'Call TmpTable
'
'
'
'UpdateConfig ("12")
'
'MsgBox "Version 12 Installed successfully"
'
'End Function
'
'Public Function Version11()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Date"
'   c = "COLUMN"
'Else
'   TextType = "DateTime"
'   c = ""
'End If
'tmp = "alter table Mst_GSL add " & c & " KRDate " & TextType
'Call TmpTable
'
'UpdateConfig ("11")
'MsgBox "Version 11 Installed successfully"
'End Function
'
'Public Function Version10()
'
'tmp = "alter table Txn_CxTF_Details add AgentID decimal (10,0) "
'Call TmpTable
'
'UpdateConfig ("10")
'MsgBox "Version 10 Installed successfully"
'End Function
'
'Public Function Version9()
'
'Dim TextType, c As Variant
'
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
'   tmp = " CREATE TABLE Mst_Agent( "
'   tmp = tmp & "  AgentID decimal (10,0) NOT Null  , "
'   tmp = tmp & "  AgentName " & TextType & " (200)  , "
'   tmp = tmp & "  Address " & TextType & " (255)  , "
'   tmp = tmp & "  ContactNo " & TextType & " (15)  , "
'   tmp = tmp & "  MobileNo " & TextType & " (15) , "
'   tmp = tmp & "  EmailID " & TextType & " (100)  , "
'   tmp = tmp & "  Fax " & TextType & " (15) , "
'   tmp = tmp & "  G_UID " & TextType & " (36) ) "
'   Call TmpTable
'If Gen_DBType <> "MDB" Then
'   tmp = " ALTER TABLE Mst_Agent WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Mst_Agent] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [AgentID] ) ON [PRIMARY]  "
'   Call TmpTable
'
'End If
'
'
'UpdateConfig ("9")
'MsgBox "Version 9 Installed successfully"
'
'End Function
'
'Public Function Version8()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'tmp = " CREATE TABLE Log_Print_History ( "
'tmp = tmp & "   SM_Prefix " & TextType & "(5)  , "
'tmp = tmp & "   DocType " & TextType & "(10)  , "
'tmp = tmp & "   DocNo decimal(18, 0)  , "
'tmp = tmp & "   FileName " & TextType & " (100)  , "
'tmp = tmp & "   UserId " & TextType & " (15)  , "
'tmp = tmp & "   Concurrency datetime Not Null ) "
'Call TmpTable
'
'If Gen_DBType = "MDB" Then
'   TextType = "Date"
'   c = "COLUMN"
'Else
'   TextType = "datetime"
'   c = ""
'End If
'tmp = "alter table Mst_GSL  add " & c & " LCDTF_Date " & TextType
'Call TmpTable
'tmp = "alter table Mst_GSL  add " & c & " GrossWt decimal(18, 3)  "
'Call TmpTable
'tmp = "alter table Mst_GSL  add " & c & " KRNo decimal(18, 0)  "
'Call TmpTable
'tmp = "alter table Mst_GSL  add " & c & " KRNoBatchID decimal(18, 0)  "
'Call TmpTable
'
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'   tmp = " CREATE TABLE Txn_GSL_OpeningBalance ( "
'   tmp = tmp & "   GSLID decimal (18) Not Null  , "
'   tmp = tmp & "   CMPID decimal (18)  , "
'   tmp = tmp & "   GodownID decimal (18)  , "
'   tmp = tmp & "   ExchangeTypeID decimal(18)  , "
'   tmp = tmp & "   CommodityID decimal(18)  , "
'   tmp = tmp & "   ISINID " & TextType & " (25)  , "
'   tmp = tmp & "   ClientIDRow decimal(18), "
'   tmp = tmp & "   StackNo " & TextType & " (10), "
'   tmp = tmp & "   LotNo " & TextType & " (10), "
'   tmp = tmp & "   DepositBags decimal(18), "
'   tmp = tmp & "   DepositWeight decimal(18,3) , "
'   tmp = tmp & "   SM_Prefix " & TextType & " (5) Not Null "
'   tmp = tmp & " ) ON [PRIMARY] "
'   Call TmpTable
'
'   tmp = " ALTER TABLE Txn_GSL_OpeningBalance WITH NOCHECK ADD "
'   tmp = tmp & "   CONSTRAINT [PK_Txn_GSL_OpeningBalance] PRIMARY KEY  CLUSTERED "
'   tmp = tmp & " ( [GSLID],[SM_Prefix] ) ON [PRIMARY]  "
'   Call TmpTable
'End If
'UpdateConfig ("8")
'MsgBox "Version 8 Installed successfully"
'End Function
'
'Public Function Version7()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'tmp = "alter table Txn_WHR_Detail add " & c & " PolicyNo " & TextType & "(255)"
'Call TmpTable
'
'tmp = "alter table Txn_WHR_Detail add " & c & " CoverNoteNo " & TextType & "(255)"
'Call TmpTable
'
'tmp = "alter table Mst_CMP add " & c & " CMPAlias " & TextType & "(255)"
'Call TmpTable
'tmp = "alter table Mst_Godown add " & c & " GodownAlias " & TextType & "(255)"
'Call TmpTable
'
'UpdateConfig ("7")
'MsgBox "Version 7 Installed successfully"
'End Function
'Public Function Version6()
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Date"
'   c = "COLUMN"
'Else
'   TextType = "smalldatetime"
'   c = ""
'End If
'tmp = "alter table Mst_Users add " & c & " PassExpiryDate " & TextType & " "
'Call TmpTable
'UpdateConfig ("6")
'MsgBox "Version 5 Installed successfully"
'
'End Function
'
'Public Function Version5()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Date"
'   c = "COLUMN"
'Else
'   TextType = "smalldatetime"
'   c = ""
'End If
'
'tmp = "alter table Txn_CashReceipt add " & c & " HologramNo " & TextType & " "
'Call TmpTable
'
'UpdateConfig ("5")
'MsgBox "Version 5 Installed successfully"
'End Function
'
'Public Function Version3()
'If Gen_DBType = "MDB" Then
'
'Else
'    tmp = " CREATE TABLE Mst_Hologram ( "
'    tmp = tmp & " HologramNo varchar (25) not Null , "
'    tmp = tmp & "   DocType varchar (20)  , "
'    tmp = tmp & "   DocNo decimal(18, 0)  , "
'    tmp = tmp & "   Concurrency datetime  , "
'    tmp = tmp & "   CreatedBy nvarchar (30), "
'    tmp = tmp & "   CreatedDate datetime  , "
'    tmp = tmp & "   UpdatedDate datetime , "
'    tmp = tmp & "   UpdatedOn nvarchar (30)   "
'    tmp = tmp & " ) ON [PRIMARY] "
'    Call TmpTable
'
'    tmp = " ALTER TABLE Mst_Hologram WITH NOCHECK ADD "
'    tmp = tmp & "   CONSTRAINT [PK_Mst_Hologram] PRIMARY KEY  CLUSTERED "
'    tmp = tmp & " ( [HologramNo] ) ON [PRIMARY]  "
'    Call TmpTable
'
'    tmp = "ALTER TABLE Mst_Hologram ADD CONSTRAINT DF_Mst_Hologram_Concurrency DEFAULT (getdate()) FOR Concurrency "
'    Call TmpTable
'End If
'
'UpdateConfig ("3")
'MsgBox "Version 3 Installed successfully"
'
'End Function
'Public Function Version4()
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Date"
'   c = "COLUMN"
'Else
'   TextType = "smalldatetime"
'   c = ""
'End If
'
'tmp = "alter table Mst_CMP add " & c & " CurrentDate " & TextType & " "
'Call TmpTable
'
'
'UpdateConfig ("4")
'MsgBox "Version 4 Installed successfully"
'End Function
'
'
'Public Function Version1()
''-------------------------------------------------------------------------
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'
'tmp = "alter table Mst_Bank add " & c & " G_UID " & TextType & "(36)  "
'Call TmpTable
'tmp = "alter table Mst_Branch add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Category add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Client add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_CMP add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Commodity  add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_CommodityDetail  add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Depository  add " & c & " G_UID " & TextType & "(36)"
'Call TmpTable
'tmp = "alter table Mst_ExchangeType  add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Godown  add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_GSL  add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_InsuranceCoMaster add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_InsuranceDetails add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_ISIN add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_License add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Location add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_MenuDetail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_MenuHasUserRole add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_RRN add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_State add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_StorageContract add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_StoreManLocations add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_UnitType add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_UserRole add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Mst_Users add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CashReceipt add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CDC add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CDC_Assayer_Detail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CDC_GSL_Detail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CDC_Insurance_Detail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CxTF_ClientDetail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CxTF_Details add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CxTF_GSL add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_CxTF_TruckDetail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_ISINGSLRevalidation add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_ISINRevalidation add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_PledgeCreation add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_PledgeCreation_GSL add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_PledgeRelease add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_PledgeRelease_GSL add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_Spillage add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'tmp = "alter table Txn_WHR_Detail add " & c & " G_UID " & TextType & "(36) "
'Call TmpTable
'
'UpdateConfig ("1")
'MsgBox "Version 1 Installed successfully"
'
''-------------------------------------------------------------------------
'End Function
'
'
'Public Function Version2()
''-------------------------------------------------------------------------
'
'Dim TextType, c As Variant
'If Gen_DBType = "MDB" Then
'   TextType = "Text"
'   c = "COLUMN"
'Else
'   TextType = "Varchar"
'   c = ""
'End If
'tmp = "alter table Txn_CxTF_Details add HologramNo " & TextType & "(25)  "
'Call TmpTable
'tmp = "alter table Txn_WHR_Detail add HologramNo " & TextType & "(25)  "
'Call TmpTable
'UpdateConfig ("2")
'MsgBox "Version 2 Installed successfully"
'
''-------------------------------------------------------------------------
'End Function



Public Function VersionUnknown()
MsgBox "Version not known", vbInformation, "Versioning"
End Function

Public Function UpdateConfig(verNum As String)
RsMst_Config!Version_No = verNum
RsMst_Config.Update
End Function




''----------------  Add By Akbar On 01/09/2007----------
''--------------- For to Check File Already Exit or Not to Export at Selectes Location
'Returns a boolean - True if the file exists
Public Function FExists(OrigFile As String)
Dim fs
Set fs = CreateObject("Scripting.FileSystemObject")
FExists = fs.FileExists(OrigFile)
End Function
