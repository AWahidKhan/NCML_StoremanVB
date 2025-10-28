VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIFrmMain 
   BackColor       =   &H00FFC0C0&
   Caption         =   "StoreMan"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   1065
   ClientWidth     =   11400
   Icon            =   "MDIFrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7635
      ScaleWidth      =   3180
      TabIndex        =   1
      Top             =   0
      Width           =   3240
      Begin VB.PictureBox Picture2 
         Height          =   1575
         Left            =   1320
         ScaleHeight     =   1515
         ScaleWidth      =   4830
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   4890
         Begin VB.Label Label1 
            BackColor       =   &H00FF8080&
            Height          =   600
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   4815
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Connected Menu only"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   75
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2520
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   45
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":05A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":0940
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":0CDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":1074
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":26CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":2B20
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":2F72
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":33C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":3816
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":3C68
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":4002
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":4454
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":476E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":4B08
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":4EA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":523C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":568E
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":5A28
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":7082
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":7AE0
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":7DFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":8114
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":842E
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":8748
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":9086
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":93A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":97F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":B704
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":BFDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":C378
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":C712
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":CA2C
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":D48A
               Key             =   ""
            EndProperty
            BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":D7A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":E0E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":E47C
               Key             =   ""
            EndProperty
            BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":E816
               Key             =   ""
            EndProperty
            BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":EB30
               Key             =   ""
            EndProperty
            BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":1018A
               Key             =   ""
            EndProperty
            BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":15450
               Key             =   ""
            EndProperty
            BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":16AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":16EFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":1748E
               Key             =   ""
            EndProperty
            BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "MDIFrmMain.frx":17D68
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   9975
         Left            =   0
         TabIndex        =   2
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   17595
         _Version        =   393217
         LabelEdit       =   1
         Style           =   5
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "ImageList1"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7695
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu Mst 
      Caption         =   "Master                             "
      Visible         =   0   'False
      Begin VB.Menu StateIndia 
         Caption         =   "State India"
      End
      Begin VB.Menu State 
         Caption         =   "State"
      End
      Begin VB.Menu Location 
         Caption         =   "Location"
      End
      Begin VB.Menu CommodityControl 
         Caption         =   "Commodity"
         Begin VB.Menu Commodity 
            Caption         =   "Commodity"
         End
         Begin VB.Menu Commodityvariety 
            Caption         =   "Commodity Variety"
         End
         Begin VB.Menu BagEquivalent 
            Caption         =   "Bags Equivalent"
         End
      End
      Begin VB.Menu Depository 
         Caption         =   "DP"
      End
      Begin VB.Menu UnitType 
         Caption         =   "Unit Type"
      End
      Begin VB.Menu Category 
         Caption         =   "Category"
      End
      Begin VB.Menu Mst_License 
         Caption         =   "License Master"
      End
      Begin VB.Menu CMP 
         Caption         =   "CMP"
      End
      Begin VB.Menu GSLMaster 
         Caption         =   "GSL Master"
      End
      Begin VB.Menu Agent 
         Caption         =   "Agent"
      End
      Begin VB.Menu Client 
         Caption         =   "Client"
         Begin VB.Menu ClientGroup 
            Caption         =   "Client Group"
         End
         Begin VB.Menu ClientMaster 
            Caption         =   "Client Master"
         End
      End
      Begin VB.Menu ISIN 
         Caption         =   "ISIN Master"
      End
      Begin VB.Menu RRNMaster 
         Caption         =   "RRN Master"
      End
      Begin VB.Menu mnuInsurance 
         Caption         =   "Insurance Company"
      End
      Begin VB.Menu mnuInsuranceDetails 
         Caption         =   "Insurance Policy"
      End
      Begin VB.Menu Bank 
         Caption         =   "Bank"
      End
      Begin VB.Menu Branch 
         Caption         =   "Branch"
      End
      Begin VB.Menu FAG 
         Caption         =   "FAG"
         Begin VB.Menu ExpenseType 
            Caption         =   "Expense Type"
         End
         Begin VB.Menu ExpenseGroup 
            Caption         =   "Expense Group"
         End
         Begin VB.Menu ServiceTax 
            Caption         =   "Service Tax"
         End
         Begin VB.Menu TDSMaster 
            Caption         =   "TDS Master"
         End
         Begin VB.Menu VendorMaster 
            Caption         =   "Vendor Master"
         End
         Begin VB.Menu NavisionCodeMapping 
            Caption         =   "Navision Code Mapping"
         End
         Begin VB.Menu FAGBank 
            Caption         =   "FAG Bank"
         End
         Begin VB.Menu IncomeExpenseType 
            Caption         =   "Income/Expense Type"
         End
      End
      Begin VB.Menu Employee 
         Caption         =   "Employee"
         Begin VB.Menu EmployeeMaster 
            Caption         =   "Employee"
         End
         Begin VB.Menu Designation 
            Caption         =   "Designation"
         End
         Begin VB.Menu StateDRCMapping 
            Caption         =   "State DRC Mapping"
         End
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Security"
         Begin VB.Menu SecurityDetails 
            Caption         =   "Security Details"
         End
         Begin VB.Menu SecurityAgency 
            Caption         =   "Security Agency"
         End
         Begin VB.Menu SecurityDesignation 
            Caption         =   "Security Designation"
         End
         Begin VB.Menu SecurityAssets 
            Caption         =   "Security Assets"
            Visible         =   0   'False
         End
         Begin VB.Menu SecurityShifts 
            Caption         =   "Security Shifts"
         End
      End
      Begin VB.Menu Audit 
         Caption         =   "Audit"
         Begin VB.Menu AuditSegment 
            Caption         =   "Audit Segment"
         End
         Begin VB.Menu AuditType 
            Caption         =   "Audit Type"
         End
         Begin VB.Menu AuditPoint 
            Caption         =   "Audit Point"
         End
      End
      Begin VB.Menu BillingCycle 
         Caption         =   "Billing Cycle"
      End
      Begin VB.Menu WHRUnitOfMeasurement 
         Caption         =   "WHR Unit Of Measurement"
      End
      Begin VB.Menu LessorMaster 
         Caption         =   "Lessor Master"
      End
      Begin VB.Menu ITVendor 
         Caption         =   "IT Vendor"
      End
      Begin VB.Menu Quality 
         Caption         =   "Quality"
         Begin VB.Menu QualityParameter 
            Caption         =   "Quality Parameter"
         End
         Begin VB.Menu CommodityQualityParameter 
            Caption         =   "Commodity Quality Parameter"
         End
         Begin VB.Menu TestingAgency 
            Caption         =   "Testing Agency"
         End
      End
   End
   Begin VB.Menu Txn 
      Caption         =   "Transaction                           "
      Visible         =   0   'False
      Begin VB.Menu EndofDay 
         Caption         =   "Beginning of Day (BOD)"
      End
      Begin VB.Menu CMPMonthEndProcess 
         Caption         =   "CMP Month End Process"
      End
      Begin VB.Menu Deposit 
         Caption         =   "Deposit"
      End
      Begin VB.Menu Withdrawal 
         Caption         =   "Withdrawal"
      End
      Begin VB.Menu WearhouseReceipt 
         Caption         =   "Warehouse Receipt "
      End
      Begin VB.Menu StorageResevation 
         Caption         =   "Storage Reservation "
         Begin VB.Menu StorageContract 
            Caption         =   "Storage Reservation"
         End
         Begin VB.Menu SRGM 
            Caption         =   "Godown Mapping"
         End
         Begin VB.Menu SCGSLMapping 
            Caption         =   "GSL Mapping"
         End
      End
      Begin VB.Menu AdjustmentTxn 
         Caption         =   "Adjustment"
      End
      Begin VB.Menu NCDEX 
         Caption         =   "NCDEX"
         Begin VB.Menu ISINMapping 
            Caption         =   "ISIN Mapping"
         End
         Begin VB.Menu ISINRevalidation 
            Caption         =   "ISIN Revalidation"
         End
         Begin VB.Menu PendingISINforRevalidation 
            Caption         =   "Pending ISIN for Revalidation"
         End
      End
      Begin VB.Menu Spot 
         Caption         =   "Spot"
         Begin VB.Menu CDC 
            Caption         =   "CDC "
         End
         Begin VB.Menu CDCTransfer 
            Caption         =   "CDC Payout"
         End
         Begin VB.Menu ReIssueCDC 
            Caption         =   "Re Issue CDC"
         End
      End
      Begin VB.Menu Pledge 
         Caption         =   "Pledge"
         Begin VB.Menu PledgeCreation 
            Caption         =   "Pledge Creation"
         End
         Begin VB.Menu PledgeRelease 
            Caption         =   "Pledge Release"
         End
      End
      Begin VB.Menu Spillage 
         Caption         =   "Spillage  "
      End
      Begin VB.Menu Transfer 
         Caption         =   "Transfer"
         Begin VB.Menu GodownTransfer 
            Caption         =   "Godown Transfer"
         End
         Begin VB.Menu GSLTransfer 
            Caption         =   "GSL Tranfer"
         End
      End
      Begin VB.Menu CashReceiptmenu 
         Caption         =   "Cash Receipt"
         Begin VB.Menu CashReceipt 
            Caption         =   "Issue of Cash Receipt"
         End
         Begin VB.Menu CRDepositDetails 
            Caption         =   "Details of cash/cheques deposited in Bank"
         End
         Begin VB.Menu CRSightedDetailsinBankStatement 
            Caption         =   "CR Sighted Details in Bank Statement"
         End
         Begin VB.Menu CITFCRMapping 
            Caption         =   "Storage Charges received against withdrawals"
         End
         Begin VB.Menu OldCashReceipt 
            Caption         =   "Old Cash Receipt"
         End
      End
      Begin VB.Menu RateCardMaster 
         Caption         =   "Rate Card"
         Begin VB.Menu NonNCDEX 
            Caption         =   "Non-NCDEX"
            Begin VB.Menu StorageRateCard 
               Caption         =   "Base Storage Rate Card"
            End
            Begin VB.Menu SpecialRateCard 
               Caption         =   "Special Storage Rate Card"
            End
         End
         Begin VB.Menu NCDEXRate 
            Caption         =   "NCDEX"
            Begin VB.Menu NCDEXStorageRateCard 
               Caption         =   "NCDEX Storage Rate Card"
            End
         End
      End
      Begin VB.Menu ClaimSubmission 
         Caption         =   "Claim Submission"
         Begin VB.Menu EmployeeClaim 
            Caption         =   "Employee"
         End
         Begin VB.Menu VendorClaim 
            Caption         =   "Vendor"
         End
         Begin VB.Menu ITVendorClaim 
            Caption         =   "IT Vendor Claim"
         End
      End
      Begin VB.Menu AuditDetails 
         Caption         =   "Audit Details"
      End
      Begin VB.Menu LessorAgreement 
         Caption         =   "Lessor Agreement"
      End
      Begin VB.Menu GodownRent 
         Caption         =   "Godown Rent"
         Begin VB.Menu Process 
            Caption         =   "Process"
         End
         Begin VB.Menu RentApproval 
            Caption         =   "Rent Approval"
         End
         Begin VB.Menu PaymentDetailUpload 
            Caption         =   "Payment Detail Upload"
         End
         Begin VB.Menu GodownRentCoverLetter 
            Caption         =   "Godown Rent CoverLetter"
         End
         Begin VB.Menu FranchiseeGodownRateCard 
            Caption         =   "Franchisee Godown Rate Card"
         End
      End
      Begin VB.Menu Invoice 
         Caption         =   "Invoice"
         Begin VB.Menu NNNR 
            Caption         =   "NNNR"
            Begin VB.Menu StockOpeningBalance 
               Caption         =   "Stock Opening Balance"
            End
            Begin VB.Menu GenerateInvoice 
               Caption         =   "Generate BOE"
            End
            Begin VB.Menu ProformaInvoice 
               Caption         =   "Approve BOE"
            End
         End
         Begin VB.Menu Reservation 
            Caption         =   "Reservation"
            Begin VB.Menu GenerateBOE 
               Caption         =   "Generate BOE"
            End
            Begin VB.Menu ApproveBOE 
               Caption         =   "Approve BOE"
            End
         End
         Begin VB.Menu AdHocInvoice 
            Caption         =   "Ad-Hoc Invoice"
         End
      End
      Begin VB.Menu mnuFumigation 
         Caption         =   "Fumigation"
         Begin VB.Menu Fumigation 
            Caption         =   "Fumigation"
         End
         Begin VB.Menu Spraying 
            Caption         =   "Spraying"
         End
         Begin VB.Menu QualityInformation 
            Caption         =   "Quality Information"
         End
      End
      Begin VB.Menu RAGMemberTransfer 
         Caption         =   "RAG Member Transfer"
      End
      Begin VB.Menu Hologram 
         Caption         =   "Hologram"
      End
      Begin VB.Menu HR 
         Caption         =   "HR"
         Begin VB.Menu MonthwiseEmployeeHistory 
            Caption         =   "Month wise Employee History"
         End
         Begin VB.Menu MonthWiseEmployeeProcess 
            Caption         =   "Month Wise Employee Process"
         End
      End
      Begin VB.Menu Security 
         Caption         =   "Security"
         Begin VB.Menu SecurityAttendance 
            Caption         =   "Security Attendance"
         End
         Begin VB.Menu mnuSecurityBilling 
            Caption         =   "Security Billing"
         End
      End
   End
   Begin VB.Menu Reports 
      Caption         =   "Reports                            "
      Visible         =   0   'False
      Begin VB.Menu Listing 
         Caption         =   "CMP Reports"
         Begin VB.Menu DTDReport 
            Caption         =   "DTD Report                         "
         End
         Begin VB.Menu TransactionListing 
            Caption         =   "Transaction Dump"
            Visible         =   0   'False
         End
         Begin VB.Menu TransactionDetailReport 
            Caption         =   "Transaction Detail Report"
         End
         Begin VB.Menu StockSummaryReport 
            Caption         =   "Stock Summary Report"
         End
         Begin VB.Menu PrintPendingList 
            Caption         =   "Print Pending List"
         End
         Begin VB.Menu CMPPV 
            Caption         =   "CMP PV"
         End
         Begin VB.Menu PendingWHRs 
            Caption         =   "Pending WHRs"
         End
      End
      Begin VB.Menu Control 
         Caption         =   "Control"
         Begin VB.Menu IncomeMISReport 
            Caption         =   "Income MIS"
            Visible         =   0   'False
         End
         Begin VB.Menu ExpenseMIS 
            Caption         =   "Expenses MIS"
            Visible         =   0   'False
         End
         Begin VB.Menu ShortagePledgeQuantity 
            Caption         =   "Shortage Pledge Quantity"
         End
         Begin VB.Menu WHRMonitoring 
            Caption         =   "WHR Monitoring"
         End
         Begin VB.Menu ControlReport 
            Caption         =   "Control Report"
         End
         Begin VB.Menu CommodityMapped 
            Caption         =   "Commodity Mapped "
         End
         Begin VB.Menu MissingEmployees 
            Caption         =   "Missing Employees"
         End
         Begin VB.Menu LoginDetails 
            Caption         =   "Login Details"
         End
         Begin VB.Menu MismatchGSLISINLocation 
            Caption         =   "Mismatch GSL ISIN Location"
         End
         Begin VB.Menu ReservationContractStock 
            Caption         =   "Reservation Contract Stock"
         End
         Begin VB.Menu CITFDiscrpancy 
            Caption         =   "CITF Discrpancy"
         End
      End
      Begin VB.Menu RentReports 
         Caption         =   "Rent Reports"
         Begin VB.Menu GodownWithoutAgreement 
            Caption         =   "Godowns without Agreements"
         End
         Begin VB.Menu UnpaidGodownRent 
            Caption         =   "Unprocessed  Godown Rent"
         End
         Begin VB.Menu GodownRentClaimReport 
            Caption         =   "Processed Godown Rent Status Report"
         End
      End
      Begin VB.Menu CMGReport 
         Caption         =   "CMG Report"
         Begin VB.Menu NewDailyMIS 
            Caption         =   "New Daily MIS"
         End
         Begin VB.Menu CMPBODCheck 
            Caption         =   "CMP BOD"
         End
         Begin VB.Menu PendingMonthBeginProcess 
            Caption         =   "Pending Month Begin Process"
         End
         Begin VB.Menu GodownCapcityMIS 
            Caption         =   "Godown Capacity MIS"
         End
         Begin VB.Menu GodownListing 
            Caption         =   "Godown Listing"
         End
         Begin VB.Menu GSLBenposMIS 
            Caption         =   "GSL Benpos MIS"
         End
         Begin VB.Menu RRNPending 
            Caption         =   "RRN Pending"
         End
         Begin VB.Menu InsuranceMIS 
            Caption         =   "Insurance MIS"
         End
         Begin VB.Menu LocationRentMIS 
            Caption         =   "Location Rent MIS"
         End
         Begin VB.Menu NCDEXAdvanceStorageCharges 
            Caption         =   "NCDEX Advance Storage Charges"
            Visible         =   0   'False
         End
         Begin VB.Menu NcdexUtilisation 
            Caption         =   "NCDEX Accreditation  Vs Stock "
         End
         Begin VB.Menu WHRFundingMIS 
            Caption         =   "WHR Pledge Detail"
         End
         Begin VB.Menu VendorClaimDetails 
            Caption         =   "Vendor Claim Details"
         End
         Begin VB.Menu ExpenseListing 
            Caption         =   "Expense Listing"
         End
         Begin VB.Menu RateCard 
            Caption         =   "Rate Card"
         End
         Begin VB.Menu FumigationReport 
            Caption         =   "Fumigation Report"
         End
         Begin VB.Menu NewFumigationReport 
            Caption         =   "New Fumigation Report"
         End
         Begin VB.Menu CITFBOE 
            Caption         =   "CITF BOE"
         End
         Begin VB.Menu NCDEXRejectionStorageAmount 
            Caption         =   "NCDEX Rejection/7 Days Storage Amount"
         End
         Begin VB.Menu ShortageMIS 
            Caption         =   "Shortage MIS"
         End
         Begin VB.Menu AMWiseIncomeMIS 
            Caption         =   "AM Wise Income MIS"
         End
         Begin VB.Menu InvoiceDetails 
            Caption         =   "Invoice Details"
         End
         Begin VB.Menu ModifiedNNNRInvoice 
            Caption         =   "Modified NNNR Invoice"
         End
      End
      Begin VB.Menu EmployeeRpt 
         Caption         =   "Employee"
         Begin VB.Menu MonitorMIS 
            Caption         =   "Employee Monitoring"
         End
         Begin VB.Menu EmployeeAllocation 
            Caption         =   "Employee Allocation"
         End
         Begin VB.Menu EmpLocationMapping 
            Caption         =   "Employee Location Mapping"
         End
         Begin VB.Menu EmployeeLedger 
            Caption         =   "Employee Ledger"
            Visible         =   0   'False
         End
         Begin VB.Menu EmployeeExpenseMIS 
            Caption         =   "Employee Expense MIS"
         End
      End
      Begin VB.Menu CorporateClients 
         Caption         =   "Corporate Clients"
         Begin VB.Menu GSLStockStatus 
            Caption         =   "GSL Stock Status"
         End
      End
      Begin VB.Menu AuditRpt 
         Caption         =   "Audit"
         Begin VB.Menu AuditReport 
            Caption         =   "Audit Report"
         End
         Begin VB.Menu AuditorGodownMIS 
            Caption         =   "Auditor Godown MIS"
         End
         Begin VB.Menu RAGGodownListing 
            Caption         =   "RAG Godown Listing"
         End
      End
      Begin VB.Menu FAGRptMenu 
         Caption         =   "FAG"
         Begin VB.Menu DuplicateGodownRent 
            Caption         =   "Duplicate Godown Rent"
         End
         Begin VB.Menu LocationWiseClientList 
            Caption         =   "Location Wise Client List"
         End
      End
   End
   Begin VB.Menu DBM 
      Caption         =   "Database Managment               "
      Visible         =   0   'False
      Begin VB.Menu SysAdmin 
         Caption         =   "Sys Admin"
         Begin VB.Menu MenuDetail 
            Caption         =   "Menu Detail"
         End
         Begin VB.Menu StoreManLocation 
            Caption         =   "StoreMan Location"
         End
         Begin VB.Menu GLAccountType 
            Caption         =   "G/L Account Type"
         End
      End
      Begin VB.Menu DBPowerUsr 
         Caption         =   "Power Users"
         Begin VB.Menu Exchange 
            Caption         =   "Exchange"
         End
         Begin VB.Menu Setting 
            Caption         =   "Setting "
         End
         Begin VB.Menu UserRole 
            Caption         =   "User Role"
         End
         Begin VB.Menu ReOpenCMP 
            Caption         =   "Re-Open CMP"
         End
         Begin VB.Menu UnmarkPrintFlag 
            Caption         =   "UnMark Print Flag"
         End
         Begin VB.Menu UpdateCxTFDetails 
            Caption         =   "Update CxTF Details"
         End
         Begin VB.Menu AR 
            Caption         =   "Access Rights"
            Begin VB.Menu AccessRight 
               Caption         =   "Access Right UR"
            End
            Begin VB.Menu AccessRightMenu 
               Caption         =   "Access Right Menu"
            End
         End
         Begin VB.Menu Login 
            Caption         =   "Login"
         End
         Begin VB.Menu CoExpense 
            Caption         =   "CO Expense"
         End
      End
   End
   Begin VB.Menu COGProcess 
      Caption         =   "CMG Process          "
      Visible         =   0   'False
      Begin VB.Menu COG_Report 
         Caption         =   "Demat Process"
         Begin VB.Menu KRPending 
            Caption         =   "KR-DR Pending"
         End
         Begin VB.Menu KRUpload 
            Caption         =   "KR Upload"
         End
         Begin VB.Menu DRUpload 
            Caption         =   "DR Upload"
         End
      End
      Begin VB.Menu RRNC 
         Caption         =   "Remat Process "
         Begin VB.Menu RRNUpload 
            Caption         =   "Upload RRN Pending Report"
         End
         Begin VB.Menu RRNConfirmation 
            Caption         =   "RRN Pending Confirmation"
         End
         Begin VB.Menu CAUpload 
            Caption         =   "CA for RRN Upload"
         End
      End
      Begin VB.Menu BENPOS 
         Caption         =   "BENPOS Upload"
      End
      Begin VB.Menu CommodityPriceUpload 
         Caption         =   "Commodity Price Upload"
      End
      Begin VB.Menu HRUpload 
         Caption         =   "HR Upload"
         Begin VB.Menu IncomeUploadFAG 
            Caption         =   "Income Upload (FAG)"
         End
         Begin VB.Menu IncomeUploadCMG 
            Caption         =   "Income Upload (CMG)"
         End
         Begin VB.Menu StockStatus 
            Caption         =   "Stock Status"
         End
         Begin VB.Menu DPUpload 
            Caption         =   "DP Upload"
         End
      End
      Begin VB.Menu FAGProcess 
         Caption         =   "FAG Process "
         Begin VB.Menu mnuExport 
            Caption         =   "Export"
            Begin VB.Menu VendorClaim_Navision 
               Caption         =   "Export Vendor Expense Claim for Navision"
            End
            Begin VB.Menu ExpGodonRentNav 
               Caption         =   "Export Godown Rent Claim for Navision"
            End
            Begin VB.Menu ExportITVendorExpenseClaimforNavision 
               Caption         =   "Export ITVendor Expense Claim for Navision"
            End
            Begin VB.Menu ExportDPInvoice 
               Caption         =   "Export DP Invoice"
            End
            Begin VB.Menu ExportNonNCDEXNonCorporateInvoice 
               Caption         =   "Export Non NCDEX/Non Corporate Invoice"
            End
            Begin VB.Menu ExportReservationInvoice 
               Caption         =   "Export Reservation Invoice"
            End
            Begin VB.Menu ExportAdHocInvoice 
               Caption         =   "Export Ad-Hoc Invoice"
            End
            Begin VB.Menu ExportCashReceipt 
               Caption         =   "Export Cash Receipt"
            End
            Begin VB.Menu ExportReturnCashReceipt 
               Caption         =   "Export Return Cash Receipt"
            End
            Begin VB.Menu ExportCRInvoice 
               Caption         =   "Export CR Invoice"
            End
            Begin VB.Menu ExportSecurityBilling 
               Caption         =   "Export Security Billing"
            End
         End
         Begin VB.Menu Import 
            Caption         =   "Import"
            Begin VB.Menu ImportDPInvoice 
               Caption         =   "Import DP Invoice"
            End
            Begin VB.Menu ImportReservationInvoice 
               Caption         =   "Import Reservation Invoice"
            End
         End
      End
   End
   Begin VB.Menu ShowHideMenu 
      Caption         =   "Show/Hide Menu"
   End
   Begin VB.Menu Search 
      Caption         =   "Search Menu"
   End
End
Attribute VB_Name = "MDIFrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const WS_THICKFRAME = &H40000
Private Const GWL_STYLE = (-16)
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Dim mSelectedNode As Double

Private Sub Assayer_Click()
'FrmMst_Assayer.Show
End Sub

Private Sub AccessRight_Click()
FrmAccessRight.Show
End Sub

Private Sub AccessRightMenu_Click()
FrmAccessRightMenu.Show
End Sub

Private Sub AdHocInvoice_Click()
FrmTxn_TempInvoice.Show
End Sub

Private Sub AdjustmentTxn_Click()
FrmTxn_AdjTransaction.Show
End Sub

Private Sub Agent_Click()
FrmMst_Agent.Show
End Sub

Private Sub AMWiseIncomeMIS_Click()
FrmRpt_NewIncomeMIS.Show
End Sub

Private Sub ApproveBOE_Click()
FrmTxn_ReservationInvoice.Show
End Sub

Private Sub AssetStatus_Click()
FrmMst_AssetStatus.Show
End Sub

Private Sub AssetType_Click()
FrmMst_AssetType.Show
End Sub

Private Sub AuditDetails_Click()
FrmTxn_AuditDetails.Show
End Sub

Private Sub AuditorGodownMIS_Click()
FrmRpt_AuditorGodownMIS.Show
End Sub

Private Sub AuditPoint_Click()
FrmMst_AuditPoint.Show
End Sub

Private Sub AuditReport_Click()
FrmRpt_Audit.Show
End Sub

Private Sub AuditSegment_Click()
FrmMst_AuditSegment.Show
End Sub

Private Sub AuditType_Click()
FrmMst_AuditType.Show
End Sub

Private Sub BagEquivalent_Click()
FrmMst_BagEquivalent.Show
End Sub

Private Sub Bank_Click()
FrmMst_Bank.Show
End Sub

Private Sub BENPOS_Click()
Frm_BenposImport.Show
End Sub

Private Sub BillingCycle_Click()
FrmMst_BillingCycle.Show
End Sub

Private Sub BOEProcess_Click()
Frm_BOEProcess.Show
End Sub

Private Sub Branch_Click()
FrmMst_Branch.Show
End Sub

Private Sub CashReceipt_Click()
'FrmTxn_CashReceipt.Show
FrmTxn_NewCashReceipt.Show

End Sub

Private Sub Category_Click()
FrmMst_Category.Show
End Sub

Private Sub CAUpload_Click()
Frm_CAUpload.Show
End Sub

Private Sub CDC_Click()
frmCDCTransaction.Show
End Sub

Private Sub CDCTransfer_Click()
Frm_CDCTransfer.Show
End Sub

Private Sub ChangePassword_Click()
FrmChangePassword.Show
End Sub

Private Sub Check1_Click()
Call AddNodes
Picture1.Visible = True
Call TableMst_MenuDetail
End Sub

Private Sub CITFBOE_Click()
FrmRpt_CITFBOE.Show
End Sub

Private Sub CITFCRMapping_Click()
FrmTxn_CITF_CR.Show
End Sub

Private Sub CITFDiscrpancy_Click()
FrmRpt_CITFDiscrepancy.Show
End Sub

Private Sub ClientGroup_Click()
FrmMst_ClientGroup.Show
End Sub

Private Sub ClientLogin_Click()
FrmMst_ClientLogin.Show
End Sub

Private Sub ClientMaster_Click()
   FrmMst_Client.Show
End Sub

Private Sub GetPathforMenu(x_ As Variant)
mMenuID_ = Val(MSHFlexGrid1.TextMatrix(x_, 0))

mMasterLabel:

tmp = "select mastermenuid, menuname from mst_menudetail where  menuid = " & mMenuID_
Call Tmp7Table
If MSHFlexGrid1.TextMatrix(x_, 2) = "" Then
   MSHFlexGrid1.TextMatrix(x_, 2) = IIf(IsNull(TmpRs7!MenuName), "", TmpRs7!MenuName)
Else
   MSHFlexGrid1.TextMatrix(x_, 2) = IIf(IsNull(TmpRs7!MenuName), "", TmpRs7!MenuName) & "-->" & MSHFlexGrid1.TextMatrix(x_, 2)
End If
If IIf(IsNull(TmpRs7!MasterMenuID), 0, TmpRs7!MasterMenuID) <> 0 Then
   mMenuID_ = TmpRs7!MasterMenuID
   GoTo mMasterLabel
End If
   
End Sub

Private Sub CMP_Click()
  FrmMst_CMP.Show
End Sub

Private Sub CMPBODCheck_Click()
FrmRpt_CMPBOD.Show
End Sub

Private Sub CMPMonthEndProcess_Click()
'FrmTxn_CMPMonthEndProcess.Show
End Sub

Private Sub CMPPV_Click()
FrmRpt_CMPPV.Show
End Sub

Private Sub CoExpense_Click()
FrmMst_CoExpense.Show
'If LCase(Gen_UserRole) = "power" Then
'   FrmTxn_UpdateStockOpBalance.Show
'End If

End Sub

Private Sub Commodity_Click()
FrmMst_CommodityGroup.Show
End Sub

Private Sub CommodityMapped_Click()
FrmRpt_GSLCommodityMap.Show
End Sub

Private Sub CommodityPriceUpload_Click()
FrmMst_CommodityPrice.Show
End Sub

Private Sub CommodityQualityParameter_Click()
FrmMst_CommodityQualityParameter.Show
End Sub

Private Sub Commodityvariety_Click()
FrmMst_Commodity.Show
End Sub

Private Sub Config_Click()
Form1.Show
End Sub

Private Sub ControlReport_Click()
FrmRpt_ControlReport.Show
End Sub

Private Sub DailyStockMIS_Click()
FrmRpt_DailyStockMIS.Show
End Sub

Private Sub CRDepositDetails_Click()
'FrmTxn_CRDepositeDetails.Show
FrmTxn_NewCRDepositeDetails.Show
End Sub

Private Sub CRSightedDetailsinBankStatement_Click()
FrmTxn_CRSightedDetails.Show
End Sub

Private Sub Deposit_Click()
FrmDeposit.Show
End Sub

Private Sub Depositor_Click()
  FrmMst_Depository.Show
End Sub

Private Sub Depository_Click()
FrmMst_Depository.Show
End Sub

Private Sub Designation_Click()
FrmMst_Designation.Show
End Sub

Private Sub DPInvoiceCoverLetter_Click()
Frm_DPInvoiceCoverLetter.Show
End Sub

Private Sub DPUpload_Click()
Frm_DPUpload.Show
End Sub

Private Sub DRUpload_Click()
Frm_DRUpload.Show
End Sub

Private Sub DTDReport_Click()
FrmRpt_DTD.Show
End Sub

Private Sub DuplicateGodownRent_Click()
FrmRpt_DuplicateGodownRent.Show
End Sub

Private Sub EHopSMMIS_Click()

If UserType <> "Power" Then
   MsgBox "Access denied!!!"
   Exit Sub
End If

FrmRpt_CombineDailyStockMIS.Show

End Sub

Private Sub EmpLocationMapping_Click()
FrmRpt_EMPLocation.Show
End Sub

Private Sub EmployeeAllocation_Click()
FrmRpt_EmpAllocation.Show
End Sub

Private Sub EmployeeAssetDetails_Click()
FrmRpt_EmpAssetMap.Show
End Sub

Private Sub EmployeeClaim_Click()
FrmTxn_ClaimSubmission.Show
End Sub

Private Sub EmployeeCTD_Click()
If UserType <> "Power" Then
   MsgBox "Access denied!!!"
   Exit Sub
End If
'Frm_EmployeeCTC.Show
Frm_ImportMonthWiseSalary.Show
End Sub

Private Sub EmployeeExpenseMIS_Click()
FrmRpt_EmployeeExpenseMIS.Show
End Sub

Private Sub EmployeeLedger_Click()
FrmRpt_EmployeeLedger.Show
End Sub

Private Sub EmployeeLedgerUpload_Click()
Frm_EmpLedger.Show
End Sub

Private Sub EmployeeMaster_Click()
FrmMst_Employee.Show
End Sub

Private Sub EndofDay_Click()
FrmTxn_DataEntryUpto.Show
End Sub

Private Sub Exchange_Click()
   FrmMst_ExchangeType.Show
End Sub

Private Sub Godown_Click()
   FrmMst_Godown.Show
End Sub

Private Sub ExpenseGroup_Click()
FrmMst_ExpenseGroup.Show
End Sub

Private Sub ExpenseListing_Click()
FrmRpt_ExpenseListing.Show
End Sub

Private Sub ExpenseMIS_Click()
If UserType <> "Power" Then
   MsgBox "Access denied!!!"
   Exit Sub
End If

FrmRpt_ExpensesMIS.Show
End Sub

Private Sub ExpenseType_Click()
FrmMst_ExpenseType.Show
End Sub

Private Sub ExpGodonRentNav_Click()
Frm_GodonRentNAVExport.Show
End Sub

Private Sub ExportAdHocInvoice_Click()
Frm_ExportTempInvoice.Show
End Sub

Private Sub ExportCashReceipt_Click()
Frm_ExportCashReceipt.Show
End Sub

Private Sub ExportCRInvoice_Click()
Frm_ExportCRInvoice.Show
End Sub

Private Sub ExportDPInvoice_Click()
Frm_ExportDPInvoice.Show
End Sub

Private Sub ExportITVendorExpenseClaimforNavision_Click()
Frm_ITVendorNAVExport.Show
End Sub

Private Sub ExportNavision_Claim_Click()
FrmNAVExport.Show
End Sub

Private Sub ExportNonNCDEXNonCorporateInvoice_Click()
Frm_ExportNonNCDEXNonCorporateInvoice.Show
End Sub

Private Sub ExportReservationInvoice_Click()
Frm_ExportReservationInvoice.Show
End Sub

Private Sub ExportReturnCashReceipt_Click()
Frm_ExportReturnCashReceipt.Show
End Sub

Private Sub ExportSecurityBilling_Click()
Frm_ExportSecurityBilling.Show
End Sub

Private Sub FAGBank_Click()
FrmMst_FAGBank.Show
End Sub

Private Sub FranchiseeGodownRateCard_Click()
FrmMst_FranchiseeGodownRateCard.Show
End Sub

Private Sub Fumigation_Click()
FrmTxn_Fumigation.Show
End Sub

Private Sub FumigationReport_Click()
FrmRpt_Fumigation.Show
End Sub

Private Sub GenerateBOE_Click()
FrmTxn_GenerateReservationInvoice.Show
End Sub

Private Sub GenerateInvoice_Click()
FrmTxn_GenerateInvoice.Show
End Sub

Private Sub GLAccountType_Click()
FrmMst_GLAccType.Show
End Sub

Private Sub GLSOpening_Click()
FrmTxn_GSL_OpeningBalance.Show
End Sub

Private Sub GodownCapcityMIS_Click()
FrmRpt_GodownUtilizationMIS.Show
End Sub

Private Sub GodownListing_Click()
FrmRpt_GodownListing.Show
End Sub

Private Sub GodownRentClaimReport_Click()
FrmRpt_GodownRentClaim.Show
End Sub

Private Sub GodownRentCoverLetter_Click()
Frm_GodownRentCoverLetter.Show
End Sub

Private Sub GodownTransfer_Click()
FrmTxn_GodownTransfer.Show
End Sub

Private Sub GodownWithoutAgreement_Click()
FrmRpt_GodownWithoutAgreement.Show
End Sub

Private Sub GSLBenposMIS_Click()
FrmRpt_GSLBenPosMIS.Show
End Sub

Private Sub GSLMaster_Click()
Gen_GSLID(0) = True
Gen_GSLID(1) = 0   ' CMP ID
Gen_GSLID(2) = 0   ' Godown ID
Gen_GSLID(3) = ""
Gen_GSLID(4) = ""
Gen_GSLID(5) = "" ' CMP Name
Gen_GSLID(6) = "" ' Godown No
FrmMst_GSL.Show
End Sub

Private Sub GSLStockStatus_Click()
FrmRpt_GSLStockStatus.Show
End Sub

Private Sub GSLTransfer_Click()
FrmTxn_GSL_Transfer.Show
End Sub

Private Sub Hologram_Click()
FrmTxn_Hologram.Show
End Sub

Private Sub ImportAdHocInvoice_Click()
Frm_UpdateTempInvoiceNo.Show
End Sub

Private Sub ImportCREntryNo_Click()
Frm_UpdateCREntryNo.Show
End Sub

Private Sub ImportDPInvoice_Click()
Frm_UploadDPInvoice.Show
End Sub

Private Sub ImportNNNRInvoice_Click()
Frm_UpdateNNNRInvoiceNo.Show
End Sub

Private Sub ImportOtherIncome_Click()
Frm_ImportOtherIncome.Show
End Sub

Private Sub ImportReservationInvoice_Click()
Frm_ImportReservationInvoice.Show
End Sub

Private Sub IncomeExpenseType_Click()
FrmMst_IncomeExpenseType.Show
End Sub

Private Sub IncomeMISReport_Click()

If UserType <> "Power" Then
   MsgBox "Access denied!!!"
   Exit Sub
End If

FrmRpt_IncomeMIS.Show
End Sub

Private Sub IncomeUpload_Click()
Frm_IncomeDetails.Show
End Sub

Private Sub IncomeUploadCMG_Click()
Frm_IncomeUpload.Show
End Sub

Private Sub IncomeUploadFAG_Click()
Frm_IncomeDetails.Show
End Sub

Private Sub InsuranceMIS_Click()
FrmRpt_AvgMonthStockMIS.Show
End Sub

Private Sub InvoiceDetails_Click()
FrmRpt_InvoiceDetails.Show
End Sub

Private Sub ISIN_Click()
FrmMst_ISIN.Show
End Sub

Private Sub ISINMapping_Click()
Frm_ISINMapping.Show
End Sub

Private Sub ISINRevalidation_Click()
Frm_ISINRevalidation.Show
End Sub

Private Sub ITVendor_Click()
FrmMst_ITVendor.Show
End Sub

Private Sub ITVendorClaim_Click()
FrmTxn_ITVendorClaim.Show
End Sub

Private Sub KRDRDelete_Click()
Frm_KRDelete.Show
End Sub

Private Sub KRPending_Click()
FrmRpt_KRPending.Show
End Sub

Private Sub KRUpload_Click()
Frm_KRPendingUpload.Show
End Sub

Private Sub LessorAgreement_Click()
FrmTxn_LessorAgreement.Show
End Sub

Private Sub LessorMaster_Click()
FrmMst_Lessor.Show
End Sub

Private Sub Location_Click()
   FrmMst_Location.Show
End Sub

Private Sub LocationRentMIS_Click()
FrmRpt_LocationRentMIS.Show
End Sub

Private Sub LocationWiseClientList_Click()
FrmRpt_LocationWiseClient.Show
End Sub

Private Sub Login_Click()
Unload Me
FrmStart.Show
End Sub

Private Sub LoginDetails_Click()
FrmRpt_LoginDetails.Show
End Sub

Private Sub MDIForm_Activate()
Unload Frm_MenuSearchPopup
End Sub

Private Sub MDIForm_Load()
'SetWindowPos frmPopup.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
'On Error Resume Next
'tmp = "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & App.Path & "\Store-Man.MDB;Data Provider=MICROSOFT.JET.OLEDB.4.0"
'DataEnvironment1.Connection1.ConnectionString = tmp
'
'
'Call TableMst_Config
'
'If RsMst_Config!Flag = "N" Then
'   MsgBox "Version Expire !!!! Contact Technology !!! "
'   End
'End If
'If Date > RsMst_Config!EndDate Then
'   RsMst_Config!Flag = "N"
'   RsMst_Config.Update
'   MsgBox "Version Expire !!!! Contact Technology !!! "
'   End
'End If

'Call StatusBarData(Me)
'Form1.Show
Call AddNodes

window_style = GetWindowLong(Picture1.hWnd, GWL_STYLE)
window_style = window_style Or WS_THICKFRAME
SetWindowLong Picture1.hWnd, GWL_STYLE, window_style
SetWindowPos Picture1.hWnd, hWnd, 0, 0, 0, 0, _
SWP_NOZORDER Or SWP_NOSIZE Or SWP_NOMOVE Or SWP_DRAWFRAME
'MsgBox "We have successfully migrated on cloud. Please enter only live data on cloud."
Call StatusBarData(Me)
Call TableMst_MenuDetail


End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Unload Frm_MenuSearchPopup
End Sub

Private Sub MenuDetail_Click()
FrmMst_MenuDetail.Show
End Sub

Private Sub MismatchGSLISINLocation_Click()
FrmRpt_GSLISINLocationMisMatch.Show
End Sub

Private Sub MissingEmployees_Click()
FrmRpt_MissingEmpID.Show
End Sub

Private Sub mnuInsurance_Click()
  FrmMstInsuranceCo.Show
End Sub

Private Sub mnuInsuranceDetails_Click()
  FrmMstInsurance.Show
End Sub

Private Sub mnuSecurityBilling_Click()
FrmTxn_SecurityBilling.Show
End Sub

Private Sub ModifiedNNNRInvoice_Click()
FrmRpt_ModifiedNNNRInvoice.Show
End Sub

Private Sub MonitorMIS_Click()
FrmRpt_MonitorMIS.Show
End Sub

Private Sub MonthlyTransactionDetails_Click()
FrmTxn_InvoiceGenProcess.Show
End Sub

Private Sub MonthwiseEmployeeHistory_Click()
FrmTxn_EmpHistory.Show
End Sub

Private Sub MonthWiseEmployeeProcess_Click()
Frm_MonthEmpProces.Show
End Sub

Private Sub MonthWiseSecurityLocation_Click()
FrmTxn_MonthwiseSecLoc.Show
End Sub

Private Sub Mst_License_Click()
FrmMst_License.Show
End Sub

Private Sub NavisionCodeMapping_Click()
FrmMst_NAVCodeMapping.Show
End Sub

Private Sub NCDEXAdvanceStorageCharges_Click()
FrmRpt_NCDEXAdvanceSC.Show
End Sub

Private Sub NCDEXRejectionStorageAmount_Click()
FrmRpt_NRejectionStorageAmount.Show
End Sub

Private Sub NCDEXStorageRateCard_Click()
FrmMst_NCDEXStorageRateCard.Show
End Sub

Private Sub NcdexUtilisation_Click()
FrmRpt_NCDEXUtilization.Show
End Sub

Private Sub NewDailyMIS_Click()
FrmRpt_NewDailyMIS.Show
End Sub

Private Sub NewFumigationReport_Click()
FrmRpt_CompFumgation.Show
End Sub

Private Sub NewISINRevalidation_Click()
Frm_NewISINRevalidation.Show
End Sub

Private Sub OldCashReceipt_Click()
FrmTxn_CashReceipt.Show
End Sub

Private Sub PaymentDetailUpload_Click()
Frm_ImportGodClaimPayDetails.Show

End Sub

Private Sub PendingISINforRevalidation_Click()
FrmRpt_ISINRevalidation.Show
End Sub

Private Sub PendingMonthBeginProcess_Click()
FrmRpt_CMPBOMDetails.Show
End Sub

Private Sub PendingWHRs_Click()
FrmRpt_PendingWHR.Show
End Sub

Private Sub PledgeCreation_Click()
Frm_Pledge_Creation.Show
End Sub

Private Sub PledgeRelease_Click()
FrmTxn_PledgeRelease.Show
End Sub

Private Sub PrintPendingList_Click()
FrmRpt_PendingPrint.Show
End Sub

Private Sub Process_Click()
FrmTxn_GodownRent.Show
End Sub

Private Sub ProcessOpeningBalance_Click()

If UserType = "Power" Then
   FrmProcess.Show
Else
   'Call GButtonLockAD(Me)
   MsgBox "Access denied !!!!!!!!!"
   Exit Sub
End If

End Sub

Private Sub ProformaInvoice_Click()
FrmTxn_ProInvoice.Show
End Sub

Private Sub QualityInformation_Click()
FrmTxn_QualityInformation.Show
End Sub

Private Sub QualityParameter_Click()
FrmMst_QualityParameter.Show
End Sub

Private Sub RAGGodownListing_Click()
FrmRpt_RAGGodownListing.Show
End Sub

Private Sub RAGMemberTransfer_Click()
FrmTxn_RAGMemberTransfer.Show
End Sub

Private Sub RateCard_Click()
FrmRpt_RateCard.Show
End Sub

Private Sub ReIssueCDC_Click()
Frm_CDCReissue.Show
End Sub

Private Sub Rpt_BalanceStk_Click()

End Sub

Private Sub RentApproval_Click()
FrmTxn_GodownApproval.Show
End Sub

Private Sub ReOpenCMP_Click()
Frm_ReOpenCMP.Show
End Sub

Private Sub ReservationContractStock_Click()
FrmRpt_ReservationContractStock.Show
End Sub

Private Sub RRNConfirmation_Click()
FrmRpt_RRNConfirmation.Show
End Sub

Private Sub RRNMaster_Click()
  FrmMst_RRN.Show
End Sub

Private Sub RRNPending_Click()
FrmRpt_RRNPending.Show
End Sub

Private Sub RRNUpload_Click()
Frm_RRNUpload.Show
End Sub

Private Sub SCGSLMapping_Click()
FrmTxn_StorageContractGslMapping.Show
End Sub

Private Sub search_Click()
Frm_MenuSearchPopup.Show
Picture1.Visible = False
End Sub

Private Sub SecurityAgency_Click()
FrmMst_SecurityAgency.Show
End Sub

Private Sub SecurityAssets_Click()
FrmMst_SecurityAsset.Show
End Sub

Private Sub SecurityAttendance_Click()
FrmTxn_MonthwiseSecLoc.Show
End Sub

Private Sub SecurityDesignation_Click()
FrmMst_SecurityDesignation.Show
End Sub

Private Sub SecurityDetails_Click()
'FrmMst_SecurityDetails.Show
FrmMst_SecurityLocationMapping.Show

End Sub

Private Sub SecurityShifts_Click()
FrmMst_SecurityShifts.Show
End Sub

Private Sub ServiceTax_Click()
FrmMst_ServiceTax.Show
End Sub

Private Sub Setting_Click()
FrmMst_Setting.Show
End Sub

Private Sub ShortageMIS_Click()
FrmRpt_ShortageMIS.Show
End Sub

Private Sub ShortagePledgeQuantity_Click()
FrmRpt_ShortageInPledgedQty.Show
End Sub

Private Sub ShowHideMenu_Click()
If Picture1.Visible = True Then
   Picture1.Visible = False
Else
   Call AddNodes
   Picture1.Visible = True
   Call TableMst_MenuDetail
End If
End Sub

Private Sub SpecialRateCard_Click()
FrmMst_SpecialRateCard.Show
End Sub

Private Sub Spillage_Click()
Frm_Spillage.Show
End Sub

Private Sub Spraying_Click()
FrmTxn_Spraying.Show
End Sub

Private Sub SRGM_Click()
FrmTxn_SRGodownMapping.Show
End Sub

Private Sub State_Click()
  FrmMst_State.Show
End Sub

Private Sub StateDRCMapping_Click()
FrmMst_StateDRCMapping.Show
End Sub

Private Sub StateIndia_Click()
FrmMst_StateIndia.Show
End Sub

Private Sub StockOpeningBalance_Click()
FrmTxn_StockOpBal.Show
End Sub

Private Sub StockStatus_Click()
frmMst_StockStatus.Show

End Sub

Private Sub StockSummaryReport_Click()
FrmRpt_StockSummaryReport.Show
End Sub

Private Sub StorageContract_Click()
   FrmMst_StorageContract.Show
End Sub

Private Sub StorageRateCard_Click()
FrmMst_StorageRateCard.Show
End Sub

Private Sub StoreManLocation_Click()
   FrmMst_StoreManLocations.Show
End Sub

Private Sub TDSMaster_Click()
FrmMst_TDS.Show
End Sub

Private Sub TelephoneUpload_Click()

If UserType <> "Power" Then
   MsgBox "Access denied!!!"
   Exit Sub
End If
Frm_TelephoneUpload.Show
'FrmTxn_ClientLedgerUpload.Show

End Sub

Private Sub TestingAgency_Click()
FrmMst_TestingAgency.Show
End Sub

Private Sub TransactionDetailReport_Click()
FrmRpt_TxnListing.Show
End Sub

Private Sub TransactionList_Click()
FrmRpt_TxnListing.Show
End Sub

Private Sub TransactionListing_Click()
'FrmRpt_CxTF.Show
End Sub

Private Sub UnitType_Click()
  FrmMst_UnitType.Show
End Sub

Private Sub UnmarkPrintFlag_Click()
FrmMst_UnmarkPrint.Show
End Sub

Private Sub UnpaidGodownRent_Click()
FrmRpt_UnpaidGodownRent.Show
End Sub

Private Sub UpdateCxTFDetails_Click()
Frm_UpdateCxtfDetails.Show
End Sub

Private Sub UpdateDPInvoiceNo_Click()
Frm_UpdateDPInvoiceNo.Show
End Sub

Private Sub UpdateStockOpBalance_Click()
If LCase(Gen_UserRole) = "power" Then
   FrmTxn_UpdateStockOpBalance.Show
End If
End Sub

Private Sub UserRole_Click()
FrmUsersRoleMaster.Show
End Sub

Private Sub Users_Click()
FrmUsersMaster.Show
End Sub

Private Sub VendorClaim_Click()
FrmTxn_VendorClaim.Show
End Sub

Private Sub VendorClaim_Navision_Click()
Frm_VendorNAVExport.Show
End Sub

Private Sub VendorClaimDetails_Click()
FrmRpt_VendorClaimDetails.Show
End Sub

Private Sub VendorMaster_Click()
FrmMst_Vendor.Show
End Sub

Private Sub WearhouseReceipt_Click()
Frm_WHRDetail.Show
End Sub

Private Sub WHRFundingMIS_Click()
FrmRpt_WHRFundingMIS.Show
End Sub

Private Sub WHRMonitoring_Click()
FrmRpt_WHRMonitor.Show
End Sub

Private Sub WHRUnitOfMeasurement_Click()
FrmMst_WHRUOM.Show
End Sub

Private Sub Withdrawal_Click()
FrmWithdrawal.Show
End Sub

Private Sub Picture1_Resize()
  TreeView1.Width = Picture1.Width
'  Dim frm As Form
'  Set frm = MDIFrmMain.ActiveForm
'  frm.Hide
  
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)

'If RsMst_MenuDetail.RecordCount > 0 Then RsMst_MenuDetail.MoveFirst
'RsMst_MenuDetail.Find "FormName='" & Node.Key & "'"
'If RsMst_MenuDetail.EOF = False Then
'   Dim frm As Form
'   Set frm = CallByName(Forms, "Add", VbMethod, RsMst_MenuDetail!FormName)
'   Picture1.Visible = False
'   frm.Show
'End If


'If RsMst_MenuDetail.RecordCount > 0 Then RsMst_MenuDetail.MoveFirst
'RsMst_MenuDetail.Find "FormName='" & Node.Key & "'"
'If RsMst_MenuDetail.EOF = False Then
'   Dim frm As Form
'   Set frm = CallByName(Forms, "Add", VbMethod, RsMst_MenuDetail!FormName)
'   Picture1.Visible = False
'   frm.Show
'End If


If TmpRs.RecordCount > 0 Then TmpRs.MoveFirst
TmpRs.Find "FormName='" & Node.Key & "'"
mSelectedNode = Node.Index
If TmpRs.EOF = False Then
   mMenuID = TmpRs!MenuID
   If TmpRs!SelfMenu = 0 Then
      mMenuName = IIf(IsNull(TmpRs!MenuName), "", TmpRs!MenuName)
      For Each lFr In Forms
          If lFr.Name = TmpRs!FormName Then
             Exit Sub
          End If
      Next
      Dim frm As Form
      Gen_DbPathFromAzure = "Y"
      Set frm = CallByName(Forms, "Add", VbMethod, TmpRs!FormName)
      Picture1.Visible = False
      frm.Caption = mMenuName 'IIf(IsNull(TmpRs!MenuName), "", TmpRs!MenuName)
      frm.Show
      If Gen_FormAccess = True Then
         Call TableLog_FormAccessDetail("Where LogID = 0")
         RsLog_FormAccessDetail.AddNew
'         RsLog_FormAccessDetail!LogID = GetMaxLogID
         RsLog_FormAccessDetail!UserId = Gen_UserID
         RsLog_FormAccessDetail!MenuID = mMenuID
         RsLog_FormAccessDetail.Update
      End If
   End If
End If

End Sub
Private Function GetMaxLogID() As Double
Dim Retval As Double
tmp = "Select max(LogID) from Log_FormAccessDetail"
Call TmpTable
Retval = IIf(IsNull(TmpRs.Fields(0)), 0, TmpRs.Fields(0))
Retval = Retval + 1
GetMaxLogID = Retval
End Function


'Commented on 05 Oct 2011
'Private Sub AddNodes()
'Dim nodX As Node  ' Create variable.
'
''TreeView1.Refresh
'TreeView1.Nodes.Clear
'
'TreeView1.ImageList = ImageList1
'TreeView1.Style = tvwTreelinesPictureText
'
'Set nodX = TreeView1.Nodes.Add(, , "Master", "Master", 1)
'
'Call TableMst_MenuDetail(" Where HMenuFlag='M' And Flag='Y' Order By MenuName")
'
'For I = 1 To RsMst_MenuDetail.RecordCount
'    Set nodX = TreeView1.Nodes.Add("Master", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
'    RsMst_MenuDetail.MoveNext
'Next
'
'Set nodX = TreeView1.Nodes.Add(, , "Transaction", "Transaction", 1)
'
'Call TableMst_MenuDetail(" Where HMenuFlag='T' And Flag='Y' Order By MenuName")
'
'For I = 1 To RsMst_MenuDetail.RecordCount
'    Set nodX = TreeView1.Nodes.Add("Transaction", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
'    RsMst_MenuDetail.MoveNext
'Next
'
'Set nodX = TreeView1.Nodes.Add(, , "Report", "Report", 1)
'
'Call TableMst_MenuDetail(" Where HMenuFlag='R' And Flag='Y' Order By MenuName")
'
'For I = 1 To RsMst_MenuDetail.RecordCount
'    Set nodX = TreeView1.Nodes.Add("Report", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
'    RsMst_MenuDetail.MoveNext
'Next
'
'Set nodX = TreeView1.Nodes.Add(, , "COG", "COG", 1)
'
'Call TableMst_MenuDetail(" Where HMenuFlag='C' And Flag='Y' Order By MenuName")
'
'For I = 1 To RsMst_MenuDetail.RecordCount
'    Set nodX = TreeView1.Nodes.Add("COG", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
'    RsMst_MenuDetail.MoveNext
'Next
'
'Set nodX = TreeView1.Nodes.Add(, , "Admin", "Admin", 1)
'
'Call TableMst_MenuDetail(" Where HMenuFlag='A' And Flag='Y' Order By MenuName")
'
'For I = 1 To RsMst_MenuDetail.RecordCount
'    Set nodX = TreeView1.Nodes.Add("Admin", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
'    RsMst_MenuDetail.MoveNext
'Next
'
'If UserType = "Power" Then
'   Set nodX = TreeView1.Nodes.Add(, , "Power", "Power", 1)
'
'   Call TableMst_MenuDetail(" Where HMenuFlag='A' And Flag='N' Order By MenuName")
'
'   For I = 1 To RsMst_MenuDetail.RecordCount
'       Set nodX = TreeView1.Nodes.Add("Power", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
'       RsMst_MenuDetail.MoveNext
'   Next
'
'End If
'
'End Sub



'''Private Sub AddNodes()
'''Dim nodX As Node  ' Create variable.
'''
'''
''''Call TableMst_MenuDetail '(" Where SelfMenu = 0 ")
'''
'''tmp = "Select MM.MenuID,MM.FormFileName,MM.MenuName,MM.FormName,MM.Flag,MM.SelfMenu,MM1.MenuName as 'MasterMenu' From Mst_MenuDetail MM"
'''tmp = tmp & " Left Join Mst_MenuDetail MM1 On MM.MasterMenuID = MM1.MenuID"
'''tmp = tmp & " Where MM.SelfMenu = 1 Or MM1.MenuName is not null "
''''tmp = tmp & " Order By MM.SelfMenu desc,MM1.MasterMenu,MM.SequenceNo "
'''tmp = tmp & " Order By  MM.SelfMenu Desc, MM.MasterMenuID, MM1.MasterMenu, MM.SequenceNo"
'''
''''tmp = tmp & " Order By MM.SelfMenu Desc "
'''
'''Call TmpTable
''''TreeView1.Refresh
'''
'''TreeView1.Nodes.Clear
'''
'''TreeView1.ImageList = ImageList1
'''TreeView1.Style = tvwTreelinesPictureText
'''
'''
'''For I = 1 To TmpRs.RecordCount
'''    If TmpRs!SelfMenu = 1 And IsNull(TmpRs!MasterMenu) = True Then
'''       Set nodX = TreeView1.Nodes.Add(, , TmpRs!FormName, TmpRs!MenuName, 1) 'Set nodX = TreeView1.Nodes.Add(, , "Master", "Master", 1)
'''    Else
'''       If TmpRs!SelfMenu = 1 And IsNull(TmpRs!MasterMenu) = False Then
'''          Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!FormName, TmpRs!MenuName, 1)
'''          'Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!MenuID, TmpRs!MenuName, 1)
'''       Else
'''          Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!FormName, TmpRs!MenuName, 9)
'''          'Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!MenuID, TmpRs!MenuName, 9)
'''       End If
'''    End If
'''    TmpRs.MoveNext
'''    'Set nodX = TreeView1.Nodes.Add(, , "Master", "Master", 1)
'''Next
'''
'''TreeView1.Nodes(1).Selected = True
'''
''''Set nodX = TreeView1.Nodes.Add(, , "Master", "Master", 1)
''''
''''Call TableMst_MenuDetail(" Where HMenuFlag='M' And Flag='Y' Order By MenuName")
''''
''''For I = 1 To RsMst_MenuDetail.RecordCount
''''    Set nodX = TreeView1.Nodes.Add("Master", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
''''    RsMst_MenuDetail.MoveNext
''''Next
''''
''''Set nodX = TreeView1.Nodes.Add(, , "Transaction", "Transaction", 1)
''''
''''Call TableMst_MenuDetail(" Where HMenuFlag='T' And Flag='Y' Order By MenuName")
''''
''''For I = 1 To RsMst_MenuDetail.RecordCount
''''    Set nodX = TreeView1.Nodes.Add("Transaction", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
''''    RsMst_MenuDetail.MoveNext
''''Next
''''
''''Set nodX = TreeView1.Nodes.Add(, , "Report", "Report", 1)
''''
''''Call TableMst_MenuDetail(" Where HMenuFlag='R' And Flag='Y' Order By MenuName")
''''
''''For I = 1 To RsMst_MenuDetail.RecordCount
''''    Set nodX = TreeView1.Nodes.Add("Report", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
''''    RsMst_MenuDetail.MoveNext
''''Next
''''
''''Set nodX = TreeView1.Nodes.Add(, , "COG", "COG", 1)
''''
''''Call TableMst_MenuDetail(" Where HMenuFlag='C' And Flag='Y' Order By MenuName")
''''
''''For I = 1 To RsMst_MenuDetail.RecordCount
''''    Set nodX = TreeView1.Nodes.Add("COG", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
''''    RsMst_MenuDetail.MoveNext
''''Next
''''
''''Set nodX = TreeView1.Nodes.Add(, , "Admin", "Admin", 1)
''''
''''Call TableMst_MenuDetail(" Where HMenuFlag='A' And Flag='Y' Order By MenuName")
''''
''''For I = 1 To RsMst_MenuDetail.RecordCount
''''    Set nodX = TreeView1.Nodes.Add("Admin", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
''''    RsMst_MenuDetail.MoveNext
''''Next
''''
''''If UserType = "Power" Then
''''   Set nodX = TreeView1.Nodes.Add(, , "Power", "Power", 1)
''''
''''   Call TableMst_MenuDetail(" Where HMenuFlag='A' And Flag='N' Order By MenuName")
''''
''''   For I = 1 To RsMst_MenuDetail.RecordCount
''''       Set nodX = TreeView1.Nodes.Add("Power", tvwChild, RsMst_MenuDetail!FormName, RsMst_MenuDetail!MenuName, 9)
''''       RsMst_MenuDetail.MoveNext
''''   Next
''''
''''End If
'''
'''End Sub
Private Sub AddNodes()
Dim nodX As Node  ' Create variable.

Gen_DbPathFromAzure = "N"
tmp = "Select MM.MenuID,MM.FormFileName,MM.MenuName,MM.FormName,MM.Flag,MM.SelfMenu,MM1.MenuName as 'MasterMenu',MR.FormAccess,MM.Detaildesc From Mst_MenuDetailVBDB MM"
tmp = tmp & " Left Join Mst_MenuDetailVBDB MM1 On MM.MasterMenuID = MM1.MenuID"
tmp = tmp & " left join Mst_MenuHasUserRoleVBDB MR on MM.MenuID=MR.MenuID And MR.UserGroupID=" & Gen_UserTypeID
tmp = tmp & " Where MM.SelfMenu = 1 Or MM1.MenuName is not null"
tmp = tmp & " Order By  MM.SelfMenu Desc, MM.MasterMenuID, MM1.MenuName, MM.SequenceNo"

Call TmpTable
'TreeView1.Refresh

TreeView1.Nodes.Clear

TreeView1.ImageList = ImageList1
TreeView1.Style = tvwTreelinesPictureText


For I = 1 To TmpRs.RecordCount
    If TmpRs!SelfMenu = 1 And IsNull(TmpRs!MasterMenu) = True Then
       Set nodX = TreeView1.Nodes.Add(, , TmpRs!FormName, TmpRs!MenuName, 1) 'Set nodX = TreeView1.Nodes.Add(, , "Master", "Master", 1)
    Else
       If TmpRs!SelfMenu = 1 And IsNull(TmpRs!MasterMenu) = False Then
          Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!FormName, TmpRs!MenuName, 1)
          'Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!MenuID, TmpRs!MenuName, 1)
       Else
          If Check1.Value = 1 Then
             If TmpRs!formaccess = "Y" Then
                Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!FormName, TmpRs!MenuName, 9)
                nodX.Tag = IIf(IsNull(TmpRs!detaildesc), "", TmpRs!detaildesc)
             End If
          Else
             Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!FormName, TmpRs!MenuName, 9)
             nodX.Tag = IIf(IsNull(TmpRs!detaildesc), "", TmpRs!detaildesc)
          End If
          
          'Set nodX = TreeView1.Nodes.Add("" & TmpRs!MasterMenu & "", tvwChild, TmpRs!MenuID, TmpRs!MenuName, 9)
       End If
    End If
    TmpRs.MoveNext
    'Set nodX = TreeView1.Nodes.Add(, , "Master", "Master", 1)
Next
If mSelectedNode > 0 Then
   TreeView1.Nodes(mSelectedNode).Selected = True
Else
   TreeView1.Nodes(1).Selected = True
End If
Gen_DbPathFromAzure = "Y"

End Sub


Private Sub TreeView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 0 Then
    SetToolTip x, y
 End If

End Sub


Private Sub SetToolTip(ByVal x As Single, ByVal y As Single)
Static last_x As Single
Static last_y As Single
Static node_over As Node
Dim new_node_over As Node

    ' See if the mouse has moved.
    If (last_x = x) And (last_y = y) Then Exit Sub
    last_x = x
    last_y = y

    ' See what node we're over.
    Set new_node_over = TreeView1.HitTest(x, y)
    If node_over Is new_node_over Then Exit Sub
    Set node_over = new_node_over
    Picture2.Visible = False
    If new_node_over Is Nothing Then
        TreeView1.ToolTipText = ""
        
    Else
        mMax = 64
        mPara = 0
        t1 = 1
        mLen = Len(node_over.Tag)
        mLen = RoundUp(mLen / mMax) '+ 1
        If mLen = 0 Then
           mLen = 1
        End If
        Picture2.Left = last_x + 300
        Picture2.Top = last_y + 300
        Picture2.Height = mLen * 225
        Picture2.Width = 4950
        Label1.Width = 4900
        Label1.Height = mLen * 225
        Label1.Refresh
        Label1.Caption = Replace(node_over.Tag, vbCrLf, " ")
        If node_over.Tag = "" Then
           Picture2.Visible = False
        Else
           Picture2.Visible = True
        End If
    End If
End Sub
