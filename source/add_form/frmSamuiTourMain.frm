VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMtpTaxMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14190
   Icon            =   "frmSamuiTourMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14190
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel SSPanel1 
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   1402
      _Version        =   131073
      BackStyle       =   1
      Begin VB.Label lblDateTime 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         Height          =   315
         Left            =   9390
         TabIndex        =   2
         Top             =   30
         Width           =   2505
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   555
         Left            =   9660
         TabIndex        =   1
         Top             =   6390
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   979
         _Version        =   131073
         PictureFrames   =   1
         Picture         =   "frmSamuiTourMain.frx":08CA
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7755
      Left            =   0
      TabIndex        =   3
      Top             =   780
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   13679
      _Version        =   131073
      BackStyle       =   1
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   5445
         Left            =   240
         TabIndex        =   4
         Top             =   0
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   9604
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":195A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":2234
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":2B0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":33E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":3542
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":3E1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":46F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":4A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":52EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":5BC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":649E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSamuiTourMain.frx":7178
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblUserName 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   9
         Top             =   5580
         Width           =   3045
      End
      Begin VB.Label lblUserGroup 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   8
         Top             =   6090
         Width           =   3045
      End
      Begin VB.Label lblVersion 
         Caption         =   "Label1"
         Height          =   465
         Left            =   360
         TabIndex        =   7
         Top             =   6600
         Width           =   3045
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   465
         Left            =   1800
         TabIndex        =   6
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPasswd 
         Height          =   465
         Left            =   330
         TabIndex        =   5
         Top             =   7170
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   820
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSPanel pnlHeader 
      Height          =   735
      Left            =   4680
      TabIndex        =   10
      Top             =   840
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1296
      _Version        =   131073
      BackStyle       =   1
   End
   Begin Threed.SSFrame fraAdmin 
      Height          =   3615
      Left            =   7560
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdUserGroup 
         Height          =   765
         Left            =   900
         TabIndex        =   14
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdUser 
         Height          =   765
         Left            =   900
         TabIndex        =   13
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdminReport 
         Height          =   765
         Left            =   900
         TabIndex        =   12
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraMain 
      Height          =   4875
      Left            =   12480
      TabIndex        =   15
      Top             =   1080
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8599
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdMainSupplier 
         Height          =   765
         Left            =   900
         TabIndex        =   20
         Top             =   2040
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainCustomer 
         Height          =   765
         Left            =   900
         TabIndex        =   19
         Top             =   1260
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainEnterprise 
         Height          =   765
         Left            =   900
         TabIndex        =   18
         Top             =   480
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainReport 
         Height          =   765
         Left            =   900
         TabIndex        =   17
         Top             =   3600
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMainEmployee 
         Height          =   765
         Left            =   900
         TabIndex        =   16
         Top             =   2820
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraCRM 
      Height          =   4875
      Left            =   4800
      TabIndex        =   21
      Top             =   2520
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8599
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdCRMReport 
         Height          =   765
         Left            =   900
         TabIndex        =   25
         Top             =   3180
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOrder 
         Height          =   765
         Left            =   900
         TabIndex        =   24
         Top             =   840
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSchedule1 
         Height          =   765
         Left            =   900
         TabIndex        =   23
         Top             =   1620
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSchedule2 
         Height          =   765
         Left            =   900
         TabIndex        =   22
         Top             =   2400
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraMaster 
      Height          =   3615
      Left            =   4800
      TabIndex        =   26
      Top             =   4080
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   6376
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdMasterReport 
         Height          =   765
         Left            =   900
         TabIndex        =   29
         Top             =   2190
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMasterCRM 
         Height          =   765
         Left            =   900
         TabIndex        =   28
         Top             =   1410
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdMasterMain 
         Height          =   765
         Left            =   900
         TabIndex        =   27
         Top             =   630
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame fraWHTax 
      Height          =   4245
      Left            =   5160
      TabIndex        =   30
      Top             =   1440
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7488
      _Version        =   131073
      BackStyle       =   1
      Begin Threed.SSCommand cmdTax3 
         Height          =   765
         Left            =   900
         TabIndex        =   34
         Top             =   2310
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdTax1 
         Height          =   765
         Left            =   900
         TabIndex        =   33
         Top             =   750
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdTax2 
         Height          =   765
         Left            =   900
         TabIndex        =   32
         Top             =   1530
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdTaxReport 
         Height          =   765
         Left            =   900
         TabIndex        =   31
         Top             =   3090
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   1349
         _Version        =   131073
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmMtpTaxMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"

Private MustAsk As Boolean
Private m_HasActivate As Boolean
Private m_Rs  As ADODB.Recordset
Private m_TableName As String

Private m_TaxDoc As CTaxDocument

Public HeaderText As String
Private m_MustAsk As Boolean
Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_CheckBoxes As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_ReportParams As Collection
Private m_FromDate As Date
Private m_ToDate As Date

Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long

   If trvMain.SelectedItem Is Nothing Then
      Exit Sub
   End If
      
   ReportKey = trvMain.SelectedItem.Key
   
   Set Rc = New CReportConfig
   Rc.REPORT_KEY = ReportKey
   Call Rc.QueryData(m_Rs, iCount)
   
   If Not m_Rs.EOF Then
      Call Rc.PopulateFromRS(1, m_Rs)
      
      frmReportConfig.ShowMode = SHOW_EDIT
      frmReportConfig.ID = Rc.REPORT_CONFIG_ID
   Else
      frmReportConfig.ShowMode = SHOW_ADD
   End If

   frmReportConfig.ReportKey = ReportKey
   frmReportConfig.HeaderText = trvMain.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdAdminReport_Click()
   If Not VerifyAccessRight("ADMIN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmSummaryReport.HeaderText = cmdAdminReport.Caption
   frmSummaryReport.MasterMode = 1
   Load frmSummaryReport
   frmSummaryReport.Show 1

   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdCRMReport_Click()
   If Not VerifyAccessRight("CRM_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSummaryReport.HeaderText = cmdCRMReport.Caption
   frmSummaryReport.MasterMode = 4
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub


Private Sub cmdMainCustomer_Click()
   If Not VerifyAccessRight("MAIN_CUSTOMER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Load frmCustomer
   frmCustomer.Show 1
   
   Unload frmCustomer
   Set frmCustomer = Nothing
End Sub

Private Sub cmdMainEmployee_Click()
   If Not VerifyAccessRight("MAIN_EMPLOYEE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If


   Load frmEmployee
   frmEmployee.Show 1
   
   Unload frmEmployee
   Set frmEmployee = Nothing
End Sub

Private Sub cmdMainEnterprise_Click()
   If Not VerifyAccessRight("MAIN_ENTERPRISE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmEnterprise.HeaderText = cmdMainEnterprise.Caption
   Load frmEnterprise
   frmEnterprise.Show 1
   
   Unload frmEnterprise
   Set frmEnterprise = Nothing
End Sub

Private Sub cmdMainReport_Click()
   If Not VerifyAccessRight("MAIN_REPORT") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmSummaryReport.HeaderText = cmdMainReport.Caption
   frmSummaryReport.MasterMode = 3
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
    Set frmSummaryReport = Nothing
End Sub

Private Sub cmdMainSupplier_Click()
   If Not VerifyAccessRight("MAIN_SUPPLIER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmSupplier
   frmSupplier.Show 1
   
   Unload frmSupplier
   Set frmSupplier = Nothing
End Sub

Private Sub cmdMasterCRM_Click()
   If Not VerifyAccessRight("MASTER_CRM") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMasterCRM.Caption
   frmMasterMain.MasterMode = 4
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdMasterMain_Click()
   If Not VerifyAccessRight("MASTER_MAIN") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmMasterMain.HeaderText = cmdMasterMain.Caption
   frmMasterMain.MasterMode = 3
   Load frmMasterMain
   frmMasterMain.Show 1
   
   Unload frmMasterMain
   Set frmMasterMain = Nothing
End Sub

Private Sub cmdOrder_Click()
   If Not VerifyAccessRight("CRM_ORDER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   frmSlipBooking.HeaderText = cmdOrder.Caption
   Load frmSlipBooking
   frmSlipBooking.Show 1
   
   Unload frmSlipBooking
   Set frmSlipBooking = Nothing
End Sub

Private Sub cmdSchedule1_Click()
   If Not VerifyAccessRight("CRM_SCHEDULE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmSchedule.HeaderText = cmdSchedule1.Caption
   frmSchedule.Area = 1
   Load frmSchedule
   frmSchedule.Show 1
   
   Unload frmSchedule
   Set frmSchedule = Nothing
End Sub

Private Sub cmdSchedule2_Click()
   If Not VerifyAccessRight("CRM_SCHEDULE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmSchedule.HeaderText = cmdSchedule2.Caption
   frmSchedule.Area = 2
   Load frmSchedule
   frmSchedule.Show 1
   
   Unload frmSchedule
   Set frmSchedule = Nothing
End Sub

Private Sub cmdTax1_Click()
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu
Dim D As CMenuItem

'TAX_TYPE_NAME = 3                                            'ภงด 3
'TAX_TYPE_NAME = 53                                          'ภงด 53
'TAX_TYPE_NAME = 4                                            'ภงด Report

If Not VerifyAccessRight("TAX_WITHOLD") Then
   Call EnableForm(Me, True)
   Exit Sub
End If

Set oMenu = New cPopupMenu
lMenuChosen = oMenu.Popup("ข้อมูล ภ.ง.ด 2", "-", "ข้อมูล ภ.ง.ด 2 ก")
 If lMenuChosen = 1 Then
       TAX_TYPE_NAME = 2                                         'ภงด 2
      frmWHTax.HeaderText = cmdTax1.Caption
      frmWHTax.TaxType = TAX_TYPE_NAME
      Load frmWHTax
      frmWHTax.Show 1
      
      Unload frmWHTax
      Set frmWHTax = Nothing
   ElseIf lMenuChosen = 3 Then
      frmWHTax.HeaderText = cmdTax1.Caption & "ก"
      TAX_TYPE_NAME = 21                                         'ภงด 2ก
      frmWHTax.TaxType = TAX_TYPE_NAME
      Load frmWHTax
      frmWHTax.Show 1
      
      Unload frmWHTax
      Set frmWHTax = Nothing
   End If
End Sub

Private Sub cmdTax2_Click()
TAX_TYPE_NAME = 3                                           'ภงด 3

   If Not VerifyAccessRight("TAX_WITHOLD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmWHTax.HeaderText = cmdTax2.Caption
   frmWHTax.TaxType = 3
   Load frmWHTax
   frmWHTax.Show 1
   
   Unload frmWHTax
   Set frmWHTax = Nothing
End Sub

Private Sub cmdTax3_Click()
'TAX_TYPE_NAME = 2                                            'ภงด 2
'TAX_TYPE_NAME = 3                                            'ภงด 3
'TAX_TYPE_NAME = 4                                            'ภงด Report
TAX_TYPE_NAME = 53                                    'ภงด 53

   If Not VerifyAccessRight("TAX_WITHOLD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmWHTax.HeaderText = cmdTax3.Caption
   frmWHTax.TaxType = 53
   Load frmWHTax
   frmWHTax.Show 1
   
   Unload frmWHTax
   Set frmWHTax = Nothing
End Sub

Private Sub cmdTaxReport_Click()

   frmSummaryReport.HeaderText = cmdTaxReport.Caption
   frmSummaryReport.MasterMode = 4
   Load frmSummaryReport
   frmSummaryReport.Show 1
   
   Unload frmSummaryReport
   Set frmSummaryReport = Nothing

End Sub

Private Sub cmdUser_Click()
   If Not VerifyAccessRight("ADMIN_USER") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   Load frmUser
   frmUser.Show 1
   
   Unload frmUser
   Set frmUser = Nothing
End Sub

Private Sub cmdUserGroup_Click()
   If Not VerifyAccessRight("ADMIN_GROUP") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   Load frmUserGroup
   frmUserGroup.Show 1
   
   Unload frmUserGroup
   Set frmUserGroup = Nothing
End Sub

Private Sub Form_Activate()
Dim OKClick As Boolean
Dim iCount As Long

   If Not m_HasActivate Then
      m_HasActivate = True
      Call PatchDB
      
      Load frmLogin
      frmLogin.Show 1
      
      OKClick = frmLogin.OKClick
      
      Unload frmLogin
      Set frmLogin = Nothing
      
      glbEnterPrise.ENTERPRISE_ID = -1
      Call glbEnterPrise.QueryData(m_Rs, iCount)
      If Not m_Rs.EOF Then
         Call glbEnterPrise.PopulateFromRS(1, m_Rs)
      End If
      
      If Not OKClick Then
         m_MustAsk = False
         Unload Me
      Else
         Call LoadAccessRight(Nothing, glbAccessRight, glbUser.GROUP_ID)
      End If
   End If
End Sub

Private Sub Form_Load()
   m_MustAsk = True
   Call InitFormLayout
   Set m_Rs = New ADODB.Recordset
   
   Set m_ReportControls = New Collection
   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
   Set m_ReportParams = New Collection
   Set m_CheckBoxes = New Collection
End Sub

Private Sub InitFormLayout()
   Call InitNormalLabel(lblUsername, MapText("ผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("เวอร์ชัน : ") & glbParameterObj.Version & " (Interbase) ", RGB(0, 0, 255))
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.BackStyle = 1
   lblDateTime.BackColor = RGB(255, 255, 255)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Me.Caption = MapText("ระบบงานภาษีอากร")
   
   cmdUserGroup.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdUser.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdAdminReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   Call InitMainButton(cmdUserGroup, MapText("ข้อมูลกลุ่มผู้ใช้งาน"))
   Call InitMainButton(cmdUser, MapText("ข้อมูลผู้ใช้งาน"))
   Call InitMainButton(cmdAdminReport, MapText("รายงานข้อมูลผู้ใช้งาน"))
   
   cmdMainEnterprise.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainCustomer.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainSupplier.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainReport.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMainEmployee.Picture = LoadPicture(glbParameterObj.MainButton)
   
   Call InitMainButton(cmdMainEnterprise, MapText("ข้อมูลบริษัท"))
   Call InitMainButton(cmdMainCustomer, MapText("ข้อมูลลูกค้า"))
   Call InitMainButton(cmdMainSupplier, MapText("ข้อมูลซัพพลายเออร์"))
   Call InitMainButton(cmdMainEmployee, MapText("ข้อมูลพนักงาน"))
   Call InitMainButton(cmdMainReport, MapText("รายงานข้อมูลกลาง"))
   
   
   cmdOrder.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSchedule1.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdSchedule2.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdCRMReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   Call InitMainButton(cmdOrder, MapText("ข้อมูลการจอง"))
   Call InitMainButton(cmdSchedule1, MapText("ข้อมูลตารางรถ"))
   Call InitMainButton(cmdSchedule2, MapText("ข้อมูลตารางเรือ"))
   Call InitMainButton(cmdCRMReport, MapText("รายงานระบบรับออเดอร์"))
   
   cmdMasterMain.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMasterCRM.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdMasterReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   Call InitMainButton(cmdMasterMain, MapText("ข้อมูลหลักส่วนกลาง"))
   Call InitMainButton(cmdMasterCRM, MapText("ข้อมูลหลักระบบภาษีอากร"))
   Call InitMainButton(cmdMasterReport, MapText("รายงานข้อมูลหลัก"))
   
   cmdTax1.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdTax2.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdTax3.Picture = LoadPicture(glbParameterObj.MainButton)
   cmdTaxReport.Picture = LoadPicture(glbParameterObj.MainButton)
   
   Call InitMainButton(cmdTax1, MapText("ระบบข้อมูล ภ.ง.ด. 2"))
   Call InitMainButton(cmdTax2, MapText("ระบบข้อมูล ภ.ง.ด. 3"))
   Call InitMainButton(cmdTax3, MapText("ระบบข้อมูล ภ.ง.ด. 53"))
   Call InitMainButton(cmdTaxReport, MapText("รายงานระบบข้อมูลภาษี"))
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
      
   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdPasswd, MapText("โปรแกรม"))
   
   Call InitMainTreeview
End Sub

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String

   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
      
   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, MapText("ระบบงานภาษีอากร"), 8)
   Node.Expanded = True
   Node.Selected = True
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", MapText("1. ระบบข้อมูลผู้ใช้งาน"), 4, 4)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-0", MapText("2. ระบบข้อมูลหลัก"), 6, 6)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-0", MapText("3. ระบบข้อมูลกลาง"), 1, 1)
   Node.Expanded = False
      
'   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-0", MapText("4. ระบบรับออเดอร์"), 7, 7)
'   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("4. ระบบข้อมูลภาษี"), 7, 7)
   Node.Expanded = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If m_MustAsk Then
      glbErrorLog.LocalErrorMsg = MapText("ท่านต้องการออกจากโปรแกรมใช่หรือไม่")
      If glbErrorLog.AskMessage = vbYes Then
         Cancel = False
      Else
         Cancel = True
      End If
   Else
      Cancel = False
   End If
End Sub

Private Sub FillReportInput(r As CReportInterface)
Dim C As CReportControl

   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call r.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "CB") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call r.AddParam(m_Combos(C.ControlIndex).ListIndex, C.Param2)
         End If
      End If
      
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call r.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call r.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call r.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param2)
         End If
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               End If
            End If
            If C.Param2 = "FROM_DATE" Then
               m_FromDate = m_Dates(C.ControlIndex).ShowDate
            ElseIf C.Param2 = "TO_DATE" Then
               m_ToDate = m_Dates(C.ControlIndex).ShowDate
            End If
            Call r.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
   
   Next C
End Sub

Private Function VerifyReportInput() As Boolean
Dim C As CReportControl

   VerifyReportInput = False
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "T") Then
         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   
      If (C.ControlType = "D") Then
         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
            Exit Function
         End If
      End If
   Next C
   VerifyReportInput = True
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
   Set m_ReportParams = Nothing
   Set m_CheckBoxes = Nothing
   Set m_Rs = Nothing
   
   Call ReleaseAll
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False

   lblDateTime.Caption = "                                                    "
   lblDateTime.Caption = DateToStringExtEx3(Now)
   lblUsername.Caption = MapText("ผู้ใช้ : ") & " " & glbUser.USER_NAME
   lblUserGroup.Caption = MapText("กลุ่มผู้ใช้ : ") & " " & glbUser.GROUP_NAME
      
  Timer1.Enabled = True
End Sub

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
   If Node Is Nothing Then
      Exit Sub
   End If
   
   fraAdmin.Visible = False
   fraMain.Visible = False
   fraCRM.Visible = False
   fraMaster.Visible = False
   fraWHTax.Visible = False
   
   pnlHeader.Caption = Node.Text
   If Node.Key = ROOT_TREE & " 1-0" Then
        fraAdmin.Left = 5280
        fraAdmin.Top = 2310
        fraAdmin.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 2-0" Then
        fraMaster.Left = 5280
        fraMaster.Top = 2310
        fraMaster.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 3-0" Then
        fraMain.Left = 5280
        fraMain.Top = 2310
        fraMain.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 4-0" Then
        fraCRM.Left = 5280
        fraCRM.Top = 2310
        fraCRM.Visible = True
   ElseIf Node.Key = ROOT_TREE & " 4-1" Then
        fraWHTax.Left = 5280
        fraWHTax.Top = 2310
        fraWHTax.Visible = True
   End If
End Sub

