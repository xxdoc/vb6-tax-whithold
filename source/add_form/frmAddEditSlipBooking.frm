VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddEditSlipBooking 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditSlipBooking.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboLanguage 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8700
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1920
         Width           =   2925
      End
      Begin Threed.SSFrame fraGeneral 
         Height          =   3135
         Left            =   150
         TabIndex        =   40
         Top             =   4590
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   5530
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjMtpTax.uctlTextBox txtVocherNo 
            Height          =   435
            Left            =   1710
            TabIndex        =   11
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtCollectPrice 
            Height          =   435
            Left            =   5700
            TabIndex        =   12
            Top             =   210
            Width           =   1395
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextLookup uctlAgencyLookup 
            Height          =   405
            Left            =   1710
            TabIndex        =   13
            Top             =   660
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   714
         End
         Begin prjMtpTax.uctlTextBox txtSender 
            Height          =   435
            Left            =   1710
            TabIndex        =   14
            Top             =   1110
            Width           =   5385
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtReserve 
            Height          =   435
            Left            =   1710
            TabIndex        =   15
            Top             =   1560
            Width           =   5385
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtNote 
            Height          =   435
            Left            =   1710
            TabIndex        =   16
            Top             =   2010
            Width           =   9375
            _ExtentX        =   4471
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtAdl 
            Height          =   435
            Left            =   1710
            TabIndex        =   17
            Top             =   2460
            Width           =   885
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtChd 
            Height          =   435
            Left            =   3270
            TabIndex        =   18
            Top             =   2460
            Width           =   945
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtInf 
            Height          =   435
            Left            =   5010
            TabIndex        =   19
            Top             =   2460
            Width           =   1065
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtFoc 
            Height          =   435
            Left            =   6900
            TabIndex        =   20
            Top             =   2460
            Width           =   1035
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin prjMtpTax.uctlTextBox txtIns 
            Height          =   435
            Left            =   8700
            TabIndex        =   21
            Top             =   2460
            Width           =   1035
            _ExtentX        =   3201
            _ExtentY        =   767
         End
         Begin VB.Label lblIns 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7980
            TabIndex        =   53
            Top             =   2550
            Width           =   645
         End
         Begin VB.Label lblFoc 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6180
            TabIndex        =   52
            Top             =   2550
            Width           =   645
         End
         Begin VB.Label lblInf 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4320
            TabIndex        =   51
            Top             =   2550
            Width           =   615
         End
         Begin VB.Label lblChd 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2670
            TabIndex        =   50
            Top             =   2550
            Width           =   525
         End
         Begin VB.Label lblAdl 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   360
            TabIndex        =   49
            Top             =   2520
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   210
            TabIndex        =   47
            Top             =   2100
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   210
            TabIndex        =   46
            Top             =   1650
            Width           =   1395
         End
         Begin VB.Label lblSender 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   210
            TabIndex        =   45
            Top             =   1200
            Width           =   1395
         End
         Begin VB.Label lblAgency 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   60
            TabIndex        =   44
            Top             =   750
            Width           =   1575
         End
         Begin VB.Label lblBaht 
            Height          =   315
            Left            =   7170
            TabIndex        =   43
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label lblCollectPrice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4350
            TabIndex        =   42
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label lblVocherNo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   450
            TabIndex        =   41
            Top             =   330
            Width           =   1155
         End
      End
      Begin VB.ComboBox cboVehicle 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1500
         Width           =   3075
      End
      Begin prjMtpTax.uctlTime uctlTime1 
         Height          =   405
         Left            =   10500
         TabIndex        =   37
         Top             =   1470
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlDate uctlBookingDate 
         Height          =   405
         Left            =   6630
         TabIndex        =   2
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlSaleByLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   10
         Top             =   3270
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   22
         Top             =   4050
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjMtpTax.uctlTextBox txtSlipBookNo 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjMtpTax.uctlDate uctlTravelDate 
         Height          =   405
         Left            =   6630
         TabIndex        =   4
         Top             =   1470
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlGuestNameLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   5
         Top             =   1920
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlSourceLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   7
         Top             =   2370
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextLookup uctlDestinationLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   9
         Top             =   2820
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtRoomNo 
         Height          =   435
         Left            =   8700
         TabIndex        =   8
         Top             =   2370
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   767
      End
      Begin VB.Label lblFM 
         Height          =   315
         Left            =   8760
         TabIndex        =   54
         Top             =   3300
         Width           =   2685
      End
      Begin VB.Label lblLanguage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7290
         TabIndex        =   48
         Top             =   1980
         Width           =   1335
      End
      Begin VB.Label lblRoomNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7350
         TabIndex        =   39
         Top             =   2430
         Width           =   1275
      End
      Begin VB.Label lblVehicle 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   150
         TabIndex        =   38
         Top             =   1560
         Width           =   1605
      End
      Begin VB.Label lblTravelDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   36
         Top             =   1470
         Width           =   1485
      End
      Begin VB.Label lblBookingDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   35
         Top             =   1020
         Width           =   1485
      End
      Begin Threed.SSCommand cmdAuto 
         Height          =   405
         Left            =   4410
         TabIndex        =   1
         Top             =   1020
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSlipBooking.frx":27A2
         ButtonStyle     =   3
      End
      Begin VB.Label lblResponseBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   34
         Top             =   3330
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   26
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSlipBooking.frx":2ABC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   27
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   24
         Top             =   7830
         Visible         =   0   'False
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   23
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSlipBooking.frx":2DD6
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   25
         Top             =   7830
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSlipBooking.frx":30F0
         ButtonStyle     =   3
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblGuestName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   31
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblSlipBookNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   30
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblDestination 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   29
         Top             =   2880
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditSlipBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Customer As CCustomer

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public Area As Long

Private m_Employees As Collection
Public m_Agencies As Collection
Public m_Sources As Collection
Public m_Dests As Collection
Public m_Customers As Collection

Private FileName As String

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_Customer.CUSTOMER_ID = ID
      If Not glbDaily.QueryCustomer(m_Customer, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If itemcount > 0 Then
      Call m_Customer.PopulateFromRS(1, m_Rs)
      
      txtSlipBookNo.Text = m_Customer.CUSTOMER_CODE
      uctlSaleByLookup.MyCombo.ListIndex = IDToListIndex(uctlSaleByLookup.MyCombo, m_Customer.RESPONSE_BY)

      Dim Name As cName
      Dim CstName As CCustomerName
      If (Not m_Customer.CstNames Is Nothing) And (m_Customer.CstNames.Count > 0) Then
         Set CstName = m_Customer.CstNames(1)
         Set Name = CstName.Name
      Else
      End If
   Else
      ShowMode = SHOW_ADD
   End If
   
   If ShowMode = SHOW_ADD Then
      Dim Acc As CAccount
      Dim Subc As CSubscriber
'      Dim Agr As CAgreement
      
      Set Acc = New CAccount
      Set Subc = New CSubscriber
'      Set Agr = New CAgreement
      
      Acc.AddEditMode = ShowMode
      Subc.AddEditMode = ShowMode
'      Agr.AddEditMode = ShowMode
      
      Acc.Flag = "A"
      Subc.Flag = "A"
'      Agr.Flag = "A"
      
      Call Acc.ActSubs.Add(Subc)
'      Call Acc.ActAgrmnts.Add(Agr)
      Call m_Customer.CstAccounts.Add(Acc)
      
      Acc.ACCOUNT_NO = "DMY000"
      Acc.ACCOUNT_NAME = "DMY000"
      Acc.ACCOUNT_STATUS = -1
      Acc.ACCOUNT_TYPE = -1
      Acc.MASTER_FLAG = "Y"
      Acc.ENABLE_FLAG = "Y"
      
      Subc.DUMMY_FLAG = "Y"
      Subc.SUBSCRIBER_NO = "DMY999"
      Subc.SUBSCRIBER_STATUS = "Y"
      
'      Agr.SOC_CODE = ""
'      Agr.SOC_FEATURE_ID = -1
'      Agr.SOC_ID = -1
'      Agr.EXCLUDE_FLAG = "N"
'      Agr.EFFECTIVE_DATE = -2
'      Agr.EXPIRE_DATE = -1
'      Agr.ISSUE_DATE = Now
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkEnable_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("CRM_ORDER_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("CRM_ORDER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblSlipBookNo, txtSlipBookNo, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblBookingDate, uctlBookingDate, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblTravelDate, uctlTravelDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblGuestName, uctlGuestNameLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSource, uctlSourceLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblDestination, uctlDestinationLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblCollectPrice, txtCollectPrice, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAdl, txtAdl, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblChd, txtChd, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblInf, txtInf, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblFoc, txtFoc, True) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblIns, txtIns, True) Then
      Exit Function
   End If
   
   If Not CheckUniqueNs(BOOKSLIP_NO, txtSlipBookNo.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("มีข้อมูล") & " " & txtSlipBookNo.Text & " " & MapText("อยู่ในระบบแล้ว")
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   m_Customer.AddEditMode = ShowMode
   m_Customer.BIRTH_DATE = -1
   m_Customer.CUSTOMER_PASSWORD = ""
   m_Customer.CUSTOMER_CODE = txtSlipBookNo.Text
   m_Customer.RESPONSE_BY = uctlSaleByLookup.MyCombo.ItemData(Minus2Zero(uctlSaleByLookup.MyCombo.ListIndex))
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditCustomer(m_Customer, IsOK, True, glbErrorLog) Then
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      SaveData = False
      Call EnableForm(Me, True)
      Exit Function
   End If
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
End Function

Private Sub cboBusinessGroup_Click()
   m_HasModify = True
End Sub

Private Sub cboBusinessType_Click()
   m_HasModify = True
End Sub

Private Sub cboEnterpriseType_Click()
   m_HasModify = True
End Sub

Private Sub cboLanguage_Click()
   m_HasModify = True
End Sub

Private Sub cboVehicle_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.ShowMode = SHOW_ADD
      frmAddEditCustomerAddress.HeaderText = MapText("เพิ่มที่อยู่")
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Set frmAddEditCustomerAccount.TempCollection = m_Customer.CstAccounts
      frmAddEditCustomerAccount.ShowMode = SHOW_ADD
      frmAddEditCustomerAccount.HeaderText = MapText("เพิ่มบัญชีลูกค้า")
      Load frmAddEditCustomerAccount
      frmAddEditCustomerAccount.Show 1

      OKClick = frmAddEditCustomerAccount.OKClick

      Unload frmAddEditCustomerAccount
      Set frmAddEditCustomerAccount = Nothing

      If OKClick Then
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdAuto_Click()
Dim No As String

   If Trim(txtSlipBookNo.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(CUSTOMER_NUMBER, No, glbErrorLog)
      txtSlipBookNo.Text = No
   End If
End Sub


Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadEmployee(uctlSaleByLookup.MyCombo, m_Employees)
      Set uctlSaleByLookup.MyCollection = m_Employees
      
      Call LoadCustomer(uctlGuestNameLookup.MyCombo, m_Customers)
      Set uctlGuestNameLookup.MyCollection = m_Customers
      
      Call LoadSupplier(uctlAgencyLookup.MyCombo, m_Agencies)
      Set uctlAgencyLookup.MyCollection = m_Agencies
      
      Call LoadMaster(uctlSourceLookup.MyCombo, m_Sources, MASTER_SOURCE)
      Set uctlSourceLookup.MyCollection = m_Sources
      
      Call LoadMaster(uctlDestinationLookup.MyCombo, m_Dests, MASTER_DEST)
      Set uctlDestinationLookup.MyCollection = m_Dests
      
      Call LoadMaster(cboVehicle, , MASTER_BRANCH)
      Call LoadMaster(cboLanguage, , MASTER_LANGUAGE)
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_Customer.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_Customer.QueryFlag = 0
         Call QueryData(False)
      End If
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Customer = Nothing
   Set m_Employees = Nothing
   Set m_Agencies = Nothing
   Set m_Sources = Nothing
   Set m_Dests = Nothing
   Set m_Customers = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
 '  'Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   fraGeneral.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblBookingDate, MapText("วันที่จอง"))
   Call InitNormalLabel(lblTravelDate, MapText("วันที่เดินทาง"))
   Call InitNormalLabel(lblSlipBookNo, MapText("เลขที่ใบจอง"))
   Call InitNormalLabel(lblVehicle, MapText("สาขา"))
   Call InitNormalLabel(lblRoomNo, MapText("หมายเลขห้อง"))
   Call InitNormalLabel(lblSource, MapText("ต้นทาง"))
   Call InitNormalLabel(lblGuestName, MapText("ชื่อแขก"))
   Call InitNormalLabel(lblDestination, MapText("ปลายทาง"))
   Call InitNormalLabel(lblResponseBy, MapText("ผู้ทำรายการ"))
   Call InitNormalLabel(lblLanguage, MapText("ภาษา"))
   
   Call InitNormalLabel(lblVocherNo, MapText("Vocher NO."))
   Call InitNormalLabel(lblCollectPrice, MapText("จำนวนเงิน"))
   Call InitNormalLabel(lblBaht, MapText("บาท"))
   Call InitNormalLabel(lblAgency, MapText("ซัพพลายเออร์"))
   Call InitNormalLabel(lblSender, MapText("ผู้ติดต่อ"))
   Call InitNormalLabel(Label1, MapText("ผู้จอง"))
   Call InitNormalLabel(Label2, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblAdl, MapText("ADL"))
   Call InitNormalLabel(lblChd, MapText("CHD"))
   Call InitNormalLabel(lblInf, MapText("INF"))
   Call InitNormalLabel(lblFoc, MapText("FOC"))
   Call InitNormalLabel(lblIns, MapText("INS"))
   Call InitNormalLabel(lblFM, MapText("FM-RES-01 Rev.00"))
   
   Call InitCombo(cboVehicle)
   Call InitCombo(cboLanguage)
   
   Call txtSlipBookNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtRoomNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtVocherNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtCollectPrice.SetTextLenType(TEXT_STRING, glbSetting.MONEY_TYPE)
   Call txtSender.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtReserve.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtAdl.SetTextLenType(TEXT_STRING, glbSetting.MONEY_TYPE)
   Call txtChd.SetTextLenType(TEXT_STRING, glbSetting.MONEY_TYPE)
   Call txtInf.SetTextLenType(TEXT_STRING, glbSetting.MONEY_TYPE)
   Call txtFoc.SetTextLenType(TEXT_STRING, glbSetting.MONEY_TYPE)
   Call txtIns.SetTextLenType(TEXT_STRING, glbSetting.MONEY_TYPE)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAuto.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายละเอียดทั่วไป")
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   OKClick = False
   Call InitFormLayout
      
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Customer = New CCustomer
   Set m_Employees = New Collection
   
   Set m_Agencies = New Collection
   Set m_Sources = New Collection
   Set m_Dests = New Collection
   Set m_Customers = New Collection
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_Customer.CstAddr Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CCustomerAddress
      Dim Addr As CAddress
      If m_Customer.CstAddr.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Customer.CstAddr, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
      Set Addr = CR.Addresses

      Values(1) = Addr.ADDRESS_ID
      Values(2) = RealIndex
      Values(3) = Addr.PackAddress
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_Customer.CstAccounts Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim Ca As CAccount
      If m_Customer.CstAccounts.Count <= 0 Then
         Exit Sub
      End If
      Set Ca = GetItem(m_Customer.CstAccounts, RowIndex, RealIndex)
      If Ca Is Nothing Then
         Exit Sub
      End If

      Values(1) = Ca.ACCOUNT_ID
      Values(2) = RealIndex
      Values(3) = Ca.ACCOUNT_NO
      Values(4) = Ca.NOTE
      Values(5) = ""
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtBusinessDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtAdl_Change()
   m_HasModify = True
End Sub

Private Sub txtChd_Change()
   m_HasModify = True
End Sub

Private Sub txtCollectPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtFoc_Change()
   m_HasModify = True
End Sub

Private Sub txtInf_Change()
   m_HasModify = True
End Sub

Private Sub txtIns_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtReserve_Change()
   m_HasModify = True
End Sub

Private Sub txtRoomNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtSlipBookNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlAgencyLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub txtVocherNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlBookingDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDestinationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlGuestNameLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSaleByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSourceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox2_Change()
   m_HasModify = True
End Sub
