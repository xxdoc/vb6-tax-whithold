VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditCustomer 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditCustomer.frx":0000
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
      TabIndex        =   18
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjMtpTax.uctlTextLookup uctlSaleByLookup 
         Height          =   465
         Left            =   1860
         TabIndex        =   10
         Top             =   3720
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   820
      End
      Begin VB.ComboBox cboEnterpriseType 
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
         Left            =   6990
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2820
         Width           =   3495
      End
      Begin VB.ComboBox cboBusinessType 
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
         TabIndex        =   7
         Top             =   2820
         Width           =   3495
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   11
         Top             =   4470
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
      Begin prjMtpTax.uctlTextBox txtName 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1470
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtShortName 
         Height          =   435
         Left            =   1860
         TabIndex        =   0
         Top             =   1020
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtEmail 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   1920
         Width           =   7005
         _ExtentX        =   12356
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtWebSite 
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   2370
         Width           =   9225
         _ExtentX        =   16960
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtBusinessDesc 
         Height          =   450
         Left            =   1860
         TabIndex        =   9
         Top             =   3270
         Width           =   9225
         _ExtentX        =   16907
         _ExtentY        =   794
      End
      Begin prjMtpTax.uctlTextBox txtCredit 
         Height          =   435
         Left            =   5700
         TabIndex        =   2
         Top             =   1020
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin prjMtpTax.uctlTextBox txtDiscountPercent 
         Height          =   435
         Left            =   8070
         TabIndex        =   3
         Top             =   1020
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   2715
         Left            =   150
         TabIndex        =   12
         Top             =   5010
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4789
         Version         =   "2.0"
         BoundColumnIndex=   ""
         ReplaceColumnIndex=   ""
         TabKeyBehavior  =   1
         MethodHoldFields=   -1  'True
         AllowColumnDrag =   0   'False
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         DataMode        =   99
         HeaderFontName  =   "AngsanaUPC"
         HeaderFontBold  =   -1  'True
         HeaderFontSize  =   14.25
         HeaderFontWeight=   700
         FontSize        =   9.75
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   480
         IntProp1        =   0
         IntProp2        =   0
         IntProp7        =   0
         ColumnsCount    =   2
         Column(1)       =   "frmAddEditCustomer.frx":27A2
         Column(2)       =   "frmAddEditCustomer.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditCustomer.frx":290E
         FormatStyle(2)  =   "frmAddEditCustomer.frx":2A6A
         FormatStyle(3)  =   "frmAddEditCustomer.frx":2B1A
         FormatStyle(4)  =   "frmAddEditCustomer.frx":2BCE
         FormatStyle(5)  =   "frmAddEditCustomer.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditCustomer.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
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
         MouseIcon       =   "frmAddEditCustomer.frx":2F36
         ButtonStyle     =   3
      End
      Begin VB.Label lblResponseBy 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   60
         TabIndex        =   30
         Top             =   3780
         Width           =   1695
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   16
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   17
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
         TabIndex        =   14
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   13
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   15
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomer.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblDiscountPercent 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6900
         TabIndex        =   28
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label Label2 
         Height          =   315
         Left            =   6570
         TabIndex        =   27
         Top             =   1080
         Width           =   345
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4710
         TabIndex        =   26
         Top             =   1110
         Width           =   885
      End
      Begin VB.Label lblBusinessDesc 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   30
         TabIndex        =   25
         Top             =   3390
         Width           =   1695
      End
      Begin VB.Label lblWebsite 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   24
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   23
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblShortName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   22
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblEnterpriseType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   5400
         TabIndex        =   21
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label lblBusinessType 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   20
         Top             =   2880
         Width           =   1485
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmAddEditCustomer"
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
Private m_Employees As Collection

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long

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
      
      txtEmail.Text = m_Customer.EMAIL
      txtWebSite.Text = m_Customer.WEBSITE
      cboBusinessType.ListIndex = IDToListIndex(cboBusinessType, m_Customer.CUSTOMER_TYPE)
      cboEnterpriseType.ListIndex = IDToListIndex(cboEnterpriseType, m_Customer.CUSTOMER_GRADE)
      txtShortName.Text = m_Customer.CUSTOMER_CODE
      txtBusinessDesc.Text = m_Customer.BUSINESS_DESC
      txtCredit.Text = m_Customer.CREDIT
      txtDiscountPercent.Text = m_Customer.NORMAL_DISCOUNT
      uctlSaleByLookup.MyCombo.ListIndex = IDToListIndex(uctlSaleByLookup.MyCombo, m_Customer.RESPONSE_BY)

      Dim Name As cName
      Dim CstName As CCustomerName
      If (Not m_Customer.CstNames Is Nothing) And (m_Customer.CstNames.Count > 0) Then
         Set CstName = m_Customer.CstNames(1)
         Set Name = CstName.Name
         txtName.Text = Name.LONG_NAME
      Else
         txtName.Text = ""
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
      If Not VerifyAccessRight("MAIN_CUSTOMER_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("MAIN_CUSTOMER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblShortName, txtShortName, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBusinessType, cboBusinessType, True) Then
      Exit Function
   End If
   If Not VerifyCombo(lblEnterpriseType, cboEnterpriseType, True) Then
      Exit Function
   End If

   If Not CheckUniqueNs(CUSTCODE_UNIQUE, txtShortName.Text, ID) Then
      glbErrorLog.LocalErrorMsg = MapText("�բ�����") & " " & txtShortName.Text & " " & MapText("������к�����")
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
   m_Customer.EMAIL = txtEmail.Text
   m_Customer.WEBSITE = txtWebSite.Text
   m_Customer.CUSTOMER_TYPE = cboBusinessType.ItemData(Minus2Zero(cboBusinessType.ListIndex))
   m_Customer.CUSTOMER_GRADE = cboEnterpriseType.ItemData(Minus2Zero(cboEnterpriseType.ListIndex))
   m_Customer.CREDIT = Val(txtCredit.Text)
   m_Customer.CUSTOMER_CODE = txtShortName.Text
   m_Customer.BUSINESS_DESC = txtBusinessDesc.Text
   m_Customer.NORMAL_DISCOUNT = Val(txtDiscountPercent.Text)
   m_Customer.RESPONSE_BY = uctlSaleByLookup.MyCombo.ItemData(Minus2Zero(uctlSaleByLookup.MyCombo.ListIndex))

   'Create Dummy account
   If m_Customer.CstAccounts.Count <= 0 Then
      Dim Acc As CAccount
      
      Set Acc = New CAccount
      
      Acc.ACCOUNT_NO = m_Customer.CUSTOMER_CODE
      Acc.Flag = "A"
      
      Call m_Customer.CstAccounts.Add(Acc)
      
      Set Acc = Nothing
   End If

   Dim CstName As CCustomerName
   If m_Customer.CstNames.Count <= 0 Then
      Set CstName = New CCustomerName
      CstName.Flag = "A"
      Call m_Customer.CstNames.Add(CstName)
   Else
      Set CstName = m_Customer.CstNames.Item(1)
      CstName.Flag = "E"
   End If
   
   Dim Name As cName
   If m_Customer.CstNames.Count <= 0 Then
      Set Name = CstName.Name
      Name.LONG_NAME = txtName.Text
      Name.SHORT_NAME = txtShortName.Text
      Name.Flag = "A"
   Else
      Set Name = CstName.Name
      Name.LONG_NAME = txtName.Text
      Name.SHORT_NAME = txtShortName.Text
      Name.Flag = "E"
   End If
   
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

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   OKClick = False
   If TabStrip1.SelectedItem.Index = 1 Then
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.ShowMode = SHOW_ADD
      frmAddEditCustomerAddress.HeaderText = MapText("�����������")
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Set frmAddEditCustomerAccount.TempCollection = m_Customer.CstAccounts
      frmAddEditCustomerAccount.ShowMode = SHOW_ADD
      frmAddEditCustomerAccount.HeaderText = MapText("�����ѭ���١���")
      Load frmAddEditCustomerAccount
      frmAddEditCustomerAccount.Show 1

      OKClick = frmAddEditCustomerAccount.OKClick

      Unload frmAddEditCustomerAccount
      Set frmAddEditCustomerAccount = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstAccounts)
         GridEX1.Rebind
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

   If Trim(txtShortName.Text) = "" Then
      Call glbDatabaseMngr.GenerateNumber(CUSTOMER_NUMBER, No, glbErrorLog)
      txtShortName.Text = No
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_Customer.CstAddr.Remove (ID2)
      Else
         m_Customer.CstAddr.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_Customer.CstAddr)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      If m_Customer.CstAccounts.Item(ID2).MASTER_FLAG = "Y" Then
         glbErrorLog.LocalErrorMsg = "������öź�ѭ�վ�鹰ҹ��"
         glbErrorLog.ShowUserError
         Exit Sub
      End If
   
      If ID1 <= 0 Then
         m_Customer.CstAccounts.Remove (ID2)
      Else
         m_Customer.CstAccounts.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_Customer.CstAccounts)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

'   If Not VerifyAccessRight("GROUP_QUERY_RIGHT") Then
'      Exit Sub
'   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditCustomerAddress.ID = ID
      Set frmAddEditCustomerAddress.TempCollection = m_Customer.CstAddr
      frmAddEditCustomerAddress.HeaderText = MapText("��䢷������")
      frmAddEditCustomerAddress.ShowMode = SHOW_EDIT
      Load frmAddEditCustomerAddress
      frmAddEditCustomerAddress.Show 1

      OKClick = frmAddEditCustomerAddress.OKClick

      Unload frmAddEditCustomerAddress
      Set frmAddEditCustomerAddress = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstAddr)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      frmAddEditCustomerAccount.ID = ID
      Set frmAddEditCustomerAccount.TempCollection = m_Customer.CstAccounts
      frmAddEditCustomerAccount.HeaderText = MapText("��䢺ѭ���١���")
      frmAddEditCustomerAccount.ShowMode = SHOW_EDIT
      Load frmAddEditCustomerAccount
      frmAddEditCustomerAccount.Show 1

      OKClick = frmAddEditCustomerAccount.OKClick

      Unload frmAddEditCustomerAccount
      Set frmAddEditCustomerAccount = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_Customer.CstAccounts)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()

   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadMaster(cboBusinessType, , MASTER_CUSTYPE)
      Call LoadMaster(cboEnterpriseType, , MASTER_CUSGRADE)
      
      Call LoadEmployee(uctlSaleByLookup.MyCombo, m_Employees)
      Set uctlSaleByLookup.MyCollection = m_Employees
      
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
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
      Call cmdEdit_Click
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
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
  ' 'Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 11550
   Col.Caption = MapText("�������")
End Sub

Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.itemcount = 0
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX1.ColumnHeaderFont.Bold = True
   GridEX1.ColumnHeaderFont.Name = GLB_FONT
   GridEX1.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 1470
   Col.Caption = MapText("�Ţ���ѭ��")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 6855
   Col.Caption = MapText("��������´")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 3240
   Col.Caption = MapText("ᾤࡨ")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblWebsite, MapText("���䫵�"))
   Call InitNormalLabel(lblShortName, MapText("�����١���"))
   Call InitNormalLabel(lblEnterpriseType, MapText("�дѺ�١���"))
   Call InitNormalLabel(lblName, MapText("�����١���"))
   Call InitNormalLabel(lblEmail, MapText("�������"))
   Call InitNormalLabel(lblBusinessType, MapText("�������١���"))
   Call InitNormalLabel(lblBusinessDesc, MapText("��������´�١���"))
   Call InitNormalLabel(lblCredit, MapText("�ôԵ"))
   Call InitNormalLabel(Label2, MapText("�ѹ"))
   Call InitNormalLabel(lblDiscountPercent, MapText("% ��ǹŴ"))
   Call InitNormalLabel(lblResponseBy, MapText("����Ѻ�Դ�ͺ"))
   
   Call InitCombo(cboBusinessType)
   Call InitCombo(cboEnterpriseType)
   
   Call txtShortName.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtEmail.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtWebSite.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtBusinessDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtCredit.SetTextLenType(TEXT_INTEGER, glbSetting.MONEY_TYPE)
   
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
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdAuto, MapText("A"))
   
   Call InitGrid1
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("�������")
   TabStrip1.Tabs.Add().Caption = MapText("�ѭ���١���")
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
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
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
      Call InitGrid1
      GridEX1.itemcount = CountItem(m_Customer.CstAddr)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
      Call InitGrid2
      GridEX1.itemcount = CountItem(m_Customer.CstAccounts)
      GridEX1.Rebind
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtBusinessDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtCredit_Change()
   m_HasModify = True
   If Val(txtCredit.Text) = 0 Then
      txtDiscountPercent.Enabled = True
   Else
      txtDiscountPercent.Enabled = False
   End If
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

Private Sub txtShortName_Change()
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

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub uctlSaleByLookup_Change()
   m_HasModify = True
End Sub
