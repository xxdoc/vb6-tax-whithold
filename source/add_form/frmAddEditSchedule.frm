VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditSchedule 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditSchedule.frx":0000
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
      TabIndex        =   12
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   6
         Top             =   3180
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
         TabIndex        =   17
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
         TabIndex        =   2
         Top             =   1020
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlGuestNameLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   3
         Top             =   1470
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlSourceLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   4
         Top             =   1920
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextLookup uctlDestinationLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   5
         Top             =   2370
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   3975
         Left            =   150
         TabIndex        =   20
         Top             =   3720
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7011
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
         Column(1)       =   "frmAddEditSchedule.frx":27A2
         Column(2)       =   "frmAddEditSchedule.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditSchedule.frx":290E
         FormatStyle(2)  =   "frmAddEditSchedule.frx":2A6A
         FormatStyle(3)  =   "frmAddEditSchedule.frx":2B1A
         FormatStyle(4)  =   "frmAddEditSchedule.frx":2BCE
         FormatStyle(5)  =   "frmAddEditSchedule.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditSchedule.frx":2D5E
      End
      Begin VB.Label lblFM 
         Height          =   315
         Left            =   8760
         TabIndex        =   19
         Top             =   3300
         Width           =   2685
      End
      Begin VB.Label lblTravelDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5040
         TabIndex        =   18
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
         MouseIcon       =   "frmAddEditSchedule.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   10
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSchedule.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   11
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSchedule.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditSchedule.frx":3884
         ButtonStyle     =   3
      End
      Begin VB.Label lblSource 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   16
         Top             =   2010
         Width           =   1575
      End
      Begin VB.Label lblGuestName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   15
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblSlipBookNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   600
         TabIndex        =   14
         Top             =   1140
         Width           =   1155
      End
      Begin VB.Label lblDestination 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   13
         Top             =   2430
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditSchedule"
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

Private m_Employees As Collection
Public m_Agencies As Collection
Public m_Sources As Collection
Public m_Dests As Collection
Public m_Customers As Collection
Public Area As Long

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
      
      Call LoadMaster(uctlGuestNameLookup.MyCombo, m_Customers, MASTER_REVENUETYPE)
      Set uctlGuestNameLookup.MyCollection = m_Customers
      
      If Area = 1 Then
         Call LoadMaster(uctlSourceLookup.MyCombo, m_Sources, MASTER_TAXRATE)
      Else
         Call LoadMaster(uctlSourceLookup.MyCombo, m_Sources, MASTER_CONDITION)
      End If
      Set uctlSourceLookup.MyCollection = m_Sources
      
      Call LoadEmployee(uctlDestinationLookup.MyCombo, m_Dests)
      Set uctlDestinationLookup.MyCollection = m_Dests
      
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
   'Debug.Print ColIndex & " " & NewColWidth
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
   Col.Width = 2130
   Col.Caption = MapText("เลขที่ใบจอง")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 3000
   Col.Caption = MapText("ชื่อแขก")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2415
   Col.Caption = MapText("ต้นทาง")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2370
   Col.Caption = MapText("หมายเลขห้อง")
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 1620
   Col.Caption = MapText("เวลา")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblTravelDate, MapText("วันที่เดินทาง"))
   Call InitNormalLabel(lblSlipBookNo, MapText("เลขที่ใบจอง"))
   If Area = 1 Then
      Call InitNormalLabel(lblSource, MapText("หมายเลขรถ"))
   ElseIf Area = 2 Then
      Call InitNormalLabel(lblSource, MapText("หมายเลขเรือ"))
   End If
   Call InitNormalLabel(lblGuestName, MapText("โซน"))
   Call InitNormalLabel(lblDestination, MapText("ผู้รับผิดชอบ"))

   Call InitNormalLabel(lblFM, MapText(""))
   
   Call txtSlipBookNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)

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
   
   Call InitGrid1
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายละเอียด")
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
