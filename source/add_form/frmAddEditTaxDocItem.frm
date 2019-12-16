VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditTaxDocItem 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditTaxDocItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10755
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   6345
      Left            =   0
      TabIndex        =   14
      Top             =   600
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   11192
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboTaxRate 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3000
         Width           =   1515
      End
      Begin prjMtpTax.uctlDate uctlPaidDate 
         Height          =   405
         Left            =   1710
         TabIndex        =   3
         Top             =   1650
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlSupplierLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   1
         Top             =   750
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboAddress 
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
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   8355
      End
      Begin prjMtpTax.uctlTextBox txtPaidAmount 
         Height          =   435
         Left            =   1710
         TabIndex        =   5
         Top             =   2550
         Width           =   2205
         _ExtentX        =   7858
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtProvince 
         Height          =   435
         Left            =   1710
         TabIndex        =   7
         Top             =   3450
         Width           =   2235
         _ExtentX        =   6535
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextLookup uctlRevenueTypeLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   4
         Top             =   2100
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextLookup uctlConditionLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   8
         Top             =   3930
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtRefNo 
         Height          =   435
         Left            =   1710
         TabIndex        =   0
         Top             =   270
         Width           =   2595
         _ExtentX        =   7858
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtNote 
         Height          =   435
         Left            =   1710
         TabIndex        =   9
         Top             =   4380
         Width           =   8325
         _ExtentX        =   7858
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextLookup uctlBranchLookup 
         Height          =   435
         Left            =   1710
         TabIndex        =   10
         Top             =   4830
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   4890
         Width           =   1485
      End
      Begin VB.Label lblNote 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   26
         Top             =   4440
         Width           =   1515
      End
      Begin VB.Label lblRefNo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   25
         Top             =   330
         Width           =   1515
      End
      Begin VB.Label lblCondition 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   3990
         Width           =   1485
      End
      Begin VB.Label Label2 
         Height          =   375
         Left            =   4050
         TabIndex        =   23
         Top             =   3030
         Width           =   1095
      End
      Begin VB.Label Label1 
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   2100
         Width           =   1095
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3750
         TabIndex        =   11
         Top             =   5580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxDocItem.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   5400
         TabIndex        =   12
         Top             =   5580
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin VB.Label lblVillage 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label lblSupplier 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label lblSoi 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   1260
         Width           =   1485
      End
      Begin VB.Label lblMoo 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblRoad 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   3060
         Width           =   1485
      End
      Begin VB.Label lblAmphur 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3510
         Width           =   1485
      End
      Begin VB.Label lblDistrict 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   2610
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmAddEditTaxDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public Header As String
Public ShowMode As SHOW_MODE_TYPE
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection
Public m_Suppliers As Collection
Public m_RevenueTypes As Collection
Public m_Conditions As Collection
Public m_TaxRates As Collection
Public m_Branches As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboAddress_Click()
   m_HasModify = True
End Sub

Private Sub chkBangkok_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cboTaxRate_Click()
Dim TempID As Long
Dim Mr As CMasterRef

   TempID = cboTaxRate.ItemData(Minus2Zero(cboTaxRate.ListIndex))
   If TempID > 0 Then
      Set Mr = m_TaxRates(Trim(Str(TempID)))
                If InStr(1, Mr.KEY_NAME, "%", 1) <> 0 Then
                    Mr.KEY_NAME = Left(Mr.KEY_NAME, InStr(1, Mr.KEY_NAME, "%", 1) - 1)
                End If
      txtProvince.Text = Val(Mr.KEY_NAME) * Val(txtPaidAmount.Text) / 100
   End If
   
   m_HasModify = True
End Sub

Private Sub cboTaxRate_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   Call InitNormalLabel(lblSoi, MapText("ที่อยู่"))
   Call InitNormalLabel(lblMoo, MapText("วันที่จ่าย"))
   Call InitNormalLabel(lblVillage, MapText("ประเภทเงินได้"))
   Call InitNormalLabel(lblRoad, MapText("อัตราภาษี"))
   Call InitNormalLabel(lblDistrict, MapText("จำนวนเงินที่จ่าย"))
   Call InitNormalLabel(lblAmphur, MapText("จำนวนเงินนำส่ง"))
   Call InitNormalLabel(lblSupplier, MapText("ซัพพลายเออร์"))
   Call InitNormalLabel(Label1, MapText("บาท"))
   Call InitNormalLabel(Label2, MapText("บาท"))
   Call InitNormalLabel(lblCondition, MapText("เงื่อนไข"))
   Call InitNormalLabel(lblNote, MapText("หมายเหตุ"))
   Call InitNormalLabel(lblRefNo, MapText("ใบสำคัญจ่าย"))
   Call InitNormalLabel(lblBranch, MapText("สาขา"))

   Call txtProvince.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtPaidAmount.SetTextLenType(TEXT_FLOAT_MONEY, glbSetting.MONEY_TYPE)
   Call txtRefNo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   Call InitCombo(cboAddress)
   Call InitCombo(cboTaxRate)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
      
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Dim EnpAddr As CTaxDocItem
         
         Set EnpAddr = TempCollection.Item(ID)
         
         uctlSupplierLookup.MyCombo.ListIndex = IDToListIndex(uctlSupplierLookup.MyCombo, EnpAddr.SUPPLIER_ID)
         cboAddress.ListIndex = IDToListIndex(cboAddress, EnpAddr.ADDRESS_ID)
         uctlPaidDate.ShowDate = EnpAddr.PAY_DATE
         uctlRevenueTypeLookup.MyCombo.ListIndex = IDToListIndex(uctlRevenueTypeLookup.MyCombo, EnpAddr.REVENUE_TYPE)
         txtPaidAmount.Text = EnpAddr.PAID_AMOUNT
         cboTaxRate.ListIndex = IDToListIndex(cboTaxRate, EnpAddr.TAX_RATE)
         txtProvince.Text = EnpAddr.WH_AMOUNT
         uctlConditionLookup.MyCombo.ListIndex = IDToListIndex(uctlConditionLookup.MyCombo, EnpAddr.CONDITION_ID)
         txtRefNo.Text = EnpAddr.REF_NO
         txtNote.Text = EnpAddr.NOTE
         uctlBranchLookup.MyCombo.ListIndex = IDToListIndex(uctlBranchLookup.MyCombo, EnpAddr.BRANCH_ID)
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long
   
'   If Not VerifyTextControl(lblRefNo, txtRefNo, False) Then
'      Exit Function
'   End If
   If Not VerifyCombo(lblSupplier, uctlSupplierLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblSoi, cboAddress, False) Then
      Exit Function
   End If
   If Not VerifyDate(lblMoo, uctlPaidDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblVillage, uctlRevenueTypeLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblRoad, cboTaxRate, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblDistrict, txtPaidAmount, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblAmphur, txtProvince, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblCondition, uctlConditionLookup.MyCombo, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblBranch, uctlBranchLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim EnpAddress As CTaxDocItem
   If ShowMode = SHOW_ADD Then
      Set EnpAddress = New CTaxDocItem
      EnpAddress.Flag = "A"
      Call TempCollection.Add(EnpAddress)
   Else
      Set EnpAddress = TempCollection.Item(ID)
      If EnpAddress.Flag <> "A" Then
         EnpAddress.Flag = "E"
      End If
   End If
      
   EnpAddress.SUPPLIER_ID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   EnpAddress.SUPPLIER_NAME = uctlSupplierLookup.MyCombo.Text
   EnpAddress.ADDRESS_ID = cboAddress.ItemData(Minus2Zero(cboAddress.ListIndex))
   EnpAddress.PAY_DATE = uctlPaidDate.ShowDate
   EnpAddress.REVENUE_TYPE = uctlRevenueTypeLookup.MyCombo.ItemData(Minus2Zero(uctlRevenueTypeLookup.MyCombo.ListIndex))
   EnpAddress.REVENUE_TYPE_NAME = uctlRevenueTypeLookup.MyCombo.Text
   EnpAddress.PAID_AMOUNT = Val(txtPaidAmount.Text)
   EnpAddress.TAX_RATE = cboTaxRate.ItemData(Minus2Zero(cboTaxRate.ListIndex))
   EnpAddress.RATETYPE_NAME = cboTaxRate.Text
   EnpAddress.WH_AMOUNT = Val(txtProvince.Text)
   EnpAddress.CONDITION_ID = uctlConditionLookup.MyCombo.ItemData(Minus2Zero(uctlConditionLookup.MyCombo.ListIndex))
   EnpAddress.CONDITION_NAME = uctlConditionLookup.MyCombo.Text
   EnpAddress.REF_NO = txtRefNo.Text
   EnpAddress.NOTE = txtNote.Text
   EnpAddress.BRANCH_ID = uctlBranchLookup.MyCombo.ItemData(Minus2Zero(uctlBranchLookup.MyCombo.ListIndex))
   
   Set EnpAddress = Nothing
   SaveData = True
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call LoadMaster(cboTaxRate, m_TaxRates, MASTER_TAXRATE)
      
      Call LoadSupplier(uctlSupplierLookup.MyCombo, m_Suppliers)
      Set uctlSupplierLookup.MyCollection = m_Suppliers
      
      Call LoadMaster(uctlRevenueTypeLookup.MyCombo, m_RevenueTypes, MASTER_REVENUETYPE)
      Set uctlRevenueTypeLookup.MyCollection = m_RevenueTypes
      
      Call LoadMaster(uctlConditionLookup.MyCombo, m_Conditions, MASTER_CONDITION)
      Set uctlConditionLookup.MyCollection = m_Conditions
      
      Call LoadMaster(uctlBranchLookup.MyCombo, m_Branches, MASTER_BRANCH)
      Set uctlBranchLookup.MyCollection = m_Branches
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
         Call QueryData(True)
      End If
      
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
'      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   Set m_RevenueTypes = New Collection
   Set m_Suppliers = New Collection
   Set m_Conditions = New Collection
   Set m_TaxRates = New Collection
   Set m_Branches = New Collection
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_RevenueTypes = Nothing
   Set m_Suppliers = Nothing
   Set m_Conditions = Nothing
   Set m_TaxRates = Nothing
   Set m_Branches = Nothing
End Sub

Private Sub txtDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtKeyName_Change()
   m_HasModify = True
End Sub

Private Sub txtThaiMsg_Change()
   m_HasModify = True
End Sub

Private Sub txtAmphur_Change()
   m_HasModify = True
End Sub

Private Sub PaidAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub txtHomeNo_Change()
   m_HasModify = True
End Sub

Private Sub txtMoo_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub Label3_Click()

End Sub

Private Sub txtBranch_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtPaidAmount_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtRoad_Change()
   m_HasModify = True
End Sub

Private Sub txtSoi_Change()
   m_HasModify = True
End Sub

Private Sub txtVillage_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub txtRefNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlBranchLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlConditionLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlPaidDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlRevenueTypeLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSupplierLookup_Change()
Dim SupplierID As Long
Dim C As CSupplier

   SupplierID = uctlSupplierLookup.MyCombo.ItemData(Minus2Zero(uctlSupplierLookup.MyCombo.ListIndex))
   If SupplierID > 0 Then
      Set C = m_Suppliers(Trim(Str(SupplierID)))
      Call LoadSupplierAddress(cboAddress, , SupplierID, True)
   Else
      cboAddress.ListIndex = -1
   End If
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub
