VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditCustomerAccount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9435
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditCustomerAccount.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   9435
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   2835
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   5001
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjMtpTax.uctlTextLookup uctlSocLookup 
         Height          =   435
         Left            =   1800
         TabIndex        =   2
         Top             =   1260
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtAccountId 
         Height          =   465
         Left            =   1800
         TabIndex        =   0
         Top             =   300
         Width           =   4245
         _ExtentX        =   3519
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtDesc 
         Height          =   465
         Left            =   1800
         TabIndex        =   1
         Top             =   780
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   820
      End
      Begin VB.Label lblSocLookup 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   9
         Top             =   1320
         Width           =   1485
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   8
         Top             =   840
         Width           =   1485
      End
      Begin VB.Label lblAccountId 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   210
         TabIndex        =   7
         Top             =   360
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   3090
         TabIndex        =   3
         Top             =   1980
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditCustomerAccount.frx":08CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4740
         TabIndex        =   4
         Top             =   1980
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditCustomerAccount"
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

Private m_Socs As Collection

Private Sub cboTextType_Click()
   m_HasModify = True
End Sub

Private Sub cboAddressType_Click()
   m_HasModify = True
End Sub

Private Sub cboCountry_Click()
   m_HasModify = True
End Sub

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Caption = HeaderText
   Me.BackColor = GLB_FORM_COLOR
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitNormalLabel(lblAccountId, MapText("เลขที่บัญชี"))
   Call InitNormalLabel(lblDesc, MapText("รายละเอียด"))
   Call InitNormalLabel(lblSocLookup, MapText("แพคเกจ"))
   
   Call txtAccountId.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   Call txtDesc.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Ca As CAccount

   If Flag Then
      Call EnableForm(Me, False)
      
      If ShowMode = SHOW_EDIT Then
         Set Ca = TempCollection.Item(ID)
         txtAccountId.Text = Ca.ACCOUNT_NO
         txtDesc.Text = Ca.NOTE
'         uctlSocLookup.MyCombo.ListIndex = IDToListIndex(uctlSocLookup.MyCombo, Ca.ActAgrmnts(1).SOC_ID)
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdOK2_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cboNamePrefix_Click()
   m_HasModify = True
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblAccountId, txtAccountId) Then
      Exit Function
   End If
   
'   If Not VerifyCombo(lblSocLookup, uctlSocLookup.MyCombo) Then
'      Exit Function
'   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Ca As CAccount
   Dim Subc As CSubscriber
'   Dim Agr As CAgreement
   If ShowMode = SHOW_ADD Then
      Set Ca = New CAccount
      Ca.Flag = "A"
      Call TempCollection.Add(Ca)
      
      Set Subc = New CSubscriber
      Subc.Flag = "A"
      Subc.AddEditMode = SHOW_ADD
      Call Ca.ActSubs.Add(Subc)
   
'      Set Agr = New CAgreement
'      Agr.Flag = "A"
'      Agr.AddEditMode = SHOW_ADD
'      Call Ca.ActAgrmnts.Add(Agr)
   Else
      Set Ca = TempCollection.Item(ID)
      If Ca.Flag <> "A" Then
         Ca.Flag = "E"
      End If
      
      Set Subc = Ca.ActSubs(1)
      
'      Set Agr = Ca.ActAgrmnts(1)
   End If
   
   Ca.ACCOUNT_TYPE = -1
   Ca.ACCTTYPE_NAME = ""
   Ca.ACCOUNT_STATUS = -1
   Ca.ACCTSTS_NAME = ""
   Ca.ACCOUNT_NO = txtAccountId.Text
   Ca.NOTE = txtDesc.Text
   If Ca.MASTER_FLAG <> "Y" Then
      Ca.MASTER_FLAG = "N"
   End If
   Ca.ENABLE_FLAG = "Y"
   
   Subc.DUMMY_FLAG = "Y"
   Subc.SUBSCRIBER_NO = "DUMMY-SUBSCRIBER"
   Subc.SUBSCRIBER_STATUS = "Y"

'   Agr.SOC_CODE = uctlSocLookup.MyTextBox.Text
'   Agr.SOC_FEATURE_ID = -1
'   Agr.SOC_ID = uctlSocLookup.MyCombo.ItemData(Minus2Zero(uctlSocLookup.MyCombo.ListIndex))
'   Agr.EXCLUDE_FLAG = "N"
'   Agr.EFFECTIVE_DATE = -2
'   Agr.EXPIRE_DATE = -1
'   Agr.ISSUE_DATE = Now
   
   SaveData = True
End Function

Private Sub cmdExit_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
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
      
'      Call LoadSoc(uctlSocLookup.MyCombo, m_Socs, "N")
'      Set uctlSocLookup.MyCollection = m_Socs
      
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
   ElseIf Shift = 1 And KeyCode = 112 Then
      If glbUser.EXCEPTION_FLAG = "Y" Then
         glbUser.EXCEPTION_FLAG = "N"
      Else
         glbUser.EXCEPTION_FLAG = "Y"
      End If
   ElseIf Shift = 0 And KeyCode = 116 Then
'      Call cmdSearch_Click
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
   ElseIf Shift = 0 And KeyCode = 118 Then
'      Call cmdAdd_Click
'   ElseIf Shift = 0 And KeyCode = 117 Then
'      Call cmdDelete_Click
   ElseIf Shift = 0 And KeyCode = 113 Then
      Call cmdOK2_Click
   ElseIf Shift = 0 And KeyCode = 114 Then
'      Call cmdEdit_Click
   ElseIf Shift = 0 And KeyCode = 121 Then
'      Call cmdPrint_Click
   End If
End Sub

Private Sub Form_Load()

   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   Set m_Socs = New Collection
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   Set m_Socs = Nothing
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

Private Sub txtDistrict_Change()
   m_HasModify = True
End Sub

Private Sub txtFax_Change()
   m_HasModify = True
End Sub

Private Sub SSCheck1_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub txtAccountID_Change()
   m_HasModify = True
End Sub

Private Sub txtNickName_Change()
   m_HasModify = True
End Sub

Private Sub txtPhone_Change()
   m_HasModify = True
End Sub

Private Sub txtProvince_Change()
   m_HasModify = True
End Sub

Private Sub txtPosition_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtZipcode_Change()
   m_HasModify = True
End Sub

Private Sub uctlSocLookup_Change()
   m_HasModify = True
End Sub
