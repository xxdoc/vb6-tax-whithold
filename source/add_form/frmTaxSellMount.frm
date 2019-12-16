VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmTaxSellMount 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTaxSellMount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSFrame SSFrame1 
      Height          =   2415
      Left            =   -90
      TabIndex        =   3
      Top             =   0
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   4260
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjMtpTax.uctlTextLookup uctlEnterprise 
         Height          =   405
         Left            =   1830
         TabIndex        =   5
         Top             =   300
         Width           =   5505
         _extentx        =   9710
         _extenty        =   714
      End
      Begin VB.ComboBox cboMonth 
         Height          =   510
         ItemData        =   "frmTaxSellMount.frx":08CA
         Left            =   1830
         List            =   "frmTaxSellMount.frx":08CC
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   870
         Width           =   2265
      End
      Begin prjMtpTax.uctlTextBox txtThYear 
         Height          =   375
         Left            =   4110
         TabIndex        =   1
         Top             =   870
         Width           =   1095
         _extentx        =   1931
         _extenty        =   661
      End
      Begin VB.Label lblMonth 
         Alignment       =   1  'Right Justify
         Caption         =   "Label2"
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   930
         Width           =   1605
      End
      Begin VB.Label lblEnterprise 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1635
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmTaxSellMount.frx":08CE
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3720
         TabIndex        =   4
         Top             =   1710
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmTaxSellMount"
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
Private m_Companies As Collection

Public HeaderText As String
Public ID As Long
Public OKClick As Boolean
Public TempCollection As Collection

Private Sub cmdCancel_Click()
   If Not ConfirmExit(m_HasModify) Then
      Exit Sub
   End If
   
   OKClick = False
   Unload Me
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   
   Me.BackColor = GLB_FORM_COLOR
   
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Me.Caption = "เลือกรายงานบัญชีพิเศษ  ประจำ เดือน / พ.ศ. "
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   Call InitNormalLabel(lblEnterprise, MapText("บริษัท ( หน่วยงาน )"))
   Call InitNormalLabel(lblMonth, MapText("รายงาน เดือน / ปี"))
   
   Call InitCombo(cboMonth)
   
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
         Dim Name As cName
         Dim CstContact As CSupplierContact
         Set CstContact = TempCollection.Item(ID)
         Set Name = CstContact.Name
         
      '  txtName.Text = Name.LONG_NAME
         txtLastName.Text = Name.LAST_NAME
         txtEmail.Text = Name.EMAIL
         txtPosition.Text = CstContact.CONTACT_POSITION
      End If
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim RealIndex As Long

   If Not VerifyTextControl(lblName, txtName, False) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Dim Name As cName
   Dim CstContact As CSupplierContact
   If ShowMode = SHOW_ADD Then
      Set Name = New cName
      Set CstContact = New CSupplierContact
      Set CstContact.Name = Name
   Else
      Set CstContact = TempCollection.Item(ID)
      Set Name = CstContact.Name
   End If
   
   Name.LONG_NAME = txtName.Text
   Name.LAST_NAME = txtLastName.Text
   Name.EMAIL = txtEmail.Text
   CstContact.CONTACT_POSITION = txtPosition.Text
   
   If ShowMode = SHOW_ADD Then
      Name.Flag = "A"
      CstContact.Flag = "A"
      Call TempCollection.Add(CstContact)
   Else
      If Name.Flag <> "A" Then
         Name.Flag = "E"
      End If
      If CstContact.Flag <> "A" Then
         CstContact.Flag = "E"
      End If
   End If
   
   Set Name = Nothing
   SaveData = True
End Function

Private Sub cmdExit_Click()
ShowMode = SHOW_VIEW
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim SelectFlag As Boolean

SelectFlag = True
ShowMode = frmAddEditTaxDocument.ShowMode
If cboMonth.Text = "" Or txtThYear.Text = "" Then
MsgBox "กรุณากรอกข้อมูลเดือนปีให้ถูกต้อง", vbOKOnly, "Imformation"
    If cboMonth.Text = "" And txtThYear.Text = "" Then cboMonth.SetFocus
    If cboMonth.Text <> "" And txtThYear.Text = "" Then txtThYear.SetFocus
    If cboMonth.Text = "" And txtThYear.Text <> "" Then cboMonth.SetFocus
Else
   OKClick = True
   Unload Me
End If
End Sub

Private Sub Form_Activate()
Dim strMonth As String

 If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      Call EnableForm(Me, False)
     
      Call LoadEnterprise(uctlEnterprise.MyCombo, m_Companies)
      Set uctlEnterprise.MyCollection = m_Companies
      
      Call EnableForm(Me, True)
      m_HasModify = False
   End If
      
    Call InitLoadThMonth(cboMonth)
    txtThYear.Text = Year(Date) + 543
  
  
      
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
Set m_Companies = New Collection
   OKClick = False
   Call InitFormLayout
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub txtEnterprise_Change()
m_HasModify = True
End Sub

Private Sub uctlEnterprise_Change()
m_HasModify = True
End Sub
