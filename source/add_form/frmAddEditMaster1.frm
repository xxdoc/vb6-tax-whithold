VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAddEditMaster1 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddEditMaster1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   615
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1085
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame Frame1 
      Height          =   2115
      Left            =   -30
      TabIndex        =   7
      Top             =   420
      Width           =   8595
      _ExtentX        =   15161
      _ExtentY        =   3731
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboGroup 
         Height          =   510
         Left            =   5040
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   450
         Visible         =   0   'False
         Width           =   2955
      End
      Begin prjMtpTax.uctlTextBox txtCode 
         Height          =   435
         Left            =   2250
         TabIndex        =   0
         Top             =   450
         Width           =   1845
         _extentx        =   4683
         _extenty        =   767
      End
      Begin prjMtpTax.uctlTextBox txtName 
         Height          =   435
         Left            =   2280
         TabIndex        =   3
         Top             =   900
         Width           =   5745
         _extentx        =   4683
         _extenty        =   767
      End
      Begin prjMtpTax.uctlTextBox txtExportKey 
         Height          =   435
         Left            =   2250
         TabIndex        =   4
         Top             =   1340
         Width           =   1845
         _extentx        =   4683
         _extenty        =   767
      End
      Begin VB.Label lblExportKey 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   60
         TabIndex        =   14
         Top             =   1400
         Width           =   1965
      End
      Begin Threed.SSCheck chkFlag 
         Height          =   435
         Left            =   4110
         TabIndex        =   1
         Top             =   420
         Visible         =   0   'False
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   767
         _Version        =   131073
         Caption         =   "SSCheck1"
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Height          =   465
         Left            =   60
         TabIndex        =   13
         Top             =   930
         Width           =   2055
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   150
         TabIndex        =   8
         Top             =   480
         Width           =   1965
      End
   End
   Begin Threed.SSPanel pnlFooter 
      Height          =   705
      Left            =   0
      TabIndex        =   9
      Top             =   2520
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   1244
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   2670
         TabIndex        =   5
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdCancel 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   4298
         TabIndex        =   6
         Top             =   90
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   615
         Index           =   0
         Left            =   11130
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   615
         Left            =   13230
         TabIndex        =   10
         Top             =   60
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditMaster1"
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
Public MasterKey As String
Public MasterArea As MASTER_TYPE

Private m_MasterRef As CMasterRef

Public MasterMode As Long
Private Sub cboGroup_Click()
   m_HasModify = True
End Sub

Private Sub chkFlag_Click(Value As Integer)
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
   Frame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
      
   lblExportKey.Visible = False
   txtExportKey.Visible = False
      
   Call InitNormalLabel(lblCode, "")
   Call InitNormalLabel(lblName, "")
   
   If MasterKey = ROOT_TREE & " 3-1" Then
      Call InitNormalLabel(lblCode, MapText("���ʻ����"))
      Call InitNormalLabel(lblName, MapText("�����"))
   ElseIf MasterKey = ROOT_TREE & " 3-2" Then
      Call InitNormalLabel(lblCode, MapText("�����дѺ�١���"))
      Call InitNormalLabel(lblName, MapText("�дѺ�١���"))
   ElseIf MasterKey = ROOT_TREE & " 3-3" Then
      Call InitNormalLabel(lblCode, MapText("���ʻ������١���"))
      Call InitNormalLabel(lblName, MapText("�������١���"))
   ElseIf MasterKey = ROOT_TREE & " 3-4" Then
      Call InitNormalLabel(lblCode, MapText("�����дѺ�Ѿ���������"))
      Call InitNormalLabel(lblName, MapText("�дѺ�Ѿ���������"))
   ElseIf MasterKey = ROOT_TREE & " 3-5" Then
      Call InitNormalLabel(lblCode, MapText("���ʻ������Ѿ���������"))
      Call InitNormalLabel(lblName, MapText("�������Ѿ���������"))
   ElseIf MasterKey = ROOT_TREE & " 3-6" Then
      Call InitNormalLabel(lblCode, MapText("����ʶҹЫѾ���������"))
      Call InitNormalLabel(lblName, MapText("ʶҹЫѾ���������"))
   ElseIf MasterKey = ROOT_TREE & " 3-7" Then
      Call InitNormalLabel(lblCode, MapText("���ʵ��˹�"))
      Call InitNormalLabel(lblName, MapText("���˹�"))
   ElseIf MasterKey = ROOT_TREE & " 4-1" Then
      Call InitNormalLabel(lblCode, MapText("���ʻ������Թ��"))
      Call InitNormalLabel(lblName, MapText("�������Թ��"))
      Call InitNormalLabel(lblExportKey, MapText("���� Export Online"))
      lblExportKey.Visible = True
      txtExportKey.Visible = True
   ElseIf MasterKey = ROOT_TREE & " 4-2" Then
      Call InitNormalLabel(lblCode, MapText("�����ѵ������"))
      Call InitNormalLabel(lblName, MapText("�ѵ������"))
   ElseIf MasterKey = ROOT_TREE & " 4-3" Then
      Call InitNormalLabel(lblCode, MapText("�������͹�"))
      Call InitNormalLabel(lblName, MapText("���͹�"))
   ElseIf MasterKey = ROOT_TREE & " 4-4" Then
      Call InitNormalLabel(lblCode, MapText("�����Ң�"))
      Call InitNormalLabel(lblName, MapText("�Ң�"))
   ElseIf MasterKey = ROOT_TREE & " 4-5" Then
      Call InitNormalLabel(lblCode, MapText("���ʺѭ��"))
      Call InitNormalLabel(lblName, MapText("���ͺѭ��"))
   ElseIf MasterKey = ROOT_TREE & " 4-6" Then
      Call InitNormalLabel(lblCode, MapText("��������"))
      Call InitNormalLabel(lblName, MapText("����"))
   ElseIf MasterKey = ROOT_TREE & " 4-7" Then
      Call InitNormalLabel(lblCode, MapText("���ʵ鹷ҧ"))
      Call InitNormalLabel(lblName, MapText("�鹷ҧ"))
   End If

   Call txtCode.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)
   Call txtName.SetTextLenType(TEXT_STRING, glbSetting.NAME_LEN)

   Call InitMainButton(cmdSave, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdCancel, MapText("¡��ԡ (ESC)"))
      
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   Frame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdCancel.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   If Flag Then
      Call EnableForm(Me, False)
      m_MasterRef.KEY_ID = ID
      Call m_MasterRef.QueryData(m_Rs, itemcount)
      If itemcount > 0 Then
         Call m_MasterRef.PopulateFromRS(1, m_Rs)
         txtCode.Text = m_MasterRef.KEY_CODE
         txtName.Text = m_MasterRef.KEY_NAME
         txtExportKey.Text = m_MasterRef.EXPORT_KEY
      End If
      Call EnableForm(Me, True)
   End If
   
   IsOK = True
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdSave_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Function SaveData() As Boolean
On Error GoTo ErrorHandler
Dim IsOK As Boolean

   If MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf MasterMode = 2 Then
           If Not VerifyAccessRight("MASTER_PERSON_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_CRM_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf MasterMode = 7 Then
      If Not VerifyAccessRight("MASTER_LEDGER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("MASTER_PRODUCTION_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Exit Function
   End If
   If Not VerifyTextControl(lblName, txtName, Not txtName.Visible) Then
      Exit Function
   End If
   
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
      
   If Not VerifyTextControl(lblCode, txtCode, False) Then
      Call EnableForm(Me, True)
      Exit Function
   End If
      
   m_MasterRef.AddEditMode = ShowMode
   m_MasterRef.KEY_NAME = txtName.Text
   m_MasterRef.KEY_CODE = txtCode.Text
   m_MasterRef.MASTER_AREA = MasterArea
   m_MasterRef.EXPORT_KEY = txtExportKey.Text
   
   Call glbDaily.AddEditMasterRef(m_MasterRef, IsOK, glbErrorLog)
   
   Call EnableForm(Me, True)
   
   If Not IsOK Then
      Call EnableForm(Me, True)
      glbErrorLog.ShowUserError
      Exit Function
   End If
   
   Call EnableForm(Me, True)
   SaveData = True
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
   
   Call EnableForm(Me, True)
   SaveData = False
End Function

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      If ShowMode = SHOW_EDIT Then
         Call QueryData(True)
      ElseIf ShowMode = SHOW_ADD Then
         ID = 0
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
      Call cmdSave_Click
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
   
   m_HasActivate = False
   m_HasModify = False
   Set m_Rs = New ADODB.Recordset
   
   Set m_MasterRef = New CMasterRef
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing

   Set m_MasterRef = Nothing
End Sub
Private Sub txtCode_Change()
   m_HasModify = True
End Sub
Private Sub txtExportKey_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub
