VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLedgerReportMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   Icon            =   "frmLedgerReportMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin prjLedgerReport.uctlTextBox txtGeneric 
      Height          =   435
      Index           =   0
      Left            =   7680
      TabIndex        =   13
      Top             =   2670
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   767
   End
   Begin VB.ComboBox cboGeneric 
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
      Index           =   0
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   3855
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
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
         Picture         =   "frmLedgerReportMain.frx":08CA
         Caption         =   "SSCommand1"
         ButtonStyle     =   3
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   7755
      Left            =   0
      TabIndex        =   3
      Top             =   780
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   13679
      _Version        =   131073
      BackStyle       =   1
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.TreeView trvMain 
         Height          =   7035
         Left            =   210
         TabIndex        =   4
         Top             =   0
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   12409
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
               Picture         =   "frmLedgerReportMain.frx":195A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":2234
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":2B0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":33E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":3542
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":3E1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":46F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":4A10
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":52EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":5BC4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":649E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmLedgerReportMain.frx":7178
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
         Left            =   1920
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
      Left            =   5820
      TabIndex        =   10
      Top             =   810
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   1296
      _Version        =   131073
      BackStyle       =   1
   End
   Begin prjLedgerReport.uctlDate uctlGenericDate 
      Height          =   405
      Index           =   0
      Left            =   7680
      TabIndex        =   11
      Top             =   1860
      Visible         =   0   'False
      Width           =   3825
      _ExtentX        =   5689
      _ExtentY        =   291
   End
   Begin Threed.SSCommand cmdConfig 
      Height          =   525
      Left            =   8370
      TabIndex        =   17
      Top             =   7890
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   926
      _Version        =   131073
      Enabled         =   0   'False
      ButtonStyle     =   3
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   525
      Left            =   10020
      TabIndex        =   16
      Top             =   7890
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   926
      _Version        =   131073
      MousePointer    =   99
      MouseIcon       =   "frmLedgerReportMain.frx":7E52
      ButtonStyle     =   3
   End
   Begin VB.Label lblGeneric 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   15
      Top             =   1860
      Visible         =   0   'False
      Width           =   1725
   End
   Begin Threed.SSCheck chkGeneric 
      Height          =   465
      Index           =   0
      Left            =   7680
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   820
      _Version        =   131073
      Caption         =   "SSCheck1"
   End
End
Attribute VB_Name = "frmLedgerReportMain"
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

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String

   Key = trvMain.SelectedItem.Key
   Name = trvMain.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
   If Not (trvMain.SelectedItem Is Nothing) Then
      Call Report.AddParam(trvMain.SelectedItem.Text, "REPORT_TEXT")
   End If
   
   If Key = ROOT_TREE & " 3-1-1" Then
'      If Not VerifyAccessRight("AP_REPORT_PRINT1", trvMaster.SelectedItem.Text) Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If

      Set Report = New CReportAP001
      SelectFlag = True
   End If

   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If
      
      Call FillReportInput(Report)
      Call Report.AddParam(Name, "REPORT_NAME")
      Call Report.AddParam(Key, "REPORT_KEY")
      
      Set frmReport.ReportObject = Report
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
   End If
End Sub

Private Sub Form_Activate()
Dim OKClick As Boolean
Dim DBPath As String

   If m_HasActivate Then
      Exit Sub
   End If
   m_HasActivate = True

   Call EnableForm(Me, False)
   Call PatchDB

   frmDatabaseSelect.ShowMode = SHOW_EDIT
   Load frmDatabaseSelect
   frmDatabaseSelect.Show 1

   OKClick = frmDatabaseSelect.OKClick
   DBPath = frmDatabaseSelect.DBPath
   Unload frmDatabaseSelect
   Set frmDatabaseSelect = Nothing
  
  If OKClick Then
      Call glbDatabaseMngr.DisConnectDatabase
      Call glbDatabaseMngr.ConnectDatabase(DBPath, "", "", glbErrorLog)
   End If
   
   Call EnableForm(Me, True)
   
   If Not OKClick Then
      m_MustAsk = False
      Unload Me
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
   Call InitNormalLabel(lblUserName, MapText("ผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblUserGroup, MapText("กลุ่มผู้ใช้ : "), RGB(0, 0, 255))
   Call InitNormalLabel(lblVersion, MapText("เวอร์ชัน : ") & glbParameterObj.Version & " (Interbase) ", RGB(0, 0, 255))
   Call InitNormalLabel(lblDateTime, "", RGB(0, 0, 255))
   lblDateTime.BackStyle = 1
   lblDateTime.BackColor = RGB(255, 255, 255)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPasswd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Me.Caption = MapText("ระบบรายงานบัญชี Express")
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
      
   Call InitMainButton(cmdExit, MapText("ออก"))
   Call InitMainButton(cmdPasswd, MapText("โปรแกรม"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   
   Call InitMainTreeview
End Sub

Private Sub InitMainTreeview()
Dim Node As Node
Dim NewNodeID As String

   trvMain.Nodes.Clear
   trvMain.Font.Name = GLB_FONT
   trvMain.Font.Size = 14
   trvMain.Font.Bold = False
      
   Set Node = trvMain.Nodes.Add(, tvwFirst, ROOT_TREE, MapText("ระบบรายงานบัญชี Express"), 8)
   Node.Expanded = True
   Node.Selected = True
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-0", MapText("1. ข้อมูลสินทรัพย์"), 3, 3)
   Node.Expanded = False
   
   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-0", MapText("2. ข้อมูลรายจ่าย"), 3, 3)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("3. ข้อมูลหนี้สิน"), 3, 3)
   Node.Expanded = False
      
      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 3-1", tvwChild, ROOT_TREE & " 3-1-1", MapText("3.1. รายงานข้อมูลเจ้าหนี้"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 3-1", tvwChild, ROOT_TREE & " 3-1-2", MapText("3.2. รายงานหนี้ค้างชำระรายเจ้าหนี้"), 12, 11)
      Node.Expanded = False

      Set Node = trvMain.Nodes.Add(ROOT_TREE & " 3-1", tvwChild, ROOT_TREE & " 3-1-3", MapText("3.3. รายงานการชำระเงินเจ้าหนี้ตามช่วงเวลา"), 12, 11)
      Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-0", MapText("4. ข้อมูลทุน"), 3, 3)
   Node.Expanded = False

   Set Node = trvMain.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-0", MapText("5. ข้อมูลรายรับ"), 3, 3)
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

Private Sub FillReportInput(R As CReportInterface)
Dim C As CReportControl

   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "CB") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ListIndex, C.Param2)
         End If
      End If
      
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "CH") Then
         If C.Param1 <> "" Then
            Call R.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(Check2Flag(m_CheckBoxes(C.ControlIndex).Value), C.Param2)
         End If
      End If
      
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
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
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
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

Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "")
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long
Dim C As CReportControl
Dim ChkIdx As Long

   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   ChkIdx = m_CheckBoxes.Count + 1

   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.Add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
   ElseIf ControlType = "C" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.Add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
   ElseIf ControlType = "CB" Then
      Load cboGeneric(CboIdx)
      Call m_Combos.Add(cboGeneric(CboIdx))
      C.ControlIndex = CboIdx
   ElseIf ControlType = "T" Then
      Load txtGeneric(TxtIdx)
      Call m_Texts.Add(txtGeneric(TxtIdx))
      C.ControlIndex = TxtIdx
   ElseIf ControlType = "D" Then
      Load uctlGenericDate(DateIdx)
      Call m_Dates.Add(uctlGenericDate(DateIdx))
      C.ControlIndex = DateIdx

      If DateIdx = 1 Then
         uctlGenericDate(DateIdx).ShowDate = m_FromDate
      ElseIf DateIdx = 2 Then
         uctlGenericDate(DateIdx).ShowDate = m_ToDate
      End If
   ElseIf ControlType = "CH" Then
      Load chkGeneric(ChkIdx)
      Call m_CheckBoxes.Add(chkGeneric(ChkIdx))
      C.ControlIndex = ChkIdx
   End If

   C.AllowNull = NullAllow
   C.ControlType = ControlType
   C.Width = Width
   C.TextMsg = TextMsg
   C.Param1 = Param2
   C.Param2 = Param1
   C.ComboLoadID = ComboLoadID
   Call m_ReportControls.Add(C)
   Set C = Nothing
End Sub

Private Sub UnloadAllControl()
Dim I As Long
Dim j As Long

   I = m_Labels.Count
   While I > 0
      Call Unload(m_Labels(I))
      Call m_Labels.Remove(I)
      I = I - 1
   Wend
   
   I = m_Texts.Count
   While I > 0
      Call Unload(m_Texts(I))
      Call m_Texts.Remove(I)
      I = I - 1
   Wend

   I = m_Dates.Count
   While I > 0
      Call Unload(m_Dates(I))
      Call m_Dates.Remove(I)
      I = I - 1
   Wend

   I = m_Combos.Count
   While I > 0
      Call Unload(m_Combos(I))
      Call m_Combos.Remove(I)
      I = I - 1
   Wend
   
   I = m_TextLookups.Count
   While I > 0
      Call Unload(m_TextLookups(I))
      Call m_TextLookups.Remove(I)
      I = I - 1
   Wend
   
   I = m_CheckBoxes.Count
   While I > 0
      Call Unload(m_CheckBoxes(I))
      Call m_CheckBoxes.Remove(I)
      I = I - 1
   Wend
   
   Set m_ReportControls = Nothing
   Set m_ReportControls = New Collection
End Sub

Private Sub ShowControl()
Dim PrevTop As Long
Dim PrevLeft As Long
Dim PrevWidth As Long
Dim CurTop As Long
Dim CurLeft As Long
Dim CurWidth As Long
Dim C As CReportControl

   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width

   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "CB") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            m_Combos(C.ControlIndex).Left = PrevLeft
            m_Combos(C.ControlIndex).Top = PrevTop
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True

            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth

            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
            PrevLeft = m_Combos(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CB" Then
            m_Combos(C.ControlIndex).Left = PrevLeft
            m_Combos(C.ControlIndex).Top = PrevTop
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True

            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth

            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).Height
            PrevLeft = m_Combos(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "D" Then
            m_Dates(C.ControlIndex).Left = PrevLeft
            m_Dates(C.ControlIndex).Top = PrevTop
            m_Dates(C.ControlIndex).Width = C.Width
            m_Dates(C.ControlIndex).Visible = True

            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth

            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).Height
            PrevLeft = m_Dates(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "T" Then
            m_Texts(C.ControlIndex).Left = PrevLeft
            m_Texts(C.ControlIndex).Left = PrevLeft
            m_Texts(C.ControlIndex).Top = PrevTop
            m_Texts(C.ControlIndex).Width = C.Width
            Call m_Texts(C.ControlIndex).SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
            m_Texts(C.ControlIndex).Visible = True

            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth

            PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).Height
            PrevLeft = m_Texts(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "LU" Then
            m_TextLookups(C.ControlIndex).Left = PrevLeft
            m_TextLookups(C.ControlIndex).Top = PrevTop
            m_TextLookups(C.ControlIndex).Width = C.Width
            m_TextLookups(C.ControlIndex).Visible = True

            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth

            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).Height
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
         ElseIf C.ControlType = "CH" Then
            m_CheckBoxes(C.ControlIndex).Left = PrevLeft
            m_CheckBoxes(C.ControlIndex).Top = PrevTop
            m_CheckBoxes(C.ControlIndex).Width = C.Width
            m_CheckBoxes(C.ControlIndex).Visible = True

            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth

            PrevTop = m_CheckBoxes(C.ControlIndex).Top + m_CheckBoxes(C.ControlIndex).Height
            PrevLeft = m_CheckBoxes(C.ControlIndex).Left
            PrevWidth = C.Width
            Call InitCheckBox(m_CheckBoxes(C.ControlIndex), C.TextMsg)
         End If
      Else 'Label
            m_Labels(C.ControlIndex).Left = lblGeneric(0).Left
            m_Labels(C.ControlIndex).Top = CurTop
            m_Labels(C.ControlIndex).Width = C.Width
            Call InitNormalLabel(m_Labels(C.ControlIndex), C.TextMsg)
            m_Labels(C.ControlIndex).Visible = True
      End If
   Next C
End Sub

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

Private Sub trvMain_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim ItemCount As Long
Dim QueryFlag As Boolean

   If LastKey = Node.Key Then
      Exit Sub
   End If
   
   pnlHeader.Caption = Node.Text
   
   Status = True
   QueryFlag = False
   
   Call UnloadAllControl
   
   If Node.Key = ROOT_TREE & " 3-1-1" Then
      Call InitReport3_1
   End If
End Sub

Private Sub InitReport3_1()
Dim C As CReportControl
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสผู้จำหน่าย"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("CB", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทผู้จำหน่าย"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub LoadComboData()
Dim C As CReportControl

'   Me.Refresh
'   DoEvents
'   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "CB") Then
         If trvMain.SelectedItem.Key = ROOT_TREE & " 3-1-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadSupplierType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      End If 'C.ControlType = "C"
   Next C
'   Call EnableForm(Me, True)
End Sub

