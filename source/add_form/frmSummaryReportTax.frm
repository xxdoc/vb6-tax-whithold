VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummaryReportTax 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmSummaryReportTax.frx":0000
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
      Height          =   10995
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   19394
      _Version        =   131073
      Begin Threed.SSFrame SSFrame2 
         Height          =   6915
         Left            =   4080
         TabIndex        =   7
         Top             =   880
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   12197
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin prjMtpTax.uctlTextLookup uctlGeneric 
            Height          =   315
            Index           =   0
            Left            =   2070
            TabIndex        =   16
            Top             =   2160
            Visible         =   0   'False
            Width           =   5355
            _ExtentX        =   9446
            _ExtentY        =   556
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   0
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   12
            Top             =   2820
            Visible         =   0   'False
            Width           =   1635
         End
         Begin prjMtpTax.uctlTextBox txtGeneric 
            Height          =   435
            Index           =   0
            Left            =   2070
            TabIndex        =   9
            Top             =   1710
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
            Left            =   2070
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1350
            Visible         =   0   'False
            Width           =   3855
         End
         Begin prjMtpTax.uctlDate uctlGenericDate 
            Height          =   435
            Index           =   0
            Left            =   2070
            TabIndex        =   8
            Top             =   930
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   767
         End
         Begin VB.Label lblGeneric 
            Alignment       =   1  'Right Justify
            Caption         =   "1"
            Height          =   375
            Index           =   0
            Left            =   -270
            TabIndex        =   10
            Top             =   1050
            Visible         =   0   'False
            Width           =   2205
         End
      End
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   4
         Top             =   7800
         Width           =   11835
         _ExtentX        =   20876
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8460
            TabIndex        =   15
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmSummaryReportTax.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10110
            TabIndex        =   14
            Top             =   90
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "JasmineUPC"
               Size            =   24
               Charset         =   222
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   0
            TabIndex        =   5
            Top             =   30
            Visible         =   0   'False
            Width           =   2145
         End
         Begin Threed.SSCommand cmdConfig 
            Height          =   525
            Left            =   6810
            TabIndex        =   13
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   615
            Left            =   2160
            TabIndex        =   0
            Top             =   60
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   615
            Left            =   2610
            TabIndex        =   1
            Top             =   60
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1085
            _Version        =   131073
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   11865
         _ExtentX        =   20929
         _ExtentY        =   1508
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   0
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   2
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReportTax.frx":2ABC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReportTax.frx":3398
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   2850
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   32
            ImageHeight     =   32
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReportTax.frx":36B4
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   6915
         Left            =   0
         TabIndex        =   6
         Top             =   870
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   12197
         _Version        =   393217
         Indentation     =   882
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "JasmineUPC"
            Size            =   15.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSummaryReportTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String

Public HeaderText As String
Public MasterMode As Long

Public m_TaxDoc As CTaxDocument
Public m_TaxDocSP As CTaxDocSP
Public m_TaxDocItem As CTaxDocItem

Private m_ReportControls As Collection
Private m_Texts As Collection
Private m_Dates As Collection
Private m_Labels As Collection
Private m_Combos As Collection
Private m_TextLookups As Collection
Private m_TaxDocs  As Collection

Private m_CyclePerMonth As Long
Private C As CReportControl

Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT
   trvMaster.Font.Size = 14
   
   If MasterMode = 0 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("รายงาน ภาษีหัก ณ ที่จ่าย ภ.ง.ด  2"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("รายงาน ภาษีหัก ณ ที่จ่าย ภ.ง.ด  3"), 1, 2)
      Node.Expanded = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("รายงาน ภาษีหัก ณ ที่จ่าย ภ.ง.ด  53"), 1, 2)
      Node.Expanded = True
      
     Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-4", MapText("รายงาน บัญชีพิเศษ ภ.ง.ด  2"), 1, 2)
      Node.Expanded = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-5", MapText("รายงาน บัญชีพิเศษ ภ.ง.ด  3"), 1, 2)
      Node.Expanded = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-6", MapText("รายงาน บัญชีพิเศษ ภ.ง.ด  53"), 1, 2)
      Node.Expanded = True
   ElseIf MasterMode = 1 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   End If
End Sub

'Private Sub QueryData2(Flag As Boolean)
'Dim IsOK As Boolean
'Dim itemcount As Long
'Dim Temp As Long
'
'   If Flag Then
'      Call EnableForm(Me, False)
'
'      If Not VerifyAccessRight("TAX_WITHOLD_QUERY") Then
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'
'      Set m_TaxDoc = New CTaxDocument
'
'      m_TaxDoc.SHORT_NAME = uctlGeneric(0).MyTextBox.Text
'      m_TaxDoc.TAX_TYPE = cboGeneric(0).ListIndex
''      m_TaxDoc.FOR_MONTH = cboGeneric(1).ListIndex
''      m_TaxDoc.OrderType = cboGeneric(2).ListIndex
'
'      If Not glbDaily.QueryTaxDocument(m_TaxDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
'         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
'         Call EnableForm(Me, True)
'         Exit Sub
'      End If
'   End If
'
'   If Not IsOK Then
'      glbErrorLog.ShowUserError
'      Call EnableForm(Me, True)
'      Exit Sub
'   End If
'
'
'
'
'   Call EnableForm(Me, True)
'End Sub

Private Sub FillReportInput(R As CReportInterface)


   Call R.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
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
            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
         End If
      End If
   
   Next C
End Sub

Private Function VerifyReportInput() As Boolean


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

Private Sub cmdConfig_Click()
Dim ReportKey As String
Dim Rc As CReportConfig
Dim iCount As Long

   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
      
   ReportKey = trvMaster.SelectedItem.Key
   
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
   frmReportConfig.HeaderText = trvMaster.SelectedItem.Text
   Load frmReportConfig
   frmReportConfig.Show 1
   
   Unload frmReportConfig
   Set frmReportConfig = Nothing
   
   Set Rc = Nothing
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String



   Key = trvMaster.SelectedItem.Key
   Name = trvMaster.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
   
  
If Not VerifyAccessRight("MAIN_REPORT_PRINT3", trvMaster.SelectedItem.Text) Then
   Call EnableForm(Me, True)
   Exit Sub
End If

If Key = "Root 1-1" Then
      Set Report = New CReportTax002
  ElseIf Key = "Root 1-2" Then
      Set Report = New CReportTax003
  ElseIf Key = "Root 1-3" Then
      Set Report = New CReportTax0053
   Else
      Set Report = New CReportTaxSending
End If
     

Call Report.AddParam(-1, "TAX_DOCUMENT_ID")
Call Report.AddParam(uctlGeneric(1).MyTextBox.Text, "E_NAME")

      SelectFlag = True
   

   If SelectFlag Then
      If glbParameterObj.Temp = 0 Then
         glbParameterObj.UsedCount = glbParameterObj.UsedCount + 1
         glbParameterObj.Temp = 1
      End If

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
Dim itemcount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
            
     
      m_HasActivate = True
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
   Set m_ReportControls = Nothing
   Set m_Texts = Nothing
   Set m_Dates = Nothing
   Set m_Labels = Nothing
   Set m_Combos = Nothing
   Set m_TextLookups = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid1()
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   SSFrame2.BackColor = GLB_FORM_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame2.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("พิมพ์ (F10)"))
   Call InitMainButton(cmdConfig, MapText("ปรับค่า"))
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdConfig.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitTreeView
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()

   Call InitFormLayout
   
   m_HasActivate = False
   Set m_Rs = New ADODB.Recordset
   

   Set m_Texts = New Collection
   Set m_Dates = New Collection
   Set m_Labels = New Collection
   Set m_Combos = New Collection
   Set m_TextLookups = New Collection
  
 
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


   PrevTop = uctlGenericDate(0).Top
   PrevLeft = uctlGenericDate(0).Left
   PrevWidth = uctlGenericDate(0).Width
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Or (C.ControlType = "D") Or (C.ControlType = "T") Or (C.ControlType = "LU") Or (C.ControlType = "CH") Then
         If C.ControlType = "C" Then
            m_Combos(C.ControlIndex).Left = PrevLeft
            m_Combos(C.ControlIndex).Top = PrevTop
            m_Combos(C.ControlIndex).Width = C.Width
            Call InitCombo(m_Combos(C.ControlIndex))
            m_Combos(C.ControlIndex).Visible = True
            
            CurTop = PrevTop
            CurLeft = PrevLeft
            CurWidth = PrevWidth
            
            PrevTop = m_Combos(C.ControlIndex).Top + m_Combos(C.ControlIndex).HEIGHT
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
         
            PrevTop = m_Dates(C.ControlIndex).Top + m_Dates(C.ControlIndex).HEIGHT
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
            
            PrevTop = m_Texts(C.ControlIndex).Top + m_Texts(C.ControlIndex).HEIGHT
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
         
            PrevTop = m_TextLookups(C.ControlIndex).Top + m_TextLookups(C.ControlIndex).HEIGHT
            PrevLeft = m_TextLookups(C.ControlIndex).Left
            PrevWidth = C.Width
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

Private Sub LoadComboData()


   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-1" Then
            If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 2) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
           If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 2) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
            If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 2) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-4" Then
            If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 2) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-5" Then
           If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 2) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-6" Then
            If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 2) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
         End If

     If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-4" Then
                  If C.ComboLoadID = 0 Then
                      Call InitTaxType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 1 Then
                        Call InitReportOR(m_Combos(C.ControlIndex))
                     ElseIf (C.ComboLoadID = 3) Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                     ElseIf C.ComboLoadID = 3 Then
                        Call InitOrderType(m_Combos(C.ControlIndex))
                      End If
             End If
         
      End If
   Next C
   Call EnableForm(Me, True)

End Sub
Private Sub LoadLookUpData()

 Set m_TaxDocs = New Collection
   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
      If (C.ControlType = "LU") Then
       ' If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-4" Then
            If C.uctlLookUpID = 0 Then
               Call LoadEnterprise(uctlGeneric(C.ControlIndex).MyCombo, m_TaxDocs, MASTER_REVENUETYPE)     'MASTER_REVENUETYPE
               Set uctlGeneric(C.ControlIndex).MyCollection = m_TaxDocs
            ElseIf C.uctlLookUpID = 1 Then
               Call InitSupplierOrderBy(m_TextLookups(C.ControlIndex))
            End If
         'End If
         
      End If
   Next C
   Call EnableForm(Me, True)
End Sub

Private Sub LoadControl(ControlType As String, Width As Long, NullAllow As Boolean, TextMsg As String, Optional ComboLoadID As Long = -1, Optional Param1 As String = "", Optional Param2 As String = "")
Dim CboIdx As Long
Dim TxtIdx As Long
Dim DateIdx As Long
Dim LblIdx As Long
Dim LkupIdx As Long


   CboIdx = m_Combos.Count + 1
   TxtIdx = m_Texts.Count + 1
   DateIdx = m_Dates.Count + 1
   LblIdx = m_Labels.Count + 1
   LkupIdx = m_TextLookups.Count + 1
   
   Set C = New CReportControl
   If ControlType = "L" Then
      Load lblGeneric(LblIdx)
      Call m_Labels.Add(lblGeneric(LblIdx))
      C.ControlIndex = LblIdx
   ElseIf ControlType = "C" Then
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
   ElseIf ControlType = "LU" Then
         Load uctlGeneric(LkupIdx)
         Call m_TextLookups.Add(uctlGeneric(LkupIdx))
         C.ControlIndex = LkupIdx
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

Private Sub InitReport1_1()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_2()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "GROUP_ID", "GROUP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อกลุ่ม"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport1_3()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "USER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อผู้ใช้"))
   
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).WIDTH, True, "", 1, "GROUP_ID", "GROUP_NAME")
'   Call LoadControl("L", lblGeneric(0).WIDTH, True, GetTextMessage("TEXT-KEY71"))

   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))

   '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '6 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub trvMaster_NodeClick(ByVal Node As MSComctlLib.Node)
Static LastKey As String
Dim Status As Boolean
Dim itemcount As Long
Dim QueryFlag As Boolean

   If LastKey = Node.Key Then
      Exit Sub
   End If
   
   Status = True
   QueryFlag = False
   
   Call UnloadAllControl
   
   If Node.Key = ROOT_TREE & " 1-1" Then
      Call InitReport3_4
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitReport3_4
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitReport3_4
   ElseIf Node.Key = ROOT_TREE & " 1-4" Then
      Call InitReport3_4
   ElseIf Node.Key = ROOT_TREE & " 1-5" Then
      Call InitReport3_4
   ElseIf Node.Key = ROOT_TREE & " 1-6" Then
      Call InitReport3_4
'  ElseIf Node.Key = ROOT_TREE & " 3-4" Then
'      Call InitReport3_4
   End If
End Sub

Private Sub InitReport3_1()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "CUSTOMER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสลูกค้า"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "CUSTOMER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "CUSTOMER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทลูกค้า"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_2()

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
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสซัพพลายเออร์"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "SUPPLIER_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อซัพพลายเออร์"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "SUPPLIER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ประเภทซัพ ฯ"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_3()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงาน"))
   
   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อพนักงาน"))

   '2 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "EMP_LASTNAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("นามสกุลพนักงาน"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "EMP_POSITION")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ตำแหน่ง"))

   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub

Private Sub InitReport3_4()

Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
   '1 =============================
'   Call LoadControl("T", txtGeneric(0).Width / 1.5, True, "", , "EMP_CODE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("รหัสพนักงาน"))
   
   '1 =============================
   Call LoadControl("LU", uctlGeneric(0).Width, True, "", 0, "BRANCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อบริษัท (หน่วยงาน)"))
   '2 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 0, "")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("  แบบที่ใช้ยื่นภาษี"))
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
 '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
    '3 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   
   Call ShowControl
   Call LoadComboData
   Call LoadLookUpData
End Sub
Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long

Set m_TaxDoc = New CTaxDocument

   If Flag Then
      Call EnableForm(Me, False)
      
      If Not VerifyAccessRight("TAX_WITHOLD_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      
      m_TaxDoc.SHORT_NAME = uctlGeneric(1).MyTextBox.Text
      m_TaxDoc.TAX_TYPE = cboGeneric(1).ListIndex      'cboGeneric(1).Text
   '  m_TaxDoc.FOR_MONTH = cboGeneric(2).ListIndex        'cboGeneric(2).Text
   '  m_TaxDoc.OrderType = cboGeneric(3).ListIndex
     m_TaxDoc.OrderType = cboGeneric(3).ItemData(Minus2Zero(cboGeneric(3).ListIndex))

      If Not glbDaily.QueryTaxDocument(m_TaxDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
  
   
   Call EnableForm(Me, True)
End Sub

Private Sub QueryData2(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long
Dim ReportType As Long

'Set m_TaxDoc = New CTaxDocument
Set m_TaxDocSP = New CTaxDocSP
Set m_TaxDocItem = New CTaxDocItem
   If Flag Then
      Call EnableForm(Me, False)
      
      If Not VerifyAccessRight("TAX_WITHOLD_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      m_TaxDocSP.SHORT_NAME = uctlGeneric(1).MyTextBox.Text
      m_TaxDocSP.REPORT_TYPE = cboGeneric(1).ListIndex
      m_TaxDocItem.FROM_DATE = uctlGenericDate(1).ShowDate
      m_TaxDocItem.TO_DATE = uctlGenericDate(2).ShowDate
      m_TaxDocItem.OrderBy = cboGeneric(2).ItemData(Minus2Zero(cboGeneric(2).ListIndex))
      m_TaxDocItem.OrderType = cboGeneric(3).ItemData(Minus2Zero(cboGeneric(3).ListIndex))

      

      If Not glbDaily.QueryTaxDocument2(m_TaxDocSP, m_Rs, itemcount, IsOK, glbErrorLog) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
End Sub

