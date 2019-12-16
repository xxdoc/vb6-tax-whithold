VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmMasterMain 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmMasterMain.frx":0000
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
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   15690
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin Threed.SSPanel pnlFooter 
         Height          =   705
         Left            =   30
         TabIndex        =   2
         Top             =   7800
         Width           =   11850
         _ExtentX        =   20902
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   8445
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":27A2
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   10095
            TabIndex        =   7
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdEdit 
            Height          =   525
            Left            =   1770
            TabIndex        =   6
            Top             =   120
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   926
            _Version        =   131073
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdAdd 
            Height          =   525
            Left            =   150
            TabIndex        =   5
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2ABC
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdDelete 
            Height          =   525
            Left            =   3420
            TabIndex        =   4
            Top             =   120
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmMasterMain.frx":2DD6
            ButtonStyle     =   3
         End
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   855
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11925
         _ExtentX        =   21034
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
                  Picture         =   "frmMasterMain.frx":30F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmMasterMain.frx":39CC
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
                  Picture         =   "frmMasterMain.frx":3CE8
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin MSComctlLib.TreeView trvMaster 
         Height          =   6945
         Left            =   0
         TabIndex        =   3
         Top             =   870
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   12250
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
      Begin GridEX20.GridEX GridEX1 
         Height          =   6915
         Left            =   4500
         TabIndex        =   9
         Top             =   900
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   12197
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
         Column(1)       =   "frmMasterMain.frx":4002
         Column(2)       =   "frmMasterMain.frx":40CA
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmMasterMain.frx":416E
         FormatStyle(2)  =   "frmMasterMain.frx":42CA
         FormatStyle(3)  =   "frmMasterMain.frx":437A
         FormatStyle(4)  =   "frmMasterMain.frx":442E
         FormatStyle(5)  =   "frmMasterMain.frx":4506
         ImageCount      =   0
         PrinterProperties=   "frmMasterMain.frx":45BE
      End
   End
End
Attribute VB_Name = "frmMasterMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Rs As ADODB.Recordset
Private m_HasActivate As Boolean
Private m_TableName As String
Private m_MasterRef As CMasterRef
Private m_TempArea As MASTER_TYPE

Public HeaderText As String
Public MasterMode As Long

Private Sub cmdAdd_Click()
Dim OKClick As Boolean

   If trvMaster.SelectedItem Is Nothing Then
      Exit Sub
   End If
   
   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PERSON_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_CRM_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 7 Then
      If Not VerifyAccessRight("MASTER_LEDGER_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("MASTER_PRODUCTION_ADD") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If trvMaster.SelectedItem.Key = ROOT_TREE Then
      glbErrorLog.LocalErrorMsg = ""
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   frmAddEditMaster1.MasterArea = m_TempArea
   frmAddEditMaster1.MasterMode = MasterMode
   frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
   frmAddEditMaster1.ShowMode = SHOW_ADD
   frmAddEditMaster1.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
   Load frmAddEditMaster1
   frmAddEditMaster1.Show 1
   
   OKClick = frmAddEditMaster1.OKClick
   
   Unload frmAddEditMaster1
   Set frmAddEditMaster1 = Nothing
      
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub


Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT
   trvMaster.Font.Size = 14
   
   If MasterMode = 3 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("ประเทศ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-2", MapText("ระดับลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-3", MapText("ประเภทลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-4", MapText("ระดับซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-5", MapText("ประเภทซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-6", MapText("สถานะซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-7", MapText("ตำแหน่ง"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("ประเภทเงินได้"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-2", MapText("อัตราภาษี"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-3", MapText("เงื่อนไข"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-4", MapText("สาขา"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5", MapText("บัญชีหัก ณ ที่จ่าย"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 5 Then
   ElseIf MasterMode = 6 Then
   ElseIf MasterMode = 7 Then
   ElseIf MasterMode = 8 Then
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long

   If Flag Then
      If MasterMode = 1 Then
         If Not VerifyAccessRight("MASTER_INVENTORY_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf MasterMode = 2 Then
         If Not VerifyAccessRight("MASTER_PERSON_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf MasterMode = 3 Then
         If Not VerifyAccessRight("MASTER_MAIN_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf MasterMode = 4 Then
         If Not VerifyAccessRight("MASTER_GOLD_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf MasterMode = 6 Then
         If Not VerifyAccessRight("MASTER_PACKAGE_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf MasterMode = 7 Then
         If Not VerifyAccessRight("MASTER_LEDGER_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf MasterMode = 8 Then
         If Not VerifyAccessRight("MASTER_PRODUCTION_QUERY") Then
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
      
      Call EnableForm(Me, False)
   End If
   
   Call EnableForm(Me, True)
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrorHandler
Dim Status As Boolean
Dim IsOK As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
      
   If MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PERSON_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_CRM_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 7 Then
      If Not VerifyAccessRight("MASTER_LEDGER_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("MASTER_PRODUCTION_DELETE") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   Status = glbDaily.DeleteMasterRef(TempID, IsOK, glbErrorLog)
   
   If Status Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   Else
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Exit Sub
   End If
   
   Exit Sub
   
ErrorHandler:
End Sub

Private Sub cmdEdit_Click()
Dim OKClick As Boolean
Dim TempID As Long

   If trvMaster.SelectedItem.Key = "" Then
      Exit Sub
   End If
   
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   TempID = GridEX1.Value(1)
   
   If MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PERSON_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_CRM_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 7 Then
      If Not VerifyAccessRight("MASTER_LEDGER_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("MASTER_PRODUCTION_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
   
   frmAddEditMaster1.MasterArea = m_TempArea
   frmAddEditMaster1.MasterMode = MasterMode
   frmAddEditMaster1.ID = TempID
   frmAddEditMaster1.MasterKey = trvMaster.SelectedItem.Key
   frmAddEditMaster1.ShowMode = SHOW_EDIT
   frmAddEditMaster1.HeaderText = MapText("ข้อมูลหลัก") & " " & trvMaster.SelectedItem.Text
   Load frmAddEditMaster1
   frmAddEditMaster1.Show 1
   
   OKClick = frmAddEditMaster1.OKClick
   
   Unload frmAddEditMaster1
   Set frmAddEditMaster1 = Nothing
      
   If OKClick Then
      Call trvMaster_NodeClick(trvMaster.SelectedItem)
   End If
End Sub

Private Sub Form_Activate()
Dim itemcount As Long

   If Not m_HasActivate Then
      Me.Refresh
      DoEvents
      
      Call QueryData(True)
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
      Call cmdAdd_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 117 Then
      Call cmdDelete_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 113 Then
'      Call cmdOK_Click
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
   
   Set m_MasterRef = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
 '  'Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub InitGrid0()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_1()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเทศ")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเทศ")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_2()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสระดับลูกค้า")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ระดับลูกค้า")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_3()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภทลูกค้า")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทลูกค้า")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_4()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2220
   Col.Caption = MapText("รหัสระดับซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ระดับซัพพลายเออร์")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_5()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2220
   Col.Caption = MapText("รหัสประเภทซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ประเภทซัพพลายเออร์")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_6()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2220
   Col.Caption = MapText("รหัสสถานะซัพพลายเออร์")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("สถานะซัพพลายเออร์")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid3_7()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2220
   Col.Caption = MapText("รหัสตำแหน่ง")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5115
   Col.Caption = MapText("ตำแหน่ง")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_1()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสประเภทเงินได้")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ประเภทเงินได้")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_2()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสอัตราภาษี")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("อัตราภาษี")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_3()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสเงื่อนไข")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("เงื่อนไข")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_4()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสสาขา")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("สาขา")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_5()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสบัญชี")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ชื่อบัญชี")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_6()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสภาษา")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ภาษา")

   GridEX1.itemcount = 0
End Sub

Private Sub InitGrid4_7()
Dim Col As JSColumn
Dim fmsTemp As JSFormatStyle

   GridEX1.Columns.Clear
   GridEX1.BackColor = GLB_GRID_COLOR
   GridEX1.BackColorHeader = GLB_GRIDHD_COLOR
   
   GridEX1.FormatStyles.Clear
   Set fmsTemp = GridEX1.FormatStyles.Add("N")
   fmsTemp.ForeColor = GLB_ALERT_COLOR
   
   Set Col = GridEX1.Columns.Add '1
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 1620
   Col.Caption = MapText("รหัสต้นทาง")
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 5715
   Col.Caption = MapText("ต้นทาง")

   GridEX1.itemcount = 0
End Sub

Private Sub InitFormLayout()
   Me.KeyPreview = True
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   pnlHeader.Caption = HeaderText
   
   Me.BackColor = GLB_FORM_COLOR
   SSFrame1.BackColor = GLB_FORM_COLOR
   pnlHeader.BackColor = GLB_HEAD_COLOR
   pnlFooter.BackColor = GLB_HEAD_COLOR
   Call InitHeaderFooter(pnlHeader, pnlFooter)

   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdExit, MapText("ออก (ESC)"))
   
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   pnlFooter.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitTreeView
   Call InitGrid0
   
'   lsvMaster.Font.NAME = GLB_FONT
'   lsvMaster.Font.Size = 14
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Call InitFormLayout
   
   m_HasActivate = False
   m_TableName = "SYSTEM_PARAM"
   Set m_Rs = New ADODB.Recordset
   
   Set m_MasterRef = New CMasterRef
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_Rs Is Nothing Then
      Exit Sub
   End If

   If m_Rs.State <> adStateOpen Then
      Exit Sub
   End If

   If m_Rs.EOF Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If
   
   Call m_Rs.Move(RowIndex - 1, adBookmarkFirst)
   Call m_MasterRef.PopulateFromRS(1, m_Rs)
   
   Values(1) = m_MasterRef.KEY_ID
   Values(2) = m_MasterRef.KEY_CODE
   Values(3) = m_MasterRef.KEY_NAME
   
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

'Private Sub LoadListView(Rs As ADODB.Recordset, FieldName As String, IDName As String)
'Dim Lst As ListItem
'
'   While Not Rs.EOF
'      Set Lst = lsvMaster.ListItems.Add(, , NVLS(Rs(FieldName), ""), 1, 1)
'      Lst.Tag = NVLI(Rs(IDName), 0)
'      Rs.MoveNext
'   Wend
'End Sub

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
   
   If MasterMode = 1 Then
      If Not VerifyAccessRight("MASTER_INVENTORY_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 2 Then
      If Not VerifyAccessRight("MASTER_PERSON_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 3 Then
      If Not VerifyAccessRight("MASTER_MAIN_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 4 Then
      If Not VerifyAccessRight("MASTER_CRM_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 6 Then
      If Not VerifyAccessRight("MASTER_PACKAGE_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 7 Then
      If Not VerifyAccessRight("MASTER_LEDGER_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   ElseIf MasterMode = 8 Then
      If Not VerifyAccessRight("MASTER_PRODUCTION_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
      
   If Node.Key = ROOT_TREE & " 3-1" Then
      m_TempArea = MASTER_COUNTRY
      Call InitGrid3_1
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      m_TempArea = MASTER_CUSGRADE
      Call InitGrid3_2
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      m_TempArea = MASTER_CUSTYPE
      Call InitGrid3_3
   ElseIf Node.Key = ROOT_TREE & " 3-4" Then
      m_TempArea = MASTER_SUPGRADE
      Call InitGrid3_4
   ElseIf Node.Key = ROOT_TREE & " 3-5" Then
      m_TempArea = MASTER_SUPTYPE
      Call InitGrid3_5
   ElseIf Node.Key = ROOT_TREE & " 3-6" Then
      m_TempArea = MASTER_SUPSTATUS
      Call InitGrid3_6
   ElseIf Node.Key = ROOT_TREE & " 3-7" Then
      m_TempArea = MASTER_EMPPOSITION
      Call InitGrid3_7
   ElseIf Node.Key = ROOT_TREE & " 4-1" Then
      m_TempArea = MASTER_REVENUETYPE
      Call InitGrid4_1
   ElseIf Node.Key = ROOT_TREE & " 4-2" Then
      m_TempArea = MASTER_TAXRATE
      Call InitGrid4_2
   ElseIf Node.Key = ROOT_TREE & " 4-3" Then
      m_TempArea = MASTER_CONDITION
      Call InitGrid4_3
   ElseIf Node.Key = ROOT_TREE & " 4-4" Then
      m_TempArea = MASTER_BRANCH
      Call InitGrid4_4
   ElseIf Node.Key = ROOT_TREE & " 4-5" Then
      m_TempArea = MASTER_ACCOUNT
      Call InitGrid4_5
   ElseIf Node.Key = ROOT_TREE & " 4-6" Then
      m_TempArea = MASTER_LANGUAGE
      Call InitGrid4_6
   ElseIf Node.Key = ROOT_TREE & " 4-7" Then
      m_TempArea = MASTER_SOURCE
      Call InitGrid4_7
   Else
      Call InitGrid0
   End If
   
   Dim Mr As CMasterRef
   Set Mr = New CMasterRef
   Mr.KEY_ID = -1
   Mr.MASTER_AREA = m_TempArea
   Status = Mr.QueryData(m_Rs, itemcount)
   GridEX1.itemcount = itemcount
   GridEX1.Rebind
   Set Mr = Nothing
   
End Sub
