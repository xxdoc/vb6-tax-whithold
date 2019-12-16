VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSummaryReport 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   Icon            =   "frmSummaryReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   13230
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
         Left            =   5400
         TabIndex        =   7
         Top             =   885
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   12197
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000009&
            Height          =   1275
            Left            =   1440
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   12
            Top             =   2880
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
            _extentx        =   6800
            _extenty        =   767
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
            ItemData        =   "frmSummaryReport.frx":27A2
            Left            =   2070
            List            =   "frmSummaryReport.frx":27A4
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1350
            Visible         =   0   'False
            Width           =   3855
         End
         Begin prjMtpTax.uctlDate uctlGenericDate 
            Height          =   435
            Index           =   0
            Left            =   2040
            TabIndex        =   8
            Top             =   840
            Visible         =   0   'False
            Width           =   3855
            _extentx        =   6800
            _extenty        =   767
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
         Width           =   13245
         _ExtentX        =   23363
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
         Begin Threed.SSCommand cmdOK 
            Height          =   525
            Left            =   9660
            TabIndex        =   15
            Top             =   90
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   926
            _Version        =   131073
            MousePointer    =   99
            MouseIcon       =   "frmSummaryReport.frx":27A6
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExit 
            Cancel          =   -1  'True
            Height          =   525
            Left            =   11310
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
            Left            =   8010
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
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Width           =   13185
         _ExtentX        =   23257
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
                  Picture         =   "frmSummaryReport.frx":2AC0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSummaryReport.frx":339C
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
                  Picture         =   "frmSummaryReport.frx":36B8
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
         Width           =   5475
         _ExtentX        =   9657
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
Attribute VB_Name = "frmSummaryReport"
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
Private m_Companies As Collection


Private m_CyclePerMonth As Long
Private C As CReportControl

Private Sub InitTreeView()
Dim Node As Node

   trvMaster.Font.Name = GLB_FONT
   trvMaster.Font.Size = 14
  
   If MasterMode = 1 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-1", MapText("รายงานข้อมูลกลุ่มผู้ใช้งาน"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-2", MapText("รายงานข้อมูลผู้ใช้งาน"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 1-3", MapText("รายงานการล็อคอินสู่ระบบ"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 2 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-1", MapText("รายงานข้อมูลสินค้า/บริการ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 2-2", MapText("รายงานข้อมูลแพคเกจสินค้า/บริการ"), 1, 2)
      Node.Expanded = False
   
   ElseIf MasterMode = 3 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-1", MapText("รายงานข้อมูลลูกค้า"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-2", MapText("รายงานข้อมูลซัพพลายเออร์"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-3", MapText("รายงานข้อมูลพนักงาน"), 1, 2)
      Node.Expanded = False
      
       Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 3-4", MapText("รายงานข้อมูลภาษีหัก ณ ที่จ่าย"), 1, 2)
      Node.Expanded = False
      
   ElseIf MasterMode = 4 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-1", MapText("รายงานภาษีหัก ณ ที่จ่าย ภ.ง.ด 2 "), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-2", MapText("รายงานภาษีหัก ณ ที่จ่าย ภ.ง.ด 3"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-3", MapText("รายงานภาษีหัก ณ ที่จ่าย ภ.ง.ด 53"), 1, 2)
'      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-4", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 2 "), 1, 2)
      Node.Expanded = False
      
     ' ภ.ง.ด.2 รายปี
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-7", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 2ก"), 1, 2)
      Node.Expanded = False

      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 4-7", tvwChild, ROOT_TREE & " 4-7-1", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 2ก สรุปเป็นเดือน"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-5", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 3"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-8", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 3ก"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 4-8", tvwChild, ROOT_TREE & " 4-8-1", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 3ก สรุปเป็นเดือน"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-6", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 53"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 4-9", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 53ก"), 1, 2)
      Node.Expanded = False
      
      Set Node = trvMaster.Nodes.Add(ROOT_TREE & " 4-9", tvwChild, ROOT_TREE & " 4-9-1", MapText("รายงานบัญชีพิเศษ ภ.ง.ด 53ก สรุปเป็นเดือน"), 1, 2)
      Node.Expanded = False
      

      
   ElseIf MasterMode = 5 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-1", MapText("รายงานใบรับงาน/สั่งงาน (ขาย)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-2", MapText("รายงานใบส่งของ (ขาย)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-3", MapText("รายงานใบกำกับภาษี (ขาย)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-4", MapText("รายงานใบเสร็จ (ขาย)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-5", MapText("รายงานใบเพิ่มหนี้ (ขาย)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-6", MapText("รายงานใบลดหนี้ (ขาย)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-7", MapText("รายงานใบส่งของ (ซื้อ)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-8", MapText("รายงานใบกำกับภาษี (ซื้อ)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-9", MapText("รายงานใบเสร็จ (ซื้อ)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-10", MapText("รายงานใบเพิ่มหนี้ (ซื้อ)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-11", MapText("รายงานใบลดหนี้ (ซื้อ)"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-12", MapText("รายงานภาษีซื้อ"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-13", MapText("รายงานภาษีขาย"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-14", MapText("รายงานภาษีถูกหัก ณ ที่จ่าย"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-15", MapText("รายงานภาษีหัก ณ ที่จ่าย"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-16", MapText("รายงานข้อมูลลูกหนี้"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-17", MapText("รายงานข้อมูลเจ้าหนี้"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-18", MapText("รายงานดิวชำระเงินเจ้าหนี้"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 5-19", MapText("รายงานดิวรับชำระเงินลูกหนี้"), 1, 2)
      Node.Expanded = False
   ElseIf MasterMode = 6 Then
      Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True

      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-1", MapText("รายงานข้อมูลพนักงาน"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-2", MapText("รายงานข้อมูลเงินเดือนพนักงาน"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-3", MapText("สลิปเงินเดือนพนักงาน"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-4", MapText("รายงานเงินยืมส่วนบุคคล"), 1, 2)
      Node.Expanded = False
   
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 6-5", MapText("รายงานสรุปเงินสะสมพนักงาน"), 1, 2)
      Node.Expanded = False
   
   
   ElseIf MasterMode = 8 Then
   Set Node = trvMaster.Nodes.Add(, tvwFirst, ROOT_TREE, HeaderText, 2)
      Node.Expanded = True
      Node.Selected = True

'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-1", MapText("รายงานสูตรการผลิต"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-2", MapText("รายงานสูตรการผลิต/วัตถุดิบที่ใช้"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-3", MapText("รายงานใบสั่งผลิต/ประเมินราคา"), 1, 2)
'      Node.Expanded = False
'
'      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-4", MapText("รายงานใบสั่งผลิต/ประเมินราคา(ละเอียด)"), 1, 2)
'      Node.Expanded = False
       
   '   Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-5", MapText("รายงานตรวจสอบสินค้า"), 1, 2)
   '   Node.Expanded = False
   
     Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-6", MapText("รายงานใบสูตรการผลิต"), 1, 2)
      Node.Expanded = False
       
     Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-7", MapText("รายงานใบสั่งผลิต"), 1, 2)
      Node.Expanded = False
  
      Set Node = trvMaster.Nodes.Add(ROOT_TREE, tvwChild, ROOT_TREE & " 8-8", MapText("รายงานใบประเมินราคา"), 1, 2)
      Node.Expanded = False
  
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
'      m_TaxDoc.SHORT_NAME = uctlTextLookup(0).MyTextBox.Text
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

Private Sub FillReportInput(r As CReportInterface)


   Call r.AddParam(Picture1.Picture, "PICTURE")
   For Each C In m_ReportControls
      If (C.ControlType = "C") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call r.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
         End If
      End If
   
      If (C.ControlType = "T") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            Call r.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
         End If
      End If
   
      If (C.ControlType = "D") Then
         If C.Param1 <> "" Then
            Call r.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
         End If
         
         If C.Param2 <> "" Then
            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
               If C.Param2 = "TO_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -1
               ElseIf C.Param2 = "FROM_DATE" Then
                  m_Dates(C.ControlIndex).ShowDate = -2
               End If
            End If
            Call r.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
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
Dim Ti As CTaxDocItem
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String
Dim ClassName As String
   
'    If Not VerifyTextControl(lblGeneric(1), txtGeneric(1), False) Then
'      Exit Sub
'   End If

 '   Set m_TaxDoc = New CTaxDocument
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

If Key = "Root 3-1" Then
      Set Report = New CReportMain001
      ClassName = "CReportMain001"
ElseIf Key = "Root 3-2" Then
      Set Report = New CReportMain002
      ClassName = "CReportMain002"
ElseIf Key = "Root 3-3" Then
      Set Report = New CReportMain003
      ClassName = "CReportMain003"
ElseIf Key = "Root 4-1" Then
   Set Report = New CReportTax002
   ClassName = "CReportTax002"
   Call Report.AddParam(2, "TAX_TYPE")
ElseIf Key = "Root 4-2" Then
      Set Report = New CReportTax003
      ClassName = "CReportTax003"
      Call Report.AddParam(3, "TAX_TYPE")
ElseIf Key = "Root 4-3" Then
   Set Report = New CReportTax0053
   ClassName = "CReportTax0053"
   Call Report.AddParam(53, "TAX_TYPE")
ElseIf Key = "Root 4-4" Then
      Set Report = New CReportTaxSendingSP
      ClassName = "CReportTaxSendingSP"
      Call Report.AddParam(2, "TAX_TYPE")
ElseIf Key = "Root 4-5" Then
      Set Report = New CReportTaxSendingSP
      ClassName = "CReportTaxSendingSP"
      Call Report.AddParam(3, "TAX_TYPE")
ElseIf Key = "Root 4-6" Then
      Set Report = New CReportTaxSendingSP
      ClassName = "CReportTaxSendingSP"
      Call Report.AddParam(53, "TAX_TYPE")
ElseIf Key = "Root 4-7" Then                                         ' ภ.ง.ด.2 รายปี
   Set Report = New CReportTaxSendingYear
   ClassName = "CReportTaxSendingYear"
   Call Report.AddParam(2, "TAX_TYPE")
ElseIf Key = "Root 4-7-1" Then                                         ' สรุปเงินกู้ รายเดือน
   Set Report = New CReportSummaryLoan
   ClassName = "CReportSummaryLoan"
   Call Report.AddParam(2, "TAX_TYPE")
ElseIf Key = "Root 4-8" Then                                         ' ภ.ง.ด.3 รายปี
   Set Report = New CReportTaxSendingYear
   ClassName = "CReportTaxSendingYear"
   Call Report.AddParam(3, "TAX_TYPE")
ElseIf Key = "Root 4-8-1" Then                                         ' สรุปเงินกู้ รายเดือน
   Set Report = New CReportSummaryLoan
   ClassName = "CReportSummaryLoan"
   Call Report.AddParam(3, "TAX_TYPE")
ElseIf Key = "Root 4-9" Then                                         ' ภ.ง.ด.53 รายปี
   Set Report = New CReportTaxSendingYear
   ClassName = "CReportTaxSendingYear"
   Call Report.AddParam(53, "TAX_TYPE")
ElseIf Key = "Root 4-9-1" Then                                         ' สรุปเงินกู้ รายเดือน
   Set Report = New CReportSummaryLoan
   ClassName = "CReportSummaryLoan"
   Call Report.AddParam(53, "TAX_TYPE")
End If

SelectFlag = True
   
   If Key = "Root 3-4" Then
        SelectFlag = False
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
      frmReport.ClassName = ClassName
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
   Set m_Enterprise = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   'Debug.Print ColIndex & " " & NewColWidth
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
   Set m_Enterprise = New Collection
  
 
End Sub

Private Sub UnloadAllControl()
Dim i As Long
Dim j As Long

   i = m_Labels.Count
   While i > 0
      Call Unload(m_Labels(i))
      Call m_Labels.Remove(i)
      i = i - 1
   Wend
   
   i = m_Texts.Count
   While i > 0
      Call Unload(m_Texts(i))
      Call m_Texts.Remove(i)
      i = i - 1
   Wend

   i = m_Dates.Count
   While i > 0
      Call Unload(m_Dates(i))
      Call m_Dates.Remove(i)
      i = i - 1
   Wend

   i = m_Combos.Count
   While i > 0
      Call Unload(m_Combos(i))
      Call m_Combos.Remove(i)
      i = i - 1
   Wend
   
   i = m_TextLookups.Count
   While i > 0
      Call Unload(m_TextLookups(i))
      Call m_TextLookups.Remove(i)
      i = i - 1
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
            If C.ComboLoadID = 1 Then
               Call InitUserGroupOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
         
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitUserOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

         If trvMaster.SelectedItem.Key = ROOT_TREE & " 1-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadUserGroup(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitLoginOrderby(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-1" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_CUSTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitCustomerOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-2" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_SUPTYPE)
            ElseIf C.ComboLoadID = 2 Then
               Call InitSupplierOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If
      
         If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-3" Then
            If C.ComboLoadID = 1 Then
               Call LoadMaster(m_Combos(C.ControlIndex), , MASTER_EMPPOSITION)
            ElseIf C.ComboLoadID = 2 Then
               Call InitEmployeeOrderBy(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 3 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            End If
         End If

      If trvMaster.SelectedItem.Key = ROOT_TREE & " 3-4" Then
         If C.ComboLoadID = 0 Then
             Call InitTaxType(m_Combos(C.ControlIndex))
            ElseIf (C.ComboLoadID = 1) Then
               Call InitOrderType(m_Combos(C.ControlIndex))
            ElseIf C.ComboLoadID = 2 Then
               Call InitOrderType(m_Combos(C.ControlIndex))
             End If
      End If
      
         If MasterMode = 4 Then
                If trvMaster.SelectedItem.Key = ROOT_TREE & " 4-7-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 4-8-1" Or trvMaster.SelectedItem.Key = ROOT_TREE & " 4-9-1" Then
                    If C.ComboLoadID = 1 Then
                      ' Call LoadEnterpriseShortName(m_Combos(C.ControlIndex), m_Enterprise)
                    ElseIf C.ComboLoadID = 2 Then
                       Call InitThaiMonth(m_Combos(C.ControlIndex))
                       'cboGeneric(C.ControlIndex).ListIndex = Month(Now)
                    ElseIf C.ComboLoadID = 3 Then
                       Call InitThaiMonth(m_Combos(C.ControlIndex))
                       'cboGeneric(C.ControlIndex).ListIndex = Month(Now)
'                    ElseIf C.ComboLoadID = 4 Then
'                        Call InitOrderType(m_Combos(C.ControlIndex))
                    End If
               Else
                       If C.ComboLoadID = 0 Then
                        Call InitTaxType(m_Combos(C.ControlIndex))
                       ElseIf (C.ComboLoadID = 1) Then
                          Call InitOrderBy(m_Combos(C.ControlIndex))
                       ElseIf C.ComboLoadID = 2 Then
                          Call InitOrderType(m_Combos(C.ControlIndex))
                       ElseIf C.ComboLoadID = 3 Then
'                          Call LoadEnterpriseShortName(m_Combos(C.ControlIndex))
                        End If
                End If
             End If
         End If
   Next C
   Call EnableForm(Me, True)

End Sub
Private Sub LoadLookUpData()

 Set m_TaxDocs = New Collection
 Set m_Companies = New Collection
   Me.Refresh
   DoEvents
   Call EnableForm(Me, False)
   
   For Each C In m_ReportControls
'      If (C.ControlType = "LU") Then
'        If MasterMode = 4 Then
'            If C.uctlLookUpID = 0 Then
''               Call LoadEnterprise(uctlTextLookup(C.ControlIndex).MyCombo, m_TaxDocs, MASTER_REVENUETYPE)     'MASTER_REVENUETYPE
''               Set uctlTextLookup(C.ControlIndex).MyCollection = m_TaxDocs
''            ElseIf C.uctlLookUpID = 1 Then
''               Call InitSupplierOrderBy(m_TextLookups(C.ControlIndex))
'            End If
'         End If
'
'      End If
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
'   ElseIf ControlType = "LU" Then
'         Load uctlGuestNameLookup(LkupIdx)
'         Call m_TextLookups.Add(uctlGuestNameLookup(LkupIdx))
'         C.ControlIndex = LkupIdx
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
      Call InitReport1_1
   ElseIf Node.Key = ROOT_TREE & " 1-2" Then
      Call InitReport1_2
   ElseIf Node.Key = ROOT_TREE & " 1-3" Then
      Call InitReport1_3
   ElseIf Node.Key = ROOT_TREE & " 3-1" Then
      Call InitReport3_1
   ElseIf Node.Key = ROOT_TREE & " 3-2" Then
      Call InitReport3_2
   ElseIf Node.Key = ROOT_TREE & " 3-3" Then
      Call InitReport3_3
  ElseIf Node.Key = ROOT_TREE & " 3-4" Then
      Call InitReport3_4
  ElseIf Node.Key = ROOT_TREE & " 4-7-1" Then
      Call InitReport4_7_1
  ElseIf Node.Key = ROOT_TREE & " 4-8-1" Then
      Call InitReport4_7_1
   ElseIf Node.Key = ROOT_TREE & " 4-9-1" Then
      Call InitReport4_7_1
  ElseIf MasterMode = 4 Then
      Call InitReport4_1
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
   
'   '1 =============================
'   Call LoadControl("LU", uctlTextLookup(0).Width, True, "", 0, "")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ชื่อบริษัท (หน่วยงาน)"))
   '2 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 0, "BRANCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("  แบบที่ใช้ยื่นภาษี"))
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", 1, "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
 '4 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", 1, "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))
   '5 =============================
    Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '64 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
   Call LoadLookUpData
End Sub
Private Sub InitReport4_7_1()
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long



   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "BRANCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หน่วยงาน"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้ถูกหัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้ถูกหัก"))
   
    '2 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 2, "FROM_MONTH_ID", "FROM_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากเดือน"))

   ' =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "FROM_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากปี"))
   
   '3 =============================
   Call LoadControl("C", cboGeneric(0).Width \ 2, False, "", 3, "TO_MONTH_ID", "TO_MONTH_NAME")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงเดือน"))

   ' =============================
   Call LoadControl("T", txtGeneric(0).Width / 2, False, "", , "TO_YEAR_NO")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงปี"))

   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 4, "ORDER_TYPE")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
End Sub
Private Sub InitReport4_1()
'Dim report As CReportInterface
Dim Top As Long
Dim Left As Long
Dim LabelWidth As Long
Dim Offset As Long

'Set report = New CReportInterface

   Top = lblGeneric(0).Top
   Left = lblGeneric(0).Left
   LabelWidth = lblGeneric(0).Width
   Offset = 100
   
'1 =============================
   Call LoadControl("T", txtGeneric(0).Width, True, "", , "BRANCH")
'   Call LoadControl("LU", uctlGuestNameLookup(0).Width, True, "", , "BRANCH")
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 3, "MASTER_BRANCH")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("หน่วยงาน"))
   
'   Call LoadMaster(uctlBranchLookup.MyCombo, m_Branches, MASTER_BRANCH)
'  Set uctlBranchLookup.MyCollection = m_Branches
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "FROM_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากรหัสผู้ถูกหัก"))
   
   Call LoadControl("T", txtGeneric(0).Width / 2, True, "", , "TO_SUPPLIER_CODE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงรหัสผู้ถูกหัก"))
   
   '2 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "FROM_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("จากวันที่"))
   
   '3 =============================
   Call LoadControl("D", uctlGenericDate(0).Width, True, "", , "TO_DATE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("ถึงวันที่"))

   '4 =============================
'   Call LoadControl("C", cboGeneric(0).Width, True, "", 1, "ORDER_BY")
'   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงตาม"))

   '5 =============================
   Call LoadControl("C", cboGeneric(0).Width, True, "", 2, "ORDER_TYPE")
   Call LoadControl("L", lblGeneric(0).Width, True, MapText("เรียงจาก"))
   
   Call ShowControl
   Call LoadComboData
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
      
      
'      m_TaxDoc.SHORT_NAME = uctlTextLookup(1).MyTextBox.Text
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
      
'      m_TaxDocSP.SHORT_NAME = uctlTextLookup(1).MyTextBox.Text
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

