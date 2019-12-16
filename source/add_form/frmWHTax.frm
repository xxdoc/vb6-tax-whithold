VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmWHTax 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmWHTax.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11955
      _ExtentX        =   21087
      _ExtentY        =   15055
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboBranch 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1410
         Width           =   2985
      End
      Begin prjMtpTax.uctlDate uctlPutDate 
         Height          =   435
         Left            =   5970
         TabIndex        =   1
         Top             =   960
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   767
      End
      Begin VB.ComboBox cboOrderType 
         Height          =   315
         Left            =   5970
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1830
         Width           =   2625
      End
      Begin VB.ComboBox cboOrderBy 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1830
         Width           =   2985
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   30
         TabIndex        =   15
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5205
         Left            =   180
         TabIndex        =   8
         Top             =   2520
         Width           =   11505
         _ExtentX        =   20294
         _ExtentY        =   9181
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
         Column(1)       =   "frmWHTax.frx":27A2
         Column(2)       =   "frmWHTax.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmWHTax.frx":290E
         FormatStyle(2)  =   "frmWHTax.frx":2A6A
         FormatStyle(3)  =   "frmWHTax.frx":2B1A
         FormatStyle(4)  =   "frmWHTax.frx":2BCE
         FormatStyle(5)  =   "frmWHTax.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmWHTax.frx":2D5E
      End
      Begin prjMtpTax.uctlTextBox txtCompanyNo 
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   2985
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin prjMtpTax.uctlTextBox txtSupplierCode 
         Height          =   435
         Left            =   5970
         TabIndex        =   3
         Top             =   1380
         Width           =   2655
         _ExtentX        =   13309
         _ExtentY        =   767
      End
      Begin VB.Label lblSupplierCode 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   21
         Top             =   1470
         Width           =   1215
      End
      Begin VB.Label lblPutDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   315
         Left            =   4650
         TabIndex        =   20
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label lblCompanyNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   19
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   4680
         TabIndex        =   18
         Top             =   1890
         Width           =   1215
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   17
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "Label1"
         Height          =   435
         Left            =   30
         TabIndex        =   16
         Top             =   1890
         Width           =   1455
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   10110
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWHTax.frx":2F36
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdClear 
         Height          =   525
         Left            =   10110
         TabIndex        =   7
         Top             =   1650
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   11
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWHTax.frx":3250
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdAdd 
         Height          =   525
         Left            =   150
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWHTax.frx":356A
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdEdit 
         Height          =   525
         Left            =   1770
         TabIndex        =   10
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10095
         TabIndex        =   13
         Top             =   7830
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8445
         TabIndex        =   12
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmWHTax.frx":3884
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmWHTax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_HasActivate As Boolean
Private m_TaxDoc As CTaxDocument
Private m_TempTaxDoc As CTaxDocument
Private m_Rs As ADODB.Recordset
Private m_Coll As Collection
Private m_TableName As String

Public HeaderText As String
Public OKClick As Boolean
Public TaxType As Long

Private Sub cmdPasswd_Click()

End Sub

Private Sub cmdAdd_Click()
Dim itemcount As Long
Dim OKClick As Boolean

   If Not VerifyAccessRight("TAX_WHITHOLD_ADD") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   frmAddEditTaxDocument.TaxType = TaxType
   frmAddEditTaxDocument.HeaderText = MapText("เพิ่มใบนำส่งภาษี")
   frmAddEditTaxDocument.ShowMode = SHOW_ADD
   Load frmAddEditTaxDocument
   frmAddEditTaxDocument.Show 1
   
   OKClick = frmAddEditTaxDocument.OKClick
   
   Unload frmAddEditTaxDocument
   Set frmAddEditTaxDocument = Nothing
   
   If OKClick Then
      Call QueryData(True)
   End If
End Sub

Private Sub cmdClear_Click()
   txtSupplierCode.Text = ""
   txtCompanyNo.Text = ""
   uctlPutDate.ShowDate = -1
   cboBranch.ListIndex = -1
   cboOrderBy.ListIndex = -1
   cboOrderType.ListIndex = -1
End Sub

Private Sub cmdDelete_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long

   If Not VerifyAccessRight("TAX_WITHOLD_DELETE") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   ID = GridEX1.Value(1)
   
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   If Not ConfirmDelete(GridEX1.Value(2)) Then
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Exit Sub
   End If

   Call EnableForm(Me, False)
   If Not glbDaily.DeleteTaxDocument(ID, IsOK, True, glbErrorLog) Then
      m_TaxDoc.TAX_DOCUMENT_ID = -1
      glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
      Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   Call QueryData(True)
   
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)
   Call EnableForm(Me, True)
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim ENT_ID As Long
Dim OKClick As Boolean

   If Not VerifyAccessRight("TAX_WITHOLD_QUERY") Then
      Call EnableForm(Me, True)
      Exit Sub
   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(1))
   ENT_ID = Val(GridEX1.Value(5))
   Call glbDatabaseMngr.LockTable(m_TableName, ID, IsCanLock, glbErrorLog)
               
   frmAddEditTaxDocument.ID = ID
   frmAddEditTaxDocument.ENT_ID = ENT_ID
   frmAddEditTaxDocument.TaxType = TaxType
   frmAddEditTaxDocument.Document_Year = Trim(GridEX1.Value(2))
  frmAddEditTaxDocument.COMPANY_ID = Val(GridEX1.Value(6))
   frmAddEditTaxDocument.HeaderText = MapText("แก้ไขใบนำส่งภาษี ภ.ง.ด ")
   frmAddEditTaxDocument.ShowMode = SHOW_EDIT
   Load frmAddEditTaxDocument
   frmAddEditTaxDocument.Show 1
   
   OKClick = frmAddEditTaxDocument.OKClick
   
   Unload frmAddEditTaxDocument
   Set frmAddEditTaxDocument = Nothing
               
   If OKClick Then
      Call QueryData(True)
   End If
   Call glbDatabaseMngr.UnLockTable(m_TableName, ID, IsCanLock, glbErrorLog)

End Sub

Private Sub cmdOK_Click()
   OKClick = True
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      
      Call LoadMaster(cboBranch, , MASTER_BRANCH)
      
      Call InitTaxDocumentOrderBy(cboOrderBy)
      Call InitOrderType(cboOrderType)
      
      Call QueryData(True)
   End If
End Sub

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
Dim Temp As Long

   If Flag Then
      Call EnableForm(Me, False)
      
      If Not VerifyAccessRight("TAX_WITHOLD_QUERY") Then
         Call EnableForm(Me, True)
         Exit Sub
      End If
      
      m_TaxDoc.TAX_DOCUMENT_ID = -1
      m_TaxDoc.TAX_TYPE = TaxType
      m_TaxDoc.FROM_DATE = uctlPutDate.ShowDate
      m_TaxDoc.TO_DATE = uctlPutDate.ShowDate
      m_TaxDoc.SHORT_NAME = PatchWildCard(txtCompanyNo.Text)
      m_TaxDoc.BRANCH_ID = cboBranch.ItemData(Minus2Zero(cboBranch.ListIndex))
      m_TaxDoc.SUPPLIER_CODE = txtSupplierCode.Text
      m_TaxDoc.OrderBy = cboOrderBy.ItemData(Minus2Zero(cboOrderBy.ListIndex))
      m_TaxDoc.OrderType = cboOrderType.ItemData(Minus2Zero(cboOrderType.ListIndex))
      If (m_TaxDoc.OrderType <= 0) Then
        m_TaxDoc.OrderType = 2
      End If
      If TaxType = 2 Or TaxType = 3 Or TaxType = 53 Then
         If Not glbDaily.QueryTaxDocument(m_TaxDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      ElseIf TaxType = 21 Then
         If Not glbDaily.QueryTax_A_Document(m_TaxDoc, m_Coll, itemcount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
   End If
   
   If Not IsOK Then
      glbErrorLog.ShowUserError
      Call EnableForm(Me, True)
      Exit Sub
   End If
   
   If TaxType = 2 Or TaxType = 3 Or TaxType = 53 Then
      GridEX1.itemcount = itemcount
      GridEX1.Rebind
   ElseIf TaxType = 21 Then
      GridEX1.itemcount = CountItem(m_Coll)
      GridEX1.Rebind
   End If
   
   
   Call EnableForm(Me, True)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If Shift = 1 And KeyCode = DUMMY_KEY Then
      glbErrorLog.LocalErrorMsg = Me.Name
      glbErrorLog.ShowUserError
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 116 Then
      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
      Call cmdClear_Click
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

Private Sub InitGrid()
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

If TaxType = 21 Then
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2355
   Col.Caption = MapText("ปีที่นำส่ง")
Else
   Set Col = GridEX1.Columns.Add '2
   Col.Width = 2355
   Col.Caption = MapText("วันที่นำส่ง")
End If
      
   Set Col = GridEX1.Columns.Add '3
   Col.Width = 9105
   Col.Caption = MapText("บริษัท")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 9105
   Col.Caption = MapText("")

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 9105
   Col.Caption = MapText("")
   
   Set Col = GridEX1.Columns.Add '6
   Col.Width = 0
   Col.Caption = MapText("Company ID")
   
   GridEX1.itemcount = 0
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitGrid
   
   Call InitNormalLabel(lblBranch, MapText("สาขา"))
   Call InitNormalLabel(lblSupplierCode, MapText("รหัสซัพ ฯ"))
   Call InitNormalLabel(lblCompanyNo, MapText("รหัสบริษัท"))
   Call InitNormalLabel(lblPutDate, MapText("วันที่นำส่ง"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงจาก"))
   
   Call InitCombo(cboBranch)
   Call InitCombo(cboOrderBy)
   Call InitCombo(cboOrderType)
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdClear.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา (F5)"))
   Call InitMainButton(cmdClear, MapText("เคลียร์ (F4)"))
   
   If TaxType = 21 Then
      cmdAdd.Enabled = False
      cmdEdit.Enabled = False
      cmdDelete.Enabled = False
   End If
End Sub

Private Sub cmdExit_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub Form_Load()
   m_TableName = "USER_GROUP"
   
   Set m_TaxDoc = New CTaxDocument
   Set m_TempTaxDoc = New CTaxDocument
   Set m_Rs = New ADODB.Recordset
'   Set m_Coll = New Collection

   Call InitFormLayout
   Call EnableForm(Me, True)
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   ''Debug.Print ColIndex & " " & NewColWidth
End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
'   RowBuffer.RowStyle = RowBuffer.Value(5)
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If TaxType = 21 Then
         If m_Coll Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CTaxDocument
      If m_Coll.Count <= 0 Then
         Exit Sub
      End If
      Set CR = GetItem(m_Coll, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If
'      Set Addr = CR.Addresses

      Values(1) = CR.TAX_DOCUMENT_ID
      Values(2) = CR.Document_Year
      Values(3) = CR.LONG_NAME
      Values(6) = CR.COMPANY_ID
   
   Else
   
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
      Call m_TempTaxDoc.PopulateFromRS(1, m_Rs)
      
      Values(1) = m_TempTaxDoc.TAX_DOCUMENT_ID
      Values(2) = DateToStringExtEx2(m_TempTaxDoc.DOCUMENT_DATE)
      Values(3) = m_TempTaxDoc.LONG_NAME
End If

   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

