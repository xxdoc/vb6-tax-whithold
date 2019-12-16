VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditLender 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14160
   Icon            =   "frmAddEditLender.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   14160
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin prjMtpTax.uctlTextLookup uctlBranchLookup 
         Height          =   495
         Left            =   1680
         TabIndex        =   14
         Top             =   1440
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   873
      End
      Begin prjMtpTax.uctlTextBox txtYear 
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
         _ExtentX        =   2990
         _ExtentY        =   873
      End
      Begin prjMtpTax.uctlTextLookup uctlCompanyLookup 
         Height          =   405
         Left            =   1680
         TabIndex        =   0
         Top             =   960
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   714
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   5295
         Left            =   150
         TabIndex        =   2
         Top             =   2430
         Width           =   6500
         _ExtentX        =   11456
         _ExtentY        =   9340
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
         Column(1)       =   "frmAddEditLender.frx":27A2
         Column(2)       =   "frmAddEditLender.frx":286A
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditLender.frx":290E
         FormatStyle(2)  =   "frmAddEditLender.frx":2A6A
         FormatStyle(3)  =   "frmAddEditLender.frx":2B1A
         FormatStyle(4)  =   "frmAddEditLender.frx":2BCE
         FormatStyle(5)  =   "frmAddEditLender.frx":2CA6
         ImageCount      =   0
         PrinterProperties=   "frmAddEditLender.frx":2D5E
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   14085
         _ExtentX        =   24844
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin GridEX20.GridEX GridEX2 
         Height          =   5295
         Left            =   7560
         TabIndex        =   5
         Top             =   2430
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   9340
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
         Column(1)       =   "frmAddEditLender.frx":2F36
         Column(2)       =   "frmAddEditLender.frx":2FFE
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditLender.frx":30A2
         FormatStyle(2)  =   "frmAddEditLender.frx":31FE
         FormatStyle(3)  =   "frmAddEditLender.frx":32AE
         FormatStyle(4)  =   "frmAddEditLender.frx":3362
         FormatStyle(5)  =   "frmAddEditLender.frx":343A
         ImageCount      =   0
         PrinterProperties=   "frmAddEditLender.frx":34F2
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   1275
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Top             =   960
         Width           =   1155
      End
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   525
         Left            =   6840
         TabIndex        =   4
         Top             =   5040
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLender.frx":36CA
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSelect 
         Height          =   525
         Left            =   6840
         TabIndex        =   3
         Top             =   4440
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLender.frx":39E4
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdSearch 
         Height          =   525
         Left            =   7200
         TabIndex        =   1
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLender.frx":3CFE
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1275
      End
      Begin VB.Label Label4 
         Height          =   315
         Left            =   11250
         TabIndex        =   10
         Top             =   3420
         Width           =   585
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   5400
         TabIndex        =   6
         Top             =   7800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditLender.frx":4018
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   7200
         TabIndex        =   7
         Top             =   7800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "frmAddEditLender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_Lender_Items As CLender_Items
Private glbDaily As clsDaily

Public HeaderText As String
Public ShowMode As SHOW_MODE_TYPE
Public OKClick As Boolean
Public ID As Long
Public TempCollection As Collection

Private FileName As String
Private m_SumUnit As Double
Private m_TempCol1 As Collection
Private m_TempCol2 As Collection
Private m_TempCol3 As CLender_Items
Private m_TempCol4 As CLender
Private m_Companies As Collection
Private m_Accounts As Collection
Public m_Branches As Collection
Dim TempID As Long

Public AccountID As Long
Public ReceiptType As Long
Public InvoiceDOType As Long
Public Area As Long

Public DocumentDate As Date

Private Sub PopulateDestColl()
Dim Ri As CTaxDocItem
Dim D As CTaxDocItem

   For Each Ri In TempCollection
      Set D = New CTaxDocItem

      If Ri.Flag <> "D" Then
         Call D.CopyObject(1, Ri)
         Call m_TempCol2.Add(D)
      End If

      Set D = Nothing
   Next Ri
End Sub

Private Function IsIn(TempCol As Collection, VOUCHER As String) As Boolean
Dim D As CTaxDocItem
Dim Found As Boolean

   Found = False
   For Each D In TempCol
      If D.REF_NO = VOUCHER Then
         Found = True
      End If
   Next D

   IsIn = Found
End Function

Private Sub GenerateSourceItem(Rs As ADODB.Recordset, TempCol As Collection)
Dim BD As CLender_Items

   Set m_TempCol1 = Nothing
   Set m_TempCol1 = New Collection
   While Not Rs.EOF
      Set BD = New CLender_Items
      If ShowMode = SHOW_ADD Then
            Call BD.PopulateFromRS(3, Rs)
      ElseIf ShowMode = SHOW_EDIT Then
           Call BD.PopulateFromRS(2, Rs)
      End If

      If Not IsIn(m_TempCol2, BD.LENDER_ITEMS_ID) Then
         Call TempCol.Add(BD)
      End If

      Set BD = Nothing
      Rs.MoveNext
   Wend
End Sub
'
'Private Function GenerateAccountSet() As String
'Dim Mr As CMasterRef
'Dim TempStr As String
'Dim i As Long
'
'   If uctlAccountLookup.MyCombo.ListIndex > 0 Then
'      TempStr = "'" & uctlAccountLookup.MyTextBox.Text & "'"
'   Else
'      i = 0
'      TempStr = ""
'      For Each Mr In m_Accounts
'         i = i + 1
'         If i = m_Accounts.Count Then
'            TempStr = TempStr & "'" & Mr.KEY_CODE & "' "
'         Else
'            TempStr = TempStr & "'" & Mr.KEY_CODE & "', "
'         End If
'      Next Mr
'   End If
'   GenerateAccountSet = "(" & TempStr & ")"
'End Function

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long
 Dim TempID As Long
   IsOK = True
   If ShowMode = SHOW_ADD Then
   If Flag Then
      Call EnableForm(Me, False)
      m_Lender_Items.LENDER_ID = ID
      If Not glbDaily.QueryClenItems(m_Lender_Items, m_Rs, itemcount, IsOK, glbErrorLog, 1) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If itemcount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.itemcount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.itemcount = 0
      GridEX1.Rebind
   End If
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
ElseIf ShowMode = SHOW_EDIT Then
 If Flag Then
      Call EnableForm(Me, False)
      m_Lender_Items.LENDER_ID = ID
      If Not glbDaily.QueryClenItems(m_Lender_Items, m_Rs, itemcount, IsOK, glbErrorLog, 2) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If
      If itemcount > 0 Then
         Call GenerateSourceItem(m_Rs, m_TempCol1)
         Call CopyAllItem(m_TempCol1, m_TempCol2)
         GridEX2.itemcount = m_TempCol2.Count
         GridEX2.Rebind
     Else
         GridEX2.itemcount = 0
         GridEX2.Rebind
      End If
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
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
   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
      If Not glbDaily.QueryClenItems(m_Lender_Items, m_Rs, itemcount, IsOK, glbErrorLog, 2) Then
         glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
         Call EnableForm(Me, True)
         Exit Sub
      End If
   End If

   If itemcount > 0 Then
      Call GenerateSourceItem(m_Rs, m_TempCol1)
      GridEX1.itemcount = m_TempCol1.Count
      GridEX1.Rebind
   Else
      GridEX1.itemcount = 0
      GridEX1.Rebind
   End If

   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind

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

'Private Function SaveData() As Boolean
'Dim IsOK As Boolean
'
''   If ShowMode = SHOW_ADD Then
''      If Not VerifyAccessRight("INVENTORY_EXPORT_ADD") Then
''         Call EnableForm(Me, True)
''         Exit Function
''      End If
''   ElseIf ShowMode = SHOW_EDIT Then
''      If Not VerifyAccessRight("INVENTORY_EXPORT_EDIT") Then
''         Call EnableForm(Me, True)
''         Exit Function
''      End If
''   End If
'
'   If Not m_HasModify Then
'      SaveData = True
'      Exit Function
'   End If
'
'   Call PopulateTempColl
'
'   Call EnableForm(Me, True)
'   SaveData = True
'End Function
Private Function SaveData() As Boolean
Dim IsOK As Boolean
Dim TempData As New CLender_Items

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("LENDER_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("LENDER_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyCombo(lblCompany, uctlCompanyLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblBranch, uctlBranchLookup.MyCombo, False) Then
      Exit Function
   End If
   
   If Not VerifyCombo(lblFromDate, txtYear, False) Then
      Exit Function
   End If
   
''''   If Not VerifyTextControl(lblFromDate, txtYear, False) Then
''''      Exit Function
''''   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   Call EnableForm(Me, False)
   Set glbDaily = New clsDaily
      Set m_TempCol4 = New CLender
      m_TempCol4.AddEditMode = SHOW_ADD
      m_TempCol4.COMPANY_ID = uctlCompanyLookup.MyCombo.ItemData((Minus2Zero(uctlCompanyLookup.MyCombo.ListIndex)))
      m_TempCol4.BRANCH_ID = uctlBranchLookup.MyCombo.ItemData((Minus2Zero(uctlBranchLookup.MyCombo.ListIndex)))
      m_TempCol4.COMPANY_SHORTNAME = uctlCompanyLookup.MyTextBox.Text
      m_TempCol4.COMPANY_NAME = uctlCompanyLookup.MyCombo.Text
      m_TempCol4.BUDGET_YEAR = txtYear.Text
      
         If Not glbDaily.AddEditLender(m_TempCol4, IsOK, True, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            SaveData = False
            Call EnableForm(Me, True)
            Exit Function
         End If
'         If Not IsOK Then
'            Call EnableForm(Me, True)
'            glbErrorLog.ShowUserError
'            Exit Function
'         End If
   
      Set m_TempCol3 = Nothing
      
      
   
   
    For Each TempData In m_TempCol2
     Set m_TempCol3 = New CLender_Items
      m_TempCol3.AddEditMode = SHOW_ADD
      m_TempCol3.LENDER_ITEMS_NO = TempData.LENDER_ITEMS_NO
      m_TempCol3.LENDER_ITEMS_NAME = TempData.LENDER_ITEMS_NAME
      m_TempCol3.LENDER_ITEMS_AMOUNT = TempData.LENDER_ITEMS_AMOUNT
      m_TempCol3.LENDER_ID = m_TempCol4.LENDER_ID
      
         If Not glbDaily.AddEditLenderItems(m_TempCol3, IsOK, True, glbErrorLog) Then
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
   
      Set m_TempCol3 = Nothing
   Next TempData
  Set glbDaily = Nothing
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

Private Sub chkCommit_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkCommit_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub

Private Sub chkSaleFlag_Click(Value As Integer)
   m_HasModify = True
End Sub

Private Sub chkSaleFlag_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      SendKeys ("{TAB}")
   End If
End Sub


Private Sub cmdOK_Click()
'   If Not SaveData Then
'      Exit Sub
'   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPictureAdd_Click()
On Error Resume Next
Dim strDescription As String
   
   'edit the filter to support more image types
   dlgAdd.Filter = "Picture Files (*.jpg, *.gif)|*.jpg;*.gif"
   dlgAdd.DialogTitle = "Select Picture to Add to Database"
   dlgAdd.ShowOpen
   If dlgAdd.FileName = "" Then
      Exit Sub
   End If
    
   m_HasModify = True
End Sub

Private Sub cmdSearch_Click()
Dim Sc As CSCComp

Dim NewPath As String
Dim L As Long

   If Not VerifyCombo(lblCompany, uctlCompanyLookup.MyCombo, False) Then
      Exit Sub
   End If
   
   TempID = uctlCompanyLookup.MyCombo.ItemData((Minus2Zero(uctlCompanyLookup.MyCombo.ListIndex)))
   Set Sc = m_Companies(Trim(Str(TempID)))
   
   L = InStr(1, glbParameterObj.LegacyDBFile, "Secure", vbTextCompare)
   NewPath = Mid(glbParameterObj.LegacyDBFile, 1, L - 1) & Sc.Path
   
   Call glbDatabaseMngr.DisConnectLegacyDatabase
   
   If Not glbDatabaseMngr.ConnectLegacyDatabase(NewPath, "", "", glbErrorLog) Then
      Exit Sub
   End If
   
   Call QueryData(True)
   Call glbDatabaseMngr.DisConnectLegacyDatabase
End Sub

Public Sub CopyItem(TempCol1 As Collection, TempCol2 As Collection, ID As Long)
Dim L As CLender_Items
Dim TempGnl As CLender_Items

   If ID > 0 Then
      Set TempGnl = TempCol1(ID)

      Set L = New CLender_Items
'      L.PAY_DATE = TempGnl.VOUDAT
'      L.REF_NO = TempGnl.VOUCHER
      
'      Set frmAddEditLender.TaxDocItem = L
'      frmAddEditTaxDocItemEx.DefaultAccount = uctlAccountLookup.MyTextBox.Text
'      frmAddEditTaxDocItemEx.ShowMode = SHOW_EDIT
'      Load frmAddEditTaxDocItemEx
'      frmAddEditTaxDocItemEx.Show 1
'
'      OKClick = frmAddEditTaxDocItemEx.OKClick
'
'      Unload frmAddEditTaxDocItemEx
'      Set frmAddEditTaxDocItemEx = Nothing

      'If OKClick Then
         L.Flag = "A"
         Call TempCol2.Add(L)
         TempCol1.Remove (ID)
      'End If
      Set L = Nothing
   End If
End Sub

Public Sub CopyAllItem(TempCol1 As Collection, TempCol2 As Collection)
Dim j As Long

   For j = 1 To TempCol1.Count
      TempCol1(j).Flag = "A"
      Call TempCol2.Add(TempCol1(j))
   Next j
   Set TempCol1 = Nothing
   Set TempCol1 = New Collection
End Sub

Private Sub cmdSelect_Click()
Dim TempID As Long

   m_HasModify = True
   
   TempID = GridEX1.Row
   Call CopyItem(m_TempCol1, m_TempCol2, TempID)

   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Private Sub cmdSelectAll_Click()
   m_HasModify = True
   Call CopyAllItem(m_TempCol1, m_TempCol2)
  
   If Not SaveData Then
      Exit Sub
   End If
   
   GridEX1.itemcount = m_TempCol1.Count
   GridEX1.Rebind
   
   GridEX2.itemcount = m_TempCol2.Count
   GridEX2.Rebind
End Sub

Public Sub PopulateTempColl()
Dim D2 As CLender_Items
Dim Ri As CLender_Items

   For Each D2 In m_TempCol2
      Set Ri = New CLender_Items
      If D2.Flag = "A" Then
         Call Ri.CopyObject(1, D2)
         Ri.Flag = "A"
         Call TempCollection.Add(Ri)
      End If

      Set Ri = Nothing
   Next D2
End Sub

Private Sub Form_Activate()
'Dim FirstDate As Date
'Dim LastDate As Date

   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      Call LoadScComp(uctlCompanyLookup.MyCombo, m_Companies)
      Set uctlCompanyLookup.MyCollection = m_Companies

      Call LoadMaster(uctlBranchLookup.MyCombo, m_Branches, MASTER_BRANCH)
      Set uctlBranchLookup.MyCollection = m_Branches
      
      If (ShowMode = SHOW_EDIT) Then
'         m_GlJnl.QueryFlag = 1
         
         Call QueryData(True)
'      ElseIf ShowMode = SHOW_ADD Then
'         m_GlJnl.QueryFlag = 0
'         Call QueryData(True)
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
      Call cmdSearch_Click
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

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_Lender_Items = Nothing
   Set m_TempCol1 = Nothing
   Set m_TempCol2 = Nothing
   Set m_Companies = Nothing
   Set m_Accounts = Nothing
   Set m_Branches = Nothing
   Call glbDatabaseMngr.DisConnectLegacyDatabase
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
   Col.Width = 1000
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 2355
   Col.Caption = MapText("ชื่อผู้ให้กู้")

   Set Col = GridEX1.Columns.Add '4
   Col.Width = 2520
   Col.Caption = MapText("ยอดเงินกู้")
End Sub


Private Sub InitGrid2()
Dim Col As JSColumn

   GridEX2.Columns.Clear
   GridEX2.BackColor = GLB_GRID_COLOR
   GridEX2.itemcount = 0
   GridEX2.BackColorHeader = GLB_GRIDHD_COLOR
   GridEX2.ColumnHeaderFont.Bold = True
   GridEX2.ColumnHeaderFont.Name = GLB_FONT
   GridEX2.TabKeyBehavior = jgexControlNavigation
   
   Set Col = GridEX2.Columns.Add '1
   Col.Width = 1000
   Col.Caption = "ID"

   Set Col = GridEX2.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX2.Columns.Add '3
   Col.Width = 2355
   Col.Caption = MapText("ชื่อผู้ให้กู้")

   Set Col = GridEX2.Columns.Add '4
   Col.Width = 2520
   Col.Caption = MapText("ยอดเงินกู้")
End Sub

Private Sub GetTotalPrice()
'Dim II As CExportItem
'Dim Sum As Double
'
'   Sum = 0
'   For Each II In m_GlJnl.ImportExports
'      If II.Flag <> "D" Then
'         Sum = Sum + CDbl(Format(II.EXPORT_AVG_PRICE, "0.00")) * CDbl(Format(II.EXPORT_AMOUNT, "0.00"))
'      End If
'   Next II
''
''   txtDeliveryFee.Text = Format(Sum, "0.00")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
   Me.Caption = HeaderText
   pnlHeader.Caption = Me.Caption
   
   Call InitNormalLabel(lblFromDate, MapText("ปีที่ให้กู้"))
   Call InitNormalLabel(lblCompany, MapText("บริษัท"))
   Call InitNormalLabel(lblBranch, MapText("สาขา"))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSearch.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelect.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSelectAll.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdSearch, MapText("ค้นหา(F5)"))
   Call InitMainButton(cmdSelect, MapText(">"))
   Call InitMainButton(cmdSelectAll, MapText("COPY"))
   
   Call InitGrid1
   Call InitGrid2
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
   Set m_Lender_Items = New CLender_Items
   Set m_TempCol1 = New Collection
   Set m_TempCol2 = New Collection
   Set m_Companies = New Collection
   Set m_Accounts = New Collection
   Set m_Branches = New Collection
   Set glbDaily = New clsDaily
   
   If Not glbDatabaseMngr.ConnectLegacyDatabase(glbParameterObj.LegacyDBFile, "", "", glbErrorLog) Then
      Exit Sub
   End If
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"


   If m_TempCol1 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CLender_Items
   If m_TempCol1.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol1, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.LENDER_ITEMS_NO
   Values(2) = RealIndex
   Values(3) = CR.LENDER_ITEMS_NAME
   Values(4) = FormatNumber(CR.LENDER_ITEMS_AMOUNT)
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub GridEX2_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If m_TempCol2 Is Nothing Then
      Exit Sub
   End If

   If RowIndex <= 0 Then
      Exit Sub
   End If

   Dim CR As CLender_Items
   If m_TempCol2.Count <= 0 Then
      Exit Sub
   End If
   Set CR = GetItem(m_TempCol2, RowIndex, RealIndex)
   If CR Is Nothing Then
      Exit Sub
   End If

   Values(1) = CR.LENDER_ITEMS_NO
   Values(2) = RealIndex
   Values(3) = CR.LENDER_ITEMS_NAME
   Values(4) = FormatNumber(CR.LENDER_ITEMS_AMOUNT)
   Exit Sub

ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub txtDoNo_Change()
   m_HasModify = True
End Sub

Private Sub txtDeliveryNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtDocumentNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtTotal_Change()
   m_HasModify = True
End Sub

Private Sub txtTruckNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlDeliveryLookup_Change()
   m_HasModify = True
End Sub

Private Sub SSCommand2_Click()

End Sub

Private Sub uctlDocumentDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlEmployeeLookup_Change()
   m_HasModify = True
End Sub
