VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Begin VB.Form frmAddEditTaxDocument 
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "frmAddEditTaxDocument.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame1 
      Height          =   8520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15028
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   555
         Left            =   150
         TabIndex        =   4
         Top             =   2790
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   979
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComDlg.CommonDialog dlgAdd 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin Threed.SSPanel pnlHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   11925
         _ExtentX        =   21034
         _ExtentY        =   1244
         _Version        =   131073
         PictureBackgroundStyle=   2
      End
      Begin prjMtpTax.uctlDate uctlTravelDate 
         Height          =   405
         Left            =   1860
         TabIndex        =   0
         Top             =   1110
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlGuestNameLookup 
         Height          =   405
         Left            =   1860
         TabIndex        =   1
         Top             =   1560
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextLookup uctlDestinationLookup 
         Height          =   435
         Left            =   1860
         TabIndex        =   2
         Top             =   2010
         Width           =   5355
         _ExtentX        =   9446
         _ExtentY        =   767
      End
      Begin GridEX20.GridEX GridEX1 
         Height          =   4365
         Left            =   150
         TabIndex        =   17
         Top             =   3330
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7699
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
         Column(1)       =   "frmAddEditTaxDocument.frx":08CA
         Column(2)       =   "frmAddEditTaxDocument.frx":0992
         FormatStylesCount=   5
         FormatStyle(1)  =   "frmAddEditTaxDocument.frx":0A36
         FormatStyle(2)  =   "frmAddEditTaxDocument.frx":0B92
         FormatStyle(3)  =   "frmAddEditTaxDocument.frx":0C42
         FormatStyle(4)  =   "frmAddEditTaxDocument.frx":0CF6
         FormatStyle(5)  =   "frmAddEditTaxDocument.frx":0DCE
         ImageCount      =   0
         PrinterProperties=   "frmAddEditTaxDocument.frx":0E86
      End
      Begin Threed.SSCommand cmdSave 
         Height          =   525
         Left            =   6840
         TabIndex        =   8
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxDocument.frx":105E
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   525
         Left            =   10170
         TabIndex        =   3
         Top             =   1890
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxDocument.frx":1378
         ButtonStyle     =   3
      End
      Begin VB.Label lblFM 
         Height          =   315
         Left            =   8760
         TabIndex        =   16
         Top             =   3300
         Width           =   2685
      End
      Begin VB.Label lblTravelDate 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   270
         TabIndex        =   15
         Top             =   1110
         Width           =   1485
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   8475
         TabIndex        =   9
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxDocument.frx":1692
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   10125
         TabIndex        =   10
         Top             =   7830
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
         Top             =   7830
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
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxDocument.frx":19AC
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdDelete 
         Height          =   525
         Left            =   3420
         TabIndex        =   7
         Top             =   7830
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAddEditTaxDocument.frx":1CC6
         ButtonStyle     =   3
      End
      Begin VB.Label lblGuestName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   1650
         Width           =   1575
      End
      Begin VB.Label lblDestination 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   270
         TabIndex        =   12
         Top             =   2070
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmAddEditTaxDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const ROOT_TREE = "Root"
Private m_HasActivate As Boolean
Private m_HasModify As Boolean
Private m_Rs As ADODB.Recordset
Private m_TaxDoc As CTaxDocument
Private m_Coll As Collection

Public HeaderText As String

Public ShowMode As SHOW_MODE_TYPE
Public TaxReType As TAXWH_TYPE
Public Daily As clsDaily

Public OKClick As Boolean
Public ID As Long
Public ENT_ID As Long

Private m_Combos As Collection
Private m_ReportControls As Collection
Private m_Employees As Collection
Public m_Agencies As Collection
Public m_Sources As Collection
Public m_Dests As Collection
Public m_TaxDocs As Collection
Public Area As Long
Public TaxType As Long

Private m_Texts As Collection
Private m_Dates As Collection
Private m_Labels As Collection
Private m_TextLookups As Collection
Private m_CyclePerMonth As Long

Private m_Companies As Collection
Private FileName As String
Public Document_Year As String
Public COMPANY_ID As Long

Private Sub QueryData(Flag As Boolean)
Dim IsOK As Boolean
Dim itemcount As Long

   IsOK = True
   If Flag Then
      Call EnableForm(Me, False)
            
      m_TaxDoc.TAX_DOCUMENT_ID = ID
      m_TaxDoc.QueryFlag = 1
      m_TaxDoc.TAX_TYPE = TaxType
      m_TaxDoc.COMPANY_ID = COMPANY_ID
      m_TaxDoc.Document_Year = Document_Year
      If TaxType = 21 Then
         If Not glbDaily.QueryTax_A_Document(m_TaxDoc, m_Coll, itemcount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      Else
         If Not glbDaily.QueryTaxDocument(m_TaxDoc, m_Rs, itemcount, IsOK, glbErrorLog) Then
            glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
            Call EnableForm(Me, True)
            Exit Sub
         End If
      End If
   End If
   
   If TaxType = 21 Then
       If CountItem(m_TaxDoc.DocumentItems) > 0 Then
'         Call m_TaxDoc.PopulateFromRS(2, m_TaxDoc)

         uctlTravelDate.ShowDate = m_TaxDoc.DOCUMENT_DATE
         uctlGuestNameLookup.MyCombo.ListIndex = IDToListIndex(uctlGuestNameLookup.MyCombo, m_TaxDoc.COMPANY_ID)
         uctlDestinationLookup.MyCombo.ListIndex = IDToListIndex(uctlDestinationLookup.MyCombo, m_TaxDoc.RESPONSE_ID)
         
         GridEX1.itemcount = CountItem(m_TaxDoc.DocumentItems)
         GridEX1.Rebind
      End If
   Else
      If itemcount > 0 Then
         Call m_TaxDoc.PopulateFromRS(1, m_Rs)
         
         uctlTravelDate.ShowDate = m_TaxDoc.DOCUMENT_DATE
         uctlGuestNameLookup.MyCombo.ListIndex = IDToListIndex(uctlGuestNameLookup.MyCombo, m_TaxDoc.COMPANY_ID)
         uctlDestinationLookup.MyCombo.ListIndex = IDToListIndex(uctlDestinationLookup.MyCombo, m_TaxDoc.RESPONSE_ID)
         
         GridEX1.itemcount = CountItem(m_TaxDoc.DocumentItems)
         GridEX1.Rebind
      End If
   End If
      
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

Private Function SaveData() As Boolean
Dim IsOK As Boolean

   If ShowMode = SHOW_ADD Then
      If Not VerifyAccessRight("TAX_WITHOLD_ADD") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   ElseIf ShowMode = SHOW_EDIT Then
      If Not VerifyAccessRight("TAX_WITHOLD_EDIT") Then
         Call EnableForm(Me, True)
         Exit Function
      End If
   End If

   If Not VerifyDate(lblTravelDate, uctlTravelDate, False) Then
      Exit Function
   End If
   If Not VerifyCombo(lblGuestName, uctlGuestNameLookup.MyCombo, False) Then
      Exit Function
   End If
'   If Not VerifyCombo(lblDestination, uctlDestinationLookup.MyCombo, False) Then
'      Exit Function
'   End If
      
   If Not m_HasModify Then
      SaveData = True
      Exit Function
   End If
   
   m_TaxDoc.AddEditMode = ShowMode
   m_TaxDoc.DOCUMENT_DATE = uctlTravelDate.ShowDate
   m_TaxDoc.COMPANY_ID = uctlGuestNameLookup.MyCombo.ItemData(Minus2Zero(uctlGuestNameLookup.MyCombo.ListIndex))
   m_TaxDoc.RESPONSE_ID = uctlDestinationLookup.MyCombo.ItemData(Minus2Zero(uctlDestinationLookup.MyCombo.ListIndex))
   m_TaxDoc.TAX_TYPE = TaxType
   
   Call EnableForm(Me, False)
   If Not glbDaily.AddEditTaxDocument(m_TaxDoc, IsOK, True, glbErrorLog) Then
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

Private Sub cboLanguage_Click()
   m_HasModify = True
End Sub

Private Sub cboVehicle_Click()
   m_HasModify = True
End Sub

Private Sub cmdAdd_Click()
Dim OKClick As Boolean
Dim M As cPopupMenu
Dim lMenuChoosen As Long

   Set M = New cPopupMenu
   lMenuChoosen = M.Popup("อิมพอร์ตข้อมูลจาก Express", "-", "เพิ่มทีละรายการ")
   Set M = Nothing
   
   OKClick = False
   If (TabStrip1.SelectedItem.Index = 1) And (lMenuChoosen = 3) Then
      Set frmAddEditTaxDocItem.TempCollection = m_TaxDoc.DocumentItems
      frmAddEditTaxDocItem.ShowMode = SHOW_ADD
      frmAddEditTaxDocItem.HeaderText = MapText("เพิ่มรายการภาษี")
      Load frmAddEditTaxDocItem
      frmAddEditTaxDocItem.Show 1

      OKClick = frmAddEditTaxDocItem.OKClick

      Unload frmAddEditTaxDocItem
      Set frmAddEditTaxDocItem = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_TaxDoc.DocumentItems)
         GridEX1.Rebind
      End If
   ElseIf (TabStrip1.SelectedItem.Index = 1) And (lMenuChoosen = 1) Then
      Set frmAddWHVoucher.TempCollection = m_TaxDoc.DocumentItems
      frmAddWHVoucher.DocumentDate = uctlTravelDate.ShowDate
      frmAddWHVoucher.HeaderText = MapText("เลือกรายการภาษีจาก Express")
      Load frmAddWHVoucher
      frmAddWHVoucher.Show 1

      OKClick = frmAddWHVoucher.OKClick

      Unload frmAddWHVoucher
      Set frmAddWHVoucher = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_TaxDoc.DocumentItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdDelete_Click()
Dim ID1 As Long
Dim ID2 As Long

   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If
   
   If Not ConfirmDelete(GridEX1.Value(3)) Then
      Exit Sub
   End If
   
   ID2 = GridEX1.Value(2)
   ID1 = GridEX1.Value(1)
   
   If TabStrip1.SelectedItem.Index = 1 Then
      If ID1 <= 0 Then
         m_TaxDoc.DocumentItems.Remove (ID2)
      Else
         m_TaxDoc.DocumentItems.Item(ID2).Flag = "D"
      End If

      GridEX1.itemcount = CountItem(m_TaxDoc.DocumentItems)
      GridEX1.Rebind
      m_HasModify = True
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub cmdEdit_Click()
Dim IsOK As Boolean
Dim itemcount As Long
Dim IsCanLock As Boolean
Dim ID As Long
Dim OKClick As Boolean

'   If Not VerifyAccessRight("GROUP_QUERY_RIGHT") Then
'      Exit Sub
'   End If
      
   If Not VerifyGrid(GridEX1.Value(1)) Then
      Exit Sub
   End If

   ID = Val(GridEX1.Value(2))
   OKClick = False
   
   If TabStrip1.SelectedItem.Index = 1 Then
      frmAddEditTaxDocItem.ID = ID
      Set frmAddEditTaxDocItem.TempCollection = m_TaxDoc.DocumentItems
      frmAddEditTaxDocItem.HeaderText = MapText("แก้ไขรายการภาษี")
      frmAddEditTaxDocItem.ShowMode = SHOW_EDIT
      Load frmAddEditTaxDocItem
      frmAddEditTaxDocItem.Show 1

      OKClick = frmAddEditTaxDocItem.OKClick

      Unload frmAddEditTaxDocItem
      Set frmAddEditTaxDocItem = Nothing

      If OKClick Then
         GridEX1.itemcount = CountItem(m_TaxDoc.DocumentItems)
         GridEX1.Rebind
      End If
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
   
   If OKClick Then
      m_HasModify = True
   End If
End Sub

Private Sub cmdOK_Click()
   If Not SaveData Then
      Exit Sub
   End If
   
   OKClick = True
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim Report As CReportInterface
Dim SelectFlag As Boolean
Dim Key As String
Dim Name As String
Dim lMenuChosen As Long
Dim oMenu As cPopupMenu

Dim LocationSave As String
Dim FileID As Long
Dim i As Long

Dim CR As CTaxDocItem
Dim TempStr As String


DATA_SELC = 0
Set oMenu = New cPopupMenu
If TAX_TYPE_NAME = 2 Then
   lMenuChosen = oMenu.Popup("รายงานภาษีหัก ณ ที่จ่าย ภ.ง.ด 2", "-", "รายงานบัญชีพิเศษ ภาษีหัก ณ ที่จ่าย  ภ.ง.ด 2", "-", "EXPORT ไฟล์ ภ.ง.ด 2 ยื่นแบบ Online")
ElseIf TAX_TYPE_NAME = 3 Then
    lMenuChosen = oMenu.Popup("รายงานภาษีหัก ณ ที่จ่าย ภ.ง.ด 3", "-", "รายงานบัญชีพิเศษ ภาษีหัก ณ ที่จ่าย  ภ.ง.ด 3", "-", "EXPORT ไฟล์ ภ.ง.ด 3 ยื่นแบบ Online")
 ElseIf TAX_TYPE_NAME = 53 Then
   lMenuChosen = oMenu.Popup("รายงานภาษีหัก ณ ที่จ่าย ภ.ง.ด 53", "-", "รายงานบัญชีพิเศษ ภาษีหัก ณ ที่จ่าย  ภ.ง.ด 53", "-", "EXPORT ไฟล์ ภ.ง.ด 53 ยื่นแบบ Online")
 ElseIf TAX_TYPE_NAME = 4 Then
 
 ElseIf TAX_TYPE_NAME = 21 Then
   lMenuChosen = oMenu.Popup("EXPORT ไฟล์ ภ.ง.ด 2ก ยื่นแบบ Online")
End If
   
   Key = m_TaxDoc.TAX_DOCUMENT_ID 'ส่งพารามิเตอร์
   Name = Me.Caption  'trvMaster.SelectedItem.Text
      
   SelectFlag = False
   
   If Not VerifyReportInput Then
      Exit Sub
   End If
   
   Set Report = New CReportInterface
                                                                                               

If lMenuChosen = 1 And TAX_TYPE_NAME = 2 Then
   Set Report = New CReportTax002
'  Set Report = New CReportTaxSendingYear
ElseIf lMenuChosen = 1 And TAX_TYPE_NAME = 3 Then
         Set Report = New CReportTax003
 ElseIf lMenuChosen = 1 And TAX_TYPE_NAME = 53 Then
      Set Report = New CReportTax0053
 ElseIf lMenuChosen = 3 Then
     Set Report = New CReportTaxSending
ElseIf TAX_TYPE_NAME = 2 And lMenuChosen = 5 Then 'Export ออกไป Text File
   LocationSave = "C:\ExportOnline\Tax2\"
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & m_TaxDoc.SHORT_NAME
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_P02.txt"
      
On Error GoTo xxx
   Call Kill(LocationSave)
xxx:

   
   
   FileID = FreeFile
   Open LocationSave For Append As #FileID
   i = 0
   
   For Each CR In m_TaxDoc.DocumentItems
      i = i + 1
      TempStr = Format(i, "00") & "|"                                                                                                           'ลำดับ
      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                             'TAX ID ผู้หัก
      TempStr = TempStr & CR.IDENT_ID & "|"                                                                             'TAX ID ผู้ถูกหัก
      TempStr = TempStr & "" & "|"                                                                                                           'คำนำหน้าชื่อ
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                                                       'ชื่อ
      TempStr = TempStr & "" & "|"                                                                                                           'นามสกุล
      TempStr = TempStr & "" & "|"                                                                                                           'บัญชีเงินฝาก
      TempStr = TempStr & CR.EXPORT_KEY & "|"                                                                                                           'เงินได้ตามมาตรา
      TempStr = TempStr & DateToStringToOnline1(CR.PAY_DATE) & "|"                                                                                                           'นามสกุล
      
      If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|"                                'อัตราภาษี
      End If
      TempStr = TempStr & FormatNumberToNull(CR.PAID_AMOUNT, , False) & "|"                                                                     'จำนวนเงินที่จ่าย
      TempStr = TempStr & FormatNumberToNull(CR.WH_AMOUNT, , False) & "|"                                                                     'จำนวนเงินที่หักไว้
      TempStr = TempStr & CR.CONDITION_NAME & ""                                                                                                           'เงินได้ตามมาตรา
      
      Print #FileID, TempStr
   Next CR
   
   Close #FileID
   
   glbErrorLog.LocalErrorMsg = "COMPLETE " & i & " Record"
   glbErrorLog.ShowUserError
   
   Exit Sub
ElseIf TAX_TYPE_NAME = 21 And lMenuChosen = 1 Then 'ภงด2ก
Dim FileName As String
   LocationSave = "C:\ExportOnline\Tax2a\"
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   FileName = "PND2A_" & m_TaxDoc.TAX_ID & "_" & "000000" & "_" & Format(Year(uctlTravelDate.ShowDate) + 543, "0000") & "_00_00_00"
   LocationSave = LocationSave & FileName & ".txt"
   
   
Dim RegCode As String
RegCode = Trim(InputBox("กรุณากรอกรหัสอนุมัติลงทะเบียนที่ได้รับจากกรมสรรพากร", "ป้อนค่า", ""))
If Not Len(RegCode) = 20 Then
   glbErrorLog.LocalErrorMsg = "จำนวนตัวเลขต้องมี 20 หลัก"
   glbErrorLog.ShowUserError
   Exit Sub
End If
On Error GoTo xxx2
   Call Kill(LocationSave)
xxx2:

''Dim s As Variant
    
   FileID = FreeFile
   
   Open LocationSave For Append As #FileID


   TempStr = ""
   '####### Header############
      TempStr = "H" & "|"                                                                                                           'Header=1
      TempStr = TempStr & "0000" & "|"                                                                              'รหัสผู้นำส่ง=2
      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                             'TAX ID ผู้หัก=3
      TempStr = TempStr & "000000" & "|"                                                                              'รหัสสาขาผู้ถูกหัก=4
       TempStr = TempStr & "1" & "|"                                                                                                           'ประเภทการนำส่ง=5
       TempStr = TempStr & "PND2A" & "|"                                                                                                           'ประเภทแบบภาษี=6
       TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                                                            'TAX_ID ผู้มีหน้าที่หักภาษี=7
       TempStr = TempStr & "000000" & "|"                                                                              'รหัสสาขา ผู้ถูกหัก=8
       TempStr = TempStr & "สำนักงานใหญ่" & "|"                                                                                                           'ชื่อแผนก/ส่วน/ฝ่าย=9
       TempStr = TempStr & "1" & "|"                                                                                                           'สถานะผู้ประกอบการรายใหญ่=10
       TempStr = TempStr & "00" & "|"                                                                                                           'เดือนภาษี=11
        TempStr = TempStr & Format(Year(uctlTravelDate.ShowDate) + 543, "0000") & "|"                                                                                                             'ปีภาษี=12
        TempStr = TempStr & "|"                                                                                                            'ประเภทสาขา=13
      TempStr = TempStr & "00" & "|"                                                                                                            'ประเภทการยื่นแบบ=14

      Dim SumPaidAmount As Double
      Dim SumWhAmount As Double
      Dim SumAddAmount As Double
      i = 0
      For Each CR In m_TaxDoc.DocumentItems
         i = i + 1
         SumPaidAmount = SumPaidAmount + CDbl(FormatNumberToNull(CR.PAID_AMOUNT, , False))
         SumWhAmount = SumWhAmount + CDbl(FormatNumberToNull(CR.WH_AMOUNT, , False))

      Next CR
      TempStr = TempStr & i & "|"                                                                                                            'รวมจำนวนราย=15
      TempStr = TempStr & FormatNumberToNull(SumPaidAmount, , False) & "|"                                                                                                             'รวมจำนวนเงินได้ทั้งสิ้น=16
      TempStr = TempStr & FormatNumberToNull(SumWhAmount, , False) & "|"                                                                                                             'รวมจำนวนเงินภาษีที่นำส่งทั้งสิ้น=17
      SumAddAmount = 0
      TempStr = TempStr & FormatNumber(SumAddAmount) & "|"                                                                                                           'จำนวนเงินเพิ่ม=18
      TempStr = TempStr & FormatNumberToNull(SumWhAmount + SumAddAmount, , False) & "|"                                                                                                          'จำนวนเงินรวมยอดภาษีนำส่งทั้งสิ้นและเงินเพิ่ม=19
      TempStr = TempStr & "0.00" & "|"                                                                                                            'จำนวนเงินโอนผ่านธนาคาร=20
      TempStr = TempStr & RegCode & "|"                                                                                                             'รหัสลงทะเบียน=21
      TempStr = TempStr & "1"                                                                                                        'ช่องทางการยื่นแบบ=22

'''        s = AToW(TempStr, CP_UTF8)
'''      Put #FileID, , s
     
     Print #FileID, TempStr
   '#######End Herder###########

   '#######Detail###########
   i = 0
   For Each CR In m_TaxDoc.DocumentItems
      i = i + 1
      TempStr = "D" & "|"                                                                                                           'Detail=1
      TempStr = TempStr & Format(i, "00") & "|"                                                                                                           'ลำดับที่ =2
      TempStr = TempStr & "000000" & "|"                                                                                                           'สาขาผู้หักภาษี =3
      TempStr = TempStr & CR.IDENT_ID & "|"                                                                             'TAX ID ผู้ถูกหัก=4
      TempStr = TempStr & "0000000000" & "|"                                                                             'TAX ID ผู้หัก=5
      TempStr = TempStr & "999999999" & "|"                                                                                                            'เลขที่บัญชีเงินฝาก=6
      TempStr = TempStr & "-" & "|"                                                                                                           'คำนำหน้าชื่อ=7
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                                                       'ชื่อผู้มีเงินได้=8
      TempStr = TempStr & "" & "|"                                                                                                           'นามสกุลผู้มีเงินได้=9
      TempStr = TempStr & "31122560" & "|"                                                                                                           'วันเดือนปีที่จ่าย=10
       If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Format(Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1), "0.00") & "|"                               'อัตราภาษี=11
      Else
         TempStr = TempStr & "0.00" & "|"                                'กรณีไม่มี อัตราภาษี=11
      End If
      TempStr = TempStr & FormatNumberToNull(CR.PAID_AMOUNT, , False) & "|"                                                                     'จำนวนเงินที่จ่าย=12
      TempStr = TempStr & FormatNumberToNull(CR.WH_AMOUNT, , False) & "|"                                                                     'จำนวนเงินภาษีที่หักไว้=13
      TempStr = TempStr & CR.EXPORT_KEY & "|"                                                                                                           'เงินได้ตามมาตรา=14
      TempStr = TempStr & "1" & "|"                                                                                                           'เงื่อนไขการหักภาษี=15
      TempStr = TempStr & "" & "|"                                                                                                            'ที่อยู่=16
      TempStr = TempStr & CR.HOME & "|"                                                                                                            'ห้องเลขที่=17
      TempStr = TempStr & "" & "|"                                                                                                            'ชั้นที่=18
      TempStr = TempStr & CR.VILLAGE & "|"                                                                                                             'หมู่บ้าน=19
      TempStr = TempStr & "" & "|"                                                                                                            'เลขที่=20
      TempStr = TempStr & CR.MOO & "|"                                                                                                             'หมู่ที่=21
      TempStr = TempStr & CR.SOI & "|"                                                                                                             'ตรอก/ซอย=22
      TempStr = TempStr & CR.ROAD & "|"                                                                                                             'ถนน=23
      TempStr = TempStr & CR.DISTRICT & "|"                                                                                                            'ตำบล/แขวง=24
      TempStr = TempStr & CR.AMPHUR & "|"                                                                                                            'อำเภอ/เขต=25
      TempStr = TempStr & CR.PROVINCE & "|"                                                                                                             'จังหวัด=26
      TempStr = TempStr & CR.ZIPCODE                                                                                                          'รหัสไปรษณีย์=27


      Print #FileID, TempStr
   Next CR

   '#########EndDetail###########


   Close #FileID

   glbErrorLog.LocalErrorMsg = "COMPLETE " & i & " Record"
   glbErrorLog.ShowUserError

   Exit Sub

ElseIf TAX_TYPE_NAME = 3 And lMenuChosen = 5 Then
   LocationSave = "C:\ExportOnline\Tax3\"
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & m_TaxDoc.SHORT_NAME
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_P03.txt"
      
   On Error GoTo xxxx
      Call Kill(LocationSave)
xxxx:
   
   FileID = FreeFile
   Open LocationSave For Append As #FileID
   i = 0
   
   For Each CR In m_TaxDoc.DocumentItems
       i = i + 1
      TempStr = Format(i, "00") & "|"                                                                     'ลำดับ
      If Len(CR.IDENT_ID) > 0 Then
         TempStr = TempStr & CR.IDENT_ID & "|"                                                'เลขประจำตัวประชาชนผู้เสียภาษี
      ElseIf Len(CR.TAX_ID) > 0 Then
         TempStr = TempStr & CR.TAX_ID & "|"                                                    'เลขประจำตัวผู้เสียภาษี
      End If
      TempStr = TempStr & "|"                                                                              'คำนำหน้าชื่อ
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                   'ชื่อ
      TempStr = TempStr & "|"                                                                              'นามสกุล
      TempStr = TempStr & CR.PackAddress & "|"                                              'ที่อยู่
      TempStr = TempStr & Format(CR.PAY_DATE, "DD/MM/YYYY") & "|"   'วันเดือนปีที่จ่ายเงินได้
      TempStr = TempStr & CR.REVENUE_NAME & "|"                                      'ประเภทเงินได้
       If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|" 'อัตราภาษี
      End If
      TempStr = TempStr & CR.PAID_AMOUNT & "|"                                        'จำนวนเงินที่จ่าย
      TempStr = TempStr & CR.WH_AMOUNT & "|"                                           'จำนวนเงินที่หัก
      TempStr = TempStr & CR.CONDITION_NAME & "|"                                         'เงื่อนไขการหัก
      Print #FileID, TempStr
   Next CR
   
   Close #FileID
   
   glbErrorLog.LocalErrorMsg = "COMPLETE " & i & " Record"
   glbErrorLog.ShowUserError
   
   Exit Sub
ElseIf TAX_TYPE_NAME = 53 And lMenuChosen = 5 Then
      LocationSave = "C:\ExportOnline\Tax53\"
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   
   LocationSave = LocationSave & m_TaxDoc.SHORT_NAME
   LocationSave = LocationSave & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00")
   LocationSave = LocationSave & "_P53.txt"
      
   On Error GoTo xxxxx
      Call Kill(LocationSave)
xxxxx:
   
   FileID = FreeFile
   Open LocationSave For Append As #FileID
   i = 0
   
   For Each CR In m_TaxDoc.DocumentItems
      i = i + 1
      TempStr = Format(i, "00") & "|"                                                                     'ลำดับ
      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                          'เลขประจำตัวผู้เสียภาษีผู้หัก
      TempStr = TempStr & CR.TAX_ID & "|"                                                      'เลขประจำตัวผู้เสียภาษีผู้ถูกหัก
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                   'ชื่อ
      TempStr = TempStr & CR.PackAddress & "|"                                              'ที่อยู่
      TempStr = TempStr & Format(CR.PAY_DATE, "DD/MM/YYYY") & "|"   'วันเดือนปีที่จ่ายเงินได้
      TempStr = TempStr & CR.REVENUE_NAME & "|"                                      'ประเภทเงินได้
       If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|" 'อัตราภาษี
      End If
      TempStr = TempStr & CR.PAID_AMOUNT & "|"                                        'จำนวนเงินที่จ่าย
      TempStr = TempStr & CR.WH_AMOUNT & "|"                                           'จำนวนเงินที่หัก
      TempStr = TempStr & CR.CONDITION_NAME & "|"                                         'เงื่อนไขการหัก
      Print #FileID, TempStr
   Next CR
   
   Close #FileID
   
   glbErrorLog.LocalErrorMsg = "COMPLETE " & i & " Record"
   glbErrorLog.ShowUserError
   
   Exit Sub
Else
    Exit Sub
End If
     
      SelectFlag = True
   
        Call Report.AddParam(m_TaxDoc.TAX_DOCUMENT_ID, "TAX_DOCUMENT_ID")
        Call Report.AddParam(Name, "REPORT_NAME")
        Call Report.AddParam("TAX001", "REPORT_KEY")
        Call Report.AddParam(uctlTravelDate.ShowDate, "FROM_DATE")
        Call Report.AddParam(uctlGuestNameLookup.MyTextBox.Text, "E_NAME")
        Call Report.AddParam(m_TaxDoc.IDENT_ID, "IDENT_ID")
        Call Report.AddParam(m_TaxDoc.TAX_ID, "TAX_ID")
        Call FillReportInput(Report)
       

      Set frmReport.ReportObject = Report
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing
   
End Sub
Private Function CalData()
Dim CR As CTaxDocItem
Dim i As Long
' For Each CR In m_TaxDoc.DocumentItems
'      i = i + 1
''      TempStr = Format(i, "00") & "|"                                                                                                           'ลำดับ
''      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                             'TAX ID ผู้หัก
''      TempStr = TempStr & CR.IDENT_ID & "|"                                                                             'TAX ID ผู้ถูกหัก
''      TempStr = TempStr & "" & "|"                                                                                                           'คำนำหน้าชื่อ
''      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                                                       'ชื่อ
''      TempStr = TempStr & "" & "|"                                                                                                           'นามสกุล
''      TempStr = TempStr & "" & "|"                                                                                                           'บัญชีเงินฝาก
''      TempStr = TempStr & CR.EXPORT_KEY & "|"                                                                                                           'เงินได้ตามมาตรา
''      TempStr = TempStr & DateToStringToOnline1(CR.PAY_DATE) & "|"                                                                                                           'นามสกุล
'
'      If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
'         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|"                                'อัตราภาษี
'      End If
'      TempStr = TempStr & FormatNumberToNull(CR.PAID_AMOUNT, , False) & "|"                                                                     'จำนวนเงินที่จ่าย
'      TempStr = TempStr & FormatNumberToNull(CR.WH_AMOUNT, , False) & "|"                                                                     'จำนวนเงินที่หักไว้
'      TempStr = TempStr & CR.CONDITION_NAME & ""                                                                                                           'เงินได้ตามมาตรา
'
'      Print #FileID, TempStr
'   Next CR
End Function
Private Function VerifyReportInput() As Boolean
Dim C As CReportControl

   VerifyReportInput = False
'   For Each C In m_ReportControls
'      If (C.ControlType = "C") Then
'         If Not VerifyCombo(Nothing, m_Combos(C.ControlIndex), C.AllowNull) Then
'            Exit Function
'         End If
'      End If
'
'      If (C.ControlType = "T") Then
'         If Not VerifyTextControl(Nothing, m_Texts(C.ControlIndex), C.AllowNull) Then
'            Exit Function
'         End If
'      End If
'
'      If (C.ControlType = "D") Then
'         If Not VerifyDate(Nothing, m_Dates(C.ControlIndex), C.AllowNull) Then
'            Exit Function
'         End If
'      End If
'   Next C
   VerifyReportInput = True
End Function
Private Sub FillReportInput(r As CReportInterface)
Dim C As CReportControl

'   Call R.AddParam(Picture1.Picture, "PICTURE")
'   For Each C In m_ReportControls
'      If (C.ControlType = "C") Then
'         If C.Param1 <> "" Then
'            Call R.AddParam(m_Combos(C.ControlIndex).Text, C.Param1)
'         End If
'
'         If C.Param2 <> "" Then
'            Call R.AddParam(m_Combos(C.ControlIndex).ItemData(Minus2Zero(m_Combos(C.ControlIndex).ListIndex)), C.Param2)
'         End If
'      End If
'
'      If (C.ControlType = "T") Then
'         If C.Param1 <> "" Then
'            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param1)
'         End If
'
'         If C.Param2 <> "" Then
'            Call R.AddParam(m_Texts(C.ControlIndex).Text, C.Param2)
'         End If
'      End If
'
'      If (C.ControlType = "D") Then
'         If C.Param1 <> "" Then
'            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param1)
'         End If
'
'         If C.Param2 <> "" Then
'            If m_Dates(C.ControlIndex).ShowDate <= 0 Then
'               If C.Param2 = "TO_DATE" Then
'                  m_Dates(C.ControlIndex).ShowDate = -1
'               ElseIf C.Param2 = "FROM_DATE" Then
'                  m_Dates(C.ControlIndex).ShowDate = -2
'               End If
'            End If
'            Call R.AddParam(m_Dates(C.ControlIndex).ShowDate, C.Param2)
'         End If
'      End If
'
'   Next C
End Sub
Private Sub cmdSave_Click()
Dim Result As Boolean
   If Not SaveData Then
      Exit Sub
   End If
   
   ShowMode = SHOW_EDIT
   ID = m_TaxDoc.TAX_DOCUMENT_ID
   m_TaxDoc.QueryFlag = 1
   QueryData (True)
   m_HasModify = False
End Sub

Private Sub Form_Activate()
   If Not m_HasActivate Then
      m_HasActivate = True
      Me.Refresh
      DoEvents
      
      Call EnableForm(Me, False)
      
      Call LoadMaster(uctlGuestNameLookup.MyCombo, m_TaxDocs, MASTER_REVENUETYPE)
      Set uctlGuestNameLookup.MyCollection = m_TaxDocs
      
      Call loadEnterprise(uctlGuestNameLookup.MyCombo, m_Companies)
      Set uctlGuestNameLookup.MyCollection = m_Companies
      
      Call LoadEmployee(uctlDestinationLookup.MyCombo, m_Dests)
      Set uctlDestinationLookup.MyCollection = m_Dests
      
      If (ShowMode = SHOW_EDIT) Or (ShowMode = SHOW_VIEW_ONLY) Then
         m_TaxDoc.QueryFlag = 1
         Call QueryData(True)
         Call TabStrip1_Click
      ElseIf ShowMode = SHOW_ADD Then
         m_TaxDoc.QueryFlag = 0
         Call QueryData(False)
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
'      Call cmdSearch_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 115 Then
'      Call cmdClear_Click
      KeyCode = 0
   ElseIf Shift = 0 And KeyCode = 118 Then
      Call cmdAdd_Click
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
      Call cmdPrint_Click
      KeyCode = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If m_Rs.State = adStateOpen Then
      m_Rs.Close
   End If
   Set m_Rs = Nothing
   
   Set m_TaxDoc = Nothing
   Set m_Employees = Nothing
   Set m_Agencies = Nothing
   Set m_Sources = Nothing
   Set m_Dests = Nothing
   Set m_TaxDocs = Nothing
   Set m_Companies = Nothing
End Sub

Private Sub GridEX1_ColResize(ByVal ColIndex As Integer, ByVal NewColWidth As Long, ByVal Cancel As GridEX20.JSRetBoolean)
   Debug.Print ColIndex & " " & NewColWidth
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
   Col.Width = 0
   Col.Caption = "ID"

   Set Col = GridEX1.Columns.Add '2
   Col.Width = 0
   Col.Caption = "Real ID"

   Set Col = GridEX1.Columns.Add '3
   Col.Width = 6270
   Col.Caption = MapText("ชื่อผู้มีเงินได้ (ซัพพลายเออร์)")
  
  If TaxType = 21 Then
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 0
      Col.Caption = MapText("")
  Else
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 1935
      Col.Caption = MapText("วันเดือนปีที่จ่าย")
   End If

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2085
   Col.Caption = MapText("ประเภทเงินได้")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2370
   Col.Caption = MapText("อัตราภาษี %")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 1620
   Col.Caption = MapText("จำนวนเงินที่จ่าย")
   Col.TextAlignment = jgexAlignRight

   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1620
   Col.Caption = MapText("จำนวนเงินนำส่ง")
   Col.TextAlignment = jgexAlignRight

   Set Col = GridEX1.Columns.Add '9
   Col.Width = 1620
   Col.Caption = MapText("เงื่อนไข")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

If TaxType = 21 Then
   Me.Caption = HeaderText & "ก"
Else
   Me.Caption = HeaderText & TAX_TYPE_NAME
End If
pnlHeader.Caption = Me.Caption

   
   Call InitNormalLabel(lblTravelDate, MapText("วันที่นำส่ง"))
   Call InitNormalLabel(lblGuestName, MapText("บริษัท"))
   Call InitNormalLabel(lblDestination, MapText("ผู้รับผิดชอบ"))

   Call InitNormalLabel(lblFM, MapText(""))
   
   Me.Picture = LoadPicture(glbParameterObj.MainPicture)
   
   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19
   
   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdAdd.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdEdit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdDelete.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdPrint.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdSave.Picture = LoadPicture(glbParameterObj.NormalButton1)
   
   Call InitMainButton(cmdExit, MapText("ยกเลิก (ESC)"))
   Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
   Call InitMainButton(cmdAdd, MapText("เพิ่ม (F7)"))
   Call InitMainButton(cmdEdit, MapText("แก้ไข (F3)"))
   Call InitMainButton(cmdDelete, MapText("ลบ (F6)"))
   Call InitMainButton(cmdPrint, MapText("พิมพ์ (F8)"))
   Call InitMainButton(cmdSave, MapText("บันทึก (F10)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Call InitGrid1
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("รายละเอียด")
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
   Set m_TaxDoc = New CTaxDocument
   Set m_TaxDocs = New Collection
   Set m_Companies = New Collection
   Set m_Employees = New Collection
   Set m_Agencies = New Collection
   Set m_Sources = New Collection
   Set m_Dests = New Collection
   Set m_Coll = New Collection
 
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub GridEX1_DblClick()
   Call cmdEdit_Click
End Sub

Private Sub GridEX1_RowFormat(RowBuffer As GridEX20.JSRowData)
   If TabStrip1.SelectedItem.Index = 5 Then
      RowBuffer.RowStyle = RowBuffer.Value(7)
   End If
End Sub

Private Sub GridEX1_UnboundReadData(ByVal RowIndex As Long, ByVal Bookmark As Variant, ByVal Values As GridEX20.JSRowData)
On Error GoTo ErrorHandler
Dim RealIndex As Long

   glbErrorLog.ModuleName = Me.Name
   glbErrorLog.RoutineName = "UnboundReadData"

   If TabStrip1.SelectedItem.Index = 1 Then
      If m_TaxDoc.DocumentItems Is Nothing Then
         Exit Sub
      End If

      If RowIndex <= 0 Then
         Exit Sub
      End If

      Dim CR As CTaxDocItem
      Set CR = GetItem(m_TaxDoc.DocumentItems, RowIndex, RealIndex)
      If CR Is Nothing Then
         Exit Sub
      End If

      Values(1) = CR.TAXDOC_ITEM_ID
      Values(2) = RealIndex
      Values(3) = CR.SUPPLIER_NAME
      Values(4) = DateToStringExtEx2(CR.PAY_DATE)
      Values(5) = CR.REVENUE_TYPE_NAME
      Values(6) = CR.RATETYPE_NAME
      Values(7) = FormatNumber(CR.PAID_AMOUNT)
      Values(8) = FormatNumber(CR.WH_AMOUNT)
      Values(9) = CR.CONDITION_NAME
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
      
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Private Sub TabStrip1_Click()
   If TabStrip1.SelectedItem.Index = 1 Then
   ElseIf TabStrip1.SelectedItem.Index = 2 Then
   ElseIf TabStrip1.SelectedItem.Index = 3 Then
   ElseIf TabStrip1.SelectedItem.Index = 4 Then
   ElseIf TabStrip1.SelectedItem.Index = 5 Then
   End If
End Sub

Private Sub txtBusinessDesc_Change()
   m_HasModify = True
End Sub

Private Sub txtDiscountPercent_Change()
   m_HasModify = True
End Sub

Private Sub txtEmail_Change()
   m_HasModify = True
End Sub

Private Sub txtName_Change()
   m_HasModify = True
End Sub

Private Sub txtSellBy_Change()
   m_HasModify = True
End Sub

Private Sub txtAdl_Change()
   m_HasModify = True
End Sub

Private Sub txtChd_Change()
   m_HasModify = True
End Sub

Private Sub txtCollectPrice_Change()
   m_HasModify = True
End Sub

Private Sub txtFoc_Change()
   m_HasModify = True
End Sub

Private Sub txtInf_Change()
   m_HasModify = True
End Sub

Private Sub txtIns_Change()
   m_HasModify = True
End Sub

Private Sub txtNote_Change()
   m_HasModify = True
End Sub

Private Sub txtReserve_Change()
   m_HasModify = True
End Sub

Private Sub txtRoomNo_Change()
   m_HasModify = True
End Sub

Private Sub txtSender_Change()
   m_HasModify = True
End Sub

Private Sub txtSlipBookNo_Change()
   m_HasModify = True
End Sub

Private Sub txtPassword_Change()
   m_HasModify = True
End Sub

Private Sub txtWebSite_Change()
   m_HasModify = True
End Sub

Private Sub uctlSetupDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlAgencyLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox1_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextLookup1_Change()
   m_HasModify = True
End Sub

Private Sub txtVocherNo_Change()
   m_HasModify = True
End Sub

Private Sub uctlBookingDate_HasChange()
   m_HasModify = True
End Sub

Private Sub uctlDestinationLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlGuestNameLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSaleByLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlSourceLookup_Change()
   m_HasModify = True
End Sub

Private Sub uctlTextBox2_Change()
   m_HasModify = True
End Sub

Private Sub uctlTravelDate_HasChange()
   m_HasModify = True
End Sub
