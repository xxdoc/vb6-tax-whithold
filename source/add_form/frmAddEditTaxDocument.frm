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
   lMenuChoosen = M.Popup("������쵢����Ũҡ Express", "-", "����������¡��")
   Set M = Nothing
   
   OKClick = False
   If (TabStrip1.SelectedItem.Index = 1) And (lMenuChoosen = 3) Then
      Set frmAddEditTaxDocItem.TempCollection = m_TaxDoc.DocumentItems
      frmAddEditTaxDocItem.ShowMode = SHOW_ADD
      frmAddEditTaxDocItem.HeaderText = MapText("������¡������")
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
      frmAddWHVoucher.HeaderText = MapText("���͡��¡�����ըҡ Express")
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
      frmAddEditTaxDocItem.HeaderText = MapText("�����¡������")
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
   lMenuChosen = oMenu.Popup("��§ҹ�����ѡ � ������ �.�.� 2", "-", "��§ҹ�ѭ�վ���� �����ѡ � ������  �.�.� 2", "-", "EXPORT ��� �.�.� 2 ���Ẻ Online")
ElseIf TAX_TYPE_NAME = 3 Then
    lMenuChosen = oMenu.Popup("��§ҹ�����ѡ � ������ �.�.� 3", "-", "��§ҹ�ѭ�վ���� �����ѡ � ������  �.�.� 3", "-", "EXPORT ��� �.�.� 3 ���Ẻ Online")
 ElseIf TAX_TYPE_NAME = 53 Then
   lMenuChosen = oMenu.Popup("��§ҹ�����ѡ � ������ �.�.� 53", "-", "��§ҹ�ѭ�վ���� �����ѡ � ������  �.�.� 53", "-", "EXPORT ��� �.�.� 53 ���Ẻ Online")
 ElseIf TAX_TYPE_NAME = 4 Then
 
 ElseIf TAX_TYPE_NAME = 21 Then
   lMenuChosen = oMenu.Popup("EXPORT ��� �.�.� 2� ���Ẻ Online")
End If
   
   Key = m_TaxDoc.TAX_DOCUMENT_ID '�觾���������
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
ElseIf TAX_TYPE_NAME = 2 And lMenuChosen = 5 Then 'Export �͡� Text File
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
      TempStr = Format(i, "00") & "|"                                                                                                           '�ӴѺ
      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                             'TAX ID ����ѡ
      TempStr = TempStr & CR.IDENT_ID & "|"                                                                             'TAX ID ���١�ѡ
      TempStr = TempStr & "" & "|"                                                                                                           '�ӹ�˹�Ҫ���
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                                                       '����
      TempStr = TempStr & "" & "|"                                                                                                           '���ʡ��
      TempStr = TempStr & "" & "|"                                                                                                           '�ѭ���Թ�ҡ
      TempStr = TempStr & CR.EXPORT_KEY & "|"                                                                                                           '�Թ�����ҵ��
      TempStr = TempStr & DateToStringToOnline1(CR.PAY_DATE) & "|"                                                                                                           '���ʡ��
      
      If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|"                                '�ѵ������
      End If
      TempStr = TempStr & FormatNumberToNull(CR.PAID_AMOUNT, , False) & "|"                                                                     '�ӹǹ�Թ������
      TempStr = TempStr & FormatNumberToNull(CR.WH_AMOUNT, , False) & "|"                                                                     '�ӹǹ�Թ����ѡ���
      TempStr = TempStr & CR.CONDITION_NAME & ""                                                                                                           '�Թ�����ҵ��
      
      Print #FileID, TempStr
   Next CR
   
   Close #FileID
   
   glbErrorLog.LocalErrorMsg = "COMPLETE " & i & " Record"
   glbErrorLog.ShowUserError
   
   Exit Sub
ElseIf TAX_TYPE_NAME = 21 And lMenuChosen = 1 Then '���2�
Dim FileName As String
   LocationSave = "C:\ExportOnline\Tax2a\"
   If Dir(LocationSave, vbDirectory) = "" Then
      MkDir (LocationSave)
   End If
   FileName = "PND2A_" & m_TaxDoc.TAX_ID & "_" & "000000" & "_" & Format(Year(uctlTravelDate.ShowDate) + 543, "0000") & "_00_00_00"
   LocationSave = LocationSave & FileName & ".txt"
   
   
Dim RegCode As String
RegCode = Trim(InputBox("��سҡ�͡����͹��ѵ�ŧ����¹������Ѻ�ҡ�����þҡ�", "��͹���", ""))
If Not Len(RegCode) = 20 Then
   glbErrorLog.LocalErrorMsg = "�ӹǹ����Ţ��ͧ�� 20 ��ѡ"
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
      TempStr = TempStr & "0000" & "|"                                                                              '���ʼ�����=2
      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                             'TAX ID ����ѡ=3
      TempStr = TempStr & "000000" & "|"                                                                              '�����ҢҼ��١�ѡ=4
       TempStr = TempStr & "1" & "|"                                                                                                           '��������ù���=5
       TempStr = TempStr & "PND2A" & "|"                                                                                                           '������Ẻ����=6
       TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                                                            'TAX_ID �����˹�ҷ���ѡ����=7
       TempStr = TempStr & "000000" & "|"                                                                              '�����Ң� ���١�ѡ=8
       TempStr = TempStr & "�ӹѡ�ҹ�˭�" & "|"                                                                                                           '����Ἱ�/��ǹ/����=9
       TempStr = TempStr & "1" & "|"                                                                                                           'ʶҹм���Сͺ�������˭�=10
       TempStr = TempStr & "00" & "|"                                                                                                           '��͹����=11
        TempStr = TempStr & Format(Year(uctlTravelDate.ShowDate) + 543, "0000") & "|"                                                                                                             '������=12
        TempStr = TempStr & "|"                                                                                                            '�������Ң�=13
      TempStr = TempStr & "00" & "|"                                                                                                            '������������Ẻ=14

      Dim SumPaidAmount As Double
      Dim SumWhAmount As Double
      Dim SumAddAmount As Double
      i = 0
      For Each CR In m_TaxDoc.DocumentItems
         i = i + 1
         SumPaidAmount = SumPaidAmount + CDbl(FormatNumberToNull(CR.PAID_AMOUNT, , False))
         SumWhAmount = SumWhAmount + CDbl(FormatNumberToNull(CR.WH_AMOUNT, , False))

      Next CR
      TempStr = TempStr & i & "|"                                                                                                            '����ӹǹ���=15
      TempStr = TempStr & FormatNumberToNull(SumPaidAmount, , False) & "|"                                                                                                             '����ӹǹ�Թ�������=16
      TempStr = TempStr & FormatNumberToNull(SumWhAmount, , False) & "|"                                                                                                             '����ӹǹ�Թ���շ����觷�����=17
      SumAddAmount = 0
      TempStr = TempStr & FormatNumber(SumAddAmount) & "|"                                                                                                           '�ӹǹ�Թ����=18
      TempStr = TempStr & FormatNumberToNull(SumWhAmount + SumAddAmount, , False) & "|"                                                                                                          '�ӹǹ�Թ����ʹ���չ��觷���������Թ����=19
      TempStr = TempStr & "0.00" & "|"                                                                                                            '�ӹǹ�Թ�͹��ҹ��Ҥ��=20
      TempStr = TempStr & RegCode & "|"                                                                                                             '����ŧ����¹=21
      TempStr = TempStr & "1"                                                                                                        '��ͧ�ҧ������Ẻ=22

'''        s = AToW(TempStr, CP_UTF8)
'''      Put #FileID, , s
     
     Print #FileID, TempStr
   '#######End Herder###########

   '#######Detail###########
   i = 0
   For Each CR In m_TaxDoc.DocumentItems
      i = i + 1
      TempStr = "D" & "|"                                                                                                           'Detail=1
      TempStr = TempStr & Format(i, "00") & "|"                                                                                                           '�ӴѺ��� =2
      TempStr = TempStr & "000000" & "|"                                                                                                           '�ҢҼ���ѡ���� =3
      TempStr = TempStr & CR.IDENT_ID & "|"                                                                             'TAX ID ���١�ѡ=4
      TempStr = TempStr & "0000000000" & "|"                                                                             'TAX ID ����ѡ=5
      TempStr = TempStr & "999999999" & "|"                                                                                                            '�Ţ���ѭ���Թ�ҡ=6
      TempStr = TempStr & "-" & "|"                                                                                                           '�ӹ�˹�Ҫ���=7
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                                                       '���ͼ�����Թ��=8
      TempStr = TempStr & "" & "|"                                                                                                           '���ʡ�ż�����Թ��=9
      TempStr = TempStr & "31122560" & "|"                                                                                                           '�ѹ��͹�շ�����=10
       If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Format(Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1), "0.00") & "|"                               '�ѵ������=11
      Else
         TempStr = TempStr & "0.00" & "|"                                '�ó������ �ѵ������=11
      End If
      TempStr = TempStr & FormatNumberToNull(CR.PAID_AMOUNT, , False) & "|"                                                                     '�ӹǹ�Թ������=12
      TempStr = TempStr & FormatNumberToNull(CR.WH_AMOUNT, , False) & "|"                                                                     '�ӹǹ�Թ���շ���ѡ���=13
      TempStr = TempStr & CR.EXPORT_KEY & "|"                                                                                                           '�Թ�����ҵ��=14
      TempStr = TempStr & "1" & "|"                                                                                                           '���͹䢡���ѡ����=15
      TempStr = TempStr & "" & "|"                                                                                                            '�������=16
      TempStr = TempStr & CR.HOME & "|"                                                                                                            '��ͧ�Ţ���=17
      TempStr = TempStr & "" & "|"                                                                                                            '��鹷��=18
      TempStr = TempStr & CR.VILLAGE & "|"                                                                                                             '�����ҹ=19
      TempStr = TempStr & "" & "|"                                                                                                            '�Ţ���=20
      TempStr = TempStr & CR.MOO & "|"                                                                                                             '������=21
      TempStr = TempStr & CR.SOI & "|"                                                                                                             '��͡/���=22
      TempStr = TempStr & CR.ROAD & "|"                                                                                                             '���=23
      TempStr = TempStr & CR.DISTRICT & "|"                                                                                                            '�Ӻ�/�ǧ=24
      TempStr = TempStr & CR.AMPHUR & "|"                                                                                                            '�����/ࢵ=25
      TempStr = TempStr & CR.PROVINCE & "|"                                                                                                             '�ѧ��Ѵ=26
      TempStr = TempStr & CR.ZIPCODE                                                                                                          '������ɳ���=27


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
      TempStr = Format(i, "00") & "|"                                                                     '�ӴѺ
      If Len(CR.IDENT_ID) > 0 Then
         TempStr = TempStr & CR.IDENT_ID & "|"                                                '�Ţ��Шӵ�ǻ�ЪҪ������������
      ElseIf Len(CR.TAX_ID) > 0 Then
         TempStr = TempStr & CR.TAX_ID & "|"                                                    '�Ţ��Шӵ�Ǽ����������
      End If
      TempStr = TempStr & "|"                                                                              '�ӹ�˹�Ҫ���
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                   '����
      TempStr = TempStr & "|"                                                                              '���ʡ��
      TempStr = TempStr & CR.PackAddress & "|"                                              '�������
      TempStr = TempStr & Format(CR.PAY_DATE, "DD/MM/YYYY") & "|"   '�ѹ��͹�շ������Թ��
      TempStr = TempStr & CR.REVENUE_NAME & "|"                                      '�������Թ��
       If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|" '�ѵ������
      End If
      TempStr = TempStr & CR.PAID_AMOUNT & "|"                                        '�ӹǹ�Թ������
      TempStr = TempStr & CR.WH_AMOUNT & "|"                                           '�ӹǹ�Թ����ѡ
      TempStr = TempStr & CR.CONDITION_NAME & "|"                                         '���͹䢡���ѡ
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
      TempStr = Format(i, "00") & "|"                                                                     '�ӴѺ
      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                          '�Ţ��Шӵ�Ǽ���������ռ���ѡ
      TempStr = TempStr & CR.TAX_ID & "|"                                                      '�Ţ��Шӵ�Ǽ���������ռ��١�ѡ
      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                   '����
      TempStr = TempStr & CR.PackAddress & "|"                                              '�������
      TempStr = TempStr & Format(CR.PAY_DATE, "DD/MM/YYYY") & "|"   '�ѹ��͹�շ������Թ��
      TempStr = TempStr & CR.REVENUE_NAME & "|"                                      '�������Թ��
       If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|" '�ѵ������
      End If
      TempStr = TempStr & CR.PAID_AMOUNT & "|"                                        '�ӹǹ�Թ������
      TempStr = TempStr & CR.WH_AMOUNT & "|"                                           '�ӹǹ�Թ����ѡ
      TempStr = TempStr & CR.CONDITION_NAME & "|"                                         '���͹䢡���ѡ
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
      frmReport.HeaderText = MapText("�������§ҹ")
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
''      TempStr = Format(i, "00") & "|"                                                                                                           '�ӴѺ
''      TempStr = TempStr & m_TaxDoc.TAX_ID & "|"                                                                             'TAX ID ����ѡ
''      TempStr = TempStr & CR.IDENT_ID & "|"                                                                             'TAX ID ���١�ѡ
''      TempStr = TempStr & "" & "|"                                                                                                           '�ӹ�˹�Ҫ���
''      TempStr = TempStr & CR.SUPPLIER_NAME & "|"                                                                       '����
''      TempStr = TempStr & "" & "|"                                                                                                           '���ʡ��
''      TempStr = TempStr & "" & "|"                                                                                                           '�ѭ���Թ�ҡ
''      TempStr = TempStr & CR.EXPORT_KEY & "|"                                                                                                           '�Թ�����ҵ��
''      TempStr = TempStr & DateToStringToOnline1(CR.PAY_DATE) & "|"                                                                                                           '���ʡ��
'
'      If InStr(1, CR.RATETYPE_NAME, "%", 1) <> 0 Then
'         TempStr = TempStr & Left(CR.RATETYPE_NAME, InStr(1, CR.RATETYPE_NAME, "%", 1) - 1) & "|"                                '�ѵ������
'      End If
'      TempStr = TempStr & FormatNumberToNull(CR.PAID_AMOUNT, , False) & "|"                                                                     '�ӹǹ�Թ������
'      TempStr = TempStr & FormatNumberToNull(CR.WH_AMOUNT, , False) & "|"                                                                     '�ӹǹ�Թ����ѡ���
'      TempStr = TempStr & CR.CONDITION_NAME & ""                                                                                                           '�Թ�����ҵ��
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
   Col.Caption = MapText("���ͼ�����Թ�� (�Ѿ���������)")
  
  If TaxType = 21 Then
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 0
      Col.Caption = MapText("")
  Else
      Set Col = GridEX1.Columns.Add '4
      Col.Width = 1935
      Col.Caption = MapText("�ѹ��͹�շ�����")
   End If

   Set Col = GridEX1.Columns.Add '5
   Col.Width = 2085
   Col.Caption = MapText("�������Թ��")

   Set Col = GridEX1.Columns.Add '6
   Col.Width = 2370
   Col.Caption = MapText("�ѵ������ %")
   Col.TextAlignment = jgexAlignRight
   
   Set Col = GridEX1.Columns.Add '7
   Col.Width = 1620
   Col.Caption = MapText("�ӹǹ�Թ������")
   Col.TextAlignment = jgexAlignRight

   Set Col = GridEX1.Columns.Add '8
   Col.Width = 1620
   Col.Caption = MapText("�ӹǹ�Թ����")
   Col.TextAlignment = jgexAlignRight

   Set Col = GridEX1.Columns.Add '9
   Col.Width = 1620
   Col.Caption = MapText("���͹�")
End Sub

Private Sub InitFormLayout()
   pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   SSFrame1.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)

If TaxType = 21 Then
   Me.Caption = HeaderText & "�"
Else
   Me.Caption = HeaderText & TAX_TYPE_NAME
End If
pnlHeader.Caption = Me.Caption

   
   Call InitNormalLabel(lblTravelDate, MapText("�ѹ������"))
   Call InitNormalLabel(lblGuestName, MapText("����ѷ"))
   Call InitNormalLabel(lblDestination, MapText("����Ѻ�Դ�ͺ"))

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
   
   Call InitMainButton(cmdExit, MapText("¡��ԡ (ESC)"))
   Call InitMainButton(cmdOK, MapText("��ŧ (F2)"))
   Call InitMainButton(cmdAdd, MapText("���� (F7)"))
   Call InitMainButton(cmdEdit, MapText("��� (F3)"))
   Call InitMainButton(cmdDelete, MapText("ź (F6)"))
   Call InitMainButton(cmdPrint, MapText("����� (F8)"))
   Call InitMainButton(cmdSave, MapText("�ѹ�֡ (F10)"))
   
   TabStrip1.Font.Bold = True
   TabStrip1.Font.Name = GLB_FONT
   TabStrip1.Font.Size = 16
   
   Call InitGrid1
   
   TabStrip1.Tabs.Clear
   TabStrip1.Tabs.Add().Caption = MapText("��������´")
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
