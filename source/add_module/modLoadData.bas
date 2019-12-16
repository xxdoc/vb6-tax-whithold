Attribute VB_Name = "modLoadData"
Option Explicit
' Test test test

Public Sub InitSupplierOrderBy(C As ComboBox)
   C.Clear

   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("รหัสผู้จำหน่าย")
   C.ItemData(1) = 1
   
   C.AddItem ("ชื่อผู้จำหน่าย")
   C.ItemData(2) = 2
End Sub

Public Sub LoadDBPath(C As ComboBox, Optional Cl As Collection = Nothing)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("D:\Express\Secure")
   C.ItemData(1) = 1
End Sub


Public Sub LoadUserGroup(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CUserGroup
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CUserGroup
Dim I As Long

   Set D = New CUserGroup
   Set Rs = New ADODB.Recordset
   
   D.GROUP_ID = -1
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CUserGroup
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.GROUP_NAME)
         C.ItemData(I) = TempData.GROUP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Str(TempData.GROUP_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitUserGroupOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(1) = 1
End Sub

Public Sub InitUserStatus(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ใช้งานได้")
   C.ItemData(1) = 1

   C.AddItem ("ถูกระงับ")
   C.ItemData(2) = 2
End Sub

Public Sub InitUserOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem ("ชื่อผู้ใช้")
   C.ItemData(1) = 1

   C.AddItem ("ชื่อกลุ่ม")
   C.ItemData(2) = 2
End Sub

Public Sub LoadAccessRight(C As ComboBox, Optional Cl As Collection = Nothing, Optional GroupID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CGroupRight
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGroupRight
Dim I As Long

   Set D = New CGroupRight
   Set Rs = New ADODB.Recordset
   
   D.GROUP_RIGHT_ID = -1
   D.GROUP_ID = GroupID
   Call D.QueryData3(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGroupRight
      Call TempData.PopulateFromRS3(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.RIGHT_ITEM_NAME)
         C.ItemData(I) = TempData.GROUP_RIGHT_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitLoginOrderby(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ล็อคอิน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อผู้ใช้"))
   C.ItemData(2) = 2
End Sub
Public Sub InitReportOR(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่ เอกสาร"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อซัพพลายเออร์"))
   C.ItemData(2) = 2

   C.AddItem (MapText("ชื่อบริษัท"))
   C.ItemData(3) = 3
End Sub
Public Sub InitTaxType(C As ComboBox)
C.Clear

   C.AddItem (MapText(""))
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รายการภาษีหัก ณ ที่จ่าย ภ.ง.ด 2"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รายการภาษีหัก ณ ที่จ่าย ภ.ง.ด. 3"))
   C.ItemData(2) = 2

C.AddItem (MapText("รายการภาษีหัก ณ ที่จ่าย ภ.ง.ด. 53"))
   C.ItemData(3) = 3

C.AddItem (MapText("รายการบัญชีพิเศษ ภาษีหัก ณ ที่จ่าย ภ.ง.ด 2"))
   C.ItemData(4) = 4

   C.AddItem (MapText("รายการบัญชีพิเศษ ภาษีหัก ณ ที่จ่าย ภ.ง.ด. 3"))
   C.ItemData(5) = 5

C.AddItem (MapText("รายการบัญชีพิเศษ ภาษีหัก ณ ที่จ่าย ภ.ง.ด. 53"))
   C.ItemData(6) = 6



End Sub

Public Sub InitThaiMonth(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("มกราคม"))
   C.ItemData(1) = 1

   C.AddItem (MapText("กุมภาพันธ์"))
   C.ItemData(2) = 2

C.AddItem (MapText("มีนาคม"))
   C.ItemData(3) = 3

C.AddItem (MapText("เมษายน"))
   C.ItemData(4) = 4

C.AddItem (MapText("พฤษภาคม"))
   C.ItemData(5) = 5

C.AddItem (MapText("มิถุนายน"))
   C.ItemData(6) = 6

C.AddItem (MapText("กรกฎาคม"))
   C.ItemData(7) = 7

C.AddItem (MapText("สิงหาคม"))
   C.ItemData(8) = 8

C.AddItem (MapText("กันยายน"))
   C.ItemData(9) = 9

C.AddItem (MapText("ตุลาคม"))
   C.ItemData(10) = 10

C.AddItem (MapText("พฤษศจิกายน"))
   C.ItemData(11) = 11

   C.AddItem (MapText(" ธันวาคม"))
   C.ItemData(12) = 12
End Sub

Public Sub LoadMaster(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterType As MASTER_TYPE)
On Error GoTo ErrorHandler
Dim D As CMasterRef
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CMasterRef
Dim I As Long

   Set D = New CMasterRef
   Set Rs = New ADODB.Recordset
   
   D.KEY_ID = -1
   D.MASTER_AREA = MasterType
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CMasterRef
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.KEY_NAME)
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadLender(C As ComboBox, Optional Cl As Collection = Nothing, Optional MasterType As MASTER_TYPE, Optional DOCUMENT_DATE As Date, Optional BRANCH_ID As Long)
On Error GoTo ErrorHandler
Dim D As CLender
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CLender
Dim I As Long

   Set D = New CLender
   Set Rs = New ADODB.Recordset

   D.LENDER_ID = -1
   D.BUDGET_YEAR = Year(DOCUMENT_DATE) + 543
   D.BRANCH_ID = BRANCH_ID
   Call D.QueryData(2, Rs, itemcount)

   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If

   While Not Rs.EOF
      I = I + 1
         Set TempData = New CLender
        Call TempData.PopulateFromRS(2, Rs)

      If Not (C Is Nothing) Then
         C.AddItem (TempData.LENDER_ITEMS_NAME)
         C.ItemData(I) = TempData.KEY_ID
      End If

      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If

      Set TempData = Nothing
      Rs.MoveNext
   Wend

   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadScComp(C As ComboBox, Optional Cl As Collection = Nothing)
On Error GoTo ErrorHandler
Dim D As CSCComp
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSCComp
Dim I As Long

   Set D = New CSCComp
   Set Rs = New ADODB.Recordset
   
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSCComp
      Call TempData.PopulateFromRS(1, Rs)
       TempData.KEY_ID = I
       
      If Not (C Is Nothing) Then
         C.AddItem (TempData.COMPNAM)
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Function LoadJournalAccount(C As ComboBox, Optional Cl As Collection = Nothing, Optional VoucherNo As String = "") As String
On Error GoTo ErrorHandler
Dim D As CGLJnl
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CGLJnl
Dim I As Long
Dim ShowAllAmount As String
   
   Set D = New CGLJnl
   Set Rs = New ADODB.Recordset
   
   D.VOUCHER = VoucherNo
   Call D.QueryData(1, Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CGLJnl
      Call TempData.PopulateFromRS(1, Rs)
      TempData.KEY_ID = I
       
      ShowAllAmount = ShowAllAmount & "-->" & TempData.AMOUNT
       
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ACCNAM)
         C.ItemData(I) = TempData.KEY_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.KEY_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   
   LoadJournalAccount = ShowAllAmount
   Exit Function
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Function

Public Sub LoadEmployee(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CEmployee
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEmployee
Dim I As Long
   Set D = New CEmployee
   Set Rs = New ADODB.Recordset
   
   D.EMP_ID = -1
   D.CURRENT_POSITION = ID
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEmployee
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.Name & " " & TempData.LASTNAME)
         C.ItemData(I) = TempData.EMP_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.EMP_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadSupplier(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CSupplier
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CSupplier
Dim I As Long

   Set D = New CSupplier
   Set Rs = New ADODB.Recordset
   
   D.SUPPLIER_ID = -1
   Call D.QueryData2(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CSupplier
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.SUPPLIER_NAME)
         C.ItemData(I) = TempData.SUPPLIER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.SUPPLIER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub LoadCustomer(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CCustomer
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CCustomer
Dim I As Long

   Set D = New CCustomer
   Set Rs = New ADODB.Recordset
   
   D.CUSTOMER_ID = -1
   Call D.QueryData1(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CCustomer
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.CUSTOMER_NAME)
         C.ItemData(I) = TempData.CUSTOMER_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.CUSTOMER_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub

Public Sub InitSlipBookingOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ใบจอง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่จอง"))
   C.ItemData(2) = 2

   C.AddItem (MapText("วันที่เดินทาง"))
   C.ItemData(3) = 3
End Sub

Public Sub InitTaxDocumentOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("วันที่นำส่ง"))
   C.ItemData(1) = 1

   C.AddItem (MapText("รหัสบริษัท"))
   C.ItemData(2) = 2

End Sub

Public Sub InitScheduleOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("เลขที่ใบรับ"))
   C.ItemData(1) = 1

   C.AddItem (MapText("วันที่เดินทาง"))
   C.ItemData(2) = 2
End Sub

Public Sub InitEnterpriseOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสบริษัท"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อบริษัท"))
   C.ItemData(2) = 2
End Sub

Public Sub InitCustomerOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสลูกค้า"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อลูกค้า"))
   C.ItemData(2) = 2
End Sub

Public Sub InitEmployeeOrderBy(C As ComboBox)
   C.Clear
   
   C.AddItem ("")
   C.ItemData(0) = 0
   
   C.AddItem (MapText("รหัสพนักงาน"))
   C.ItemData(1) = 1

   C.AddItem (MapText("ชื่อ"))
   C.ItemData(2) = 2

   C.AddItem (MapText("นามสกุล"))
   C.ItemData(3) = 3

   C.AddItem (MapText("ตำแหน่ง"))
   C.ItemData(4) = 4
End Sub

Public Sub LoadSupplierAddress(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long

   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   D.SUPPLIER_ID = SupplierID
   Call D.QueryData4(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst And Not (C Is Nothing) Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
     Call Cl.Add(TempData)
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub

Public Sub loadEnterprise(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CEnterprise
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEnterprise
Dim I As Long

   Set D = New CEnterprise
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEnterprise
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.ENTERPRISE_NAME)
         C.ItemData(I) = TempData.ENTERPRISE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.ENTERPRISE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadEnterpriseShortName(C As ComboBox, Optional Cl As Collection = Nothing, Optional ID As Long = -1)
On Error GoTo ErrorHandler
Dim D As CEnterprise
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CEnterprise
Dim I As Long

   Set D = New CEnterprise
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   Call D.QueryData(Rs, itemcount)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      Set TempData = New CEnterprise
      Call TempData.PopulateFromRS(1, Rs)
   
      If Not (C Is Nothing) Then
         C.AddItem ("(" & TempData.SHORT_NAME & ") " & TempData.ENTERPRISE_NAME)
         C.ItemData(I) = TempData.ENTERPRISE_ID
      End If
      
      If Not (Cl Is Nothing) Then
         Call Cl.Add(TempData, Trim(Str(TempData.ENTERPRISE_ID)))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
End Sub
Public Sub LoadSupAddr(C As ComboBox, Optional Cl As Collection = Nothing, Optional SupplierID As Long = -1, Optional ShowFirst As Boolean = True)
On Error GoTo ErrorHandler
Dim D As CAddress
Dim itemcount As Long
Dim Rs As ADODB.Recordset
Dim TempData As CAddress
Dim I As Long

   Set D = New CAddress
   Set Rs = New ADODB.Recordset
   
   D.ENTERPRISE_ID = -1
   D.SUPPLIER_ID = SupplierID
   Call D.QueryData4(Rs, itemcount, 2)
   
   If Not (C Is Nothing) Then
      C.Clear
      I = 0
      C.AddItem ("")
   End If

   If Not (Cl Is Nothing) Then
      Set Cl = Nothing
      Set Cl = New Collection
   End If
   While Not Rs.EOF
      I = I + 1
      
      Set TempData = New CAddress
      Call TempData.PopulateFromRS(Rs, 2)
   
      If Not (C Is Nothing) Then
         C.AddItem (TempData.PackAddress)
         C.ItemData(I) = TempData.ADDRESS_ID
      End If
      If (I > 0) And ShowFirst And Not (C Is Nothing) Then
         C.ListIndex = 1
      End If
      
      If Not (Cl Is Nothing) Then
         
'    If TempData.SUPPLIER_ID = 97 Then
'         'Debug.Print TempData.SUPPLIER_ID
'  End If
     Call Cl.Add(TempData, Trim(TempData.SUPPLIER_ID & "-" & TempData.ADDRESS_ID))
      End If
      
      Set TempData = Nothing
      Rs.MoveNext
   Wend
   
   Set Rs = Nothing
   Set D = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub

Public Sub getTaxIDfromBranch(Optional C As ComboBox = Nothing, Optional Cl As Collection = Nothing, Optional BRANCH As String = "", Optional ID As String)
   Dim EnterPrise As CEnterprise
   Dim itemcount As Long
   Dim Rs As ADODB.Recordset

   Set EnterPrise = New CEnterprise
    Set Rs = New ADODB.Recordset
    
   EnterPrise.SHORT_NAME = BRANCH
   Call EnterPrise.QueryData(Rs, itemcount)
   Call EnterPrise.PopulateFromRS(1, Rs)
   
    ID = EnterPrise.TAX_ID
    
  Set Rs = Nothing
   Set EnterPrise = Nothing
   Exit Sub
   
ErrorHandler:
   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)

End Sub
