VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_ModuleName As String
Private m_ADOConn As ADODB.Connection
Private m_Rs1 As ADODB.Recordset
Private m_Rs2 As ADODB.Recordset
Private m_Rs3 As ADODB.Recordset
'Private TAXRE_TYPE As TAXWH_TYPE

Private Sub Class_Initialize()
   Set m_ADOConn = glbDatabaseMngr.DBConnection
   m_ModuleName = "clsDaily"
   Set m_Rs1 = New ADODB.Recordset
   Set m_Rs2 = New ADODB.Recordset
   Set m_Rs3 = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
   If m_Rs1.State = adStateOpen Then
      m_Rs1.Close
   End If
   Set m_Rs1 = Nothing

   If m_Rs2.State = adStateOpen Then
      m_Rs2.Close
   End If
   Set m_Rs2 = Nothing

   If m_Rs3.State = adStateOpen Then
      m_Rs3.Close
   End If
   Set m_Rs3 = Nothing
End Sub

Public Sub StartTransaction()
   m_ADOConn.BeginTrans
End Sub

Public Sub CommitTransaction()
   m_ADOConn.CommitTrans
End Sub

Public Sub RollbackTransaction()
   m_ADOConn.RollbackTrans
End Sub

Public Function QueryReportConfig(Ua As CReportConfig, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryReportConfig"
   QueryReportConfig = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)
   
   QueryReportConfig = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryReportConfig = False
End Function

Public Function QueryGlJnl(Ua As CGLJnl, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryGlJnl"
   QueryGlJnl = False

   IsOK = True
   Call Ua.QueryData(3, Rs, itemcount)
   
   QueryGlJnl = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryGlJnl = False
End Function
Public Function QueryClenItems(Ua As CLender_Items, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog, Optional dbType As Integer) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryClenItems"
   QueryClenItems = False

   If dbType = 1 Then
      Call Ua.QueryData(3, Rs, itemcount)
   Else
      Call Ua.QueryData2(2, Rs, itemcount)
   End If
    IsOK = True
   QueryClenItems = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
    IsOK = False
   QueryClenItems = False
End Function

Public Function AddEditReportConfig(Ug As CReportConfig, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditReportConfig"
   AddEditReportConfig = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditReportConfig = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditReportConfig = False
End Function

'==
Public Function QueryEnterprise(Ua As CEnterprise, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryEnterprise"
   QueryEnterprise = False
   
   IsOK = True
   
   Call Ua.QueryData(Rs, itemcount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Contact ++++++++++++++++++++++++++
      Dim cCstContact As CEnterprisePerson
      Set cCstContact = New CEnterprisePerson
      cCstContact.ENTERPRISE_PERSON_ID = -1
      cCstContact.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cCstContact.OrderBy = 1
      Call cCstContact.QueryData(m_Rs1, iCount)
      Set cCstContact = Nothing

      Set Ua.EnpPersons = Nothing
      Set Ua.EnpPersons = New Collection
      While Not m_Rs1.EOF
         Set cCstContact = New CEnterprisePerson
         Call cCstContact.PopulateFromRS(m_Rs1)

         Dim N As cName
         Set N = New cName
         N.NAME_ID = cCstContact.NAME_ID
         Call N.QueryData(m_Rs2, iCount)
         Set N = Nothing
         While Not m_Rs2.EOF
            Set N = New cName
            Call N.PopulateFromRS(m_Rs2)
            N.Flag = "I"
            Set cCstContact.Name = N
            Set N = Nothing

            m_Rs2.MoveNext
         Wend

         If cCstContact.ADDRESS_ID > 0 Then
            Dim A As CAddress
            Set A = New CAddress
            A.ADDRESS_ID = cCstContact.ADDRESS_ID
            Call A.QueryData(m_Rs2, iCount)
            Set A = Nothing
            If Not m_Rs2.EOF Then
               Set A = New CAddress
               Call A.PopulateFromRS(m_Rs2)
               A.Flag = "I"
               Set cCstContact.Address = A
               Set A = Nothing
            End If
         End If

         cCstContact.Flag = "I"
         Call Ua.EnpPersons.Add(cCstContact)
         Set cCstContact = Nothing
         m_Rs1.MoveNext
      Wend
'      Coontact ++++++++++++++++++++++++++
         
      'Name ++++++++++++++++++++++++++
      Dim cEnpname As CEnterpriseName
      Set cEnpname = New CEnterpriseName
      cEnpname.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cEnpname.OrderBy = 1
      Call cEnpname.QueryData(m_Rs1, iCount)
      Set cEnpname = Nothing
      
      Set Ua.EnpNames = Nothing
      Set Ua.EnpNames = New Collection
      While Not m_Rs1.EOF
         Set cEnpname = New CEnterpriseName
         Call cEnpname.PopulateFromRS(m_Rs1)

         Dim cName As cName
         Dim iCount2 As Long
         Set cName = New cName
         cName.NAME_ID = cEnpname.NAME_ID
         Call cName.QueryData(m_Rs2, iCount2)
         Set cName = Nothing
         While Not m_Rs2.EOF
            Set cName = New cName
            Call cName.PopulateFromRS(m_Rs2)
            cName.Flag = "I"
            Call cEnpname.Names.Add(cName)
            Set cName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpname.Flag = "I"
         Call Ua.EnpNames.Add(cEnpname)
         Set cEnpname = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
      
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CEnterpriseAddress
      Set cEnpAddr = New CEnterpriseAddress
      cEnpAddr.ENTERPRISE_ID = Ua.ENTERPRISE_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.EnpAddresses = Nothing
      Set Ua.EnpAddresses = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CEnterpriseAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            Call cEnpAddr.Addresses.Add(cAddr)
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.EnpAddresses.Add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If

   QueryEnterprise = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryEnterprise = False
End Function

Public Function AddEditEnterprise(Ug As CEnterprise, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim Name As cName
Dim Addr1 As CAddress

   RName = "AddEditEnterprise"
   AddEditEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
         
   Dim EnpName As CEnterpriseName
   Dim EmpAddr As CEnterpriseAddress
   Dim Addr As CAddress
   
   For Each EmpAddr In Ug.EnpAddresses
      If EmpAddr.Flag = "A" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData
                  
         EmpAddr.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EmpAddr.ADDRESS_ID = Addr.ADDRESS_ID
         EmpAddr.AddEditMode = SHOW_ADD
         EmpAddr.AddEditData
      ElseIf EmpAddr.Flag = "E" Then
         Set Addr = EmpAddr.Addresses(1)
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData
         
      ElseIf EmpAddr.Flag = "D" Then
         Call EmpAddr.DeleteData
         
         Set Addr = EmpAddr.Addresses(1)
         Call Addr.DeleteData
      End If
   Next EmpAddr
   
   For Each EnpName In Ug.EnpNames
      If EnpName.Flag = "A" Then
         Set Name = EnpName.Names(1)
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         EnpName.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EnpName.NAME_ID = Name.NAME_ID
         EnpName.AddEditMode = SHOW_ADD
         EnpName.AddEditData
      ElseIf EnpName.Flag = "E" Then
         Set Name = EnpName.Names(1)
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next EnpName
   
   Dim EnpPerson As CEnterprisePerson
   For Each EnpPerson In Ug.EnpPersons
      If EnpPerson.Flag = "A" Then
         Set Name = EnpPerson.Name
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         EnpPerson.ENTERPRISE_ID = Ug.ENTERPRISE_ID
         EnpPerson.NAME_ID = Name.NAME_ID
         EnpPerson.AddEditMode = SHOW_ADD
         EnpPerson.AddEditData
      ElseIf EnpPerson.Flag = "E" Then
         Set Name = EnpPerson.Name
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      
         EnpPerson.AddEditMode = SHOW_EDIT
         EnpPerson.AddEditData
      ElseIf EnpPerson.Flag = "D" Then
         Call EnpPerson.DeleteData

         Set Name = EnpPerson.Name
         Call Name.DeleteData
      End If
   Next EnpPerson
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEnterprise = False
End Function

Public Function QueryCustomerInfo(Ua As CCustomer, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCustomerInfo"
   QueryCustomerInfo = False

   IsOK = True
   Call Ua.QueryData1(Rs, itemcount)
         
   QueryCustomerInfo = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomerInfo = False
End Function

'==
Public Function QueryCustomer(Ua As CCustomer, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QueryCustomer"
   QueryCustomer = False
   
   IsOK = True
   Call Ua.QueryData2(Rs, itemcount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CCustomerName
      Set cCstName = New CCustomerName
      cCstName.CUSTOMER_NAME_ID = -1
      cCstName.CUSTOMER_ID = Ua.CUSTOMER_ID
      cCstName.OrderBy = 1
      Call cCstName.QueryData(m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CCustomerName
         Call cCstName.PopulateFromRS(m_Rs1)

         Dim cName As cName
         Dim iCount2 As Long
         Set cName = New cName
         cName.NAME_ID = cCstName.NAME_ID
         Call cName.QueryData(m_Rs2, iCount2)
         Set cName = Nothing
         While Not m_Rs2.EOF
            Set cName = New cName
            Call cName.PopulateFromRS(m_Rs2)
            cName.Flag = "I"
            Set cCstName.Name = cName
            Set cName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.Add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CCustomerAddress
      Set cEnpAddr = New CCustomerAddress
      cEnpAddr.CUSTOMER_ID = Ua.CUSTOMER_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddr = Nothing
      Set Ua.CstAddr = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CCustomerAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            cAddr.ADDRESS_TYPE = cEnpAddr.ADDRESS_TYPE
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddr.Add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   
      'Account ++++++++++++++++++++++++++
      Dim cAcc As CAccount
      Set cAcc = New CAccount
      cAcc.ACCOUNT_ID = -1
      cAcc.CUSTOMER_ID = Ua.CUSTOMER_ID
      cAcc.ENABLE_FLAG = "Y"
      cAcc.OrderBy = 1
      Call cAcc.QueryData(2, m_Rs1, iCount)
      Set cAcc = Nothing
      
      Dim cSub As CSubscriber
'      Dim cAgr As CAgreement
      
      Set Ua.CstAccounts = Nothing
      Set Ua.CstAccounts = New Collection
      While Not m_Rs1.EOF
         Set cAcc = New CAccount
         Call cAcc.PopulateFromRS(1, m_Rs1)
      
         cAcc.Flag = "I"
         Call Ua.CstAccounts.Add(cAcc)
         
         Set cSub = New CSubscriber
         cSub.ACCOUNT_ID = cAcc.ACCOUNT_ID
         cSub.SUBSCRIBER_ID = -1
         Call cSub.QueryData(m_Rs2, iCount)
         Set cSub = Nothing
         Set cAcc.ActSubs = Nothing
         Set cAcc.ActSubs = New Collection
         While Not m_Rs2.EOF
            Set cSub = New CSubscriber
            Call cSub.PopulateFromRS(m_Rs2)
            cSub.Flag = "I"
            
            Call cAcc.ActSubs.Add(cSub)
            Set cSub = Nothing
            m_Rs2.MoveNext
         Wend
         
'         Set cAgr = New CAgreement
'         cAgr.ACCOUNT_ID = cAcc.ACCOUNT_ID
'         cAgr.SUBSCRIBER_ID = -1
'         cAgr.AGREEMENT_ID = -1
'         Call cAgr.QueryData(m_Rs2, iCount)
'         Set cAgr = Nothing
'         Set cAcc.ActAgrmnts = Nothing
'         Set cAcc.ActAgrmnts = New Collection
'         While Not m_Rs2.EOF
'            Set cAgr = New CAgreement
'            Call cAgr.PopulateFromRS(m_Rs2)
'            cAgr.Flag = "I"
'
'            Call cAcc.ActAgrmnts.Add(cAgr)
'            Set cAgr = Nothing
'            m_Rs2.MoveNext
'         Wend
         
         Set cAcc = Nothing
         m_Rs1.MoveNext
      Wend
      'Account ++++++++++++++++++++++++++
   End If
      
   QueryCustomer = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QueryCustomer = False
End Function

Public Function DeleteCustomer(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CCustomer

   RName = "DeleteCustomer"
   DeleteCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CCustomer
   
   cCust.CUSTOMER_ID = UID
   Call cCust.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteCustomer = True
End Function

Public Function AddEditCustomer(Ug As CCustomer, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditCustomer"
   AddEditCustomer = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
      
   Dim CstName As CCustomerName
   Dim Name As cName
   Dim CstAddr As CCustomerAddress
   Dim Addr As CAddress
   
   For Each CstAddr In Ug.CstAddr
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         CstName.CUSTOMER_ID = Ug.CUSTOMER_ID
         CstName.NAME_ID = Name.NAME_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next CstName
   
   Dim cAcc As CAccount
   Dim cSubc As CSubscriber
'   Dim cAgr As CAgreement
   For Each cAcc In Ug.CstAccounts
      If cAcc.Flag = "A" Then
         cAcc.AddEditMode = SHOW_ADD
         cAcc.CUSTOMER_ID = Ug.CUSTOMER_ID
         Call cAcc.AddEditData
         
         Set cSubc = cAcc.ActSubs(1)
         cSubc.ACCOUNT_ID = cAcc.ACCOUNT_ID
         cSubc.AddEditMode = SHOW_ADD
         Call cSubc.AddEditData
      
'         Set cAgr = cAcc.ActAgrmnts(1)
'         cAgr.SUBSCRIBER_ID = cSubc.SUBSCRIBER_ID
'         cAgr.AddEditMode = SHOW_ADD
'         Call cAgr.AddEditData
      ElseIf cAcc.Flag = "E" Then
         cAcc.AddEditMode = SHOW_EDIT
         Call cAcc.AddEditData
         
         Set cSubc = cAcc.ActSubs(1)
         cSubc.AddEditMode = SHOW_EDIT
         Call cSubc.AddEditData
      
'         Set cAgr = cAcc.ActAgrmnts(1)
'         cAgr.AddEditMode = SHOW_EDIT
'         Call cAgr.AddEditData
      ElseIf cAcc.Flag = "D" Then
         cAcc.AddEditMode = SHOW_EDIT
         cAcc.ENABLE_FLAG = "N"
         Call cAcc.AddEditData
         
         Set cSubc = cAcc.ActSubs(1)
         cSubc.AddEditMode = SHOW_EDIT
         Call cSubc.AddEditData
      End If
   Next cAcc
         
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditCustomer = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditCustomer = False
End Function

'==
Public Function QuerySupplier(Ua As CSupplier, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long

   RName = "QuerySupplier"
   QuerySupplier = False
   
   IsOK = True
   Call Ua.QueryData2(Rs, itemcount)
   
   Dim bIsOk As Boolean
   Dim iCount As Long
   
   If Ua.QueryFlag = 1 Then
      Dim cCstContact As CSupplierContact
      Set cCstContact = New CSupplierContact
      cCstContact.SUPPLIER_CONTACT_ID = -1
      cCstContact.SUPPLIER_ID = Ua.SUPPLIER_ID
      cCstContact.OrderBy = 1
      Call cCstContact.QueryData(m_Rs1, iCount)
      Set cCstContact = Nothing
      
      Set Ua.CstContacts = Nothing
      Set Ua.CstContacts = New Collection
      While Not m_Rs1.EOF
         Set cCstContact = New CSupplierContact
         Call cCstContact.PopulateFromRS(m_Rs1)

         Dim N As cName
         Set N = New cName
         N.NAME_ID = cCstContact.NAME_ID
         Call N.QueryData(m_Rs2, iCount)
         Set N = Nothing
         While Not m_Rs2.EOF
            Set N = New cName
            Call N.PopulateFromRS(m_Rs2)
            N.Flag = "I"
            Set cCstContact.Name = N
            Set N = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstContact.Flag = "I"
         Call Ua.CstContacts.Add(cCstContact)
         Set cCstContact = Nothing
         m_Rs1.MoveNext
      Wend
      'Coontact ++++++++++++++++++++++++++
   
      'Name ++++++++++++++++++++++++++
      Dim cCstName As CSupplierName
      Set cCstName = New CSupplierName
      cCstName.SUPPLIER_NAME_ID = -1
      cCstName.SUPPLIER_ID = Ua.SUPPLIER_ID
      cCstName.OrderBy = 1
      Call cCstName.QueryData(m_Rs1, iCount)
      Set cCstName = Nothing
      
      Set Ua.CstNames = Nothing
      Set Ua.CstNames = New Collection
      While Not m_Rs1.EOF
         Set cCstName = New CSupplierName
         Call cCstName.PopulateFromRS(m_Rs1)

         Dim cName As cName
         Dim iCount2 As Long
         Set cName = New cName
         cName.NAME_ID = cCstName.NAME_ID
         Call cName.QueryData(m_Rs2, iCount2)
         Set cName = Nothing
         While Not m_Rs2.EOF
            Set cName = New cName
            Call cName.PopulateFromRS(m_Rs2)
            cName.Flag = "I"
            Set cCstName.Name = cName
            Set cName = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cCstName.Flag = "I"
         Call Ua.CstNames.Add(cCstName)
         Set cCstName = Nothing
         m_Rs1.MoveNext
      Wend
      'Name ++++++++++++++++++++++++++
            
      'Address ++++++++++++++++++++++++++
      Dim cEnpAddr As CSupplierAddress
      Set cEnpAddr = New CSupplierAddress
      cEnpAddr.SUPPLIER_ID = Ua.SUPPLIER_ID
      cEnpAddr.OrderBy = 1
      Call cEnpAddr.QueryData(m_Rs1, iCount)
      Set cEnpAddr = Nothing
      
      Set Ua.CstAddr = Nothing
      Set Ua.CstAddr = New Collection
      While Not m_Rs1.EOF
         Set cEnpAddr = New CSupplierAddress
         Call cEnpAddr.PopulateFromRS(m_Rs1)

         Dim cAddr As CAddress
         Dim iCount3 As Long
         Set cAddr = New CAddress
         cAddr.ADDRESS_ID = cEnpAddr.ADDRESS_ID
         Call cAddr.QueryData(m_Rs2, iCount3)
         Set cAddr = Nothing
         While Not m_Rs2.EOF
            Set cAddr = New CAddress
            Call cAddr.PopulateFromRS(m_Rs2)
            cAddr.Flag = "I"
            cAddr.ADDRESS_TYPE = cEnpAddr.ADDRESS_TYPE
            Set cEnpAddr.Addresses = cAddr
            Set cAddr = Nothing
            
            m_Rs2.MoveNext
         Wend
      
         cEnpAddr.Flag = "I"
         Call Ua.CstAddr.Add(cEnpAddr)
         Set cEnpAddr = Nothing
         m_Rs1.MoveNext
      Wend
      'Address ++++++++++++++++++++++++++
   End If
   
   QuerySupplier = True
   Exit Function
   
ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)
   
   QuerySupplier = False
End Function

Public Function AddEditSupplier(Ug As CSupplier, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditSupplier"
   AddEditSupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
   
   Call Ug.AddEditData
   
   Dim CstName As CSupplierName
   Dim Name As cName
   Dim CstAddr As CSupplierAddress
   Dim Addr As CAddress
   
   
   For Each CstAddr In Ug.CstAddr
      If CstAddr.Flag = "A" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_ADD
         Call Addr.AddEditData

         CstAddr.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditMode = SHOW_ADD
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "E" Then
         Set Addr = CstAddr.Addresses
         Addr.AddEditMode = SHOW_EDIT
         Call Addr.AddEditData

         CstAddr.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstAddr.ADDRESS_ID = Addr.ADDRESS_ID
         CstAddr.AddEditMode = SHOW_EDIT
         CstAddr.ADDRESS_TYPE = Addr.ADDRESS_TYPE
         CstAddr.AddEditData
      ElseIf CstAddr.Flag = "D" Then
         Call CstAddr.DeleteData

         Set Addr = CstAddr.Addresses
         Call Addr.DeleteData
      End If
   Next CstAddr
   
   For Each CstName In Ug.CstNames
      If CstName.Flag = "A" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         CstName.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstName.NAME_ID = Name.NAME_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         Set Name = CstName.Name
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      End If
   Next CstName
   
   Dim CstContact As CSupplierContact
   For Each CstContact In Ug.CstContacts
      If CstContact.Flag = "A" Then
         Set Name = CstContact.Name
         Name.AddEditMode = SHOW_ADD
         Call Name.AddEditData
                  
         CstContact.SUPPLIER_ID = Ug.SUPPLIER_ID
         CstContact.NAME_ID = Name.NAME_ID
         CstContact.AddEditMode = SHOW_ADD
         CstContact.AddEditData
      ElseIf CstContact.Flag = "E" Then
         Set Name = CstContact.Name
         Name.AddEditMode = SHOW_EDIT
         Call Name.AddEditData
      
         CstContact.AddEditMode = SHOW_EDIT
         CstContact.AddEditData
      ElseIf CstContact.Flag = "D" Then
         Call CstContact.DeleteData

         Set Name = CstContact.Name
         Call Name.DeleteData
      End If
   Next CstContact
            
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditSupplier = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditSupplier = False
End Function

Public Function DeleteSupplier(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CSupplier

   RName = "DeleteSupplier"
   DeleteSupplier = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CSupplier
   
   cCust.SUPPLIER_ID = UID
   Call cCust.DeleteData
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteSupplier = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)
   
   IsOK = False

   DeleteSupplier = True
End Function

'==
Public Function QueryEmployee(Ua As CEmployee, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

Dim Qp As CEmployee

   RName = "QueryEmployee"
   QueryEmployee = False

   IsOK = True
   Call Ua.QueryData(Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Set Ua.EmpName = Nothing
      Set Ua.EmpName = New CEmployeeName
      Ua.EmpName.EMPLOYEE_NAME_ID = -1
      Ua.EmpName.EMP_ID = Ua.EMP_ID
      Call Ua.EmpName.QueryData(m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EmpName.PopulateFromRS(m_Rs1)
      End If
      
      Set Ua.EName = Nothing
      Set Ua.EName = New cName
      Ua.EName.NAME_ID = Ua.EmpName.NAME_ID
      Call Ua.EName.QueryData(m_Rs1, iCount)
      If Not m_Rs1.EOF Then
         Call Ua.EName.PopulateFromRS(m_Rs1)
      End If
   End If
   
   QueryEmployee = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryEmployee = False
End Function

Public Function DeleteEmployee(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CEmployee

   RName = "DeleteEmployee"
   DeleteEmployee = False
   HasBegin = False

   Set B = New CEmployee

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.EMP_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteEmployee = True
End Function

Public Function AddEditEmployee(Ug As CEmployee, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditEmployee"
   AddEditEmployee = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Call Ug.EName.AddEditData
   
   Ug.EmpName.EMP_ID = Ug.EMP_ID
   Ug.EmpName.NAME_ID = Ug.EName.NAME_ID
   Call Ug.EmpName.AddEditData
      
   '   Dim CstName As CCustomerName
   'Dim Name As cName
'   Dim CstAddr As CEmpAddress
   Dim Addr As CAddress
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditEmployee = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditEmployee = False
End Function

Public Function DeleteMasterRef(UID As Long, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cData As CMasterRef

   RName = "DeleteMasterRef"
   DeleteMasterRef = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Set cData = New CMasterRef
   cData.KEY_ID = UID
   Call cData.DeleteData
   Set cData = Nothing
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   DeleteMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   Set cData = Nothing

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   IsOK = False

   DeleteMasterRef = True
End Function

Public Function AddEditMasterRef(At As CMasterRef, IsOK As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditMasterRef"
   AddEditMasterRef = False
   HasBegin = False

   m_ADOConn.BeginTrans
   HasBegin = True

   Call At.AddEditData
   
   Call m_ADOConn.CommitTrans
   HasBegin = False
   IsOK = True

   AddEditMasterRef = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      m_ADOConn.RollbackTrans
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditMasterRef = False
End Function

Public Function QueryTaxDocument(Ua As CTaxDocument, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim TaxReType As TAXWH_TYPE

   RName = "QueryTaxDocument"
   QueryTaxDocument = False

   IsOK = True

    Call Ua.QueryData(1, Rs, itemcount)


   If Ua.QueryFlag = 1 Then
      Dim Yw As CTaxDocItem

      Set Yw = New CTaxDocItem
      Yw.TAXDOC_ITEM_ID = -1
      Yw.TAX_DOCUMENT_ID = Ua.TAX_DOCUMENT_ID
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing

      Set Ua.DocumentItems = Nothing
      Set Ua.DocumentItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CTaxDocItem
         Call Yw.PopulateFromRS(1, m_Rs1)

         Yw.Flag = "I"
         Call Ua.DocumentItems.Add(Yw)

         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryTaxDocument = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTaxDocument = False
End Function
Public Function QueryTax_A_Document(Ua As CTaxDocument, Cl As Collection, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim TaxReType As TAXWH_TYPE
Dim Rs As ADODB.Recordset
Dim OldTaxType As Long

   RName = "QueryTax_A_Document"
   QueryTax_A_Document = False
   
   Set Rs = New ADODB.Recordset

   IsOK = True
   
   Dim TD As CTaxDocument
   Dim T_TD As CTaxDocument
   Dim str_Year As String
   Dim Key As String
   OldTaxType = Ua.TAX_TYPE
   Ua.TAX_TYPE = 2
   Call Ua.QueryData(2, Rs, itemcount)
   
     If Ua.QueryFlag = 0 Then
     Set Cl = New Collection
      While Not Rs.EOF
         Set TD = New CTaxDocument
         
         Call TD.PopulateFromRS(2, Rs)
         str_Year = Right(Year(TD.DOCUMENT_DATE) + 543, 4)
         Key = Trim(Str(TD.COMPANY_ID)) & "-" & Trim(str_Year)
         Set T_TD = GetObject("CTaxDocument", Cl, Key, True)
         If (T_TD Is Nothing) Then
            TD.Document_Year = str_Year
            Call Cl.Add(TD, Key)
         End If

         Set TD = Nothing
         Rs.MoveNext
      Wend
   End If
   
   If Ua.QueryFlag = 1 Then
   
   Call Ua.PopulateFromRS(2, Rs)
   
      Dim Yw As CTaxDocItem

      Set Yw = New CTaxDocItem
      Yw.TAXDOC_ITEM_ID = -1
      Yw.TAX_TYPE = 2
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Yw.Document_Year = Trim(Str(Val(Ua.Document_Year) - 543))
      Yw.COMPANY_ID = Ua.COMPANY_ID
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing
      
      Dim T_TDI As CTaxDocItem
      Set Ua.DocumentItems = Nothing
      Set Ua.DocumentItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CTaxDocItem
         Call Yw.PopulateFromRS(1, m_Rs1)

          Key = Trim(Yw.SUPPLIER_CODE)
         Set T_TDI = GetObject("CTaxDocItem", Ua.DocumentItems, Key, False)
         If Not (T_TDI Is Nothing) Then
'             T_TDI.PAID_AMOUNT = T_TDI.PAID_AMOUNT + Yw.PAID_AMOUNT
'             T_TDI.WH_AMOUNT = T_TDI.WH_AMOUNT + Yw.WH_AMOUNT
             T_TDI.PAID_AMOUNT = T_TDI.PAID_AMOUNT + CDbl(FormatNumberToNull(Yw.PAID_AMOUNT, , False))
             T_TDI.WH_AMOUNT = T_TDI.WH_AMOUNT + CDbl(FormatNumberToNull(Yw.WH_AMOUNT, , False))
         Else
            Yw.Flag = "I"
            Call Ua.DocumentItems.Add(Yw, Key)
         End If
         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryTax_A_Document = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTax_A_Document = False
End Function
Public Function QueryTaxDocument2(Ua As CTaxDocSP, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long

   RName = "QueryTaxDocument2"
   QueryTaxDocument2 = False

   IsOK = True
   Call Ua.QueryData(1, Rs, itemcount)

   If Ua.QueryFlag = 1 Then
      Dim Yw As CTaxDocItem

      Set Yw = New CTaxDocItem
      Yw.TAXDOC_ITEM_ID = -1
      Yw.TAX_DOCUMENT_ID = 0
      Yw.FROM_DATE = 1
      Yw.TO_DATE = 1
      Yw.OrderBy = -1
      Yw.OrderType = 1
      Call Yw.QueryData(1, m_Rs1, iCount)
      Set Yw = Nothing

      Set Ua.DocumentItems = Nothing
      Set Ua.DocumentItems = New Collection
      While Not m_Rs1.EOF
         Set Yw = New CTaxDocItem
         Call Yw.PopulateFromRS(1, m_Rs1)

         Yw.Flag = "I"
         Call Ua.DocumentItems.Add(Yw)

         Set Yw = Nothing
         m_Rs1.MoveNext
      Wend
   End If
      
   QueryTaxDocument2 = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryTaxDocument2 = False
End Function
Public Function AddEditTaxDocument(Ug As CTaxDocument, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean
Dim iCount As Long

   RName = "AddEditTaxDocument"
   AddEditTaxDocument = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData

   Dim CstName As CTaxDocItem
   For Each CstName In Ug.DocumentItems
      If CstName.Flag = "A" Then
         CstName.TAX_DOCUMENT_ID = Ug.TAX_DOCUMENT_ID
         CstName.AddEditMode = SHOW_ADD
         CstName.AddEditData
      ElseIf CstName.Flag = "E" Then
         CstName.AddEditMode = SHOW_EDIT
         CstName.AddEditData
      ElseIf CstName.Flag = "D" Then
         CstName.DeleteData
      End If
   Next CstName
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditTaxDocument = True
  Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditTaxDocument = False
End Function

Public Function DeleteTaxDocument(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim itemcount As Long

Dim B As CTaxDocument

   RName = "DeleteTaxDocument"
   DeleteTaxDocument = False
   HasBegin = False

   Set B = New CTaxDocument

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   B.TAX_DOCUMENT_ID = UID
   Call B.DeleteData

   Set B = Nothing

   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteTaxDocument = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False

   DeleteTaxDocument = True
End Function

Public Function DeleteEnterprise(UID As Long, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL As String
Dim HasBegin As Boolean
Dim cCust As CEnterprise

   RName = "DeleteEnterprise"
   DeleteEnterprise = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True
      
   Set cCust = New CEnterprise
   
   cCust.ENTERPRISE_ID = UID
   Call cCust.DeleteData(1)
   
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   DeleteEnterprise = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.LocalErrorMsg = "ไม่สามารถลบข้อมูลได้เนื่องจากข้อมูลนี้ยังถูกอ้างถึงอยู่จากส่วนอื่น"
   ErrorObj.ShowErrorLog (LOG_MSGBOX)
   
   IsOK = False

   DeleteEnterprise = True
End Function
Public Function QueryLender(Ua As CLender, Rs As ADODB.Recordset, itemcount As Long, IsOK As Boolean, ErrorObj As clsErrorLog) As Boolean
On Error GoTo ErrorHandler
Dim RName As String
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderStr As String
Dim OrderType As String
Dim NewStr As String
Dim SubLen As Long
Dim iCount As Long
Dim TaxReType As TAXWH_TYPE

   RName = "QueryLender"
   QueryLender = False

   IsOK = True
    Set Rs = New ADODB.Recordset
    Call Ua.QueryData(1, Rs, itemcount)


'   If Ua.QueryFlag = 1 Then
'      Dim Yw As CLender
'
'      Set Yw = New CLender
'      Yw.LENDER_ID = -1
'      Yw.OrderBy = -1
'      Yw.OrderType = 1
'      Call Yw.QueryData(1, m_Rs1, iCount)
'      Set Yw = Nothing
'
'
'      While Not m_Rs1.EOF
'         Set Yw = New CLender
'         Call Yw.PopulateFromRS(1, m_Rs1)
'
'         Yw.Flag = "I"
'
'         Set Yw = Nothing
'         m_Rs1.MoveNext
'      Wend
'   End If
      
   QueryLender = True
   Exit Function

ErrorHandler:
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_FILE_MSGBOX)

   QueryLender = False
End Function
Public Function AddEditLenderItems(Ug As CLender_Items, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLenderItems"
   AddEditLenderItems = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditLenderItems = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLenderItems = False
End Function
Public Function AddEditLender(Ug As CLender, IsOK As Boolean, AutoCommit As Boolean, ErrorObj As clsErrorLog)
On Error GoTo ErrorHandler
Dim RName As String
Dim HasBegin As Boolean

   RName = "AddEditLender"
   AddEditLender = False
   HasBegin = False

   If AutoCommit Then
      m_ADOConn.BeginTrans
   End If
   HasBegin = True

   Call Ug.AddEditData
      
   If AutoCommit Then
      Call m_ADOConn.CommitTrans
   End If
   HasBegin = False
   IsOK = True

   AddEditLender = True
   Exit Function

ErrorHandler:
   If HasBegin Then
      If AutoCommit Then
         m_ADOConn.RollbackTrans
      End If
   End If

   ErrorObj.LocalErrorMsg = "Runtime error."
   ErrorObj.SystemErrorMsg = Err.DESCRIPTION
   ErrorObj.RoutineName = RName
   ErrorObj.ModuleName = m_ModuleName
   ErrorObj.ShowErrorLog (LOG_TO_FILE)

   IsOK = False
   AddEditLender = False
End Function

