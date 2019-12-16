VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
End Sub

Public Function IsPatch(VersionStr As String) As Boolean
On Error GoTo ErrorHandler
Dim SQL As String
Dim TempRs As ADODB.Recordset
Dim itemcount As Long
Dim Result As Boolean

   Set TempRs = New ADODB.Recordset
   Result = True
   
   SQL = "SELECT * FROM PATCH_DELIVERY WHERE PATCH_VERSION = '" & VersionStr & "' "
   Call glbDatabaseMngr.GetRs(SQL, "", False, itemcount, TempRs, glbErrorLog)

   If TempRs.EOF Then
      Result = fasle
   End If
   
   If TempRs.State = adStateOpen Then
      TempRs.Close
   End If
   Set TempRs = Nothing
   
   IsPatch = Result
   Exit Function
   
ErrorHandler:
   IsPatch = False
End Function

Public Function Patch_1_0_1_2() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_1_2"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "CREATE DOMAIN IDENT_TYPE AS VARCHAR(20)"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUPPLIER ADD IDENT_ID IDENT_TYPE"
   m_Conn.Execute (SQL)
   
   SQL = "CREATE DOMAIN DPACC_TYPE AS VARCHAR(20)"
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE SUPPLIER ADD DPACC_ID DPACC_TYPE"
   m_Conn.Execute (SQL)
   
    SQL = "ALTER TABLE ENTERPRISE ADD IDENT_ID IDENT_TYPE"
     m_Conn.Execute (SQL)

   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_1_2 = True
   Exit Function
   
End Function

Public Function Patch_1_0_1_13() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_1_13"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE TAXDOC_ITEM ADD BRANCH CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_1_13 = True
   Exit Function
   
End Function

Public Function Patch_1_0_1_14() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_1_14"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE TAXDOC_ITEM ADD BRANCH_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE TAXDOC_ITEM ADD CONSTRAINT TAXDOC_ITEM_BRANCH_ID_FK FOREIGN KEY (BRANCH_ID) REFERENCES MASTER_REF "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_1_14 = True
   Exit Function
   
End Function

Public Function Patch_1_0_1_15() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "1_0_1_15"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_1_0_1_15 = True
   Exit Function
End Function
Public Function Patch_2016_08_15_1_jew() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "2016_08_15_1_jew"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE MASTER_REF ADD EXPORT_KEY CODE_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2016_08_15_1_jew = True
   Exit Function
   
End Function
Public Function Patch_2017_02_22_1_lek() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "2017_02_22_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "CREATE TABLE LENDER "
   SQL = SQL & "( "
   SQL = SQL & "LENDER_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "COMPANY_ID         ID_TYPE NOT NULL, "
   SQL = SQL & "BRANCH_ID          ID_TYPE NOT NULL, "
   SQL = SQL & "COMPANY_SHORTNAME         SNAME_TYPE NOT NULL, "
   SQL = SQL & "COMPANY_NAME         NAME_TYPE NOT NULL, "
   SQL = SQL & "BUDGET_YEAR         YEAR_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
 
   SQL = "ALTER TABLE LENDER ADD CONSTRAINT LENDER_ID_PK PRIMARY KEY (LENDER_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR LENDER_SEQ  "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE TABLE LENDER_ITEMS "
   SQL = SQL & "( "
   SQL = SQL & "LENDER_ITEMS_ID      ID_TYPE NOT NULL, "
   SQL = SQL & "LENDER_ITEMS_NO           CODE_TYPE NOT NULL, "
   SQL = SQL & "LENDER_ITEMS_NAME          CODE_TYPE NOT NULL, "
   SQL = SQL & "LENDER_ITEMS_AMOUNT          MONEY_TYPE NOT NULL, "
   SQL = SQL & "LENDER_ID         ID_TYPE NOT NULL, "
   SQL = SQL & "CREATE_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "CREATE_BY             ID_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_DATE           DATE_TYPE NOT NULL, "
   SQL = SQL & "MODIFY_BY ID_TYPE NOT NULL "
   SQL = SQL & ") "
   m_Conn.Execute (SQL)
   
   SQL = "ALTER TABLE LENDER_ITEMS ADD CONSTRAINT LENDER_ITEMS_ID_PK PRIMARY KEY (LENDER_ITEMS_ID) "
   m_Conn.Execute (SQL)
   
   SQL = "CREATE GENERATOR LENDER_ITEMS_SEQ  "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_02_22_1_lek = True
   Exit Function
   
End Function
Public Function Patch_2017_03_02_1_lek() As Boolean
On Error Resume Next
Dim SQL As String
Dim ServerDateTime As String
Dim HasBegin As Boolean
Dim VersionStr As String
Dim IsOK As Boolean

   VersionStr = "2017_03_02_1_lek"
   HasBegin = False
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   SQL = "ALTER TABLE TAXDOC_ITEM ADD LENDER_ITEMS_ID ID_TYPE "
   m_Conn.Execute (SQL)
   
   m_Conn.BeginTrans
   HasBegin = True

   SQL = "INSERT INTO PATCH_DELIVERY "
   SQL = SQL & "( "
   SQL = SQL & "PATCH_DELIVERY_ID, "
   SQL = SQL & "PATCH_VERSION, "
   SQL = SQL & "PATCH_DATE, "
   SQL = SQL & "PATCH_BY "
   SQL = SQL & ") VALUES "
   SQL = SQL & "( "
   SQL = SQL & "GEN_ID(PATCH_DELIVERY_SEQ, 1), "
   SQL = SQL & "'" & VersionStr & "', "
   SQL = SQL & "'" & ServerDateTime & "', "
   SQL = SQL & "" & glbUser.USER_ID & " "
   SQL = SQL & ")"
   m_Conn.Execute (SQL)

   m_Conn.CommitTrans
   HasBegin = False
   
   glbErrorLog.LocalErrorMsg = "Patch to version " & VersionStr & " done."
   glbErrorLog.ShowUserError
   
   Patch_2017_03_02_1_lek = True
   Exit Function
   
End Function


