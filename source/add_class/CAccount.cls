VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member

Public ACCOUNT_ID As Long
Public ACCOUNT_NO As String
Public ACCOUNT_TYPE As Long
Public ACCOUNT_STATUS As Long
Public CUSTOMER_ID As Long
Public NOTE As String
Public CREDIT As Long
Public ENABLE_FLAG As String
Public MASTER_FLAG As String

Public ACCOUNT_NAME As String
Public DOCUMENT_NO As String
Public DOCUMENT_TYPE As Long
Public ACCTTYPE_NAME As String
Public ACCTSTS_NAME As String
Public CUSTOMER_CODE As String
Public CUSTOMER_NAME As String

Public ACCTYPE_NAME As String
Public ACCSTATUS_NAME As String

Public ActDocs As Collection
Public ActAddr As Collection
Public ActSubs As Collection
Public ActAgrmnts As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set ActDocs = New Collection
   Set ActAddr = New Collection
   Set ActSubs = New Collection
   Set ActAgrmnts = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
Dim SQL1 As String
Dim SQL2 As String
Dim SelectStr1 As String
Dim SelectStr2 As String
Dim FromStr As String
Dim WhereStr As String
Dim OrderByStr As String
Dim OrderTypeStr As String
Dim SubLen As Long
Dim NewStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT ADR.*, NM.LONG_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM ACCOUNT ADR, NAME NM, CUSTOMER_NAME CN "
      
      WhereStr = "WHERE (ADR.CUSTOMER_ID = CN.CUSTOMER_ID) AND (CN.NAME_ID = NM.NAME_ID) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT AC.*, NULL ACCTYPE_NAME,  NULL ACCTSTS_NAME, CM.CUSTOMER_CODE "
      SelectStr2 = "SELECT COUNT(*)  "
   
      FromStr = "FROM ACCOUNT AC "
      FromStr = FromStr & "LEFT OUTER JOIN CUSTOMER CM ON (AC.CUSTOMER_ID = CM.CUSTOMER_ID) "
      
      WhereStr = ""
   End If
   
   If ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
      
   If CUSTOMER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (AC.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (AC.CUSTOMER_ID = " & CUSTOMER_ID & ") "
      End If
   End If
   
   If Len(ENABLE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      End If
   End If
      
   If ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
   
   If ACCOUNT_STATUS > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ACCOUNT_STATUS = " & ACCOUNT_STATUS & ") "
      Else
         WhereStr = WhereStr & "AND (ACCOUNT_STATUS = " & ACCOUNT_STATUS & ") "
      End If
   End If

   If Len(ENABLE_FLAG) > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      Else
         WhereStr = WhereStr & "AND (ENABLE_FLAG = '" & ENABLE_FLAG & "') "
      End If
   End If

   If Trim(ACCOUNT_NO) <> "" Then
      If WildCard(ACCOUNT_NO, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(ACCOUNT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(ACCOUNT_NO,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ACCOUNT_NO = '" & ChangeQuote(Trim(ACCOUNT_NO)) & "')"
         Else
            WhereStr = WhereStr & "AND (ACCOUNT_NO = '" & ChangeQuote(Trim(ACCOUNT_NO)) & "')"
         End If
      End If
   End If

   If Trim(CUSTOMER_CODE) <> "" Then
      If WildCard(CUSTOMER_CODE, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(CUSTOMER_CODE,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         Else
            WhereStr = WhereStr & "AND (CUSTOMER_CODE = '" & ChangeQuote(Trim(CUSTOMER_CODE)) & "')"
         End If
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""
      
   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ACCOUNT_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ACCOUNT_ID "
   Else
      OrderByStr = "ORDER BY ACCOUNT_ID "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "DESC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
      ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
      ACCOUNT_STATUS = NVLI(Rs("ACCOUNT_STATUS"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      NOTE = NVLS(Rs("NOTE"), "")
      ACCOUNT_TYPE = NVLI(Rs("ACCOUNT_TYPE"), -1)
      CREDIT = NVLI(Rs("CREDIT"), -1)
      ENABLE_FLAG = NVLS(Rs("ENABLE_FLAG"), "")
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "")
   ElseIf Ind = 2 Then
      ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
      ACCOUNT_NO = NVLS(Rs("ACCOUNT_NO"), "")
      ACCOUNT_STATUS = NVLI(Rs("ACCOUNT_STATUS"), -1)
      CUSTOMER_ID = NVLI(Rs("CUSTOMER_ID"), -1)
      NOTE = NVLS(Rs("NOTE"), "")
      ACCOUNT_TYPE = NVLI(Rs("ACCOUNT_TYPE"), -1)
      CREDIT = NVLI(Rs("CREDIT"), -1)
      ENABLE_FLAG = NVLS(Rs("ENABLE_FLAG"), "")
      MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "")
      
      ACCTTYPE_NAME = NVLS(Rs("ACCTYPE_NAME"), "")
      ACCTSTS_NAME = NVLS(Rs("ACCTSTS_NAME"), "")
   End If
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = ACCOUNT_TYPE
   If ACCOUNT_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = ACCOUNT_STATUS
   If ACCOUNT_STATUS <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("ACCOUNT_SEQ", ID, glbErrorLog)
      Else
         ID = ACCOUNT_ID
      End If
      
      SQL1 = "INSERT INTO ACCOUNT " & _
                     "( " & _
                     "   ACCOUNT_ID, " & _
                     "   ACCOUNT_NO, " & _
                     "   ACCOUNT_STATUS, " & _
                     "   CUSTOMER_ID, " & _
                     "   NOTE, " & _
                     "   ACCOUNT_TYPE, " & _
                     "   CREDIT, " & _
                     "   ENABLE_FLAG, " & _
                     "   MASTER_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(ACCOUNT_NO)) & "'," & _
                         "" & TempID2 & "," & _
                         " " & CUSTOMER_ID & "," & _
                         "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                         "" & TempID1 & "," & _
                         "" & CREDIT & "," & _
                         "'" & ChangeQuote(Trim(ENABLE_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(MASTER_FLAG)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ACCOUNT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ACCOUNT SET " & _
                     "   ACCOUNT_NO = '" & ChangeQuote(Trim(ACCOUNT_NO)) & "', " & _
                     "   ACCOUNT_STATUS = " & TempID2 & ", " & _
                     "   CUSTOMER_ID = '" & CUSTOMER_ID & "', " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "   ACCOUNT_TYPE = " & TempID1 & ", " & _
                     "   CREDIT = " & CREDIT & ", " & _
                     "   ENABLE_FLAG = '" & ChangeQuote(Trim(ENABLE_FLAG)) & "', " & _
                     "   MASTER_FLAG = '" & ChangeQuote(Trim(MASTER_FLAG)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ACCOUNT_ID = " & ACCOUNT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
   
End Function

Public Function DeleteData()
Dim SQL1 As String
            
   SQL1 = "DELETE FROM AGREEMENT WHERE SUBSCRIBER_ID IN (SELECT SUBSCRIBER_ID FROM SUBSCRIBER WHERE ACCOUNT_ID = " & ACCOUNT_ID & " ) "
   m_Conn.Execute (SQL1)
         
   SQL1 = "DELETE FROM SUBSCRIBER WHERE ACCOUNT_ID = " & ACCOUNT_ID
   m_Conn.Execute (SQL1)
   
   SQL1 = "DELETE FROM ACCOUNT WHERE ACCOUNT_ID = " & ACCOUNT_ID
   m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set ActDocs = Nothing
   Set ActAddr = Nothing
   Set ActSubs = Nothing
   Set ActAgrmnts = Nothing
End Sub




