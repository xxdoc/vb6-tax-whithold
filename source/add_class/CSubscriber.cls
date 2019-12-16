VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSubscriber"
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

Public SUBSCRIBER_ID As Long
Public SUBSCRIBER_NO As String
Public SUBSCRIBER_STATUS As String
Public ACCOUNT_ID As Long
Public DUMMY_FLAG As String
Public SUBSCRIBER_DESC As String

Public Agreements As Collection
Public Quoatas As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Agreements = New Collection
   Set Quoatas = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, ItemCount As Long)
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

   SelectStr1 = "SELECT ADR.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM SUBSCRIBER ADR "
   
   If SUBSCRIBER_ID >= 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (SUBSCRIBER_ID = " & SUBSCRIBER_ID & ") "
      End If
   End If
      
   If ACCOUNT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ACCOUNT_ID = " & ACCOUNT_ID & ") "
      End If
   End If
      
   If SUBSCRIBER_STATUS <> "" Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (SUBSCRIBER_STATUS = '" & SUBSCRIBER_STATUS & "') "
      Else
         WhereStr = WhereStr & "AND (SUBSCRIBER_STATUS = '" & SUBSCRIBER_STATUS & "') "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SUBSCRIBER_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SUBSCRIBER_ID "
   Else
      OrderByStr = "ORDER BY SUBSCRIBER_ID "
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
   
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, ItemCount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Rs As ADODB.Recordset)
   SUBSCRIBER_ID = NVLI(Rs("SUBSCRIBER_ID"), -1)
   SUBSCRIBER_NO = NVLS(Rs("SUBSCRIBER_NO"), "")
   SUBSCRIBER_STATUS = NVLS(Rs("SUBSCRIBER_STATUS"), "")
   ACCOUNT_ID = NVLI(Rs("ACCOUNT_ID"), -1)
   DUMMY_FLAG = NVLS(Rs("DUMMY_FLAG"), "")
   SUBSCRIBER_DESC = NVLS(Rs("SUBSCRIBER_DESC"), "")
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("SUBSCRIBER_SEQ", ID, glbErrorLog)
      Else
         ID = SUBSCRIBER_ID
      End If
      
      SQL1 = "INSERT INTO SUBSCRIBER " & _
                     "( " & _
                     "   SUBSCRIBER_ID, " & _
                     "   SUBSCRIBER_NO, " & _
                     "   SUBSCRIBER_STATUS, " & _
                     "   ACCOUNT_ID, " & _
                     "   DUMMY_FLAG, " & _
                     "   SUBSCRIBER_DESC, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(SUBSCRIBER_NO)) & "'," & _
                         "'" & SUBSCRIBER_STATUS & "'," & _
                         " " & ACCOUNT_ID & "," & _
                         "'" & ChangeQuote(Trim(DUMMY_FLAG)) & "'," & _
                         "'" & ChangeQuote(Trim(SUBSCRIBER_DESC)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         SUBSCRIBER_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE SUBSCRIBER SET " & _
                     "   SUBSCRIBER_NO = '" & ChangeQuote(Trim(SUBSCRIBER_NO)) & "', " & _
                     "   SUBSCRIBER_STATUS = '" & SUBSCRIBER_STATUS & "', " & _
                     "   ACCOUNT_ID = '" & ACCOUNT_ID & "', " & _
                     "   DUMMY_FLAG = '" & ChangeQuote(Trim(DUMMY_FLAG)) & "', " & _
                     "   SUBSCRIBER_DESC = '" & ChangeQuote(Trim(SUBSCRIBER_DESC)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE SUBSCRIBER_ID = " & SUBSCRIBER_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM AGREEMENT WHERE SUBSCRIBER_ID = " & SUBSCRIBER_ID & " "
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM QUOATA_AGREEMENT WHERE SUBSCRIBER_ID = " & SUBSCRIBER_ID & " "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM SUBSCRIBER WHERE SUBSCRIBER_ID = " & SUBSCRIBER_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Agreements = Nothing
   Set Quoatas = Nothing
End Sub


