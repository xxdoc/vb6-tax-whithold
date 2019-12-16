VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnterprisePerson"
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

Public ENTERPRISE_PERSON_ID As Long
Public ENTERPRISE_ID As Long
Public NAME_ID As Long
Public MASTER_FLAG As String
Public ADDRESS_ID As Long

Public Name As cName
Public Address As CAddress

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set Name = New cName
   Set Address = New CAddress
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

   SelectStr1 = "SELECT UG.*  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ENTERPRISE_PERSON UG "

   WhereStr = ""
   If ENTERPRISE_PERSON_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_PERSON_ID = " & ENTERPRISE_PERSON_ID & ") "
      Else
         WhereStr = "AND (ENTERPRISE_PERSON_ID = " & ENTERPRISE_PERSON_ID & ") "
      End If
   End If
   
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      Else
         WhereStr = "AND (ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      End If
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY ENTERPRISE_PERSON_ID "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY ENTERPRISE_PERSON_ID "
   Else
      OrderByStr = "ORDER BY ENTERPRISE_PERSON_ID "
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
   ENTERPRISE_PERSON_ID = NVLI(Rs("ENTERPRISE_PERSON_ID"), -1)
   ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
   NAME_ID = NVLI(Rs("NAME_ID"), -1)
   ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
   MASTER_FLAG = NVLS(Rs("MASTER_FLAG"), "")
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = ADDRESS_ID
   If ADDRESS_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("ENTERPRISE_PERSON_SEQ", ID, glbErrorLog)
      Else
         ID = ENTERPRISE_PERSON_ID
      End If
      
      SQL1 = "INSERT INTO ENTERPRISE_PERSON " & _
                     "( " & _
                     "   ENTERPRISE_PERSON_ID, " & _
                     "   ENTERPRISE_ID, " & _
                     "   NAME_ID, " & _
                     "   ADDRESS_ID, " & _
                     "   MASTER_FLAG, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & ENTERPRISE_ID & "," & _
                         "" & NAME_ID & "," & _
                         "" & TempID1 & "," & _
                         "'" & MASTER_FLAG & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ENTERPRISE_PERSON_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ENTERPRISE_PERSON SET " & _
                     "   ENTERPRISE_ID = " & ENTERPRISE_ID & ", " & _
                     "   NAME_ID = " & NAME_ID & ", " & _
                     "   ADDRESS_ID = " & TempID1 & ", " & _
                     "   MASTER_FLAG = '" & MASTER_FLAG & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ENTERPRISE_PERSON_ID = " & ENTERPRISE_PERSON_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM ENTERPRISE_PERSON WHERE ENTERPRISE_PERSON_ID = " & ENTERPRISE_PERSON_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set Name = Nothing
   Set Address = Nothing
End Sub


