VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEnterprise"
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

Public ENTERPRISE_ID As Long
Public BUSINESS_TYPE As Long
Public EMAIL As String
Public TAX_ID As String
Public IDENT_ID As String
Public SETUP_DATE As Date
Public POLICY As String
 Public ENTERPRISE_TYPE As Long
Public WEBSITE As String
Public BRANCH_CODE As String
Public BRANCH_NAME As String

Public ENTERPRISE_NAME As String
Public SHORT_NAME As String

Public EnpAddresses As Collection
Public EnpNames As Collection
Public EnpPersons As Collection

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set EnpAddresses = New Collection
   Set EnpNames = New Collection
   Set EnpPersons = New Collection
End Sub

'In this class we will not handle any error or any transaction
Public Function QueryData(Rs As ADODB.Recordset, itemcount As Long)
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

   SelectStr1 = "SELECT ETP.*, NM.LONG_NAME, NM.SHORT_NAME  "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM ENTERPRISE ETP "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (ETP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
   
   WhereStr = ""
   
   If ENTERPRISE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (ETP.ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (ETP.ENTERPRISE_ID = " & ENTERPRISE_ID & ") "
      End If
   End If
      
   If Len(Trim(SHORT_NAME)) > 0 Then
      If WildCard(SHORT_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(SHORT_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(SHORT_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "')"
         End If
      End If
   End If
      
   If Len(Trim(ENTERPRISE_NAME)) > 0 Then
      If WildCard(ENTERPRISE_NAME, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (SUBSTR(ENTERPRISE_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         Else
            WhereStr = WhereStr & "AND (SUBSTR(ENTERPRISE_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (ENTERPRISE_NAME = '" & ChangeQuote(Trim(ENTERPRISE_NAME)) & "')"
         Else
            WhereStr = WhereStr & "AND (ENTERPRISE_NAME = '" & ChangeQuote(Trim(ENTERPRISE_NAME)) & "')"
         End If
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

   If OrderBy = 1 Then
      OrderByStr = "ORDER BY SHORT_NAME "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY LONG_NAME "
   Else
      OrderByStr = "ORDER BY ENTERPRISE_ID "
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
   ''Debug.Print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   ENTERPRISE_ID = NVLI(Rs("ENTERPRISE_ID"), -1)
   BUSINESS_TYPE = NVLI(Rs("BUSINESS_TYPE"), -1)
   EMAIL = NVLS(Rs("EMAIL"), "")
   TAX_ID = NVLS(Rs("TAX_ID"), "")
  IDENT_ID = NVLS(Rs("IDENT_ID"), "")
   SETUP_DATE = InternalDateToDate(NVLS(Rs("SETUP_DATE"), ""))
   POLICY = NVLS(Rs("POLICY"), "")
   ENTERPRISE_TYPE = NVLI(Rs("ENTERPRISE_TYPE"), -1)
   ENTERPRISE_NAME = NVLS(Rs("LONG_NAME"), "")
   BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
   BRANCH_NAME = NVLS(Rs("SHORT_NAME"), "")
   
   SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
   
   KEY_ID = ENTERPRISE_ID
   KEY_LOOKUP = SHORT_NAME
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String
Dim TempID2 As String

   TempID1 = ENTERPRISE_TYPE
   If ENTERPRISE_TYPE <= 0 Then
      TempID1 = "NULL"
   End If
   
   TempID2 = BUSINESS_TYPE
   If BUSINESS_TYPE <= 0 Then
      TempID2 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("ENTERPRISE_SEQ", ID, glbErrorLog)
      Else
         ID = ENTERPRISE_ID
      End If
      
      SQL1 = "INSERT INTO ENTERPRISE " & _
                     "( " & _
                     "   ENTERPRISE_ID, " & _
                     "   BUSINESS_TYPE, " & _
                     "   EMAIL, " & _
                     "   WEBSITE, " & _
                     "   TAX_ID, " & _
                     "   SETUP_DATE, " & _
                     "   POLICY, " & _
                     "   ENTERPRISE_TYPE, " & _
                     "   BRANCH_CODE, " & _
                     "   BRANCH_NAME, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & TempID2 & "," & _
                         "'" & ChangeQuote(Trim(EMAIL)) & "'," & _
                         "'" & ChangeQuote(Trim(WEBSITE)) & "'," & _
                         "'" & ChangeQuote(Trim(TAX_ID)) & "'," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(SETUP_DATE))) & "'," & _
                         "'" & ChangeQuote(Trim(POLICY)) & "'," & _
                         "" & TempID1 & "," & _
                         "'" & ChangeQuote(Trim(BRANCH_CODE)) & "'," & _
                         "'" & ChangeQuote(Trim(BRANCH_NAME)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         ENTERPRISE_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE ENTERPRISE SET " & _
                     "   BUSINESS_TYPE = " & TempID2 & ", " & _
                     "   EMAIL = '" & ChangeQuote(Trim(EMAIL)) & "', " & _
                     "   WEBSITE = '" & ChangeQuote(Trim(WEBSITE)) & "', " & _
                     "   TAX_ID = '" & ChangeQuote(Trim(TAX_ID)) & "', " & _
                     "   SETUP_DATE = '" & ChangeQuote(Trim(DateToStringInt(SETUP_DATE))) & "', " & _
                     "   POLICY = '" & ChangeQuote(Trim(POLICY)) & "', " & _
                     "   ENTERPRISE_TYPE = " & TempID1 & ", " & _
                     "   BRANCH_CODE = '" & ChangeQuote(Trim(BRANCH_CODE)) & "', " & _
                     "   BRANCH_NAME = '" & ChangeQuote(Trim(BRANCH_NAME)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData(Ind As Long)
Dim SQL1 As String
      
      SQL1 = "DELETE FROM ENTERPRISE_NAME WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM ENTERPRISE_ADDRESS WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
      m_Conn.Execute (SQL1)
      
      SQL1 = "DELETE FROM ENTERPRISE WHERE ENTERPRISE_ID = " & ENTERPRISE_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set EnpAddresses = Nothing
   Set EnpNames = Nothing
   Set EnpPersons = Nothing
End Sub
