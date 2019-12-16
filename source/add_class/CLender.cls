VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Private Const MODULE_NAME = "CLender"

Public QueryFlag As Long
Public OrderBy As Long
Public OrderType As Long
Public Flag As String
' End of generic data member
Public LENDER_ID As Long
Public COMPANY_ID As Long
Public BRANCH_ID As Long
Public COMPANY_SHORTNAME As String
Public COMPANY_NAME As String
Public BUDGET_YEAR As Long

Public LENDER_ITEMS_ID As Long
Public LENDER_ITEMS_NO As String
Public LENDER_ITEMS_NAME As String
Public LENDER_ITEMS_AMOUNT As Double

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
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
Dim tempDate As String
  If Ind = 1 Then
   SelectStr1 = "SELECT * "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM  LENDER "
 ElseIf Ind = 2 Then
   SelectStr1 = "SELECT LD.*,LDI.* "
   SelectStr2 = "SELECT COUNT(*)  "

   FromStr = "FROM  LENDER_ITEMS LDI "
   FromStr = FromStr & "LEFT OUTER JOIN LENDER LD ON (LDI.LENDER_ID = LD.LENDER_ID) "
 End If

   WhereStr = ""

   If COMPANY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LD.COMPANY_ID = " & COMPANY_ID & ")  "
      Else
         WhereStr = WhereStr & "AND (LD.COMPANY_ID = " & COMPANY_ID & ")  "
      End If
   End If
   
   If BRANCH_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LD.BRANCH_ID = '" & BRANCH_ID & "')  "
      Else
         WhereStr = WhereStr & "AND (LD.BRANCH_ID = '" & BRANCH_ID & "')  "
      End If
   End If
   
   If BUDGET_YEAR > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LD.BUDGET_YEAR = '" & BUDGET_YEAR & "')  "
      Else
         WhereStr = WhereStr & "AND (LD.BUDGET_YEAR = '" & BUDGET_YEAR & "')  "
      End If
   End If
   

'   If FROM_DATE > -1 Then
'      TempDate = DateToStringIntLow(Trim(FROM_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
'
'   If TO_DATE > -1 Then
'      TempDate = DateToStringIntHi(Trim(TO_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If

'   If TAX_TYPE > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (TAX_TYPE = " & TAX_TYPE & ") "
'      Else
'         WhereStr = WhereStr & "AND (TAX_TYPE = " & TAX_TYPE & ") "
'      End If
'   End If

   OrderByStr = ""
   OrderTypeStr = ""

'   If OrderBy = 1 Then
'      OrderByStr = "ORDER BY DOCUMENT_DATE "
'   ElseIf OrderBy = 2 Then
'      OrderByStr = "ORDER BY SHORT_NAME "
'   Else
'      OrderByStr = "ORDER BY DOCUMENT_DATE "
'   End If

'   If OrderType = 1 Then
'      OrderTypeStr = "ASC "
'   ElseIf OrderType = 2 Then
'      OrderTypeStr = "DESC "
'   Else
'      OrderTypeStr = "ASC "
'   End If

   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function
Public Function QueryData2(Ind As Long, Rs As ADODB.Recordset, itemcount As Long)
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
Dim tempDate As String

'   SelectStr1 = "SELECT ADR.*, EP.BRANCH_NAME BRANCH_NAME,EP.TAX_TYPE , EP.IDENT_ID , NM.LONG_NAME, NM.SHORT_NAME "
'   SelectStr2 = "SELECT COUNT(*)  "
'
'   FromStr = "FROM TAX_DOCUMENT ADR "
'   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE EP ON (ADR.COMPANY_ID = EP.ENTERPRISE_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (EP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
'   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
'
'
''   If TAX_DOCUMENT_ID > 0 Then
''      If Len(WhereStr) <= 0 Then
''         WhereStr = "WHERE (TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
''      Else
''         WhereStr = WhereStr & "AND (TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
''      End If
''   End If
'
'   If COMPANY_ID > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (COMPANY_ID = " & COMPANY_ID & ") "
'      Else
'         WhereStr = WhereStr & "AND (COMPANY_ID = " & COMPANY_ID & ") "
'      End If
'   End If
'
'
'
'   If FROM_DATE > -1 Then
'      TempDate = DateToStringIntLow(Trim(FROM_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
'
'   If TO_DATE > -1 Then
'      TempDate = DateToStringIntHi(Trim(TO_DATE))
'      If WhereStr = "" Then
'         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
'      Else
'         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(TempDate) & "')"
'      End If
'   End If
'
'   If TAX_TYPE > 0 Then
'      If Len(WhereStr) <= 0 Then
'         WhereStr = "WHERE (TAX_TYPE = " & TAX_TYPE & ") "
'      Else
'         WhereStr = WhereStr & "AND (TAX_TYPE = " & TAX_TYPE & ") "
'      End If
'   End If
'
'   If Len(Trim(SHORT_NAME)) > 0 Then
'         If WildCard(SHORT_NAME, SubLen, NewStr) Then
'            If WhereStr = "" Then
'               WhereStr = "WHERE (SUBSTR(SHORT_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
'            Else
'               WhereStr = WhereStr & "AND (SUBSTR(SHORT_NAME,1," & SubLen & ") = '" & ChangeQuote(Trim(NewStr)) & "')"
'            End If
'         Else
'            If WhereStr = "" Then
'               WhereStr = "WHERE (SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "')"
'            Else
'               WhereStr = WhereStr & "AND (SHORT_NAME = '" & ChangeQuote(Trim(SHORT_NAME)) & "')"
'            End If
'         End If
'   Else
'            WhereStr = "WHERE (SHORT_NAME = 'HO' ) OR (SHORT_NAME = 'MP' ) OR (SHORT_NAME = 'MA2' )"
'   End If
'
'   OrderByStr = ""
'   OrderTypeStr = ""
'
' If OrderBy = 1 Then
'      OrderByStr = "ORDER BY DOCUMENT_DATE"
'   ElseIf OrderBy = 2 Then
'      OrderByStr = "ORDER BY SHORT_NAME "
'   Else
'      OrderByStr = "ORDER BY DOCUMENT_DATE "
'   End If
'
'   If OrderType = 1 Then
'      OrderTypeStr = "ASC "
'   ElseIf OrderType = 2 Then
'      OrderTypeStr = "DESC "
'   Else
'      OrderTypeStr = "DESC "
'   End If
'
'   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
'   SQL2 = SelectStr2 & FromStr & WhereStr
''   'Debug.Print SQL1
''   'Debug.Print "-----"
'   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      LENDER_ID = NVLS(Rs("LENDER_ID"), "")
      COMPANY_ID = NVLI(Rs("COMPANY_ID"), -1)
      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      COMPANY_SHORTNAME = NVLS(Rs("COMPANY_SHORTNAME"), "")
      COMPANY_NAME = NVLS(Rs("COMPANY_NAME"), "")
      BUDGET_YEAR = NVLI(Rs("BUDGET_YEAR"), -1)
   ElseIf Ind = 2 Then
      LENDER_ID = NVLI(Rs("LENDER_ID"), -1)
      COMPANY_ID = NVLI(Rs("COMPANY_ID"), -1)
      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      COMPANY_SHORTNAME = NVLS(Rs("COMPANY_SHORTNAME"), "")
      COMPANY_NAME = NVLS(Rs("COMPANY_NAME"), "")
      BUDGET_YEAR = NVLI(Rs("BUDGET_YEAR"), -1)
      LENDER_ITEMS_ID = NVLI(Rs("LENDER_ITEMS_ID"), -1)
      LENDER_ITEMS_NO = NVLS(Rs("LENDER_ITEMS_NO"), "")
      LENDER_ITEMS_NAME = NVLS(Rs("LENDER_ITEMS_NAME"), "")
      LENDER_ITEMS_AMOUNT = NVLD(Rs("LENDER_ITEMS_AMOUNT"), 0)
      
      KEY_ID = NVLI(Rs("LENDER_ITEMS_ID"), -1)
      KEY_LOOKUP = NVLS(Rs("LENDER_ITEMS_ID"), "")
   End If
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String



   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)

   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("LENDER_SEQ", ID, glbErrorLog)
      Else
         ID = LENDER_ID
      End If

      SQL1 = "INSERT INTO LENDER " & _
                     "( " & _
                     "   LENDER_ID, " & _
                     "   COMPANY_ID, " & _
                     "   BRANCH_ID, " & _
                     "   COMPANY_SHORTNAME, " & _
                     "   COMPANY_NAME, " & _
                     "   BUDGET_YEAR, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "" & COMPANY_ID & "," & _
                         "" & BRANCH_ID & "," & _
                         "'" & COMPANY_SHORTNAME & "'," & _
                         "'" & COMPANY_NAME & "'," & _
                         "" & BUDGET_YEAR & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         LENDER_ID = ID
   Else 'Show edit
'      SQL1 = "UPDATE TAX_DOCUMENT SET " & _
'                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
'                     "   RESPONSE_ID = " & TempID1 & ", " & _
'                     "   COMPANY_ID = '" & COMPANY_ID & "', " & _
'                     "   TAX_TYPE = " & TAX_TYPE & ", " & _
'                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
'                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
'                     "WHERE TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID
   End If

   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String

      SQL1 = "DELETE FROM LENDER WHERE LENDER_ID = " & LENDER_ID & " "
      m_Conn.Execute (SQL1)
End Function

