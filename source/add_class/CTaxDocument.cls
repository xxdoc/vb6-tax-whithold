VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaxDocument"
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
Public TAX_DOCUMENT_ID As Long
Public DOCUMENT_DATE As Date
Public RESPONSE_ID As Long
Public COMPANY_ID As Long
Public TAX_TYPE As Long
Public TAX_ID As String
Public IDENT_ID As String
Public LENDER_ITEMS_ID As Long

Public SUPPLIER_CODE As String
Public BRANCH_ID As Long
Public BRANCH_NAME As String
Public FROM_DATE As Date
Public TO_DATE As Date
Public SHORT_NAME As String
Public LONG_NAME As String
Public FOR_MONTH As String
Public DOCUMENT_YEAR As String

Public DocumentItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set DocumentItems = New Collection
   FROM_DATE = -1
   TO_DATE = -1
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
   SelectStr1 = "SELECT ADR.*, EP.BRANCH_NAME BRANCH_NAME,EP.TAX_ID , EP.IDENT_ID,NM.LONG_NAME, NM.SHORT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM TAX_DOCUMENT ADR "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE EP ON (ADR.COMPANY_ID = EP.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (EP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
ElseIf Ind = 2 Then
   SelectStr1 = "SELECT DISTINCT(ADR.COMPANY_ID),ADR.TAX_DOCUMENT_ID,ADR.TAX_TYPE,ADR.DOCUMENT_DATE, EP.BRANCH_NAME BRANCH_NAME,EP.TAX_ID , EP.IDENT_ID,NM.LONG_NAME, NM.SHORT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM TAX_DOCUMENT ADR "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE EP ON (ADR.COMPANY_ID = EP.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (EP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
End If
   
   
   If TAX_DOCUMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
      End If
   End If
      
   If COMPANY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMPANY_ID = " & COMPANY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COMPANY_ID = " & COMPANY_ID & ") "
      End If
   End If
      
   If RESPONSE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RESPONSE_ID = " & RESPONSE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RESPONSE_ID = " & RESPONSE_ID & ") "
      End If
   End If
      
   If FROM_DATE > -1 Then
      tempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(tempDate) & "')"
      End If
   End If
   
   If TO_DATE > -1 Then
      tempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(tempDate) & "')"
      End If
   End If
      
   If TAX_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TAX_TYPE = " & TAX_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (TAX_TYPE = " & TAX_TYPE & ") "
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
   
   OrderByStr = ""
   OrderTypeStr = ""

 If OrderBy = 1 Then
      OrderByStr = "ORDER BY DOCUMENT_DATE "
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SHORT_NAME "
   Else
      OrderByStr = "ORDER BY DOCUMENT_DATE "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
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

   SelectStr1 = "SELECT ADR.*, EP.BRANCH_NAME BRANCH_NAME,EP.TAX_TYPE , EP.IDENT_ID , NM.LONG_NAME, NM.SHORT_NAME "
   SelectStr2 = "SELECT COUNT(*)  "
   
   FromStr = "FROM TAX_DOCUMENT ADR "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE EP ON (ADR.COMPANY_ID = EP.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN ENTERPRISE_NAME EN ON (EP.ENTERPRISE_ID = EN.ENTERPRISE_ID) "
   FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (EN.NAME_ID = NM.NAME_ID) "
   
   
   If TAX_DOCUMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
      End If
   End If
      
   If COMPANY_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (COMPANY_ID = " & COMPANY_ID & ") "
      Else
         WhereStr = WhereStr & "AND (COMPANY_ID = " & COMPANY_ID & ") "
      End If
   End If
      
   If RESPONSE_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (RESPONSE_ID = " & RESPONSE_ID & ") "
      Else
         WhereStr = WhereStr & "AND (RESPONSE_ID = " & RESPONSE_ID & ") "
      End If
   End If
      
   If FROM_DATE > -1 Then
      tempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE >= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE >= '" & ChangeQuote(tempDate) & "')"
      End If
   End If
   
   If TO_DATE > -1 Then
      tempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DOCUMENT_DATE <= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DOCUMENT_DATE <= '" & ChangeQuote(tempDate) & "')"
      End If
   End If
      
   If TAX_TYPE > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (TAX_TYPE = " & TAX_TYPE & ") "
      Else
         WhereStr = WhereStr & "AND (TAX_TYPE = " & TAX_TYPE & ") "
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
   Else
            WhereStr = "WHERE (SHORT_NAME = 'HO' ) OR (SHORT_NAME = 'MP' ) OR (SHORT_NAME = 'MA2' )"
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

 If OrderBy = 1 Then
      OrderByStr = "ORDER BY DOCUMENT_DATE"
   ElseIf OrderBy = 2 Then
      OrderByStr = "ORDER BY SHORT_NAME "
   Else
      OrderByStr = "ORDER BY DOCUMENT_DATE "
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
'   'Debug.Print SQL1
'   'Debug.Print "-----"
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TAX_DOCUMENT_ID = NVLI(Rs("TAX_DOCUMENT_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      RESPONSE_ID = NVLI(Rs("RESPONSE_ID"), -1)
      COMPANY_ID = NVLI(Rs("COMPANY_ID"), -1)
      TAX_TYPE = NVLI(Rs("TAX_TYPE"), -1)
      BRANCH_NAME = NVLS(Rs("BRANCH_NAME"), "")
      LONG_NAME = NVLS(Rs("LONG_NAME"), "")
      SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      IDENT_ID = NVLS(Rs("IDENT_ID"), "")
    ElseIf Ind = 2 Then
      TAX_DOCUMENT_ID = NVLI(Rs("TAX_DOCUMENT_ID"), -1)
      DOCUMENT_DATE = InternalDateToDate(NVLS(Rs("DOCUMENT_DATE"), ""))
      COMPANY_ID = NVLI(Rs("COMPANY_ID"), -1)
      TAX_TYPE = NVLI(Rs("TAX_TYPE"), -1)
      BRANCH_NAME = NVLS(Rs("BRANCH_NAME"), "")
      LONG_NAME = NVLS(Rs("LONG_NAME"), "")
      SHORT_NAME = NVLS(Rs("SHORT_NAME"), "")
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      IDENT_ID = NVLS(Rs("IDENT_ID"), "")
   End If
     
End Function

Public Function AddEditData(Optional AutoGenerateFlag As Boolean = True)
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim TempID1 As String

   TempID1 = RESPONSE_ID
   If RESPONSE_ID <= 0 Then
      TempID1 = "NULL"
   End If
   
   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
   
   If AddEditMode = SHOW_ADD Then
      If AutoGenerateFlag Then
         Call glbDatabaseMngr.GetSeqID("TAX_DOCUMENT_SEQ", ID, glbErrorLog)
      Else
         ID = TAX_DOCUMENT_ID
      End If
      
      SQL1 = "INSERT INTO TAX_DOCUMENT " & _
                     "( " & _
                     "   TAX_DOCUMENT_ID, " & _
                     "   DOCUMENT_DATE, " & _
                     "   RESPONSE_ID, " & _
                     "   COMPANY_ID, " & _
                     "   TAX_TYPE, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "'," & _
                         "" & TempID1 & "," & _
                         "" & COMPANY_ID & "," & _
                         "" & TAX_TYPE & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         TAX_DOCUMENT_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE TAX_DOCUMENT SET " & _
                     "   DOCUMENT_DATE = '" & ChangeQuote(Trim(DateToStringInt(DOCUMENT_DATE))) & "', " & _
                     "   RESPONSE_ID = " & TempID1 & ", " & _
                     "   COMPANY_ID = '" & COMPANY_ID & "', " & _
                     "   TAX_TYPE = " & TAX_TYPE & ", " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      
      SQL1 = "DELETE FROM TAXDOC_ITEM WHERE TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & " "
      m_Conn.Execute (SQL1)

      SQL1 = "DELETE FROM TAX_DOCUMENT WHERE TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID
      m_Conn.Execute (SQL1)
End Function

Private Sub Class_Terminate()
   Set DocumentItems = Nothing
End Sub




