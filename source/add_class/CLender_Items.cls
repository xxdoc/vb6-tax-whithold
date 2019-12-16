VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLender_Items"
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

Public LENDER_ITEMS_ID As Long
Public LENDER_ITEMS_NO As String
Public LENDER_ITEMS_NAME As String
Public LENDER_ITEMS_AMOUNT As Double
Public LENDER_ID As Long

Public DocumentItems As Collection

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   Set DocumentItems = New Collection

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

      SelectStr1 = "SELECT accnum,accnam,level,parent,group,acctyp,bal1cy+bal2cy+bal3cy+bal4cy+bal5cy+bal6cy+bal7cy+bal8cy+bal9cy+bal10cy+bal11cy+bal12cy+beglas+bal1ly+bal2ly+bal3ly+bal4ly+bal5ly+bal6ly+bal7ly+bal8ly+bal9ly+bal10ly+bal11ly+bal12ly as Sum "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GLACC "
      WhereStr = " WHERE ACCNUM LIKE '220-%' "
      
   OrderByStr = ""
   OrderTypeStr = ""

  OrderByStr = "ORDER BY accnum "

   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If

   SQL1 = SelectStr1 & FromStr & WhereStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr

   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
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
If Ind = 2 Then
      SelectStr1 = "SELECT * "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM LENDER_ITEMS LI "
      WhereStr = ""
End If
    If LENDER_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (LI.LENDER_ID = " & LENDER_ID & ") "
      Else
         WhereStr = WhereStr & "AND (LI.LENDER_ID = " & LENDER_ID & ") "
      End If
   End If
      
   OrderByStr = ""
   OrderTypeStr = ""

  OrderByStr = "ORDER BY LI.LENDER_ITEMS_NAME "

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
Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 2 Then
         LENDER_ITEMS_ID = NVLI(Rs("LENDER_ITEMS_ID"), -1)
         LENDER_ITEMS_NO = NVLS(Rs("LENDER_ITEMS_NO"), "")
         LENDER_ITEMS_NAME = NVLS(Rs("LENDER_ITEMS_NAME"), "")
         LENDER_ITEMS_AMOUNT = NVLD(Rs("LENDER_ITEMS_AMOUNT"), 0)
         LENDER_ID = NVLI(Rs("LENDER_ID"), -1)
   ElseIf Ind = 3 Then
         LENDER_ITEMS_NO = NVLS(Rs("accnum"), "")
         LENDER_ITEMS_NAME = NVLS(Rs("accnam"), "")
         LENDER_ITEMS_AMOUNT = NVLD(Rs("sum"), 0)
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
         Call glbDatabaseMngr.GetSeqID("LENDER_ITEMS_SEQ", ID, glbErrorLog)
      Else
         ID = LENDER_ID
      End If

      SQL1 = "INSERT INTO LENDER_ITEMS " & _
                     "( " & _
                     "   LENDER_ITEMS_ID, " & _
                     "   LENDER_ITEMS_NO, " & _
                     "   LENDER_ITEMS_NAME, " & _
                     "   LENDER_ITEMS_AMOUNT, " & _
                     "   LENDER_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES "
         SQL1 = SQL1 & "( " & _
                         ID & "," & _
                         "'" & LENDER_ITEMS_NO & "'," & _
                         "'" & LENDER_ITEMS_NAME & "'," & _
                         "" & LENDER_ITEMS_AMOUNT & "," & _
                         "" & LENDER_ID & "," & _
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

Private Sub Class_Terminate()
   Set DocumentItems = Nothing
End Sub
Public Function CopyObject(Ind As Long, O As CLender_Items)
'      TAXDOC_ITEM_ID = O.TAXDOC_ITEM_ID
      LENDER_ITEMS_NO = O.LENDER_ITEMS_NO
      LENDER_ITEMS_NAME = O.LENDER_ITEMS_NAME
      LENDER_ITEMS_AMOUNT = O.LENDER_ITEMS_AMOUNT
End Function





