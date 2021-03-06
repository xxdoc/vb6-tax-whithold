VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGLJnl"
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

Public VOUCHER As String
Public VOUDAT As Date
Public ACCNUM As String
Public DEPCOD As String
Public DESCRP As String
Public TRNTYP As String
Public AMOUNT As Double
Public JNLTYP As String

Public ACCNAM As String

Public FROM_VOUCHER_DATE As Date
Public TO_VOUCHER_DATE As Date
Public AccNumSet As String

Public KEY_ID As Long
Public KEY_LOOKUP As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBLegacyConnection
   FROM_VOUCHER_DATE = -1
   TO_VOUCHER_DATE = -1
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
Dim TempDate As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT ART.*, DTOS(ART.VOUDAT) VOUDAT_STR, GLA.ACCNAM "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GLJNL ART "
      FromStr = FromStr & "LEFT OUTER JOIN GLACC GLA ON (ART.ACCNUM = GLA.ACCNUM) "
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT DISTINCT ART.VOUCHER, DTOS(ART.VOUDAT) VOUDAT_STR "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GLJNL ART "
      WhereStr = "WHERE (ART.VOUCHER NOT IN (SELECT APT.DOCNUM FROM APTRN APT WHERE RECTYP = '9')) "
   ElseIf Ind = 3 Then
      SelectStr1 = "SELECT DISTINCT ART.VOUCHER, DTOS(ART.VOUDAT) VOUDAT_STR "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = "FROM GLJNL ART "
      WhereStr = ""
   End If
   
   If Trim(JNLTYP) <> "" Then
      If WildCard(JNLTYP, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (JNLTYP LIKE '" & ChangeQuote(Trim(JNLTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (JNLTYP LIKE '" & ChangeQuote(Trim(JNLTYP)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (JNLTYP = '" & ChangeQuote(Trim(JNLTYP)) & "')"
         Else
            WhereStr = WhereStr & "AND (JNLTYP = '" & ChangeQuote(Trim(JNLTYP)) & "')"
         End If
      End If
   End If
   
   If Trim(VOUCHER) <> "" Then
      If WildCard(VOUCHER, SubLen, NewStr) Then
         If WhereStr = "" Then
            WhereStr = "WHERE (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         Else
            WhereStr = WhereStr & "AND (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         End If
      Else
         If WhereStr = "" Then
            WhereStr = "WHERE (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         Else
            WhereStr = WhereStr & "AND (VOUCHER = '" & ChangeQuote(Trim(VOUCHER)) & "')"
         End If
      End If
   End If
                                       
    If FROM_VOUCHER_DATE > -1 Then
      TempDate = DateToStringIntLowLeg(Trim(FROM_VOUCHER_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(VOUDAT) >= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(VOUDAT) >= '" & ChangeQuote(TempDate) & "')"
      End If
   End If

   If TO_VOUCHER_DATE > -1 Then
      TempDate = DateToStringIntHiLeg(Trim(TO_VOUCHER_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (DTOS(VOUDAT) <= '" & ChangeQuote(TempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (DTOS(VOUDAT) <= '" & ChangeQuote(TempDate) & "')"
      End If
   End If
                             
   If Len(AccNumSet) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (ACCNUM IN " & AccNumSet & ")"
      Else
         WhereStr = WhereStr & "AND (ACCNUM IN " & AccNumSet & ")"
      End If
   End If
   
   GroupByStr = ""
   
   OrderByStr = ""
   OrderTypeStr = ""
   
   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY VOUCHER "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY VOUCHER "
      Else
         OrderByStr = "ORDER BY VOUCHER "
      End If
   ElseIf Ind = 2 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY VOUDAT ASC, VOUCHER "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY VOUDAT ASC, VOUCHER "
      Else
         OrderByStr = "ORDER BY VOUDAT ASC, VOUCHER "
      End If
   ElseIf Ind = 3 Then
      'OrderByStr = "ORDER BY VOUDAT ASC, VOUCHER "
      OrderByStr = " ORDER BY VOUCHER "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   If Len(OrderByStr) > 0 Then
      SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   Else
      SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr
   End If
   SQL2 = SelectStr2 & FromStr & WhereStr
   
   Call glbDatabaseMngr.GetLegacyRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      VOUDAT = InternalDateToDate(NVLS(Rs("VOUDAT_STR"), ""))
      VOUCHER = NVLS(Rs("VOUCHER"), "")
      ACCNUM = NVLS(Rs("ACCNUM"), "")
      DEPCOD = NVLS(Rs("DEPCOD"), "")
      DESCRP = NVLS(Rs("DESCRP"), "")
      TRNTYP = NVLS(Rs("TRNTYP"), "")
      AMOUNT = NVLD(Rs("AMOUNT"), 0)
      JNLTYP = NVLS(Rs("JNLTYP"), "")
      ACCNAM = NVLS(Rs("ACCNAM"), "")
      
      KEY_LOOKUP = ACCNUM
   ElseIf Ind = 2 Then
      VOUDAT = InternalDateToDate(NVLS(Rs("VOUDAT_STR"), ""))
      VOUCHER = NVLS(Rs("VOUCHER"), "")
   ElseIf Ind = 3 Then
      VOUDAT = InternalDateToDateEx3(NVLS(Rs("VOUDAT_STR"), ""))
      VOUCHER = NVLS(Rs("VOUCHER"), "")
   End If
End Function



