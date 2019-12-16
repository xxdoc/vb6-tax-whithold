VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTaxDocItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Conn As ADODB.Connection

Public AddEditMode As SHOW_MODE_TYPE
Public QueryFlag As Long
Public Flag As String
' End of generic data member

Public TAXDOC_ITEM_ID As Long
Public TAX_DOCUMENT_ID As Long
Public SUPPLIER_ID As Long
Public ADDRESS_ID As Long
Public REVENUE_TYPE As Double
Public TAX_RATE As Double
Public PAY_DATE As Date
Public CONDITION_ID As Long
Public PAID_AMOUNT As Double
Public WH_AMOUNT As Double
Public REF_NO As String
Public NOTE As String
Public BRANCH As String
Public BRANCH_ID As Long
Public KEY_CODE As String
Public MASTER_AREA As Long
Public LENDER_ITEMS_ID As Long

Public SUPPLIER_CODE As String
Public SUPPLIER_NAME As String
Public REVENUE_TYPE_NAME As String
Public REVENUE_NAME As String
Public RATETYPE_NAME As String
Public CONDITION_NAME As String

Public TAX_ID As String
Public IDENT_ID As String
Public KEY_NAME As String
Public DPACC_ID As String
Public PS_TYPE As String

Public EXPORT_KEY As String

Public Address As String
Public HOME As String
Public SOI As String
Public MOO As String
Public VILLAGE As String
Public ROAD As String
Public DISTRICT As String
Public AMPHUR As String
Public PROVINCE As String
Public ZIPCODE As String
Private BANGKOK_FLAG As String
Public BRANCH_CODE As String

Public FROM_DATE As Date
Public TO_DATE As Date
Public OrderBy As Long
Public OrderType As Long
Public TAX_TYPE As Long
Public FROM_PAY_DATE As Date
Public TO_PAY_DATE As Date

Public FROM_SUPPLIER_CODE As String
Public TO_SUPPLIER_CODE As String


Public TotalInterest As String
Public TotalVat As String
Public TotalInterestYear As String
Public TempDiff As String
Public TotalVatYear As Double
Public YYYYMM As String

Public Addr As CAddress
Public Document_Year As String
Public Company_ID As String

Private Sub Class_Initialize()
   Set m_Conn = glbDatabaseMngr.DBConnection
   FROM_DATE = -1
   TO_DATE = -1
   FROM_PAY_DATE = -1
   TO_PAY_DATE = -1
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
Dim tempDate As String
Dim SubLen As Long
Dim NewStr As String
Dim GroupByStr As String

   If Ind = 1 Then
      SelectStr1 = "SELECT UG.*, TD.*,ADDR.*,SP.*,NM.LONG_NAME SUPPLIER_NAME, CT.KEY_NAME CONDITION_NAME," & _
      " RT.KEY_NAME RATETYPE_NAME, RV.KEY_NAME REVENUE_NAME,RV.EXPORT_KEY EXPORT_KEY , BR.KEY_NAME BRANCH_CODE, " & _
      " MSRF.KEY_NAME PS_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM TAXDOC_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN TAX_DOCUMENT TD ON (UG.TAX_DOCUMENT_ID = TD.TAX_DOCUMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (UG.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SP.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (SN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF CT ON (UG.CONDITION_ID = CT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF RT ON (UG.TAX_RATE = RT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF RV ON (UG.REVENUE_TYPE = RV.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ADDRESS ADDR  ON (ADDR.ADDRESS_ID = UG.ADDRESS_ID) "
     FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MSRF ON (SP.SUPPLIER_TYPE = MSRF.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF BR ON (UG.BRANCH_ID = BR.KEY_ID) "
      
      WhereStr = ""
      
   ElseIf Ind = 2 Then
      SelectStr1 = "SELECT UG.*, TD.*,ADDR.*,SP.*,NM.LONG_NAME SUPPLIER_NAME, CT.KEY_NAME CONDITION_NAME," & _
      " RT.KEY_NAME RATETYPE_NAME, RV.KEY_NAME REVENUE_NAME , BR.KEY_NAME BRANCH_CODE, " & _
      " MSRF.KEY_NAME PS_TYPE "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM TAXDOC_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN TAX_DOCUMENT TD ON (UG.TAX_DOCUMENT_ID = TD.TAX_DOCUMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (UG.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SP.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (SN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF CT ON (UG.CONDITION_ID = CT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF RT ON (UG.TAX_RATE = RT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF RV ON (UG.REVENUE_TYPE = RV.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ADDRESS ADDR  ON (ADDR.ADDRESS_ID = UG.ADDRESS_ID) "
     FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MSRF ON (SP.SUPPLIER_TYPE = MSRF.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF BR ON (UG.BRANCH_ID = BR.KEY_ID) "
      
      WhereStr = ""
  ElseIf Ind = 3 Then
      SelectStr1 = "SELECT NM.LONG_NAME SUPPLIER_NAME, MSRF.KEY_NAME PS_TYPE , RT.KEY_NAME RATETYPE_NAME , " & _
      " UG.SUPPLIER_ID , SP.IDENT_ID,SP.TAX_ID,  UG.ADDRESS_ID , " & _
      " SUM(UG.PAID_AMOUNT) PAID_AMOUNT , SUM(UG.WH_AMOUNT) WH_AMOUNT  "         ' µÈÕß¡’‡≈¢°”°—∫¡—È¬ CT.KEY_NAME CONDITION_NAME,
'      " RT.KEY_NAME RATETYPE_NAME, RV.KEY_NAME REVENUE_NAME , BR.KEY_NAME BRANCH_CODE, " & _
'      " MSRF.KEY_NAME PS_TYPE , TD.*   ,  ADDR.*   ,  SP.*  ,  NM.LONG_NAME SUPPLIER_NAME  "
   
'      SelectStr1 = "SELECT CT.KEY_NAME CONDITION_NAME,  NM.LONG_NAME SUPPLIER_NAME, " & _
'      " UG.TAXDOC_ITEM_ID  , UG.TAX_DOCUMENT_ID , UG.SUPPLIER_ID , UG.ADDRESS_ID , UG.REVENUE_TYPE , UG.TAX_RATE , UG.PAY_DATE, UG.CONDITION_ID , UG.REF_NO , UG.NOTE , UG.BRANCH ,UG.BRANCH_ID , " & _
'      " SUM(UG.PAID_AMOUNT) PAID_AMOUNT , SUM(UG.WH_AMOUNT) WH_AMOUNT  "
''      " RT.KEY_NAME RATETYPE_NAME, RV.KEY_NAME REVENUE_NAME , BR.KEY_NAME BRANCH_CODE, " & _
''      " MSRF.KEY_NAME PS_TYPE , TD.*   ,  ADDR.*   ,  SP.*  ,  NM.LONG_NAME SUPPLIER_NAME  "
      SelectStr2 = "SELECT COUNT(*)  "
      
      FromStr = " FROM TAXDOC_ITEM UG "
      FromStr = FromStr & "LEFT OUTER JOIN TAX_DOCUMENT TD ON (UG.TAX_DOCUMENT_ID = TD.TAX_DOCUMENT_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER SP ON (UG.SUPPLIER_ID = SP.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN SUPPLIER_NAME SN ON (SP.SUPPLIER_ID = SN.SUPPLIER_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN NAME NM ON (SN.NAME_ID = NM.NAME_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF CT ON (UG.CONDITION_ID = CT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF RT ON (UG.TAX_RATE = RT.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF RV ON (UG.REVENUE_TYPE = RV.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN ADDRESS ADDR  ON (ADDR.ADDRESS_ID = UG.ADDRESS_ID) "
     FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF MSRF ON (SP.SUPPLIER_TYPE = MSRF.KEY_ID) "
      FromStr = FromStr & "LEFT OUTER JOIN MASTER_REF BR ON (UG.BRANCH_ID = BR.KEY_ID) "

      
      WhereStr = ""

   End If
   
   If TAXDOC_ITEM_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.TAXDOC_ITEM_ID = " & TAXDOC_ITEM_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.TAXDOC_ITEM_ID = " & TAXDOC_ITEM_ID & ") "
      End If
   End If
   
   If TAX_DOCUMENT_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ") "
      End If
   End If
   
   If ADDRESS_ID > 0 Then
      If Len(WhereStr) <= 0 Then
         WhereStr = "WHERE (UG.ADDRESS_ID = " & ADDRESS_ID & ") "
      Else
         WhereStr = WhereStr & "AND (UG.ADDRESS_ID = " & ADDRESS_ID & ") "
      End If
   End If

   If Len(Document_Year) > 0 Then
      If WhereStr = "" Then
         WhereStr = "WHERE (TD.DOCUMENT_DATE  LIKE  '" & Document_Year & "%')"
      Else
         WhereStr = WhereStr & "AND (TD.DOCUMENT_DATE  LIKE  '" & Document_Year & "%')"
      End If
   End If
   

   If Len(Company_ID) > 0 Then
      If WhereStr = "" Then
      WhereStr = "WHERE (TD.COMPANY_ID  =  '" & Company_ID & "')"
      Else
      WhereStr = WhereStr & "AND (TD.COMPANY_ID  =  '" & Company_ID & "')"
      End If
   End If
   
   If FROM_DATE > -1 Then
      tempDate = DateToStringIntLow(Trim(FROM_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (TD.DOCUMENT_DATE >= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (TD.DOCUMENT_DATE >= '" & ChangeQuote(tempDate) & "')"
      End If
   End If

   If TO_DATE > -1 Then
      tempDate = DateToStringIntHi(Trim(TO_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (TD.DOCUMENT_DATE <= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (TD.DOCUMENT_DATE <= '" & ChangeQuote(tempDate) & "')"
      End If
   End If

   If FROM_PAY_DATE > -1 Then
      tempDate = DateToStringIntLow(Trim(FROM_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.PAY_DATE >= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & "AND (UG.PAY_DATE >= '" & ChangeQuote(tempDate) & "')"
      End If
   End If

   If TO_PAY_DATE > -1 Then
      tempDate = DateToStringIntHi(Trim(TO_PAY_DATE))
      If WhereStr = "" Then
         WhereStr = "WHERE (UG.PAY_DATE <= '" & ChangeQuote(tempDate) & "')"
      Else
         WhereStr = WhereStr & " AND  (UG.PAY_DATE <= '" & ChangeQuote(tempDate) & "')"
      End If
   End If

   If BRANCH_CODE <> "" Then
         If WhereStr = "" Then
            WhereStr = WhereStr & "WHERE (BR.KEY_NAME = '" & BRANCH_CODE & "') "
         Else
            WhereStr = WhereStr & "AND (BR.KEY_NAME = '" & BRANCH_CODE & "') "
         End If
   End If
   
   If KEY_CODE <> "" Then
         If WhereStr = "" Then
            WhereStr = WhereStr & "WHERE (BR.KEY_CODE = " & KEY_CODE & ") "
         Else
            WhereStr = WhereStr & "AND (BR.KEY_CODE = " & KEY_CODE & ") "
         End If
   End If
   
   If MASTER_AREA > 0 Then
         If WhereStr = "" Then
            WhereStr = WhereStr & "WHERE (BR.MASTER_AREA = " & MASTER_AREA & ") "
         Else
            WhereStr = WhereStr & "AND (BR.MASTER_AREA = " & MASTER_AREA & ") "
         End If
   End If
   
   If TAX_TYPE > 0 Then
      If WhereStr = "" Then
         WhereStr = WhereStr & "WHERE (TAX_TYPE= " & TAX_TYPE & " )"
      Else
         WhereStr = WhereStr & "AND (TAX_TYPE = " & TAX_TYPE & " )"
      End If
   End If
   
 If Trim(FROM_SUPPLIER_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SUPPLIER_CODE >= '" & ChangeQuote(Trim(FROM_SUPPLIER_CODE)) & "')"
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_CODE >= '" & ChangeQuote(Trim(FROM_SUPPLIER_CODE)) & "')"
      End If
   End If
   
   If Trim(TO_SUPPLIER_CODE) <> "" Then
      If WhereStr = "" Then
         WhereStr = "WHERE (SUPPLIER_CODE <= '" & ChangeQuote(Trim(TO_SUPPLIER_CODE)) & "')"
      Else
         WhereStr = WhereStr & "AND (SUPPLIER_CODE <= '" & ChangeQuote(Trim(TO_SUPPLIER_CODE)) & "')"
      End If
   End If
   
  If WhereStr <> "" Then WhereStr = WhereStr & " AND (BR.MASTER_AREA =  13) "

   GroupByStr = ""
   If Ind = 3 Then
      GroupByStr = " GROUP BY UG.SUPPLIER_ID , NM.LONG_NAME ,MSRF.KEY_NAME  , RT.KEY_NAME , SP.IDENT_ID , SP.TAX_ID, UG.ADDRESS_ID "      ' , CT.KEY_NAME
   End If
   
   OrderByStr = ""
   OrderTypeStr = ""

   If Ind = 1 Then
      If OrderBy = 1 Then
         OrderByStr = "ORDER BY PAY_DATE "
      ElseIf OrderBy = 2 Then
         OrderByStr = "ORDER BY SUPLYER_NAME,PAY_DATE "
      ElseIf OrderBy = 3 Then
         OrderByStr = "ORDER BY BRANCH,PAY_DATE "
      Else
         OrderByStr = "ORDER BY  PAY_DATE ASC,REF_NO "
      End If
   ElseIf Ind = 2 Then
         OrderByStr = "ORDER BY UG.SUPPLIER_ID,UG.PAY_DATE  "
   End If
   
   If OrderType = 1 Then
      OrderTypeStr = "ASC "
   ElseIf OrderType = 2 Then
      OrderTypeStr = "DESC "
   Else
      OrderTypeStr = "ASC "
   End If
   
   SQL1 = SelectStr1 & FromStr & WhereStr & GroupByStr & OrderByStr & OrderTypeStr
   SQL2 = SelectStr2 & FromStr & WhereStr
'   'Debug.Print SQL1
   Call glbDatabaseMngr.GetRs(SQL1, SQL2, True, itemcount, Rs, glbErrorLog)                            'sql1
End Function

Public Function CopyObject(Ind As Long, O As CTaxDocItem)
   If Ind = 1 Then
      TAXDOC_ITEM_ID = O.TAXDOC_ITEM_ID
      TAX_DOCUMENT_ID = O.TAX_DOCUMENT_ID
      SUPPLIER_ID = O.SUPPLIER_ID
      ADDRESS_ID = O.ADDRESS_ID
      REVENUE_TYPE = O.REVENUE_TYPE
      TAX_RATE = O.TAX_RATE
      PAY_DATE = O.PAY_DATE
      WH_AMOUNT = O.WH_AMOUNT
      PAID_AMOUNT = O.PAID_AMOUNT
      CONDITION_ID = O.CONDITION_ID
      REF_NO = O.REF_NO
      NOTE = O.NOTE
     BRANCH = O.BRANCH
      BRANCH_ID = O.BRANCH_ID
      
      SUPPLIER_NAME = O.SUPPLIER_NAME
      REVENUE_TYPE_NAME = O.REVENUE_TYPE_NAME
      REVENUE_NAME = O.REVENUE_NAME
      RATETYPE_NAME = O.RATETYPE_NAME
      CONDITION_NAME = O.CONDITION_NAME

      HOME = O.HOME
      SOI = O.SOI
      MOO = O.MOO
      VILLAGE = O.VILLAGE
      ROAD = O.ROAD
      DISTRICT = O.DISTRICT
      AMPHUR = O.AMPHUR
      PROVINCE = O.PROVINCE
      ZIPCODE = O.ZIPCODE
'      BANGKOK_FLAG = O.BANGKOK_FLAG
      
      TAX_ID = O.TAX_ID
      IDENT_ID = O.IDENT_ID
      DPACC_ID = O.DPACC_ID
                                                                                             
      PS_TYPE = O.PS_TYPE
      BRANCH_CODE = O.BRANCH_CODE
      TAX_TYPE = O.TAX_TYPE
  ElseIf Ind = 2 Then
      TAXDOC_ITEM_ID = O.TAXDOC_ITEM_ID
      TAX_DOCUMENT_ID = O.TAX_DOCUMENT_ID
      SUPPLIER_ID = O.SUPPLIER_ID
      ADDRESS_ID = O.ADDRESS_ID
      REVENUE_TYPE = O.REVENUE_TYPE
      TAX_RATE = O.TAX_RATE
      PAY_DATE = O.PAY_DATE
      WH_AMOUNT = O.WH_AMOUNT
      PAID_AMOUNT = O.PAID_AMOUNT
      CONDITION_ID = O.CONDITION_ID
      REF_NO = O.REF_NO
      NOTE = O.NOTE
     BRANCH = O.BRANCH
      BRANCH_ID = O.BRANCH_ID
      
      SUPPLIER_NAME = O.SUPPLIER_NAME
      REVENUE_TYPE_NAME = O.REVENUE_TYPE_NAME
      REVENUE_NAME = O.REVENUE_NAME
      RATETYPE_NAME = O.RATETYPE_NAME
      CONDITION_NAME = O.CONDITION_NAME

      HOME = O.HOME
      SOI = O.SOI
      MOO = O.MOO
      VILLAGE = O.VILLAGE
      ROAD = O.ROAD
      DISTRICT = O.DISTRICT
      AMPHUR = O.AMPHUR
      PROVINCE = O.PROVINCE
      ZIPCODE = O.ZIPCODE
'      BANGKOK_FLAG = O.BANGKOK_FLAG
      
      TAX_ID = O.TAX_ID
      IDENT_ID = O.IDENT_ID
      DPACC_ID = O.DPACC_ID
                                                                                             
      PS_TYPE = O.PS_TYPE
      BRANCH_CODE = O.BRANCH_CODE
      TAX_TYPE = O.TAX_TYPE
    ElseIf Ind = 3 Then
      TAXDOC_ITEM_ID = O.TAXDOC_ITEM_ID
      TAX_DOCUMENT_ID = O.TAX_DOCUMENT_ID
      SUPPLIER_ID = O.SUPPLIER_ID
      ADDRESS_ID = O.ADDRESS_ID
      REVENUE_TYPE = O.REVENUE_TYPE
      TAX_RATE = O.TAX_RATE
      PAY_DATE = O.PAY_DATE
      WH_AMOUNT = O.WH_AMOUNT
      PAID_AMOUNT = O.PAID_AMOUNT
      CONDITION_ID = O.CONDITION_ID
      REF_NO = O.REF_NO
      NOTE = O.NOTE
     BRANCH = O.BRANCH
      BRANCH_ID = O.BRANCH_ID
      
      SUPPLIER_NAME = O.SUPPLIER_NAME
'      REVENUE_TYPE_NAME = O.REVENUE_TYPE_NAME
'      REVENUE_NAME = O.REVENUE_NAME
'      RATETYPE_NAME = O.RATETYPE_NAME
'      CONDITION_NAME = O.CONDITION_NAME
'
'      HOME = O.HOME
'      SOI = O.SOI
'      MOO = O.MOO
'      VILLAGE = O.VILLAGE
'      ROAD = O.ROAD
'      DISTRICT = O.DISTRICT
'      AMPHUR = O.AMPHUR
'      PROVINCE = O.PROVINCE
'      ZIPCODE = O.ZIPCODE
''      BANGKOK_FLAG = O.BANGKOK_FLAG
'
'      TAX_ID = O.TAX_ID
'      IDENT_ID = O.IDENT_ID
'      DPACC_ID = O.DPACC_ID
'
'      PS_TYPE = O.PS_TYPE
'      BRANCH_CODE = O.BRANCH_CODE
'      TAX_TYPE = O.TAX_TYPE
   End If
End Function

Public Function PopulateFromRS(Ind As Long, Rs As ADODB.Recordset)
   If Ind = 1 Then
      TAXDOC_ITEM_ID = NVLI(Rs("TAXDOC_ITEM_ID"), -1)
      TAX_DOCUMENT_ID = NVLI(Rs("TAX_DOCUMENT_ID"), -1)
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      REVENUE_TYPE = NVLD(Rs("REVENUE_TYPE"), 0)
      TAX_RATE = NVLD(Rs("TAX_RATE"), 0)
      PAY_DATE = InternalDateToDate(NVLS(Rs("PAY_DATE"), ""))
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CONDITION_ID = NVLI(Rs("CONDITION_ID"), -1)
      REF_NO = NVLS(Rs("REF_NO"), "")
      NOTE = NVLS(Rs("NOTE"), "")
     BRANCH = NVLS(Rs("CONDITION_NAME"), "")
      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      REVENUE_TYPE_NAME = NVLS(Rs("REVENUE_NAME"), "")
      REVENUE_NAME = NVLS(Rs("REVENUE_NAME"), "")
      EXPORT_KEY = NVLS(Rs("EXPORT_KEY"), "")
      RATETYPE_NAME = NVLS(Rs("RATETYPE_NAME"), "")
      CONDITION_NAME = NVLS(Rs("CONDITION_NAME"), "")
      LENDER_ITEMS_ID = NVLI(Rs("LENDER_ITEMS_ID"), -1)
      'ADDRESS_ID = NVLI(Rs("SUPPLIER_ADDRESS_ID"), -1)
      
      
      HOME = NVLS(Rs("HOME"), "")
      SOI = NVLS(Rs("SOI"), "")
      MOO = NVLS(Rs("MOO"), "")
      VILLAGE = NVLS(Rs("VILLAGE"), "")
      ROAD = NVLS(Rs("ROAD"), "")
      DISTRICT = NVLS(Rs("DISTRICT"), "")
      AMPHUR = NVLS(Rs("AMPHUR"), "")
      PROVINCE = NVLS(Rs("PROVINCE"), "")
      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "N")
      Address = PackAddress
   '  ADDRESS = HOME & SOI & MOO & VILLAGE & ROAD & DISTRICT & AMPHUR & PROVINCE & ZIPCODE
      
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      IDENT_ID = NVLS(Rs("IDENT_ID"), "")
      DPACC_ID = NVLS(Rs("DPACC_ID"), "")
                                                                                             ' From master referent
      PS_TYPE = NVLS(Rs("PS_TYPE"), "")
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      TAX_TYPE = NVLI(Rs("TAX_TYPE"), -1)
  ElseIf Ind = 2 Then
      TAXDOC_ITEM_ID = NVLI(Rs("TAXDOC_ITEM_ID"), -1)
      TAX_DOCUMENT_ID = NVLI(Rs("TAX_DOCUMENT_ID"), -1)
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
      REVENUE_TYPE = NVLD(Rs("REVENUE_TYPE"), 0)
      TAX_RATE = NVLD(Rs("TAX_RATE"), 0)
      PAY_DATE = InternalDateToDate(NVLS(Rs("PAY_DATE"), ""))
      YYYYMM = DateToStringYYYYMM(NVLS(Rs("PAY_DATE"), ""))
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
      CONDITION_ID = NVLI(Rs("CONDITION_ID"), -1)
      REF_NO = NVLS(Rs("REF_NO"), "")
      NOTE = NVLS(Rs("NOTE"), "")
      BRANCH = NVLS(Rs("CONDITION_NAME"), "")
      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      
      SUPPLIER_CODE = NVLS(Rs("SUPPLIER_CODE"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
      REVENUE_TYPE_NAME = NVLS(Rs("REVENUE_NAME"), "")
      REVENUE_NAME = NVLS(Rs("REVENUE_NAME"), "")
      RATETYPE_NAME = NVLS(Rs("RATETYPE_NAME"), "")
      CONDITION_NAME = NVLS(Rs("CONDITION_NAME"), "")                                                                                              '
      'ADDRESS_ID = NVLI(Rs("SUPPLIER_ADDRESS_ID"), -1)
      HOME = NVLS(Rs("HOME"), "")
      SOI = NVLS(Rs("SOI"), "")
      MOO = NVLS(Rs("MOO"), "")
      VILLAGE = NVLS(Rs("VILLAGE"), "")
      ROAD = NVLS(Rs("ROAD"), "")
      DISTRICT = NVLS(Rs("DISTRICT"), "")
      AMPHUR = NVLS(Rs("AMPHUR"), "")
      PROVINCE = NVLS(Rs("PROVINCE"), "")
      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "N")
      
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      IDENT_ID = NVLS(Rs("IDENT_ID"), "")
      DPACC_ID = NVLS(Rs("DPACC_ID"), "")
                                                                                             ' From master referent
      PS_TYPE = NVLS(Rs("PS_TYPE"), "")
      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
      TAX_TYPE = NVLI(Rs("TAX_TYPE"), -1)
   ElseIf Ind = 3 Then
'      TAXDOC_ITEM_ID = NVLI(Rs("TAXDOC_ITEM_ID"), -1)
'      TAX_DOCUMENT_ID = NVLI(Rs("TAX_DOCUMENT_ID"), -1)
      SUPPLIER_ID = NVLI(Rs("SUPPLIER_ID"), -1)
      ADDRESS_ID = NVLI(Rs("ADDRESS_ID"), -1)
'      REVENUE_TYPE = NVLD(Rs("REVENUE_TYPE"), 0)
'      TAX_RATE = NVLD(Rs("TAX_RATE"), 0)
'      PAY_DATE = InternalDateToDate(NVLS(Rs("PAY_DATE"), ""))
      WH_AMOUNT = NVLD(Rs("WH_AMOUNT"), 0)
      PAID_AMOUNT = NVLD(Rs("PAID_AMOUNT"), 0)
'      CONDITION_ID = NVLI(Rs("CONDITION_ID"), -1)
'      REF_NO = NVLS(Rs("REF_NO"), "")
'      NOTE = NVLS(Rs("NOTE"), "")
'     BRANCH = NVLS(Rs("CONDITION_NAME"), "")
'      BRANCH_ID = NVLI(Rs("BRANCH_ID"), -1)
      TAX_ID = NVLS(Rs("TAX_ID"), "")
      SUPPLIER_NAME = NVLS(Rs("SUPPLIER_NAME"), "")
'      REVENUE_TYPE_NAME = NVLS(Rs("REVENUE_NAME"), "")
'      REVENUE_NAME = NVLS(Rs("REVENUE_NAME"), "")
      RATETYPE_NAME = NVLS(Rs("RATETYPE_NAME"), "")
      IDENT_ID = NVLS(Rs("IDENT_ID"), "")
'      CONDITION_NAME = NVLS(Rs("CONDITION_NAME"), "")                                                                                              '
'      'ADDRESS_ID = NVLI(Rs("SUPPLIER_ADDRESS_ID"), -1)
'      HOME = NVLS(Rs("HOME"), "")
'      SOI = NVLS(Rs("SOI"), "")
'      MOO = NVLS(Rs("MOO"), "")
'      VILLAGE = NVLS(Rs("VILLAGE"), "")
'      ROAD = NVLS(Rs("ROAD"), "")
'      DISTRICT = NVLS(Rs("DISTRICT"), "")
'      AMPHUR = NVLS(Rs("AMPHUR"), "")
'      PROVINCE = NVLS(Rs("PROVINCE"), "")
'      ZIPCODE = NVLS(Rs("ZIPCODE"), "")
'      BANGKOK_FLAG = NVLS(Rs("BANGKOK_FLAG"), "N")
'
'      TAX_ID = NVLS(Rs("TAX_ID"), "")
'      IDENT_ID = NVLS(Rs("IDENT_ID"), "")
'      DPACC_ID = NVLS(Rs("DPACC_ID"), "")
'                                                                                             ' From master referent
      PS_TYPE = NVLS(Rs("PS_TYPE"), "")
'      BRANCH_CODE = NVLS(Rs("BRANCH_CODE"), "")
'      TAX_TYPE = NVLI(Rs("TAX_TYPE"), -1)
   End If
End Function

Public Function AddEditData()
Dim ServerDateTime As String
Dim SQL1 As String
Dim ID As Long
Dim ID2 As Long
Dim TempID1 As String

   Call glbDatabaseMngr.GetServerDateTime(ServerDateTime, glbErrorLog)
      
   If AddEditMode = SHOW_ADD Then
      Call glbDatabaseMngr.GetSeqID("TAXDOC_ITEM_SEQ", ID, glbErrorLog)
      
      SQL1 = "INSERT INTO TAXDOC_ITEM " & _
                     "( " & _
                     "   TAXDOC_ITEM_ID, " & _
                     "   TAX_DOCUMENT_ID, " & _
                     "   PAY_DATE, " & _
                     "   SUPPLIER_ID, " & _
                     "   ADDRESS_ID, " & _
                     "   REVENUE_TYPE, " & _
                     "   TAX_RATE, " & _
                     "   PAID_AMOUNT, " & _
                     "   WH_AMOUNT, " & _
                     "   CONDITION_ID, "
            SQL1 = SQL1 & _
                     "   REF_NO, " & _
                     "   NOTE, " & _
                     "   BRANCH_ID, " & _
                     "   LENDER_ITEMS_ID, " & _
                     "   CREATE_DATE, " & _
                    "    CREATE_BY, " & _
                    "    MODIFY_DATE, " & _
                    "    MODIFY_BY " & _
                     ") " & _
                     "VALUES " & _
                     "( " & _
                         ID & "," & _
                         "" & TAX_DOCUMENT_ID & "," & _
                         "'" & ChangeQuote(Trim(DateToStringInt(PAY_DATE))) & "'," & _
                         "" & SUPPLIER_ID & "," & _
                         "" & ADDRESS_ID & "," & _
                         "" & REVENUE_TYPE & "," & _
                         "" & TAX_RATE & "," & _
                         "" & PAID_AMOUNT & "," & _
                         "" & WH_AMOUNT & ","
            SQL1 = SQL1 & _
                        "" & CONDITION_ID & "," & _
                        "'" & ChangeQuote(Trim(REF_NO)) & "'," & _
                        "'" & ChangeQuote(Trim(NOTE)) & "'," & _
                        "'" & ChangeQuote(Trim(BRANCH_ID)) & "'," & _
                        "'" & ChangeQuote(Trim(LENDER_ITEMS_ID)) & "'," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & "," & _
                         "'" & ServerDateTime & "'," & _
                         glbUser.USER_ID & _
                     ") "
         TAXDOC_ITEM_ID = ID
   Else 'Show edit
      SQL1 = "UPDATE TAXDOC_ITEM SET " & _
                     "   TAX_DOCUMENT_ID = " & TAX_DOCUMENT_ID & ", " & _
                     "   PAY_DATE = '" & ChangeQuote(Trim(DateToStringInt(PAY_DATE))) & "', " & _
                     "   SUPPLIER_ID = " & SUPPLIER_ID & ", " & _
                     "   ADDRESS_ID = " & ADDRESS_ID & ", " & _
                     "   REVENUE_TYPE = " & REVENUE_TYPE & ", " & _
                     "   TAX_RATE = " & TAX_RATE & ", " & _
                     "   PAID_AMOUNT = " & PAID_AMOUNT & ", " & _
                     "   WH_AMOUNT = " & WH_AMOUNT & ", " & _
                     "   CONDITION_ID = " & CONDITION_ID & ", " & _
                     "   REF_NO = '" & ChangeQuote(Trim(REF_NO)) & "', " & _
                     "   NOTE = '" & ChangeQuote(Trim(NOTE)) & "', " & _
                     "   BRANCH_ID = '" & ChangeQuote(Trim(BRANCH_ID)) & "', " & _
                     "   LENDER_ITEMS_ID = '" & ChangeQuote(Trim(LENDER_ITEMS_ID)) & "', " & _
                    "    MODIFY_DATE = '" & ServerDateTime & "', " & _
                    "    MODIFY_BY = " & glbUser.USER_ID & " " & _
                     "WHERE TAXDOC_ITEM_ID = " & TAXDOC_ITEM_ID
   End If
   
   If (AddEditMode = SHOW_ADD) Or (AddEditMode = SHOW_EDIT) Then
      m_Conn.Execute (SQL1)
   End If
End Function

Public Function DeleteData()
Dim SQL1 As String
      SQL1 = "DELETE FROM TAXDOC_ITEM WHERE TAXDOC_ITEM_ID = " & TAXDOC_ITEM_ID
      m_Conn.Execute (SQL1)
End Function

Public Function PackAddress() As String
Dim AddressStr As String

   AddressStr = ""
   
   If Trim(HOME) <> "" Then
      AddressStr = AddressStr & HOME & " "
   End If

   If Trim(SOI) <> "" Then
      AddressStr = AddressStr & "´." & SOI & " "
   End If

   If Trim(MOO) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË") & MOO & " "
   End If

   If Trim(VILLAGE) <> "" Then
      AddressStr = AddressStr & MapText("À¡ŸË∫È“π") & VILLAGE & " "
   End If

   If Trim(ROAD) <> "" Then
      AddressStr = AddressStr & MapText("∂.") & ROAD & " "
   End If

   If Trim(DISTRICT) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("·¢«ß") & DISTRICT & " "
      Else
         AddressStr = AddressStr & MapText("µ”∫≈") & DISTRICT & " "
      End If
   End If

   If Trim(AMPHUR) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & MapText("‡¢µ") & AMPHUR & " "
      Else
         AddressStr = AddressStr & MapText("Õ”‡¿Õ") & AMPHUR & " "
      End If
   End If

   If Trim(PROVINCE) <> "" Then
      If BANGKOK_FLAG = "Y" Then
         AddressStr = AddressStr & "" & PROVINCE & " "
      Else
         AddressStr = AddressStr & MapText("®.") & PROVINCE & " "
      End If
   End If

   If Trim(ZIPCODE) <> "" Then
      AddressStr = AddressStr & " " & ZIPCODE & " "
   End If

   PackAddress = AddressStr
End Function

