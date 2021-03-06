VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportTaxSendingYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportTaxSendingYear"
Public MODULE_DESC As String

Private WithEvents Vsp As VSPrinter
Attribute Vsp.VB_VarHelpID = -1
Private mstrErrMsg As String
Private mcolParam As Collection
Private mblnNewPage As Boolean
Private mblnHeader As Boolean
Private mblnEndPage As Boolean
Private mdblWid   As Double
Private mdteDate As Date
Private mdY As Double
Private Rs As ADODB.Recordset
Private m_Rs As ADODB.Recordset

Private Const TITLE_SIZE = 11
Private Const HEADER_SIZE = 11
Private Const DETAIL_SIZE = 11

Private Const TITLE_FORMAT = "<28|<3|<5;"

Private m_TempCol As Collection
Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers1_1 As CFieldList
Private m_Headers1_2 As CFieldList
Private m_Headers1_3 As CFieldList
Private m_Headers1_4 As CFieldList
Private m_Headers1_5 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection

Private Total1(100) As Double
Private Total2(100) As Double
Private Total3(100) As Double

Private PageAll As Long
Private Ct As CTaxDocument
Private Ti As CTaxDocItem
Private Tii As CTaxDocItem
Private RCount As Long
Private I2 As Long
Private iCount As Long
Private IsOK As Boolean
Private SupAddr As CAddress
Private m_SupAddr As Collection

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
      Set m_Headers1_5 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
   Set SupAddr = New CAddress
   Set m_SupAddr = New Collection
End Sub

Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   Set m_TempCol = Nothing
   
   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
   Set m_Headers1_1 = Nothing
   Set m_Headers1_2 = Nothing
   Set m_Headers1_3 = Nothing
   Set m_Headers1_4 = Nothing
      Set m_Headers1_5 = Nothing
   Set m_Details2 = Nothing
      Set SupAddr = Nothing
   Set m_SupAddr = Nothing
End Sub

Private Function CReportInterface_AddParam(varItem As Variant, strIndex As String) As Boolean
   Call mcolParam.Add(varItem, strIndex)
   CReportInterface_AddParam = True
End Function

Private Sub CReportInterface_ClearParam()
   Call ClearParam
End Sub

Private Property Get CReportInterface_ErrorMsg() As String
   CReportInterface_ErrorMsg = mstrErrMsg
End Property

Private Function CReportInterface_Preview() As Boolean
   CReportInterface_Preview = genDoc(True)
End Function

Private Function CReportInterface_PrintDoc() As Boolean
   CReportInterface_PrintDoc = genDoc(False)
End Function

Private Property Set CReportInterface_VsPrint(RHS As VSPrinter7LibCtl.IVSPrinter)
   Set Vsp = RHS
End Property

Private Sub ClearParam()
   Dim i As Long
   
   For i = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next i

End Sub
'============================= Codes above used in every report ==========================

Private Sub printHeader()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   
   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers1_1.ClearField
   Call m_Headers1_2.ClearField
   Call m_Headers1_3.ClearField
   
   Set Cf = New CReportField
                                                                                                                    'Head 1
           Call Cf.SetFieldValue(1, "<", "(����¡��͡��¡����Ṻ������Թ�����л�����  ���������ͧ����  ' / ' ŧ㹪�ͧ��ҧ ˹�Ң�ͤ���������ó� ��§�������)" & vbCrLf & _
            "�������Թ��           ......( 1 )  �������Թ�����ҵ�� 40 (4)(�) �͡�����Թ�ҡ  �͡���¾ѹ��ѵ�  �͡���µ����Թ  ���                ......( 3 )  �Թ����ҵ�� 40 (4) (�) �Ż���ª������ҡ����͹��� ���" & vbCrLf & _
            "                                 ......( 2 )  �Թ�����ҵ��  40 (4) (�) �Թ�ѹ�� ���                                                                                          ......( 4 )  �Թ�����ҵ�� 40 (4)  ����", "^")
   Call m_Headers1_2.AddField(Cf)
   Set Cf = Nothing
  Call m_Headers1_2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)

        Set Cf = New CReportField
              Call Cf.SetFieldValue(1, "^", "��", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(10, "^", "�Ţ��Шӵ�Ǽ�����������ҡ�(13��ѡ)* �ͧ����������", "<")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(1, "^", "�ѵ��", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ�����·�駻�", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "�ӹǹ���շ���ѡ", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(1, "^", "���͹", "^")
                Call m_Headers1_1.AddField(Cf)
   Set Cf = Nothing
  Call m_Headers1_1.GetString(2, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
                                                                                                                            'Head 2
        Set Cf = New CReportField
              Call Cf.SetFieldValue(1, "^", "�Ѻ", "^")
                Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(5, "^", "�Ţ���ѭ���Թ�ҡ(�ͧ������Թ��)", "<")
                Call m_Headers1.AddField(Cf)
          Call Cf.SetFieldValue(5, "^", "���ͼ�����Թ�� (����к����Ѵਹ��ҹ�� ��� �ҧ �ҧ��� ������)", "<")
                Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(1, "^", "����", "^")
                Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "", ">")
                Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "��й��觷�駻�", ">")
                Call m_Headers1.AddField(Cf)
            Call Cf.SetFieldValue(1, "^", "�", "^")
                Call m_Headers1.AddField(Cf)

    Set Cf = Nothing
   
Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)


           Set Cf = New CReportField
              Call Cf.SetFieldValue(1, "^", "���", "^")
                Call m_Headers1_3.AddField(Cf)
            Call Cf.SetFieldValue(10, "<", "�������ͧ������Թ�� (����к��Ţ��� ��͡/��� ��� �Ӻ�/�ǧ �����/ࢵ �ѧ��Ѵ)", "<")
                Call m_Headers1_3.AddField(Cf)
            Call Cf.SetFieldValue(1, "^", "������", "^")
                Call m_Headers1_3.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "", "^")
                Call m_Headers1_3.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "", "^")
                Call m_Headers1_3.AddField(Cf)
            Call Cf.SetFieldValue(1, "^", "", "^")
                Call m_Headers1_3.AddField(Cf)
   Set Cf = Nothing
  Call m_Headers1_3.GetString(2, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   
     Set TempRs = Nothing
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
End Sub

Private Function initDoc() As Boolean
Dim strDate As String
Dim TempStr1 As String
Dim TempStr2 As String
Dim TempStr3 As String
Dim FromDate As String
Dim ToDate As String
Dim TAX_ID As String
Dim IDENT_ID As String


   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orLandscape
   Vsp.MarginBottom = 200
   Vsp.MarginFooter = 100
   Vsp.MarginHeader = 200
   Vsp.MarginLeft = 700
   Vsp.MarginRight = 700
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

    
    Call LoadSupAddr(Nothing, m_SupAddr, -1, True)
    
   Call getTaxIDfromBranch(, , mcolParam("BRANCH"), TAX_ID)
                                                                                                                                                                                                                                                               
    If Len(TAX_ID) = 13 Then
         TAX_ID = toTax_ID(TAX_ID)
   Else
       TAX_ID = TAX_ID
   End If
    MODULE_DESC = vbCrLf & "�Ţ��Шӵ�Ǽ�����������ҡ�(13��ѡ)*  (�ͧ�����˹�ҷ���ѡ���� � ������)     " & TAX_ID & vbCrLf & _
 "                                                                                                                                                                                                                                                                 �Ңҷ��        00000    "

      Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
    
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function

Private Function genDoc(isPreview As Boolean) As Boolean
On Error GoTo ErrHandler
Dim RName As String
Dim i As Long
Dim j As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim Ug  As CUserGroup
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim strPrint As String

'Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As Long
Dim PrevKey2 As String
Dim PrevKey3 As Long
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim PrevKeyi As Long
Dim ID As Long
Dim ReportKey As String
Dim Report As CReportInterface
Dim TType As Long
Dim RsLine As Integer
Dim E_NAME As String
      Dim Addr As String
      
Dim TAXID As String
Dim r As Long

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
   Set m_Rs = New ADODB.Recordset
     Set Rs = New ADODB.Recordset
   If Not initDoc Then Exit Function
   
   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------

   For j = 1 To UBound(Total1)
      Total1(j) = 0
      Total2(j) = 0
      Total3(j) = 0
   Next j

    r = 0

   Set Ti = New CTaxDocItem
   Set Tii = New CTaxDocItem

      Tii.BRANCH_CODE = mcolParam("BRANCH")
'      Tii.MASTER_AREA = 13
      Tii.FROM_PAY_DATE = mcolParam("FROM_DATE")
      Tii.TO_PAY_DATE = mcolParam("TO_DATE")
      
      Tii.FROM_SUPPLIER_CODE = mcolParam("FROM_SUPPLIER_CODE")
      Tii.TO_SUPPLIER_CODE = mcolParam("TO_SUPPLIER_CODE")
      
      Tii.OrderType = mcolParam("ORDER_TYPE")
      Tii.TAX_TYPE = mcolParam("TAX_TYPE")
      
      Call Tii.QueryData(2, m_Rs, iCount)
   
      i = 0
  ' 'Debug.Print (m_Rs.RecordCount)
   Vsp.StartDoc
   Set BodyCf = New CReportField
  If m_Rs.RecordCount > 0 Then
     While Not m_Rs.EOF
        Call Tii.PopulateFromRS(2, m_Rs)
           If PrevKey1 <> Tii.SUPPLIER_ID Then
                       i = i + 1
                       PrevKey1 = Tii.SUPPLIER_ID
            End If
            
            If (i Mod 5) = 0 And i <> PrevKeyi Then
                  PageAll = PageAll + 1
                  PrevKeyi = i
            End If
               If Not m_Rs.EOF Then
               m_Rs.MoveNext
            End If
    Wend
   
         If i / 5 > PageAll Then
         PageAll = PageAll + 1
         End If
         
         m_Rs.MoveFirst
         
          i = 0
          PrevKeyi = 0
          PrevKey1 = 0
'         Vsp.StartDoc
'      Set BodyCf = New CReportField
   
         While Not m_Rs.EOF
            Call Ti.PopulateFromRS(2, m_Rs)
            
             If PrevKey1 = Ti.SUPPLIER_ID Then
                       Total2(4) = Total2(4) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�繤��
                       Total2(5) = Total2(5) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�繤��
                        Total1(2) = Total1(2) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�Թ��駻�
                         Total1(3) = Total1(3) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�Թ��駻�
   '                   PrevKey6 = Ti.PAID_AMOUNT
   '                   PrevKey7 = Ti.WH_AMOUNT
            Else                       ' ������ѹ��ҧ�ѹ��������ѹ���
   If PrevKey1 <> 0 Then
          i = i + 1
            I2 = I2 + 1
            Call m_Details1.ClearField
   
            Set HeadCf = m_Headers1_1.Fields(1)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         
        '  'Debug.Print Ti.TAX_DOCUMENT_ID
          If PrevKey2 <> "" Then
               TAXID = toTax_ID(PrevKey2)
         Else
               TAXID = PrevKey6
         End If
   
        Set SupAddr = GetSupAddr(m_SupAddr, PrevKey1 & "-" & PrevKey3, False)
        If SupAddr Is Nothing Then
               Addr = ""
        Else
               Addr = SupAddr.PackAddress
         End If
            Set HeadCf = m_Headers1_1.Fields(2)
          TempStr = TAXID & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & PrevKey4 & vbCrLf & Addr                                   ' ���ʻ�ЪҪ�    '  ����  ��鹺�÷Ѵ����  ' �������
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(3)
            TempStr = PrevKey5             ' ���� 15 %
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
   
            Set HeadCf = m_Headers1_1.Fields(4)                                                         '�ؤ�Ÿ�����
            If PrevKey7 = "�ؤ�Ÿ�����" Or PrevKey7 = "�ԵԺؤ��" Then
                     TempStr = FormatNumber(Total2(4))
                     'Debug.Print CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", ""))
      '               Total1(2) = Total1(2) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�Թ��駻�
      '               Total2(4) = Total2(4) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�繤��
   
            Else
                     TempStr = "-"
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
           
   
            Set HeadCf = m_Headers1_1.Fields(5)
            TempStr = FormatNumber(Total2(5))                '�����ѡ��»�
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
    '        Total1(3) = Total1(3) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "0"), ",", ""))   'Ti.WH_AMOUNT
   '         Total2(4) = Total2(4) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�繤��
   
   
             Set HeadCf = m_Headers1_1.Fields(6)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(9) = Total1(9) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "0"), ",", ""))   'Ti.WH_AMOUNT
   
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            Vsp.FontBold = False
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            'Rs.MoveNext
            
      End If
               For j = 1 To UBound(Total1)
                  Total2(j) = 0
               Next j
            
                       PrevKey1 = Ti.SUPPLIER_ID
                       PrevKey2 = Ti.IDENT_ID
                       PrevKey3 = Ti.ADDRESS_ID
                       PrevKey4 = Ti.SUPPLIER_NAME
                       PrevKey5 = Ti.RATETYPE_NAME
                       PrevKey6 = Ti.TAX_ID
                       PrevKey7 = Ti.PS_TYPE
                      
                      Total2(4) = Total2(4) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�繤��
                       Total2(5) = Total2(5) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�繤��
                        Total1(2) = Total1(2) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�Թ��駻�
                         Total1(3) = Total1(3) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT         ' �ӹǹ�Թ��駻�
     End If
     
           If (i Mod 5) = 0 And i <> 0 And i <> PrevKeyi Then
               If r < PageAll - 1 Then 'Still left
   '            'Debug.Print i
                   r = r + 1
                  Vsp.NewPage
                  I2 = 0
                  PrevKeyi = i
               End If
           End If
   
           If Not m_Rs.EOF Then
               m_Rs.MoveNext
            End If
            
            
         Wend
         
    i = i + 1
            I2 = I2 + 1
            Call m_Details1.ClearField
   
            Set HeadCf = m_Headers1_1.Fields(1)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
         
        '  'Debug.Print Ti.TAX_DOCUMENT_ID
          If PrevKey2 <> "" Then
               TAXID = toTax_ID(PrevKey2)
         Else
               TAXID = PrevKey6
         End If
   
        Set SupAddr = GetSupAddr(m_SupAddr, PrevKey1 & "-" & PrevKey3, False)
        If SupAddr Is Nothing Then
               Addr = ""
        Else
               Addr = SupAddr.PackAddress
         End If
            Set HeadCf = m_Headers1_1.Fields(2)
          TempStr = TAXID & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & PrevKey4 & vbCrLf & Addr                                   ' ���ʻ�ЪҪ�    '  ����  ��鹺�÷Ѵ����  ' �������
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(3)
            TempStr = PrevKey5             ' ���� 15 %
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            
   
            Set HeadCf = m_Headers1_1.Fields(4)                                                         '�ؤ�Ÿ�����
            If PrevKey7 = "�ؤ�Ÿ�����" Or PrevKey7 = "�ԵԺؤ��" Then
                     TempStr = FormatNumber(Total2(4))
                     'Debug.Print CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", ""))
            Else
                     TempStr = "-"
            End If
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
           
   
            Set HeadCf = m_Headers1_1.Fields(5)
            TempStr = FormatNumber(Total2(5))                '�����ѡ��»�
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
             Set HeadCf = m_Headers1_1.Fields(6)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
            Total1(9) = Total1(9) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "0"), ",", ""))   'Ti.WH_AMOUNT
   
            Call m_Details1.GetString(1, TempStr1, TempStr2)
            Vsp.FontBold = False
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
            
            Total1(2) = Total1(2) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", ""))
             Total1(3) = Total1(3) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", ""))
            
      If I2 < 4 Then
      'Call PatchTable(Vsp, strFormat, "|||||||||", 500)
      '===
       RCount = I2
        
       For I2 = 0 To 4 - RCount
            i = i + 1
            Call m_Details1.ClearField
   
            Set HeadCf = m_Headers1_1.Fields(1)
            TempStr = i
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr & vbCrLf & "")
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(2)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(3)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(4)
            TempStr = "" & vbCrLf & "" & vbCrLf & ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(5)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
            Set HeadCf = m_Headers1_1.Fields(6)
            TempStr = ""
            Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
            Call m_Details1.AddField(BodyCf)
   
           Call m_Details1.GetString(1, TempStr1, TempStr2)
            strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
            Call Vsp.AddTable(strFormat, "", TempStr2)
         Next I2
       
   End If                                                                                                              'Footer Report
End If
    If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing

   Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", strPrint)
   Vsp.FontBold = False

   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
      
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(TXT1 As String, TXT2 As String, TXT3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers0.Fields
      j = j + 1
      If j = 1 Then
         TempStr = TXT1
      ElseIf (j = 2) Then
         TempStr = FormatNumber(Tot(j))
    ElseIf (j = 3) Then
         TempStr = FormatNumber(Tot(j))
      Else:
         TempStr = ""
      End If
      
   Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
   Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub

Private Sub VSP_EndDoc()
'This event occur when VSPrinter.EndDoc is used
End Sub

Private Sub VSP_EndPage()
Dim strFormat As String
Dim strPrint As String
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim iSize As Integer
Dim alngX() As Long
Dim PrevID As Long
Dim TempStr1 As String
Dim TempStr1_1 As String
Dim Cf As CReportField
Dim ExportWidth As Long
Dim HeadCf As CReportField
Dim j As Long
Dim TempStr As String
Dim TempRs As ADODB.Recordset
Dim ImportItems As Collection
Dim Amt As Double
Dim iCount As Long

On Error Resume Next

Call m_Headers1_4.ClearField
Call m_Headers1_5.ClearField
   
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = 12
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

Set Cf = New CReportField
                                                                                                                                                                    'Sum Detail
Call Cf.SetFieldValue(12, ">", "����ʹ�Թ��������շ����� (�������Ѻ�Ṻ �.�.�2� �����(�����)", "^")
         Call m_Headers1_4.AddField(Cf)
 Call Cf.SetFieldValue(2, ">", FormatNumber(Str(Total1(2) - CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", ""))), "-"), "^")
         Call m_Headers1_4.AddField(Cf)
Call Cf.SetFieldValue(2, ">", FormatNumber(Str(Total1(3) - CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", ""))), "-"), "^")
         Call m_Headers1_4.AddField(Cf)
Call Cf.SetFieldValue(1, "<", "", "^")
         Call m_Headers1_4.AddField(Cf)
                                                                                                                                                                        'Detail  Signature
Call Cf.SetFieldValue(9, "<", "(����͡�ӴѺ��������ͧ�ѹ价ء�蹵���Թ�����л�����)" & vbCrLf & _
"�����˵�    1. ����͡੾�Сóը��´͡�����Թ�ҡ" & vbCrLf & _
"                  2. ���͹䢡���ѡ����  ����͡�ѧ���     - �ѡ � ������  ��͡ 1   - �͡����ʹ�  ��͡ 2    - �͡����������  ��͡ 3" & vbCrLf & _
"                   �Ţ��Шӵ�Ǽ�����������ҡ�(13��ѡ)* ���¶֧" & vbCrLf & _
"                            1. �óպؤ�Ÿ����� ������Ţ��Шӵ�ǻ�ЪҪ��������û���ͧ�͡���" & vbCrLf & _
"                            2. �óչԵԺؤ�� ������Ţ�Ţ����¹�ԵԺؤ�ŷ�����Ѳ�Ҹ�áԨ��ä���͡���" & vbCrLf & _
"                            3. �ó�����͡�˹�ͨҡ 1. ��� 2. ������Ţ��Шӵ�Ǽ�����������ҡ�(13��ѡ) �������þҡ��͡���", "^")
  Call m_Headers1_5.AddField(Cf)


 Call Cf.SetFieldValue(6, "^", "(ŧ��һ�зѺ)                                                                                                                     " & vbCrLf & vbCrLf & _
                                                "                             ŧ����......................................................................�������Թ" & vbCrLf & _
                                                "                            (...........................................................................)" & vbCrLf & _
                                                "                            ���˹�...................................................................." & vbCrLf & _
                                                "                             ����ѹ���.........��͹................................�.�. ............." & vbCrLf, "^")
        Call m_Headers1_5.AddField(Cf)

Set Cf = Nothing


Call m_Headers1_4.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
Call m_Headers1_5.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
End Sub

Private Sub VSP_Error()
'Error in runtime occur here press F1 in VSP.ErrorDescription to see more information
   mstrErrMsg = Vsp.ErrorDescription
End Sub

Private Sub VSP_NewPage()
Dim talnT As TextAlignSettings
Dim tbdT As TableBorderSettings
Dim blnBold As Boolean
Dim blnUnder As Boolean
Dim blnItalic As Boolean
Dim iSize As Integer
Dim sName As String
Dim strFormat As String
Dim dY(0 To 1) As Double
Dim alngX() As Long

                                                                                                                                 '_________End Page Count____________
   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
  
  Vsp.FontSize = 30
     Vsp.FontBold = True
   Call Vsp.AddTable(strFormat, "", "�Ṻ  �.�.�.2�  ")
      Vsp.FontBold = False
   Vsp.FontSize = 12

 Call Vsp.AddTable(strFormat, "", "|�蹷��|" & Vsp.CurrentPage & "        㹨ӹǹ   " & PageAll & "     ��")
   

' �ѹ��ҧ��� ��������
'   Vsp.PenColor = Blue
'  Call Vsp.DrawRectangle(1500, 1500, 1200, 1200)
'   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))

   Vsp.CurrentY = dY(0)
   
   Vsp.FontSize = TITLE_SIZE
   Vsp.FontBold = True
   Vsp.TextAlign = taCenterBottom
   Vsp.Paragraph = MODULE_DESC
   Vsp.Paragraph = ""
   If mblnHeader Then
      Call printHeader
   End If
   mdY = Vsp.CurrentY
   
   Vsp.TextAlign = talnT
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontUnderline = blnUnder
   Vsp.FontItalic = blnItalic
   Vsp.FontSize = iSize
   Vsp.FontName = sName
End Sub

Private Sub VSP_StartDoc()
'This event occur when VSPrinter.StartDoc is used and used to initialize some information before generating printed document
End Sub

Private Function toTax_ID(ID As String) As String
Dim A As String
Dim B As String
Dim C As String
Dim D As String
Dim E As String
Dim F As Long

 F = Len(ID)
       If F = 13 Then
         A = Left(ID, 1)
         B = Mid(ID, 2, 4)
         C = Mid(ID, 6, 5)
         D = Mid(ID, 11, 2)
         E = Mid(ID, 13, 1)
         toTax_ID = A & " - " & B & " - " & C & " - " & D & " - " & E
      Else
        toTax_ID = "(�Ţ�ѵû�ЪҪ��Դ��Ҵ)"
      End If
End Function



