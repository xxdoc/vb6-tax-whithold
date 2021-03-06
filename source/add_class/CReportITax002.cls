VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportTax0053"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportTax0053"
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

Private Const TITLE_SIZE = 12
Private Const HEADER_SIZE = 10
Private Const DETAIL_SIZE = 11

Private Const TITLE_FORMAT = "<29|<3|<5;"                                       'title format

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
Private RCount As Long
Private I2 As Long
Private iCount As Long
Private IsOK As Boolean

Private Sub Class_Initialize()
   Set mcolParam = New Collection
   Set m_TempCol = New Collection
   
   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers1_1 = New CFieldList
   Set m_Headers1_2 = New CFieldList
   Set m_Headers1_3 = New CFieldList
   Set m_Headers1_4 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
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
   Set m_Details2 = Nothing
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
   Dim I As Long
   
   For I = 1 To mcolParam.Count
      mcolParam.Remove 1
   Next I

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
   

   Set Cf = New CReportField
                                                                                                                    'Head 1
        Set Cf = New CReportField
            Call Cf.SetFieldValue(6, "^", "���ͼ���Ѻ�Թ��֧�����Թ(�����" & vbCrLf & "���������ҧ�����ǹ�ӡѴ����" & vbCrLf & "��ҧ�����ǹ���ѭ�ԵԺؤ��", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(6, "^", "�Ţ��Шӵ�Ǽ����������(������Թ��) ", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(7, "^", "��������´����ǡѺ��è����Թ��֧�����Թ", "^")
                Call m_Headers1_1.AddField(Cf)
            Call Cf.SetFieldValue(3, "^", "�Թ���շ���ѡ��й��� " & vbCrLf & " 㹤��駹��", "^")
                Call m_Headers1_1.AddField(Cf)
   Set Cf = Nothing
  Call m_Headers1_1.GetString(2, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
     
                                                                                                       
    Set Cf = New CReportField
    Call Cf.SetFieldValue(1, "^", "�ӴѺ���", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(11, "^", "�������ͧ������Թ��(����к��Ţ��� ��͡/��� ��� �Ӻ�/�ǧ  �����/ࢵ �ѧ��Ѵ)", "<")
        Call m_Headers1.AddField(Cf)
  Call Cf.SetFieldValue(2, "^", "�ѹ��͹��" & vbCrLf & "������", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�������Թ��" & vbCrLf & "�֧�����Թ������", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "�ѵ�� " & vbCrLf & "����" & vbCrLf & "������", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ" & vbCrLf & "�����¤��駹��", ">")
        Call m_Headers1.AddField(Cf)
Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ", ">")
        Call m_Headers1.AddField(Cf)
Call Cf.SetFieldValue(1, "^", "���͹" & vbCrLf & "�", "^")
        Call m_Headers1.AddField(Cf)

    Set Cf = Nothing
   
Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
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
Dim E_NAME As String

IDENT_ID = mcolParam("IDENT_ID")
TAX_ID = mcolParam("TAX_ID")
E_NAME = mcolParam("E_NAME")
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 200
   Vsp.MarginFooter = 300
   Vsp.MarginHeader = 300
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 200
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE
      
MODULE_DESC = "�Ţ��Шӵ�Ǽ�����������ҡ�    " & TAX_ID & vbCrLf & " "    '& "  �Ң�   " & glbEnterPrise.BRANCH_NAME
'mcolParam("REPORT_NAME") & vbCrLf &
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
Dim I As Long
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

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim ID As Long
Dim strPrint As String
Dim RCount As String
Dim E_NAME As String

   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   
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
   
   Vsp.StartDoc
     If I = Null Then I = 0
      Set BodyCf = New CReportField

    Set Ct = New CTaxDocument
    Dim xxx As String
    xxx = mcolParam("REPORT_KEY")
      Ct.TAX_DOCUMENT_ID = mcolParam("TAX_DOCUMENT_ID")
      Ct.QueryFlag = 1
      Call glbDaily.QueryTaxDocument(Ct, Rs, iCount, IsOK, glbErrorLog)
      I = 0

      'While Not Rs.EOF
      RCount = I2
      For Each Ti In Ct.DocumentItems
         I = I + 1
         I2 = I2 + 1
         
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, vbCrLf & TempStr & vbCrLf)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Ti.SUPPLIER_NAME & "         " & Ti.TAX_ID & _
         vbCrLf & Ti.PackAddress
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

'         Set HeadCf = m_Headers1.Fields(3)
'         TempStr = Ti.ADDRESS_ID
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = DateToStringExtEx2(Ti.PAY_DATE) & vbCrLf & Ti.BRANCH_CODE
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = Ti.REVENUE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = Ti.RATETYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = FormatNumber(Ti.PAID_AMOUNT, "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(6) = Total1(6) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "0"), ",", ""))  'Ti.PAID_AMOUNT

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = FormatNumber(Ti.WH_AMOUNT, "-")
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(7) = Total1(7) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "0"), ",", "")) 'Ti.WH_AMOUNT

         Set HeadCf = m_Headers1.Fields(8)
         TempStr = Ti.CONDITION_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)



         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         'Rs.MoveNext
         
     If (I Mod 11) = 0 Then
         If Ct.DocumentItems.Count - I > 0 Then  'Still left
            Vsp.NewPage
         End If
          I2 = 0
    End If
            
            
      Next Ti
                                                                               
                                                                                                 'Footer Report

   RCount = I2
  
   'Call PatchTable(Vsp, strFormat, "|||||||||", 500)
   '===

If (I2 > 0) And (I2 <= 11) Then '�ͧ����� 11
    RCount = I2
     
    For I2 = 0 To 11 - RCount '�ͧ����� 11
            I = I + 1
       
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = I
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, vbCrLf & TempStr & vbCrLf)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(4)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(5)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(6)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         

         Set HeadCf = m_Headers1.Fields(8)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)



         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
     
     
'       If I >= 3 Then
'         Call AdjustPage(Vsp, strFormat, strPrint, 5000)
'      End If
'      Call Vsp.AddTable(strFormat, "", strPrint)
         
      Next I2
    
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

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, DateCount As Long)
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim j As Long
Dim TempStr As String

   Call Details.ClearField
   Set BodyCf = New CReportField
   
   j = 0
   For Each HeadCf In m_Headers1.Fields
      j = j + 1
      If j = 1 Then
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      ElseIf (j = 4) Or (j = 5) Then
         TempStr = ""
      ElseIf j = 6 Then
         'TempStr = Tot(j)
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 7 Then
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 8 Then
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


  Call m_Headers1_3.ClearField
  Call m_Headers1_4.ClearField
  
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize
   
   Vsp.FontSize = 10
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

Set Cf = New CReportField
                                                                                                                                                                    'Sum Detail
Call Cf.SetFieldValue(17, ">", "����ʹ�Թ��������շ����� (�������Ѻ㺵�� �.�.� " & TAX_TYPE_NAME & " ��Ѻ��� (�����))", "^")
         Call m_Headers1_3.AddField(Cf)
 Call Cf.SetFieldValue(2, ">", FormatNumber(Str(Total1(6)), "-"), "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(2, ">", FormatNumber(Str(Total1(7)), "-"), "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(1, "<", "", "^")
         Call m_Headers1_3.AddField(Cf)
                                                                                                                                                                        'Detail  Signature
Call Cf.SetFieldValue(14, "<", "(����͡�ӴѺ��������ͧ�ѹ价ء�蹵���Թ�����л�����)" & vbCrLf & _
"�����˵�    1. ����к���Ҩ����繤������ �� �������Ҥ�� ����ͺ�ѭ��" & vbCrLf & _
"                     ����͡Ẻ  ��ҡ�����ҧ�ç���¹ ��ҫ�������ͧ�����մ" & vbCrLf & _
"                     ��ҫ��;ת�ŷҧ����ɵ� (�ҧ���� �ѹ�ӻ���ѧ �� ���� ���) ��Ҩ�ҧ�Ӣͧ" & vbCrLf & _
"                     ��Ҩ�ҧ�ɳ� �͡����Թ�ҡ �͡���µ����Թ �Թ�ѹ�� �Թ��ǹ�觢ͧ����" & vbCrLf & _
"                     �ҧ���㹡�û�СǴ  ����觢ѹ  ��êԧ⪤ ���" & vbCrLf & _
"                     2 . ���͹�㹡���ѡ���� � ������ ����͡�ѧ���" & vbCrLf & _
"                            -  �ѡ � ������ ��͡ 1     - �͡��������͡ 2 ", "^")
    Call m_Headers1_4.AddField(Cf)
 Call Cf.SetFieldValue(8, "^", "(ŧ��һ�зѺ)                                                                  " & vbCrLf & vbCrLf & vbCrLf & _
                                                "                   ŧ����.............................................................�������Թ" & vbCrLf & _
                                                "                   (...........................................................................)" & vbCrLf & _
                                                "                   ���˹�...................................................................." & vbCrLf & _
                                                "                   ����ѹ���.........��͹................................�.�. ............." & vbCrLf, "^")
        Call m_Headers1_4.AddField(Cf)

Set Cf = Nothing


Call m_Headers1_3.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
Call m_Headers1_4.GetString(1, TempStr1, TempStr1_1)
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
 
                                                                                                                                 ' ________Set Page Count Page All_______
  If PageAll = 0 Then
        
        Set Rs = New ADODB.Recordset
        Set Ct = New CTaxDocument
                                                                                                                                    
        Ct.TAX_DOCUMENT_ID = mcolParam("TAX_DOCUMENT_ID")
              Ct.QueryFlag = 1
        Call glbDaily.QueryTaxDocument(Ct, Rs, iCount, IsOK, glbErrorLog)
                                                                                                                                   
        RCount = Ct.DocumentItems.Count
        PageAll = (RCount \ 11)
        If (RCount Mod 11) > 0 Then PageAll = PageAll + 1
                                                                                                                               
           If Rs.State = adStateOpen Then
              Rs.Close
           End If
     
           
End If
                                                                                                                                 '_________End Page Count____________
    If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
  Vsp.FontSize = 12
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
 ' Call Vsp.AddTable(strFormat, "", "||")
   Vsp.FontSize = 25
   Call Vsp.AddTable(strFormat, "", "㺵��  �.�.�  " & TAX_TYPE_NAME)
   Vsp.FontSize = 12
  Call Vsp.AddTable(strFormat, "", "|�蹷��|" & Vsp.CurrentPage & " 㹨ӹǹ " & PageAll & "  ��|")
  
'   Call Vsp.AddTable(strFormat, "", "|�ѹ|: " & DateToStringExt(mdteDate))
'   Call Vsp.AddTable(strFormat, "", "|����|: " & Right$(DateToStringExtEx(mdteDate), 8))
'   Call Vsp.AddTable(strFormat, "", "|Tax ID|: " & glbEnterPrise.TAX_ID)

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






