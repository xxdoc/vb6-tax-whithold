VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportTaxSending"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportTaxSending"
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

Private Const TITLE_SIZE = 13
Private Const HEADER_SIZE = 11
Private Const DETAIL_SIZE = 13

Private Const TITLE_FORMAT = "<30|<3|<4;"

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
Private strPrint As String



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
   
   Set Cf = New CReportField
                                                                                                                    'Head 1
        Set Cf = New CReportField
            
            Call Cf.SetFieldValue(1, "^", "                                   ����ӹǹ�Թ�ѡ���� � �����»�Ш��ѹ", "^")
                Call m_Headers1_1.AddField(Cf)
   Set Cf = Nothing
  Call m_Headers1_1.GetString(2, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
                                                                                                                            'Head 2
                                                                                                                            
        Set Cf = New CReportField

    Call Cf.SetFieldValue(11, "^", "", "^")
            Call m_Headers1_2.AddField(Cf)
    Call Cf.SetFieldValue(3, "^", "�ѡ�ҡ �ؤ�Ÿ�����", "^")
            Call m_Headers1_2.AddField(Cf)
    Call Cf.SetFieldValue(3, "^", "�ѡ�ҡ �ԵԺؤ��", "^")
            Call m_Headers1_2.AddField(Cf)
    Call Cf.SetFieldValue(7, "^", "", "^")
            Call m_Headers1_2.AddField(Cf)

   Set Cf = Nothing
  Call m_Headers1_2.GetString(2, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
                                                                                                                'Head 3
    Set Cf = New CReportField
    Call Cf.SetFieldValue(2, "^", "�ѹ/��͹/��", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(6, "^", "���ͺؤ������˹��§ҹ���" & vbCrLf & "�ѡ � ������", "<")
        Call m_Headers1.AddField(Cf)
  Call Cf.SetFieldValue(3, "^", "�Ţ����" & vbCrLf & "�Ӥѭ����", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�. 3 ���", ">")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(1, "^", "����", "^")
        Call m_Headers1.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "�. 3 ���", ">")
        Call m_Headers1.AddField(Cf)
Call Cf.SetFieldValue(1, "^", "����", "^")
        Call m_Headers1.AddField(Cf)
Call Cf.SetFieldValue(2, "^", "Ẻ��������" & vbCrLf & "��¡��", "^")
        Call m_Headers1.AddField(Cf)
Call Cf.SetFieldValue(2, "^", "�ӹǹ�Թ", ">")
        Call m_Headers1.AddField(Cf)
Call Cf.SetFieldValue(3, "^", "�����˵�", "<")
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
Dim E_NAME As String
Dim D As Date
Dim M_NAME As String

'M_NAME = IntToThaiMonth(Month(frmAddEditTaxDocument.uctlTravelDate.ShowDate))
'M_NAME = M_NAME & "  " & Year(frmAddEditTaxDocument.uctlTravelDate.ShowDate) + 543
'E_NAME = frmAddEditTaxDocument.uctlGuestNameLookup.MyTextBox.Text         '
   
   mstrErrMsg = ""
   mblnHeader = True
   mblnNewPage = True
   mblnEndPage = True
   Vsp.PaperSize = pprA4 '
   Vsp.ORIENTATION = orPortrait
   Vsp.MarginBottom = 300
   Vsp.MarginFooter = 700
   Vsp.MarginHeader = 1440
   Vsp.MarginLeft = 300
   Vsp.MarginRight = 300
   Vsp.MarginTop = 300
   Vsp.FontName = "AngsanaUPC"
   Vsp.FontSize = DETAIL_SIZE

      E_NAME = mcolParam("E_NAME")
      D = mcolParam("FROM_DATE")

MODULE_DESC = "�ѭ�վ�����ʴ�����ѡ���� � ������ ��С�ù�������  ��Ш�˹��§ҹ   " & E_NAME & vbCrLf & _
"�����С��͸Ժ�ա����þҡ�  ����ǡѺ�����Թ�� ������ա�ä��  ( ��Ѻ��� " & Vsp.CurrentPage & " ) " & vbCrLf & _
"����ͧ  ��˹��������˹�ҷ���ѡ�����Թ�� �������ա�ä�� � �������պѭ�վ����  ��Ш���͹    " & IntToThaiMonth(Month(D)) & "  " & Year(D) + 543
                                          'mcolParam("REPORT_NAME")
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

Dim iCount As Long
Dim TempStr As String
Dim PrevKey1 As String
Dim PrevKey2 As String
Dim ID As Long


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
      i = 0
      Set BodyCf = New CReportField

    Set Ct = New CTaxDocument
   Set Ti = New CTaxDocItem
   
    Dim xxx As String
    xxx = mcolParam("REPORT_KEY")
    
      Ct.TAX_DOCUMENT_ID = mcolParam("TAX_DOCUMENT_ID")
    
      
      Ct.QueryFlag = 1
      Call glbDaily.QueryTaxDocument(Ct, Rs, iCount, IsOK, glbErrorLog)
      i = 0

      'While Not Rs.EOF
      For Each Ti In Ct.DocumentItems
         i = i + 1
         I2 = I2 + 1
         Call m_Details1.ClearField

         Set HeadCf = m_Headers1.Fields(1)
         TempStr = DateToStringExtEx2(Ti.PAY_DATE)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(2)
         TempStr = Ti.SUPPLIER_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

'         Set HeadCf = m_Headers1.Fields(3)
'         TempStr = Ti.ADDRESS_ID
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)

         Set HeadCf = m_Headers1.Fields(3)
         TempStr = Ti.REF_NO
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

         
         Set HeadCf = m_Headers1.Fields(4)                                                         '�ؤ�Ÿ�����
         If Ti.PS_TYPE = "�ؤ�Ÿ�����" Then
                  TempStr = FormatNumber(Ti.PAID_AMOUNT)
                  Total1(4) = Total1(4) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", "")) 'Ti.PAID_AMOUNT
         Else
                  TempStr = "-"
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
        
         
            
         Set HeadCf = m_Headers1.Fields(5)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)

       
                  Set HeadCf = m_Headers1.Fields(6)                                                     '�ԵԺؤ��
         If Ti.PS_TYPE = "�ԵԺؤ��" Then
                  TempStr = FormatNumber(Ti.PAID_AMOUNT)
                  Total1(6) = Total1(6) + CDbl(Replace(FormatNumber(Ti.PAID_AMOUNT, "-"), ",", ""))  'Ti.PAID_AMOUNT
         Else
                  TempStr = "-"
         End If
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
      

         Set HeadCf = m_Headers1.Fields(7)
         TempStr = ""
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         Set HeadCf = m_Headers1.Fields(8)
         TempStr = "�.�.� " & TAX_TYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         
         Set HeadCf = m_Headers1.Fields(9)
         TempStr = FormatNumber(Ti.WH_AMOUNT)
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         Total1(9) = Total1(9) + CDbl(Replace(FormatNumber(Ti.WH_AMOUNT, "-"), ",", ""))   'Ti.WH_AMOUNT

         Set HeadCf = m_Headers1.Fields(10)
         TempStr = Ti.REVENUE_NAME & " " & Ti.RATETYPE_NAME
         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
         Call m_Details1.AddField(BodyCf)
         
         Call m_Details1.GetString(1, TempStr1, TempStr2)
         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
         Call Vsp.AddTable(strFormat, "", TempStr2)
         'Rs.MoveNext
         
         Call AdjustPage(Vsp, strFormat, strPrint, 1500)
         
      Next Ti
   
                                                                                                                        'Footer Report
 'RCount = I2
  
   'Call PatchTable(Vsp, strFormat, "|||||||||", 500)
   '===

'If I2 <= 18 Then
'    RCount = I2
'
'
'    For I2 = 0 To 18 - RCount
'            i = i + 1
'
'         Call m_Details1.ClearField
'
'         Set HeadCf = m_Headers1.Fields(1)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(2)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(3)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(4)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(5)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(6)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'
'         Set HeadCf = m_Headers1.Fields(7)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'
'         Set HeadCf = m_Headers1.Fields(8)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(9)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'         Set HeadCf = m_Headers1.Fields(10)
'         TempStr = ""
'         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr)
'         Call m_Details1.AddField(BodyCf)
'
'
'
'         Call m_Details1.GetString(1, TempStr1, TempStr2)
'         strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
'         Call Vsp.AddTable(strFormat, "", TempStr2)
'
'
'        '       If I >= 3 Then
'        '         Call AdjustPage(Vsp, strFormat, strPrint, 5000)
'        '      End If
'        '      Call Vsp.AddTable(strFormat, "", strPrint)
'
'            Next I2
'
'      End If
         

      
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
      ElseIf (j = 4) Then
         TempStr = FormatNumber(Tot(j))
    ElseIf (j = 5) Then
         TempStr = ""
      ElseIf j = 6 Then
         'TempStr = Tot(j)
         TempStr = FormatNumber(Tot(j))
      ElseIf j = 7 Then
         TempStr = ""
      ElseIf j = 8 Then
         TempStr = ""
    ElseIf j = 9 Then
         TempStr = FormatNumber(Tot(j))
    ElseIf j = 10 Then
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
   
   Vsp.FontSize = 13
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll

Set Cf = New CReportField
                                                                                                                                                                    'Sum Detail
Call Cf.SetFieldValue(11, "^", "������Թ���������͹���", "^")
         Call m_Headers1_3.AddField(Cf)
 Call Cf.SetFieldValue(2, ">", FormatNumber(Total1(4), "-"), "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(1, "^", "", "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(2, ">", FormatNumber(Total1(6), "-"), "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(1, "^", "", "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(2, "^", "", "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(2, ">", FormatNumber(Str(Total1(9)), "-"), "^")
         Call m_Headers1_3.AddField(Cf)
Call Cf.SetFieldValue(3, "<", "", "^")
         Call m_Headers1_3.AddField(Cf)

Set Cf = Nothing
Call m_Headers1_3.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
   Vsp.TableBorder = tbdT
   Vsp.FontBold = blnBold
   Vsp.FontSize = iSize
   
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

   If Not mblnNewPage Then Exit Sub
   talnT = Vsp.TextAlign
   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   blnUnder = Vsp.FontUnderline
   blnItalic = Vsp.FontItalic
   iSize = Vsp.FontSize
   sName = Vsp.FontName
   
   Vsp.FontSize = 11
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
'   Call Vsp.AddTable(strFormat, "", "|�蹷��|" & Vsp.CurrentPage & "�ҡ������ " & Vsp.PageCount & " ��" & vbCrLf)
  
'   Call Vsp.AddTable(strFormat, "", "㺵�� �.�.� ")
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










