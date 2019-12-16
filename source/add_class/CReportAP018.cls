VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CReportSummaryLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'User Group Report
Implements CReportInterface

Private Const MODULE_NAME = "CReportSummaryLoan"
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
Private Rs2 As ADODB.Recordset

Private Const TITLE_SIZE = 14
Private Const HEADER_SIZE = 12
Private Const DETAIL_SIZE = 12
Private Const ISO_SIZE = 18

Private Const TITLE_FORMAT = "<32|<3|<5;"

Private m_HeaderFormat1 As String
Private m_HeaderFormat2 As String
Private m_HeaderText1 As String
Private m_HeaderText2 As String

Private m_Headers0 As CFieldList
Private m_Headers1 As CFieldList
Private m_Headers2 As CFieldList
Private m_Details1 As CFieldList
Private m_Details2 As CFieldList
Private m_Left As Double
Private m_YearSeqs As Collection
Private Tii As CTaxDocItem
Private Ti As CTaxDocItem
Private Tiii As CTaxDocItem

Private m_TempCol As Collection
Private m_TempCols As Collection
Private m_Temp As CTaxDocItem
Private m_Temp2 As CTaxDocItem
Private TempEnterprise As CEnterprise

Private m_Data1 As Collection
Private m_PartGroup As Collection
Private Sub Class_Initialize()
   Set mcolParam = New Collection
  ' Set m_TempCol = New Collection
   Set m_TempCols = New Collection

   Set m_Data1 = New Collection
   Set m_PartGroup = New Collection
   


   Set m_Headers0 = New CFieldList
   Set m_Headers1 = New CFieldList
   Set m_Headers2 = New CFieldList
   Set m_Details1 = New CFieldList
   Set m_Details2 = New CFieldList
   Set m_YearSeqs = New Collection
End Sub
Private Sub Class_Terminate()
   Call ClearParam
   Set mcolParam = Nothing
   Set Vsp = Nothing
   'Set m_TempCol = Nothing
   Set m_TempCols = Nothing
   
   Set m_Data1 = Nothing
   Set m_PartGroup = Nothing

   Set m_Headers0 = Nothing
   Set m_Headers1 = Nothing
   Set m_Headers2 = Nothing
   Set m_Details1 = Nothing
   Set m_YearSeqs = Nothing
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
Dim FromDate   As Date
Dim ToDate  As Date
Dim TempFromDate As Date

   tbdT = Vsp.TableBorder
   blnBold = Vsp.FontBold
   iSize = Vsp.FontSize

   Vsp.FontSize = HEADER_SIZE
   Vsp.FontBold = True
   Vsp.TableBorder = tbAll
   

   Call m_Headers0.ClearField
   Call m_Headers1.ClearField
   Call m_Headers2.ClearField
   
   FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   ToDate = DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1)
   Set Cf = New CReportField
    Call Cf.SetFieldValue(1, "^", "", "^")
      Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(4, "<", "", "<")
      Call m_Headers1.AddField(Cf)
    Call Cf.SetFieldValue(2, "<", "", "<")
      Call m_Headers1.AddField(Cf)
   TempFromDate = FromDate
   While TempFromDate <= ToDate
            Call Cf.SetFieldValue(8, "^", Month(TempFromDate) & "/" & Year(TempFromDate) + 543, ">")
                Call m_Headers1.AddField(Cf)
      TempFromDate = DateAdd("M", 1, TempFromDate)
   Wend
   '***********************************************
   FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
   ToDate = DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1)
    Call Cf.SetFieldValue(1, "^", "No.", "^")
       Call m_Headers2.AddField(Cf)
    Call Cf.SetFieldValue(4, "^", "รายชื่อ", "<")
      Call m_Headers2.AddField(Cf)
   Call Cf.SetFieldValue(2, "^", "เงินกู้", "<")
      Call m_Headers2.AddField(Cf)
   While FromDate <= ToDate
           Call Cf.SetFieldValue(2, "^", "เอกสาร", "^")
                Call m_Headers2.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "ดอกเบี้ย", "^")
                Call m_Headers2.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "ภาษี15%", ">")
                Call m_Headers2.AddField(Cf)
            Call Cf.SetFieldValue(2, "^", "จ่ายสุทธิ", "^")
                Call m_Headers2.AddField(Cf)
      FromDate = DateAdd("M", 1, FromDate)
   Wend
   Set Cf = Nothing
   
   Call m_Headers1.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)
   
  Call m_Headers2.GetString(1, TempStr1, TempStr1_1)
   strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
   strPrint = TempStr1_1
   Call Vsp.AddTable(strFormat, "", strPrint)

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
Dim TempData As CEnterprise
Set TempData = New CEnterprise


   
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
   Call getTaxIDfromBranch(, , mcolParam("BRANCH"), TAX_ID)
                                                                                                                                                                                                                                                            
'    Set TempEnterprise = GetObject("CEnterprise", m_Enterprise, Trim(mcolParam("BRANCH")))
   MODULE_DESC = "ข้อมูลการจ่ายดอกเบี้ยของผู้ให้กู้" & vbCrLf & "หน่วยงาน " & mcolParam("BRANCH") & vbCrLf & _
                                       mcolParam("REPORT_NAME")
    Call SetReportConfig(Vsp, mcolParam("REPORT_KEY"))
    
   mdblWid = Vsp.PageWidth - Vsp.MarginLeft - Vsp.MarginRight
   If Not glbDatabaseMngr.GetServerDateTime(strDate, glbErrorLog) Then
      mstrErrMsg = "Error GetDateTime Error."
      Exit Function
   End If
   Set TempData = Nothing
   mdteDate = InternalDateToDate(strDate)
   initDoc = True
End Function
Private Function genDoc(isPreview As Boolean) As Boolean
 On Error GoTo ErrHandler
Dim O As Object
Dim RName As String
Dim i As Long
Dim j As Long
Dim K As Long
Dim strFormat As String
Dim alngX() As Long
Dim IsOK As Boolean
Dim Amt As Double
Dim HeadCf As CReportField
Dim BodyCf As CReportField
Dim TempStr1 As String
Dim TempStr2 As String
Dim Total1(100) As Double
Dim Total2(100) As Double
Dim iCount As Long
Dim TempStr As String
Dim TempID As Long
Dim TempDiff As Double

Dim PrevKey As Long
Dim PrevKey1 As Long
Dim PrevKey2 As String
Dim PrevKey3 As Long
Dim PrevKey4 As String
Dim PrevKey5 As String
Dim PrevKey6 As String
Dim PrevKey7 As String
Dim PrevKeyi As Long
Dim FromDate   As Date
Dim ToDate  As Date

Dim RoColl As Collection
Dim CarryForward As Double

Dim Sum1 As Double
Dim Sum2 As Double
Dim PoApproved As String
Dim TempFromDate  As Date
Dim TempToDate As Date
Dim Temp1 As String
Dim Temp2 As String
Dim Key As String
Dim YYYYMM As String
Dim tempList As Collection
Set tempList = New Collection

Dim REF_NO As String
Dim TotalInterest As String
Dim TotalVat As String
Dim Diff As String


   RName = "genDoc"
'-----------------------------------------------------------------------------------------------------
'                                             Query Here
'-----------------------------------------------------------------------------------------------------
   Set Rs = New ADODB.Recordset
   Set Rs2 = New ADODB.Recordset
   
   If Not initDoc Then Exit Function

   Vsp.Preview = isPreview
'-----------------------------------------------------------------------------------------------------
'                                         Main Operation Here
'-----------------------------------------------------------------------------------------------------
      For j = 1 To UBound(Total1)
         Total1(j) = 0
         Total2(j) = 0
      Next j
     
         Call GetFirstLastDate(DateSerial(mcolParam("TO_YEAR_NO") - 543, mcolParam("TO_MONTH_ID"), 1), FromDate, ToDate)
         FromDate = DateSerial(mcolParam("FROM_YEAR_NO") - 543, mcolParam("FROM_MONTH_ID"), 1)
         Set Ti = New CTaxDocItem
         Ti.BRANCH_CODE = mcolParam("BRANCH")
'         Ti.MASTER_AREA = 13
         Ti.FROM_PAY_DATE = FromDate
         Ti.TO_PAY_DATE = ToDate
         
         Ti.FROM_SUPPLIER_CODE = mcolParam("FROM_SUPPLIER_CODE")
         Ti.TO_SUPPLIER_CODE = mcolParam("TO_SUPPLIER_CODE")
         
         Ti.OrderType = 1
         Ti.TAX_TYPE = mcolParam("TAX_TYPE")
         
         i = 0
         Call Ti.QueryData(2, Rs, iCount)
         Set Rs2 = Rs.Clone(adLockReadOnly)
         Call StartExportFile(Vsp)
         Vsp.StartDoc
         Set BodyCf = New CReportField
         If Rs.RecordCount > 0 Then
         
            'รอบแรก ดึงค่า ของ Supplier แต่ล่ะคนก่อนแต่ล่ะ PV
             Set m_TempCol = New Collection
             While Not Rs.EOF
               Set m_Temp = New CTaxDocItem
                i = i + 1
                
                Call Ti.PopulateFromRS(2, Rs)
                           If (PrevKey1 = Ti.SUPPLIER_ID And PrevKey2 = Ti.YYYYMM) Or (i = 1) Then  'หาชื่อเดียวกัน และเดือนเดียวกันก่อน ซึ่งอาจจะมีหลาย PV
                               PrevKey1 = Ti.SUPPLIER_ID
                               PrevKey2 = Ti.YYYYMM
                               
                               TotalInterest = IIf(TotalInterest = "", FormatNumber(Ti.PAID_AMOUNT), TotalInterest & vbNewLine & FormatNumber(Ti.PAID_AMOUNT))
                               TotalVat = IIf(TotalVat = "", FormatNumber(Ti.WH_AMOUNT), TotalVat & vbNewLine & FormatNumber(Ti.WH_AMOUNT)) '
                               TempDiff = Ti.PAID_AMOUNT - Ti.WH_AMOUNT
                               Diff = IIf(Diff = "", CStr(FormatNumber(TempDiff, 2)), Diff & vbNewLine & CStr(FormatNumber(TempDiff, 2)))
                               REF_NO = IIf(REF_NO = "", Ti.REF_NO, REF_NO & vbNewLine & Ti.REF_NO)
                               If iCount = 1 Then 'ถ้ามีแค่ Record เดียวให้เก็บเลย
'                                 m_Temp.TotalInterest = TotalInterest
'                                 m_Temp.TotalVat = TotalVat
'                                 m_Temp.REF_NO = REF_NO
'                                 m_Temp.TempDiff = Diff
'                                 Key = PrevKey1 & "-" & PrevKey2
'                                 Call m_TempCol.Add(m_Temp, Key)
                               End If
                           Else
                               m_Temp.TotalInterest = TotalInterest
                               m_Temp.TotalVat = TotalVat
                               m_Temp.REF_NO = REF_NO
                               m_Temp.TempDiff = Diff
                              Key = PrevKey1 & "-" & PrevKey2
                              Call m_TempCol.Add(m_Temp, Key)
                              
                              TotalInterest = ""
                              TotalVat = ""
                              REF_NO = ""
                              Diff = ""
                              
                               PrevKey1 = Ti.SUPPLIER_ID
                               PrevKey2 = Ti.YYYYMM
                               
                              TotalInterest = IIf(TotalInterest = "", FormatNumber(Ti.PAID_AMOUNT), TotalInterest & vbNewLine & FormatNumber(Ti.PAID_AMOUNT))
                               TotalVat = IIf(TotalVat = "", FormatNumber(Ti.WH_AMOUNT), TotalVat & vbNewLine & FormatNumber(Ti.WH_AMOUNT)) '
                               TempDiff = Ti.PAID_AMOUNT - Ti.WH_AMOUNT
                               Diff = IIf(Diff = "", CStr(FormatNumber(TempDiff, 2)), Diff & vbNewLine & CStr(FormatNumber(TempDiff, 2)))
                               REF_NO = IIf(REF_NO = "", Ti.REF_NO, REF_NO & vbNewLine & Ti.REF_NO)
                           End If
                Rs.MoveNext
             Wend
                                Set m_Temp = New CTaxDocItem
                                m_Temp.TotalInterest = TotalInterest
                                 m_Temp.TotalVat = TotalVat
                                 m_Temp.REF_NO = REF_NO
                                 m_Temp.TempDiff = Diff
                                 Key = PrevKey1 & "-" & PrevKey2
                                 Call m_TempCol.Add(m_Temp, Key)
            
          i = 0
          Call m_Details1.ClearField
         TempFromDate = FromDate
          Rs.MoveFirst
          While Not Rs.EOF
            Call m_Details1.ClearField
            
                Set Tii = New CTaxDocItem
                Call Tii.PopulateFromRS(2, Rs)
                If Tii.SUPPLIER_ID <> PrevKey Or i = 0 Then
                         i = i + 1
                         PrevKey = Tii.SUPPLIER_ID
                         
                         Set HeadCf = m_Headers2.Fields(1)
                         TempStr = i
                         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                         Call m_Details1.AddField(BodyCf)
          
                         Set HeadCf = m_Headers2.Fields(2)
                         TempStr = Tii.SUPPLIER_NAME & vbNewLine & " (" & Tii.SUPPLIER_CODE & ")"
                         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                         Call m_Details1.AddField(BodyCf)
          
                         Set HeadCf = m_Headers2.Fields(3)
                         TempStr = ""
                         Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                         Call m_Details1.AddField(BodyCf)
                        
                         TempFromDate = FromDate
                         YYYYMM = DateToStringIntEx4(TempFromDate)
                         j = 4
                        While TempFromDate <= ToDate
                             Key = Tii.SUPPLIER_ID & "-" & YYYYMM
                             Set m_Temp2 = GetObject("CTaxDocItem", m_TempCol, Key, True)
                             
                                Set HeadCf = m_Headers2.Fields(j)
                               TempStr = m_Temp2.REF_NO
                               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                               Call m_Details1.AddField(BodyCf)
                               j = j + 1
                               
                               Set HeadCf = m_Headers2.Fields(j)
                               TempStr = m_Temp2.TotalInterest
                               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                               Call m_Details1.AddField(BodyCf)
                               j = j + 1
                               
                               Set HeadCf = m_Headers2.Fields(j)
                               TempStr = m_Temp2.TotalVat
                               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                               Call m_Details1.AddField(BodyCf)
                               j = j + 1
                               
                               Set HeadCf = m_Headers2.Fields(j)
                               TempStr = m_Temp2.TempDiff
                               Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.Align, TempStr, HeadCf.BodyAlign)
                               Call m_Details1.AddField(BodyCf)
                               j = j + 1
                            
                            TempFromDate = DateAdd("M", 1, TempFromDate)
                            YYYYMM = DateToStringIntEx4(TempFromDate)
                            Set m_Temp2 = Nothing
                         Wend
                        End If
                          Call m_Details1.GetString(2, TempStr1, TempStr2)
                           strFormat = VSP_CalTable(TempStr1, mdblWid, alngX)
                           Call Vsp.AddTable(strFormat, "", TempStr2)
                      Rs.MoveNext
          Wend
          Set Tii = Nothing
      End If
   Set BodyCf = Nothing
   Vsp.EndDoc
   Call CloseExportFile(Vsp)
   
   If Rs.State = adStateOpen Then
      Rs.Close
   End If
   Set Rs = Nothing
   genDoc = True
   Exit Function
   
ErrHandler:
   mstrErrMsg = "Error(" & RName & ")" & Err.Number & " : " & Err.DESCRIPTION
   Set Rs = Nothing
End Function

Private Sub GenerateFooter(Txt1 As String, Txt2 As String, Txt3 As String, Tot() As Double, Details As CFieldList, BalAmount As Double)
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
      ElseIf j = 6 Then
         TempStr = FormatNumber(Tot(j))
      Else
          TempStr = ""
'         TempStr = FormatNumber(Tot(j))
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
   Vsp.TextAlign = taRightBottom
   Vsp = "FM-PC-04 Rev.01(07/04/59)" 'การเปลี่ยนแปลงรูปแบบเอกสารแต่ล่ะครั้งให้เปลี่ยนเวอร์ชั่น ISO ของเอกสารด้วยทุกครั้ง
   If Not mblnEndPage Then Exit Sub
'   Vsp.DrawLine Vsp.MarginLeft, mdY, Vsp.MarginLeft, Vsp.CurrentY
'   Vsp.DrawLine Vsp.PageWidth - Vsp.MarginRight, mdY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
'   Vsp.DrawLine Vsp.MarginLeft, Vsp.CurrentY, Vsp.PageWidth - Vsp.MarginRight, Vsp.CurrentY
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

   Vsp.FontSize = 8
   dY(0) = Vsp.CurrentY
   Vsp.TextAlign = taLeftBottom
   strFormat = VSP_CalTable(TITLE_FORMAT, mdblWid, alngX)
   Vsp.TableBorder = tbNone
   Call Vsp.AddTable(strFormat, "", "|หน้าที่|: " & Vsp.CurrentPage)
   Call Vsp.AddTable(strFormat, "", "|วัน|: " & DateToStringExt(mdteDate))
   Call Vsp.AddTable(strFormat, "", "|เวลา|: " & Right$(DateToStringExtEx(mdteDate), 8))
   Call Vsp.AddTable(strFormat, "", "|ผู้ใช้|: " & glbUser.USER_NAME)
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
Private Sub GenerateHeader(Txt1 As String, Txt2 As String, Txt3 As String, Details As CFieldList)
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
         TempStr = Txt1
      ElseIf j = 2 Then
         TempStr = Txt2
      ElseIf j = 3 Then
         TempStr = Txt3
      Else
         TempStr = ""
      End If
      
      Call BodyCf.SetFieldValue(HeadCf.Width, HeadCf.BodyAlign, TempStr, HeadCf.Align)
      Call Details.AddField(BodyCf)
   Next HeadCf
      
   Set BodyCf = Nothing
End Sub



