VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGuiConfigs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_GuiConfigs As Collection
Public Tax2Items As Collection

Private Sub Class_Initialize()
   Set m_GuiConfigs = New Collection
   Set Tax2Items = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_GuiConfigs = Nothing
   Set Tax2Items = Nothing
End Sub

Private Sub CreateGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = New CGuiConfig
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
   
   Call m_GuiConfigs.Add(TempC, Key)
   Set TempC = Nothing
End Sub

Private Sub SetGuiConfigObj(Key As String, Flag As Boolean)
Dim TempC As CGuiConfig

   Set TempC = m_GuiConfigs(Key)
   TempC.ACCESS_RIGHT_KEY = Key
   TempC.ENABLE_FLAG = Flag
End Sub
Private Sub GenerateTax2Items(MenuMask As String, Col As Collection, Programowner As String)
Dim Mask As String
Dim I As Long
Dim TempKey(32) As String
Dim TempID(32) As Long
Dim D As CMenuItem
   
   TempKey(1) = MapText("ข้อมูล ภ.ง.ด. 2")
   TempID(1) = 1
   TempKey(2) = MapText("-")
   TempID(2) = 2
   TempKey(3) = MapText("ข้อมูลผู้กู้ ภ.ง.ด. 2")
   TempID(3) = 3
   
   For I = 1 To Len(MenuMask)
      Mask = Mid(MenuMask, I, 1)
      If Mask = "Y" Then
         Set D = New CMenuItem
         D.KEYWORD = TempKey(I)
         D.KEY_ID = TempID(I)
         Call Col.Add(D, Trim(Str(TempID(I))))
         Set D = Nothing
      End If
   Next I
End Sub
Public Sub CreateGuiConfig(Programowner As String)
'
'   '====
'   Call CreateGuiConfigObj("LOGO_VIEW", False)
'
'   '====
'   Call CreateGuiConfigObj("HR_VIEW", True)
'
'   '====
'   Call CreateGuiConfigObj("PRODUCTION_VIEW", True)
'
'   '=====
'   Call CreateGuiConfigObj("LEDGER_VIEW", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_QUOATATION", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_PO", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_DO", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_INVOICE", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_RECEIPT", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_DBN", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_CDN", True)
'   Call CreateGuiConfigObj("LEDGER_SELL_PKGLIST", True)
'
'   Call CreateGuiConfigObj("LEDGER_BUY_QUOATATION", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_PO", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_DO", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_INVOICE", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_RECEIPT", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_DBN", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_CDN", True)
'   Call CreateGuiConfigObj("LEDGER_BUY_PKGLIST", True)
'   '====
'
'Call CreateGuiConfigObj("Tax2Items", True)
Call GenerateTax2Items("YYY", Tax2Items, Programowner)
'   '====
'   Call CreateGuiConfigObj("SELL_DO_NORMAL_PRINT", True)
'   Call CreateGuiConfigObj("SELL_PO_NORMAL_PRINT", True)
'   Call CreateGuiConfigObj("SELL_QUOATATION_NORMAL_PRINT", True)
'   Call CreateGuiConfigObj("SELL_RECEIPT_NORMAL_PRINT", True)
'   Call CreateGuiConfigObj("SELL_BILLS_NORMAL_PRINT", True)
'   Call CreateGuiConfigObj("SELL_DBN_NORMAL_PRINT", True)
'   Call CreateGuiConfigObj("SELL_CDN_NORMAL_PRINT", True)
'
'   Call CreateGuiConfigObj("SELL_DO_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_INVOICE_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_PO_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_QUOATATION_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_RECEIPT_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_BILLS_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_DBN_PREFORM_PRINT", False)
'   Call CreateGuiConfigObj("SELL_CDN_PREFORM_PRINT", False)
'
'   '=====
'   Call CreateGuiConfigObj("SELL_DO_EXEC", True)
'   Call CreateGuiConfigObj("SELL_INVOICE_EXEC", True)
'   Call CreateGuiConfigObj("SELL_PO_EXEC", True)
'   Call CreateGuiConfigObj("SELL_QUOATATION_EXEC", True)
'   Call CreateGuiConfigObj("SELL_RECEIPT_EXEC", True)
'   Call CreateGuiConfigObj("SELL_BILLS_EXEC", True)
'   Call CreateGuiConfigObj("SELL_DBN_EXEC", True)
'   Call CreateGuiConfigObj("SELL_CDN_EXEC", True)
'   Call CreateGuiConfigObj("SELL_PKGLST_EXEC", False)
'
'   Call CreateGuiConfigObj("BUY_DO_EXEC", True)
'   Call CreateGuiConfigObj("BUY_INVOICE_EXEC", True)
'   Call CreateGuiConfigObj("BUY_PO_EXEC", True)
'   Call CreateGuiConfigObj("BUY_QUOATATION_EXEC", True)
'   Call CreateGuiConfigObj("BUY_RECEIPT_EXEC", True)
'   Call CreateGuiConfigObj("BUY_BILLS_EXEC", True)
'   Call CreateGuiConfigObj("BUY_DBN_EXEC", True)
'   Call CreateGuiConfigObj("BUY_CDN_EXEC", True)
'   Call CreateGuiConfigObj("BUY_PKGLST_EXEC", False)
'   '=====
'
'   '====
'   Call CreateGuiConfigObj("AGENT_SERVER_CONNECT", False)
'
'   If ProgramOwner = XEROX_OWNER Then 'ซีร็อคกราฟฟิค
'      Call SetGuiConfigObj("HR_VIEW", False)
'      Call SetGuiConfigObj("PRODUCTION_VIEW", False)
'
'      Call SetGuiConfigObj("SELL_PO_PREFORM_PRINT", True)
'      Call SetGuiConfigObj("SELL_RECEIPT_PREFORM_PRINT", True)
'   ElseIf ProgramOwner = PLAZA_OWNER Then 'พลาซ่า
'      Call SetGuiConfigObj("PRODUCTION_VIEW", False)
'
'      Call SetGuiConfigObj("SELL_DO_PREFORM_PRINT", True)
'      Call SetGuiConfigObj("SELL_INVOICE_PREFORM_PRINT", True)
'      Call SetGuiConfigObj("SELL_RECEIPT_PREFORM_PRINT", True)
'
'    '  Call SetGuiConfigObj("AGENT_SERVER_CONNECT", True)
'   ElseIf ProgramOwner = FACTORY_OWNER Then 'โรงงานถุงพลาสติก
'      Call SetGuiConfigObj("HR_VIEW", False)
''      Call SetGuiConfigObj("AGENT_SERVER_CONNECT", True)
'   ElseIf ProgramOwner = KLONGTOM_OWNER Then 'คลองถม
'      Call SetGuiConfigObj("HR_VIEW", False)
'      Call SetGuiConfigObj("PRODUCTION_VIEW", False)
'
'      Call SetGuiConfigObj("HR_VIEW", False)
'      Call SetGuiConfigObj("PRODUCTION_VIEW", False)
'
'      Call SetGuiConfigObj("SELL_DO_EXEC", False)
'      Call SetGuiConfigObj("SELL_INVOICE_EXEC", False)
'      Call SetGuiConfigObj("SELL_PO_EXEC", False)
'      Call SetGuiConfigObj("SELL_QUOATATION_EXEC", False)
'      Call SetGuiConfigObj("SELL_RECEIPT_EXEC", False)
'      Call SetGuiConfigObj("SELL_BILLS_EXEC", False)
'      Call SetGuiConfigObj("SELL_DBN_EXEC", False)
'      Call SetGuiConfigObj("SELL_CDN_EXEC", False)
'      Call SetGuiConfigObj("SELL_PKGLST_EXEC", False)
'
'      Call SetGuiConfigObj("BUY_DO_EXEC", True)
'      Call SetGuiConfigObj("BUY_INVOICE_EXEC", True)
'      Call SetGuiConfigObj("BUY_PO_EXEC", False)
'      Call SetGuiConfigObj("BUY_QUOATATION_EXEC", True)
'      Call SetGuiConfigObj("BUY_RECEIPT_EXEC", False)
'      Call SetGuiConfigObj("BUY_BILLS_EXEC", False)
'      Call SetGuiConfigObj("BUY_DBN_EXEC", False)
'      Call SetGuiConfigObj("BUY_CDN_EXEC", False)
'      Call SetGuiConfigObj("BUY_PKGLST_EXEC", True)
'   ElseIf ProgramOwner = TNKP_OWNER Then
'      Call SetGuiConfigObj("LOGO_VIEW", True)
'      Call SetGuiConfigObj("HR_VIEW", False)
'   End If
End Sub

Public Function VerifyGuiConfig(KeyCode As String, Optional PrintError As Boolean = False) As Boolean
Dim TempC As CGuiConfig

'   Set TempC = m_GuiConfigs(KeyCode)
'   VerifyGuiConfig = TempC.ENABLE_FLAG
'
'   If Not TempC.ENABLE_FLAG Then
'      If PrintError Then
'         glbErrorLog.LocalErrorMsg = "โปรแกรมไม่สนับสนุนฟังก์ชันนี้ในเวอร์ชันนี้"
'         glbErrorLog.ShowUserError
'      End If
'   End If
End Function

Public Function ShowWindowCaption(Programowner As String) As String
'   If ProgramOwner = PLAZA_OWNER Then
'      ShowWindowCaption = MapText("BAINISYS ACCOUNTING PRO by BAINISYS Co., Ltd.")
'   ElseIf ProgramOwner = FACTORY_OWNER Then
'      ShowWindowCaption = MapText("BAINISYS ACCOUNTING PRO by BAINISYS Co., Ltd.")
'   Else
'      ShowWindowCaption = MapText("ระบบบริหารงาน WIN Pricing 400")
'   End If
End Function
