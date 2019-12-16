VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'On Error GoTo ErrorHandler
'Dim lMenuChosen As Long
'Dim oMenu As cPopupMenu
'
'   Set oMenu = New cPopupMenu
'   lMenuChosen = oMenu.Popup("พิมพ์ไปเครื่องพิมพ์", "-", "บันทึกไปที่ไฟล์")
'   If lMenuChosen = 0 Then
'      Exit Sub
'   End If
'
'   If lMenuChosen = 1 Then
'      VSPrinter1.PrintDoc (True)
'      If m_ErrorFlag Then
'         glbErrorLog.LocalErrorMsg = "พบข้อผิดพลาด"
'         glbErrorLog.ShowUserError
'         Exit Sub
'      Else
'         glbErrorLog.LocalErrorMsg = "โปรแกรมได้ทำการพิมพ์รายงานเสร็จสิ้นแล้ว"
'         glbErrorLog.ShowUserError
'         Exit Sub
'      End If
'   ElseIf lMenuChosen = 3 Then
'      CommonDialog1.Filter = "Save Files (*.html, *.htm)|*.html;*.htm;"
'      CommonDialog1.DialogTitle = "Select access file to import"
'      CommonDialog1.ShowSave
'      If CommonDialog1.FileName = "" Then
'         Exit Sub
'      End If
'
'      Call FileCopy(glbParameterObj.ReportFile, CommonDialog1.FileName)
'   End If
'
'   OKClick = True
'   Unload Me
'
'   Exit Sub
'
'ErrorHandler:
'   glbErrorLog.SystemErrorMsg = Err.DESCRIPTION
'   glbErrorLog.ShowErrorLog (LOG_FILE_MSGBOX)
