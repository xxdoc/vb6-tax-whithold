VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmTaxReport 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "frmTaxReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel pnlHeader 
      Height          =   705
      Left            =   -60
      TabIndex        =   7
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   1244
      _Version        =   131073
      PictureBackgroundStyle=   2
   End
   Begin Threed.SSFrame SSFrame 
      Height          =   4215
      Left            =   -60
      TabIndex        =   8
      Top             =   630
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   7435
      _Version        =   131073
      PictureBackgroundStyle=   2
      Begin VB.ComboBox cboOrderType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   3855
      End
      Begin VB.ComboBox cboOrderBy 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2010
         Width           =   3855
      End
      Begin prjMtpTax.uctlDate uctlToDate 
         Height          =   405
         Left            =   1980
         TabIndex        =   3
         Top             =   1560
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   714
      End
      Begin prjMtpTax.uctlTextBox txtBranch 
         Height          =   405
         Left            =   1980
         TabIndex        =   0
         Top             =   300
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   714
      End
      Begin VB.ComboBox cboTaxType 
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   9
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   3855
      End
      Begin prjMtpTax.uctlDate uctlFromDate 
         Height          =   435
         Left            =   1980
         TabIndex        =   2
         Top             =   1140
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   767
      End
      Begin VB.Label lblOrderBy 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   2490
         Width           =   1935
      End
      Begin VB.Label lblOrderType 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   2070
         Width           =   1935
      End
      Begin VB.Label lblToDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   0
         TabIndex        =   13
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblTaxType 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   0
         TabIndex        =   12
         Top             =   810
         Width           =   1935
      End
      Begin VB.Label lblBranch 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   0
         TabIndex        =   11
         Top             =   420
         Width           =   1935
      End
      Begin Threed.SSCommand cmdExit 
         Cancel          =   -1  'True
         Height          =   525
         Left            =   3300
         TabIndex        =   10
         Top             =   3390
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1650
         TabIndex        =   6
         Top             =   3390
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmTaxReport.frx":030A
         ButtonStyle     =   3
      End
      Begin VB.Label lblFromDate 
         Alignment       =   1  'Right Justify
         Caption         =   "1"
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   1260
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmTaxReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public HeaderText As String


Private Sub InitFormLayout()
pnlHeader.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
SSFrame.PictureBackground = LoadPicture(glbParameterObj.NormalForm1)
   
Me.Caption = HeaderText
pnlHeader.Caption = Me.Caption

   Call InitNormalLabel(lblBranch, MapText("ชื่อบริษัท (สาขา)"))
   Call InitNormalLabel(lblTaxType, MapText("ประเภทภาษี"))
   Call InitNormalLabel(lblFromDate, MapText("จากวันที่"))
   Call InitNormalLabel(lblToDate, MapText("ถึงวันที่"))
   Call InitNormalLabel(lblOrderType, MapText("เรียงตาม"))
   Call InitNormalLabel(lblOrderBy, MapText("เรียงจาก"))

   Call txtBranch.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
'   Call cboTaxType.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
'   Call cboFromdate.SetTextLenType(TEXT_STRING, glbSetting.DEST_TYPE)
'   Call Cbo.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)
'   Call txtNote.SetTextLenType(TEXT_STRING, glbSetting.DESC_TYPE)
'   Call txtBranch.SetTextLenType(TEXT_STRING, glbSetting.CODE_TYPE)

   Me.Picture = LoadPicture(glbParameterObj.MainPicture)

   pnlHeader.Font.Name = GLB_FONT
   pnlHeader.Font.Bold = True
   pnlHeader.Font.Size = 19

   cmdExit.Picture = LoadPicture(glbParameterObj.NormalButton1)
   cmdOK.Picture = LoadPicture(glbParameterObj.NormalButton1)

Call InitMainButton(cmdOK, MapText("ตกลง (F2)"))
Call InitMainButton(cmdExit, MapText("ยกเลิก "))

 Call InitCombo(cboTaxType)
Call InitCombo(cboOrderType)
Call InitCombo(cboOrderBy)

 Call InitTaxType(cboTaxType)
 Call InitReportOR(cboOrderBy)
 Call InitOrderType(cboOrderType)

End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Report As CReportInterface

If Not VerifyTextControl(lblBranch, txtBranch, False) Then
      Exit Sub
   End If
  If Not VerifyCombo(lblTaxType, cboTaxType, False) Then
      Exit Sub
   End If
   If Not VerifyDate(lblFromDate, uctlFromDate, False) Then
      Exit Sub
   End If
If Not VerifyDate(lblToDate, uctlToDate, False) Then
      Exit Sub
   End If
Select Case cboTaxType
   Case 1:
         Set Report = New CReportTax002
   Case 2:
         Set Report = New CReportTax003
   Case 3:
         Set Report = New CReportTax0053
   Case 4:
         Set Report = New CReportTaxSending
   Case 5:
         Set Report = New CReportTaxSending
   Case 6:
        Set Report = New CReportTaxSending
   
End Select
   

Call Report.AddParam(txtBranch, "BRANCH")
Call Report.AddParam(cboTaxType, "TAXTYPE")
Call Report.AddParam(uctlFromDate, "FROM_DATE")
Call Report.AddParam(uctlToDate, "TO_DATE")
Call Report.AddParam(cboOrderBy, "ORDER_BY")
Call Report.AddParam(cboOrderType, "ORDER_TYPE")


       

      Set frmReport.ReportObject = Report
      frmReport.HeaderText = MapText("พิมพ์รายงาน")
      Load frmReport
      frmReport.Show 1

      Unload frmReport
      Set frmReport = Nothing

End Sub

Private Sub Form_Activate()
Call InitFormLayout
End Sub

