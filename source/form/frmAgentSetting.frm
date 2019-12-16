VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#2.0#0"; "THREED20.OCX"
Begin VB.Form frmAgentSetting 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "AngsanaUPC"
      Size            =   14.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgentSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6225
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   795
      Left            =   0
      TabIndex        =   5
      Top             =   -210
      Width           =   6225
      Begin VB.Label lblHeader 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   14.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   150
         TabIndex        =   6
         Top             =   300
         Width           =   5955
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6225
      Begin VB.TextBox txtIP 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1380
         TabIndex        =   0
         Top             =   420
         Width           =   3855
      End
      Begin VB.TextBox txtPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1380
         TabIndex        =   1
         Top             =   870
         Width           =   1965
      End
      Begin Threed.SSCommand cmdCancel 
         Height          =   525
         Left            =   3000
         TabIndex        =   8
         Top             =   1500
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   926
         _Version        =   131073
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdOK 
         Height          =   525
         Left            =   1380
         TabIndex        =   7
         Top             =   1500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   926
         _Version        =   131073
         MousePointer    =   99
         MouseIcon       =   "frmAgentSetting.frx":08CA
         ButtonStyle     =   3
      End
      Begin VB.Label lblIP 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   150
         TabIndex        =   4
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label lblPort 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmAgentSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public OKClick As Boolean
Public Header As String

Public IP As String
Public Port As String

Private Sub cmdCancel_Click()
   OKClick = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If Not VerifyTextData(lblIP, txtIP, False) Then
      Exit Sub
   End If
   If Not VerifyTextData(lblPort, txtPort, False) Then
      Exit Sub
   End If
   
   IP = txtIP.Text
   Port = txtPort.Text
   
   Call EnableForm(Me, False)
   If Not glbDatabaseMngr.ConnectAgentServer(IP, Port, glbErrorLog) Then
      Call EnableForm(Me, True)
      txtIP.SetFocus
      
      Exit Sub
   End If
   
   Call EnableForm(Me, True)
   
   OKClick = True
   Unload Me
End Sub

Private Sub Form_Load()

   Me.BackColor = GLB_FORM_COLOR
   Frame1.BackColor = GLB_FORM_COLOR
   lblHeader.BackColor = GLB_HEAD_COLOR
   Frame2.BackColor = GLB_HEAD_COLOR

   OKClick = False
    
   Call InitNormalLabel(lblIP, MapText("IP Address"))
   Call InitNormalLabel(lblPort, MapText("æÕ√Ïµ"))
   
   Call InitTextBox(txtPort, glbParameterObj.LicensePort)
   Call InitTextBox(txtIP, glbParameterObj.LicenseIP)

   Call InitMainButton(cmdOK, MapText("µ°≈ß"))
   Call InitMainButton(cmdCancel, MapText("¬°‡≈‘°"))
   
   Call InitDialogHeader(lblHeader, "∑”°“√‡´µ§Ë“¬Ÿ ‡´Õ√Ï‡Õ‡®πµÏ‡´‘√Ïø‡«Õ√Ï")
   Call SetTextLenType(txtIP, TEXT_STRING, glbSetting.IP_TYPE)
   Call SetTextLenType(txtPort, TEXT_INTEGER, glbSetting.PORT_TYPE)
End Sub

Private Sub txtIP_GotFocus()
   Call SetSelect(txtIP)
End Sub

Private Sub txtPort_GotFocus()
   Call SetSelect(txtPort)
End Sub
