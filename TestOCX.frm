VERSION 5.00
Object = "{1647A91A-CE99-11D4-98F0-BFE9D8EEEB63}#1.0#0"; "ASPICTURE.OCX"
Begin VB.Form TestOCX 
   Caption         =   "AutoScrollPicture Control - very easy for use!"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6810
   Icon            =   "TestOCX.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   6810
   StartUpPosition =   2  'CenterScreen
   Begin AutoScrollPictureBox.ASPictureBox ASPictureBox1 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      Picture         =   "TestOCX.frx":030A
      BackColor       =   16777215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save to file"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Picture"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   4560
      Left            =   480
      Picture         =   "TestOCX.frx":2DAB
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   6060
   End
   Begin VB.CommandButton cmdLoad4 
      Caption         =   "Load Picture4"
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad3 
      Caption         =   "Load Picture3"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad2 
      Caption         =   "Load Picture2"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdLoad1 
      Caption         =   "Load Picture1"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "TestOCX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
'AutoScrollPicture - very easy for use!
'**************************************
'Copyright © 2000 by Alexander Anikin
'e-mail: pegas@poshuk.com
'http://www.poshuk.com/pegas/index.htm
'**************************************
Option Explicit
Private Sub cmdClear_Click()
 ASPictureBox1.Picture = LoadPicture()

End Sub

Private Sub cmdLoad1_Click()
 ASPictureBox1.Picture = Picture1

End Sub

Private Sub cmdLoad2_Click()
 ASPictureBox1.Picture = LoadPicture(App.Path & "\ver.gif")

End Sub

Private Sub cmdLoad3_Click()
 ASPictureBox1.Picture = LoadPicture(App.Path & "\hor.gif")

End Sub

Private Sub cmdLoad4_Click()
 ASPictureBox1.Picture = Me.Icon

End Sub

Private Sub cmdSave_Click()

 Dim MyStr As String
 Dim MyExtension As String

 'Returns the graphic format of a Picture object
 Select Case ASPictureBox1.Picture.Type
  Case vbPicTypeBitmap
   MyExtension = "bmp"
  Case vbPicTypeEMetafile
   MyExtension = "emf"
  Case vbPicTypeIcon
   MyExtension = "ico"
  Case vbPicTypeMetafile
   MyExtension = "wmf"
  Case vbPicTypeNone
   Beep
   Exit Sub
 End Select

 MyStr = App.Path & "\MyPicture." & MyExtension
 SavePicture ASPictureBox1.Picture, MyStr
 MsgBox MyStr, , "Saved to:"

End Sub
