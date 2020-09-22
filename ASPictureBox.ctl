VERSION 5.00
Begin VB.UserControl ASPictureBox 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4215
   PropertyPages   =   "ASPictureBox.ctx":0000
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   ToolboxBitmap   =   "ASPictureBox.ctx":0010
   Begin VB.VScrollBar vsbScroll 
      Height          =   2295
      Left            =   3960
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.HScrollBar hsbScroll 
      Height          =   200
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox picTwo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   261
      TabIndex        =   1
      Top             =   0
      Width           =   3975
   End
   Begin VB.PictureBox picOne 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   81
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ASPictureBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'*************************************
'Copyright © 2000 by Alexander Anikin
'e-mail: pegas@poshuk.com
'http://www.poshuk.com/pegas/index.htm
'*************************************
Event Click()
Attribute Click.VB_MemberFlags = "200"
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event ReadProperties(PropBag As PropertyBag)
Event WriteProperties(PropBag As PropertyBag)

Private Sub hsbScroll_Change()
 UpdatePicTwo
End Sub

Private Sub hsbScroll_Scroll()
 hsbScroll_Change
End Sub

Private Sub picOne_Change()
 If picOne.ScaleWidth <= picTwo.ScaleWidth _
 Or picOne.Picture = LoadPicture() _
 Then hsbScroll.Visible = False

 If picOne.ScaleHeight <= picTwo.ScaleHeight _
 Or picOne.Picture = LoadPicture() _
 Then vsbScroll.Visible = False

 If picOne.Picture = LoadPicture() _
 Then picTwo.Picture = LoadPicture(): Exit Sub
 
 picTwo.Picture = picOne.Picture

 If picOne.ScaleWidth > picTwo.ScaleWidth Then
 hsbScroll.Visible = True
 End If

 If picOne.ScaleHeight > picTwo.ScaleHeight Then
 vsbScroll.Visible = True
 End If

 Call hsbScrollSett_Refresh
 Call vsbScrollSett_Refresh

End Sub

Private Sub picTwo_Click()
    RaiseEvent Click

End Sub

Private Sub UserControl_Initialize()
 UserControl.ScaleMode = vbPixels
 picOne.ScaleMode = vbPixels
 picTwo.ScaleMode = vbPixels

End Sub

Private Sub UserControl_Resize()

 If UserControl.Height < 1500 Then
  UserControl.Height = 1500
 ElseIf UserControl.Width < 1500 Then
  UserControl.Width = 1500
 End If
'******************
 picTwo.Height = UserControl.ScaleHeight - hsbScroll.Height
 picTwo.Width = UserControl.ScaleWidth - vsbScroll.Width
'************************
 vsbScroll.Left = picTwo.Width
 vsbScroll.Height = picTwo.Height
 hsbScroll.Top = picTwo.Height
 hsbScroll.Width = picTwo.Width

 Call picOne_Change

End Sub

Private Sub vsbScroll_Change()

 UpdatePicTwo

End Sub

Private Sub vsbScroll_Scroll()
 vsbScroll_Change
End Sub

Private Sub UpdatePicTwo()

 If hsbScroll.Visible = False _
 And vsbScroll.Visible = False Then Exit Sub

 picTwo.PaintPicture picOne.Picture, 0, 0, _
 picTwo.ScaleWidth, picTwo.ScaleHeight, _
 hsbScroll.Value, vsbScroll.Value, _
 picTwo.ScaleWidth, picTwo.ScaleHeight, _
 vbSrcCopy

End Sub

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
Attribute Picture.VB_UserMemId = 0
Attribute Picture.VB_MemberFlags = "200"
    Set Picture = picOne.Picture
End Property

Public Property Let Picture(ByVal New_Picture As IPictureDisp)
    Set Picture = New_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
 Set picOne.Picture = New_Picture
 PropertyChanged "Picture"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = picTwo.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    picTwo.BackColor() = New_BackColor
    Call UpdatePicTwo
    PropertyChanged "BackColor"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = picTwo.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    picTwo.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub picTwo_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picTwo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub picTwo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub picTwo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    RaiseEvent ReadProperties(PropBag)

    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    picTwo.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    picTwo.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    RaiseEvent WriteProperties(PropBag)

    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", picTwo.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BorderStyle", picTwo.BorderStyle, 1)
End Sub

Private Sub hsbScrollSett_Refresh()
 hsbScroll.Value = 0
 If picOne.ScaleWidth <= picTwo.ScaleWidth Then Exit Sub
 hsbScroll.Max = picOne.ScaleWidth - picTwo.ScaleWidth
 '**************
 If hsbScroll.Max < 25 Then
  hsbScroll.LargeChange = 1
  hsbScroll.SmallChange = 1
 Else
  hsbScroll.LargeChange = hsbScroll.Max \ 10
  hsbScroll.SmallChange = hsbScroll.Max \ 25
 End If
 
End Sub

Private Sub vsbScrollSett_Refresh()
 vsbScroll.Value = 0
 If picOne.ScaleHeight <= picTwo.ScaleHeight Then Exit Sub
 vsbScroll.Max = picOne.ScaleHeight - picTwo.ScaleHeight
 '****************
 If vsbScroll.Max < 25 Then
  vsbScroll.LargeChange = 1
  vsbScroll.SmallChange = 1
 Else
  vsbScroll.LargeChange = vsbScroll.Max \ 10
  vsbScroll.SmallChange = vsbScroll.Max \ 25
End If

End Sub

