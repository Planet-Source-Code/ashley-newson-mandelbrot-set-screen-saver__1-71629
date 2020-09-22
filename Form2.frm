VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSSS Configuration"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3015
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll11 
      Height          =   255
      Left            =   2040
      Max             =   255
      TabIndex        =   34
      Top             =   5040
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2040
      Max             =   255
      Min             =   1
      TabIndex        =   7
      Top             =   360
      Value           =   1
      Width           =   855
   End
   Begin VB.HScrollBar HScroll10 
      Height          =   255
      Left            =   2040
      Max             =   255
      Min             =   1
      TabIndex        =   31
      Top             =   4680
      Value           =   1
      Width           =   855
   End
   Begin VB.HScrollBar HScroll9 
      Height          =   255
      Left            =   2040
      Max             =   255
      Min             =   1
      TabIndex        =   29
      Top             =   4320
      Value           =   1
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":164A
      Left            =   1560
      List            =   "Form2.frx":165A
      TabIndex        =   25
      Text            =   "Ocean"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "No Highlight Box"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   3600
      Width           =   1575
   End
   Begin VB.HScrollBar HScroll8 
      Height          =   255
      LargeChange     =   16
      Left            =   240
      Max             =   255
      TabIndex        =   19
      Top             =   3240
      Value           =   255
      Width           =   735
   End
   Begin VB.HScrollBar HScroll7 
      Height          =   255
      LargeChange     =   16
      Left            =   1200
      Max             =   255
      TabIndex        =   18
      Top             =   3240
      Value           =   255
      Width           =   735
   End
   Begin VB.HScrollBar HScroll6 
      Height          =   255
      LargeChange     =   16
      Left            =   2160
      Max             =   255
      TabIndex        =   17
      Top             =   3240
      Value           =   255
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test and Exit"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   2520
      Width           =   1095
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   255
      Left            =   2040
      Max             =   255
      Min             =   1
      TabIndex        =   15
      Top             =   720
      Value           =   1
      Width           =   855
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   255
      LargeChange     =   16
      Left            =   2160
      Max             =   255
      TabIndex        =   10
      Top             =   2160
      Value           =   255
      Width           =   735
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   16
      Left            =   1200
      Max             =   255
      TabIndex        =   9
      Top             =   2160
      Value           =   255
      Width           =   735
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   16
      Left            =   240
      Max             =   255
      TabIndex        =   8
      Top             =   2160
      Value           =   255
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Custom Colour:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Grey Scale"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Preset Theme:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Zoom Skip: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   3000
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   3000
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   3000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   3000
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label16 
      Caption         =   "Hover mouse over text for description..."
      Height          =   855
      Left            =   0
      TabIndex        =   32
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Label Label15 
      Caption         =   "Zoom Area Size: 0 "
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Distance From Edge: 0"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label13 
      Caption         =   "Automatic Zoom Rules:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   1800
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label11 
      Caption         =   "Highlight Box Colour (black equals off):"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label10 
      Caption         =   "R"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "G"
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label8 
      Caption         =   "B"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   3240
      Width           =   135
   End
   Begin VB.Label Label7 
      Caption         =   "Detail (Accuracy): 0"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "B"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "G"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   2160
      Width           =   135
   End
   Begin VB.Label Label4 
      Caption         =   "R"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Theme:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Software made by Ashley Newson. Mandelbrot Set Screen Saver V1.2.0 (C) Ashley Newson 2009."
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   5520
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Detail (Resolution): 0"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Mandelbrot Set SS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Disables zoom area highlighting."
End Sub

Private Sub Combo1_Click()
Option1.Value = True
End Sub

Private Sub Command1_Click()
det = HScroll1.Value
Open "mandelss.ini" For Binary As #1
Put #1, 1, det
If Option2.Value = True Then
rr = 255
rg = 255
rb = 255
Else
rr = HScroll2.Value
rg = HScroll3.Value
rb = HScroll4.Value
mi = HScroll5.Value
End If
hr = HScroll8.Value
hg = HScroll7.Value
hb = HScroll6.Value
theme = Combo1.ListIndex
zd = HScroll9.Value
zl = HScroll10.Value
zskip = HScroll11.Value
If Option1.Value = True Then
rr = 0
rg = 0
rb = 0
End If
If Check1.Value = 1 Then
hr = 0
hg = 0
hb = 0
End If
Put #1, 2, rr
Put #1, 3, rg
Put #1, 4, rb
Put #1, 5, mi
Put #1, 6, hr
Put #1, 7, hg
Put #1, 8, hb
Put #1, 9, theme
Put #1, 10, zd
Put #1, 11, zl
Put #1, 12, zskip
Close #1
End
End Sub

Private Sub Command2_Click()
If Option2.Value = True Then
rr = 255
rg = 255
rb = 255
Else
rr = HScroll2.Value
rg = HScroll3.Value
rb = HScroll4.Value
hr = HScroll8.Value
hg = HScroll7.Value
hb = HScroll6.Value
theme = Combo1.ListIndex
zd = HScroll9.Value
zl = HScroll10.Value
zskip = HScroll11.Value
End If
If Option1.Value = True Then
rr = 0
rg = 0
rb = 0
End If
If Check1.Value = 1 Then
hr = 0
hg = 0
hb = 0
End If
mi = HScroll5.Value
det = HScroll1.Value
preview = True
Load Form1
Form1.Show
End Sub

Private Sub Form_Load()
Unload Form1
Open "mandelss.ini" For Binary As #1
Get #1, 1, det
Get #1, 2, rr
Get #1, 3, rg
Get #1, 4, rb
Get #1, 5, mi
Get #1, 6, hr
Get #1, 7, hg
Get #1, 8, hb
Get #1, 9, theme
Get #1, 10, zd
Get #1, 11, zl
Get #1, 12, zskip
HScroll2.Value = rr
HScroll3.Value = rg
HScroll4.Value = rb
HScroll8.Value = hr
HScroll7.Value = hg
HScroll6.Value = hb
Close #1
If det = 0 Then det = 60
If mi = 0 Then mi = 128
If zd = 0 Then zd = 15
If zl = 0 Then zl = 30
HScroll5.Value = mi
HScroll1.Value = det
HScroll9.Value = zd
HScroll10.Value = zl
HScroll11.Value = zskip
Option3.Value = True
If rr = 0 And rg = 0 And rb = 0 Then Option1.Value = True
If hr = 0 And hg = 0 And hb = 0 Then Check1.Value = 1
If rr = 255 And rg = 255 And rb = 255 Then Option2.Value = True
Combo1.ListIndex = theme
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Hover mouse over text for description..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
Label1.Caption = "Detail (Resolution):" + Str$(HScroll1.Value)
End Sub

Private Sub HScroll11_Change()
Label17.Caption = "Zoom Skip:" + Str$(HScroll11.Value)
End Sub

Private Sub HScroll2_Change()
Shape1.FillColor = RGB(HScroll2.Value, HScroll3.Value, HScroll4.Value)
End Sub

Private Sub HScroll3_Change()
Shape1.FillColor = RGB(HScroll2.Value, HScroll3.Value, HScroll4.Value)
End Sub

Private Sub HScroll4_Change()
Shape1.FillColor = RGB(HScroll2.Value, HScroll3.Value, HScroll4.Value)
End Sub

Private Sub HScroll5_Change()
Label7.Caption = "Detail (Accuracy):" + Str$(HScroll5.Value)
End Sub

Private Sub HScroll6_Change()
Shape2.FillColor = RGB(HScroll8.Value, HScroll7.Value, HScroll6.Value)
End Sub

Private Sub HScroll7_Change()
Shape2.FillColor = RGB(HScroll8.Value, HScroll7.Value, HScroll6.Value)
End Sub

Private Sub HScroll8_Change()
Shape2.FillColor = RGB(HScroll8.Value, HScroll7.Value, HScroll6.Value)
End Sub

Private Sub HScroll9_Change()
Label14.Caption = "Distance from edge:" + Str$(HScroll9.Value)
End Sub

Private Sub HScroll10_Change()
Label15.Caption = "Zoom Level:" + Str$(HScroll10.Value)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Resolution of image (for example: 15 being 1 image pixel displayed over 1 screen pixel or 30 being 1 image pixel displayed over 2 screen pixels)."
End Sub

Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Red amount."
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sets the appearence of the zoom highlight box if any."
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "This program was run with the arguments: " + Command$
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Defines the rules for the automatic zoom."
End Sub

Private Sub Label14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sets the distance from the mandelbrot set to the point of zoom in pixels."
End Sub

Private Sub Label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sets the size of the zoom area height in pixels."
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sets the number of images to skip after restart."
End Sub

Private Sub label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "This software was made by Ashley Newson. Thank you for using MSSS."
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Sets the colour scheme of the mandelbrot set. Use either a preset theme or a custom colour."
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Red amount."
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Green amount."
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Blue amount."
End Sub

Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Colour Accuracy measured in max iterations (for example: accuracy 2 is higher accuracy than accuracy 1)."
End Sub

Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Blue amount."
End Sub

Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Green amount."
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Use the preset theme on the right."
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Grey Scale Mandelbrot (from white to black)."
End Sub

Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label16.Caption = "Use a custom colour using the colour chooser below using red, green and blue amounts. (From colour to black)."
End Sub
