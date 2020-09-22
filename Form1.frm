VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "MSSS"
   ClientHeight    =   4335
   ClientLeft      =   -60
   ClientTop       =   -75
   ClientWidth     =   5775
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "Form1.frx":164A
   MousePointer    =   99  'Custom
   ScaleHeight     =   4335
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Display 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox mem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox startimage 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Mandelbrot Set Screen Saver is rendering the first image..."
      BeginProperty Font 
         Name            =   "System"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cc As Boolean
Public firstimages As Boolean
Public X1 As Double
Public Y1 As Double
Public X2 As Double
Public Y2 As Double

Private Sub Display_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Enabled = True Then End
End Sub

Private Sub Display_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = True Then End
End Sub

Private Sub Display_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If preview = False And Timer1.Enabled = True Then
If Label1.Caption <> Str$(X) + "," + Str$(Y) Then
If Label1.Caption = "" Then
Label1.Caption = Str$(X) + "," + Str$(Y)
Else
End
End If
End If
End If
End Sub

Private Sub label2_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Enabled = True Then End
End Sub

Private Sub label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = True Then End
End Sub

Private Sub label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If preview = False And Timer1.Enabled = True Then
If Label1.Caption <> Str$(X) + "," + Str$(Y) Then
If Label1.Caption = "" Then
Label1.Caption = Str$(X) + "," + Str$(Y)
Else
End
End If
End If
End If
End Sub

Private Sub Form_Load()
On Error GoTo fatalload
If 0 = 1 Then
fatalload:
MsgBox ("Fatal Error. Error" + Str$(Err.Number) + ". " + Err.Description)
End
End If
If Left$(Command$, 2) = "/p" Or Left$(Command$, 2) = "/P" Then End
If (Command$ <> "/s" And Command$ <> "/S") And preview = False Then
Load Form2
Form2.Show
Exit Sub
Else
If preview = False Then
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
Close #1
End If
If det = 0 Then det = 60
If mi = 0 Then mi = 128
If zd = 0 Then zd = 15
If zl = 0 Then zl = 30
cc = True
If rr = 0 And rg = 0 And rb = 0 Then cc = False
Timer2.Enabled = True
Randomize Timer
X1 = -2
Y1 = -1.5
X2 = 1
Y2 = 1.5
firstimages = True
End If
End Sub

Private Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Enabled = True Then End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Timer1.Enabled = True Then End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If preview = False And Timer1.Enabled = True Then
If Label1.Caption <> Str$(X) + "," + Str$(Y) Then
If Label1.Caption = "" Then
Label1.Caption = Str$(X) + "," + Str$(Y)
Else
End
End If
End If
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo fatal
If 0 = 1 Then
fatal:
End
End If
Dim z As complex
Dim tz As complex
Dim t As complex
sc = (Y2 - Y1) / Form1.Height
drr = rr / mi
drg = rg / mi
drb = rb / mi
rzd = zd * 15
rzl = zl * 15
If cc = False Then
If theme = 0 Then cm = 100 / mi
If theme = 1 Then cm = (255 / mi) * 3
If theme = 2 Then cm = (255 / mi) * 3
If theme = 3 Then cm = (255 / mi) * 3
Else
cm = 1
End If
If X1 = -2 And X2 = 1 And Display.Visible = True And zskip > 0 Then
Set mem.Picture = startimage.Image
Else
For i = 0 To Form1.Width Step det
For j = 0 To Form1.Height Step det
X = X1 + (i * sc)
Y = Y1 + (j * sc)

t.r = X
t.i = Y

z.r = 0
z.i = 0

For i2 = 1 To mi
    tz.r = (z.r * z.r - z.i * z.i) + t.r
    tz.i = (z.r * z.i * 2) + t.i

    z = tz

    If (z.r * z.r) + (z.i * z.i) >= 4 Then
        c = i2 * cm
        GoTo done
    End If
Next
c = 0
done:
If cc = False Then
If theme = 0 Then mem.Line (i, j)-(i + det, j + det), RGB((c Mod 60) * 10, (c Mod 100) * 20, (c Mod 200) * 200), BF
If theme >= 1 And theme <= 3 Then
            If c < 256 Then
                cg = c
                cr = 256 - cg
                cb = 0
            End If
            If c >= 256 And c < 512 Then
                cb = c - 256
                cg = 255 - cb
                cr = 0
            End If
            If c >= 512 And c < 768 Then
                cr = c - 512
                cb = 255 - cr
                cg = 0
            End If
            If c = 0 Then
                cr = 0
                cg = 0
                cb = 0
            End If
'If theme = 1 Then mem.PSet (i, j), RGB(cr, cg, cb)
'If theme = 2 Then mem.PSet (i, j), RGB(cb, cr, cg)
'If theme = 3 Then mem.PSet (i, j), RGB(cg, cb, cr)
If theme = 1 Then mem.Line (i, j)-(i + det, j + det), RGB(cr, cg, cb), BF
If theme = 2 Then mem.Line (i, j)-(i + det, j + det), RGB(cb, cr, cg), BF
If theme = 3 Then mem.Line (i, j)-(i + det, j + det), RGB(cg, cb, cr), BF
End If
Else
mem.Line (i, j)-(i + det, j + det), RGB(c * drr, c * drg, c * drb), BF
End If
Next
DoEvents
Next
End If
If Display.Visible = False Then
Set startimage.Picture = mem.Image
End If
Display.Visible = True

If za >= zskip Or firstimages = True Then
Set Display.Picture = mem.Image
DoEvents
rxy = Form1.Height / Form1.Width
With Display
Do
ln = 0
Do
X = Int(Rnd * Form1.Width)
Y = Int(Rnd * Form1.Height)
Blue = Int(.Point(X, Y) / 65536)
Green = Int((.Point(X, Y) - (Blue * 65536)) / 256)
Red = (.Point(X, Y) - (Blue * 65536) - (Green * 256))
If .Point(X, Y) <> 0 Then
If cc = True Then
    If (Red < (rr / 4) + 5 And Green < (rg / 4) + 5 And Blue < (rb / 4) + 5) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Red < (rr / 3) + 5 And Green < (rg / 3) + 5 And Blue < (rb / 3) + 5 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Red < (rr / 2) + 5 And Green < (rg / 2) + 5 And Blue < (rb / 2) + 5 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
Else
If (Red < 10 And Green >= 254 And Blue >= 254 And theme = 0) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
X2 = (X + (rzl / rxy) - rzl) * sc + X1
Y2 = (Y + rzl) * sc + Y1
X1 = (X - (rzl / rxy)) * sc + X1
Y1 = (Y - rzl) * sc + Y1
za = za + 1
GoTo finish
End If
If theme = 1 Then
    If (Red > 0 And Blue = 0 And Green <= 255) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Red = 0 And Blue < 64 And Green > 0 And Green <= 255 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Red = 0 And Blue < 128 And Green > 0 And Green <= 255 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
End If
If theme = 2 Then
    If (Green > 0 And Red = 0 And Blue <= 255) And (Form1.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Green = 0 And Red < 64 And Blue > 0 And Blue <= 255 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Green = 0 And Red < 128 And Blue > 0 And Blue <= 255 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
End If
If theme = 3 Then
    If (Blue > 0 And Green = 0 And Red <= 255) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Blue = 0 And Green < 64 And Red > 0 And Red <= 255 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
    If (Blue = 0 And Green < 128 And Red > 0 And Red <= 255 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish
    End If
End If
End If
End If
ln = ln + 1
If ln > 30000 Then DoEvents
Loop Until ln = 30000
DoEvents
Loop Until X1 <> -2 Or X2 <> 1
End With
X1 = -2
Y1 = -1.5
X2 = 1
Y2 = 1.5
za = 0
firstimages = False
finish:
If ln <> 30000 And (hr <> 0 Or hg <> 0 Or hb <> 0) Then Display.Line (X - (rzl / rxy) - 15, Y - rzl - 15)-(X + (rzl / rxy) + 15, Y + rzl + 15), RGB(hr, hg, hb), B
Else
'||| ZOOM SKIP |||
'vvv ZOOM SKIP vvv
DoEvents
rxy = Form1.Height / Form1.Width
With mem
Do
ln = 0
Do
X = Int(Rnd * Form1.Width)
Y = Int(Rnd * Form1.Height)
Blue = Int(.Point(X, Y) / 65536)
Green = Int((.Point(X, Y) - (Blue * 65536)) / 256)
Red = (.Point(X, Y) - (Blue * 65536) - (Green * 256))
If .Point(X, Y) <> 0 Then
If cc = True Then
    If (Red < (rr / 4) + 5 And Green < (rg / 4) + 5 And Blue < (rb / 4) + 5) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Red < (rr / 3) + 5 And Green < (rg / 3) + 5 And Blue < (rb / 3) + 5 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Red < (rr / 2) + 5 And Green < (rg / 2) + 5 And Blue < (rb / 2) + 5 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
Else
If (Red < 10 And Green >= 254 And Blue >= 254 And theme = 0) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
X2 = (X + (rzl / rxy) - rzl) * sc + X1
Y2 = (Y + rzl) * sc + Y1
X1 = (X - (rzl / rxy)) * sc + X1
Y1 = (Y - rzl) * sc + Y1
za = za + 1
GoTo finish2
End If
If theme = 1 Then
    If (Red > 0 And Blue = 0 And Green <= 255) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Red = 0 And Blue < 64 And Green > 0 And Green <= 255 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Red = 0 And Blue < 128 And Green > 0 And Green <= 255 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
End If
If theme = 2 Then
    If (Green > 0 And Red = 0 And Blue <= 255) And (Form1.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    GoTo finish2
    End If
    If (Green = 0 And Red < 64 And Blue > 0 And Blue <= 255 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Green = 0 And Red < 128 And Blue > 0 And Blue <= 255 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
End If
If theme = 3 Then
    If (Blue > 0 And Green = 0 And Red <= 255) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Blue = 0 And Green < 64 And Red > 0 And Red <= 255 And ln > 10000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
    If (Blue = 0 And Green < 128 And Red > 0 And Red <= 255 And ln > 20000) And (.Point(X - rzd, Y) = 0 Or .Point(X + rzd, Y) = 0 Or .Point(X, Y - rzd) = 0 Or .Point(X, Y + rzd) = 0) Then
    X2 = (X + (rzl / rxy) - rzl) * sc + X1
    Y2 = (Y + rzl) * sc + Y1
    X1 = (X - (rzl / rxy)) * sc + X1
    Y1 = (Y - rzl) * sc + Y1
    za = za + 1
    GoTo finish2
    End If
End If
End If
End If
ln = ln + 1
If ln > 30000 Then DoEvents
Loop Until ln = 30000
DoEvents
Loop Until X1 <> -2 Or X2 <> 1
End With
X1 = -2
Y1 = -1.5
X2 = 1
Y2 = 1.5
za = 0
firstimages = False
finish2:
End If
End Sub

Private Sub Timer2_Timer()
Timer1.Enabled = True
Form1.WindowState = 2
Form1.Show
Display.Width = Form1.Width
Display.Height = Form1.Height
mem.Width = Form1.Width
mem.Height = Form1.Height
startimage.Width = Form1.Width
startimage.Height = Form1.Height
Label2.Left = (Form1.Width - 5775) / 2
Label2.Top = (Form1.Height - 975) / 2
End Sub
