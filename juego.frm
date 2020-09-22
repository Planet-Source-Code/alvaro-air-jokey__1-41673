VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Juego 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   Icon            =   "juego.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   64
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ContinuarB 
      Caption         =   "&Continuar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   12
      Top             =   1635
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3675
      Top             =   975
   End
   Begin VB.PictureBox PicSprites 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   4155
      Picture         =   "juego.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   1935
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox J2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1230
      Left            =   3810
      Picture         =   "juego.frx":0A84
      ScaleHeight     =   1230
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   1665
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox J1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1230
      Left            =   3525
      Picture         =   "juego.frx":1A26
      ScaleHeight     =   1230
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1695
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox SalvaFondo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2925
      Left            =   120
      Picture         =   "juego.frx":29C8
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   555
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.PictureBox fondo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   2925
      Left            =   165
      ScaleHeight     =   195
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   570
      Width           =   4500
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   255
         Left            =   3975
         TabIndex        =   30
         Top             =   165
         Width           =   195
      End
      Begin VB.CommandButton BContinuar 
         Caption         =   "&Continuar"
         Height          =   315
         Left            =   1740
         TabIndex        =   25
         Top             =   2610
         Width           =   915
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00008000&
         Caption         =   "Controles"
         ForeColor       =   &H00FFFFFF&
         Height          =   2370
         Left            =   870
         TabIndex        =   15
         Top             =   255
         Width           =   2895
         Begin MSComctlLib.Slider BarraD 
            Height          =   300
            Left            =   675
            TabIndex        =   27
            Top             =   2010
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            LargeChange     =   1
            Min             =   1
            Max             =   4
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00008000&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   45
            TabIndex        =   21
            Top             =   1245
            Width           =   2190
            Begin VB.OptionButton Pcj2 
               BackColor       =   &H00008000&
               Caption         =   "Pc"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   0
               TabIndex        =   26
               Top             =   585
               Width           =   1575
            End
            Begin VB.OptionButton Tec1J2 
               BackColor       =   &H00008000&
               Caption         =   "Teclado1 (W y S)"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton Tec2J2 
               BackColor       =   &H00008000&
               Caption         =   "Teclado2 (Arriva y Abajo)"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   0
               TabIndex        =   23
               Top             =   195
               Width           =   2085
            End
            Begin VB.OptionButton MouJ2 
               BackColor       =   &H00008000&
               Caption         =   "Mouse"
               ForeColor       =   &H00FFFFFF&
               Height          =   195
               Left            =   0
               TabIndex        =   22
               Top             =   390
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.OptionButton MouJ1 
            BackColor       =   &H00008000&
            Caption         =   "Mouse"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   18
            Top             =   795
            Width           =   1575
         End
         Begin VB.OptionButton Tec2J1 
            BackColor       =   &H00008000&
            Caption         =   "Teclado2 (Arriva y Abajo)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   17
            Top             =   600
            Width           =   2085
         End
         Begin VB.OptionButton Tec1J1 
            BackColor       =   &H00008000&
            Caption         =   "Teclado1 (W y S)"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   16
            Top             =   405
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fácil"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   2310
            TabIndex        =   29
            Top             =   2085
            Width           =   330
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Difícil"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   105
            TabIndex        =   28
            Top             =   2070
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jugador2"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   90
            TabIndex        =   20
            Top             =   1005
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jugador1"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   75
            TabIndex        =   19
            Top             =   180
            Width           =   660
         End
      End
      Begin VB.PictureBox Jugador2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   4350
         ScaleHeight     =   615
         ScaleWidth      =   120
         TabIndex        =   9
         Top             =   1170
         Width           =   120
      End
      Begin VB.PictureBox Jugador1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   30
         ScaleHeight     =   615
         ScaleWidth      =   120
         TabIndex        =   8
         Top             =   1170
         Width           =   120
      End
   End
   Begin VB.Line Borde2 
      BorderColor     =   &H00FF0000&
      X1              =   250
      X2              =   159
      Y1              =   16
      Y2              =   21
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00FF0000&
      Height          =   240
      Left            =   2520
      Top             =   30
      Width           =   930
   End
   Begin VB.Image Icono 
      Height          =   240
      Left            =   150
      Picture         =   "juego.frx":2D08E
      Stretch         =   -1  'True
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Nombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Air-Jokey"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   405
      TabIndex        =   2
      Top             =   105
      Width           =   855
   End
   Begin VB.Label EndB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4275
      TabIndex        =   1
      ToolTipText     =   "Cerra"
      Top             =   105
      Width           =   240
   End
   Begin VB.Label MinB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3570
      TabIndex        =   0
      ToolTipText     =   "Minimizar"
      Top             =   105
      Width           =   240
   End
   Begin VB.Image barra 
      Height          =   285
      Left            =   150
      Picture         =   "juego.frx":2D4D0
      Stretch         =   -1  'True
      Top             =   75
      Width           =   4500
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Acerca de..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1590
      TabIndex        =   31
      Top             =   330
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Jugador2:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3105
      TabIndex        =   14
      Top             =   330
      Width           =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Jugador1:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   195
      TabIndex        =   13
      Top             =   330
      Width           =   1080
   End
   Begin VB.Label PuntosJ2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4185
      TabIndex        =   11
      Top             =   330
      Width           =   135
   End
   Begin VB.Label PuntosJ1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1290
      TabIndex        =   10
      Top             =   330
      Width           =   135
   End
End
Attribute VB_Name = "juego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MaxSprites = 1
Dim Sprites(0 To MaxSprites - 1) As T_Sprite
Dim Detener As Boolean
Const Pelota = 0

Dim MouseJ1 As Boolean
Dim MouseJ2 As Boolean
Dim MouseDemora  As Integer
Dim FlechasJ1 As Boolean
Dim FlechasJ2 As Boolean
Dim TeclasJ1 As Boolean
Dim TeclasJ2 As Boolean
Dim VsPc As Boolean

Dim DirX As Variant ' 0 izquierda
Dim DirY As Variant ' 0 arriva
Dim AngY As Integer
Dim PVel As Variant
Dim TAngY As Integer
Dim PcDifi As Integer

Dim TLuzJ1 As Integer
Dim TLuzJ2 As Integer

Private Sub AvanzarAnimacion(i As Integer)
   With Sprites(i)
      .Pas = .Pas + 1
      If .Pas = .Vel Then
         .Pas = 0
         .C = .C + 1
         If .C > .TC - 1 Then .C = 0
      End If
   End With
End Sub

Private Sub MostrarSprite(i As Integer)
  With Sprites(i)
      BitBlt fondo.hDC, .Pos.X, .Pos.Y, .Dim.X, .Dim.Y, PicSprites(i).hDC, .C * .Dim.X, .Dim.Y, vbSrcAnd
      BitBlt fondo.hDC, .Pos.X, .Pos.Y, .Dim.X, .Dim.Y, PicSprites(i).hDC, .C * .Dim.X, 0, vbSrcInvert
   End With
End Sub

Public Sub MoverSprite(i As Integer)
With Sprites(Pelota)
 If VsPc = True Then
    Dim PcError
    If PcDifi > 1 Then
        Randomize
        PcError = Int((Rnd * PcDifi) + 1)
        Select Case PcError
            Case 1: If .Pos.Y + 16 < Jugador2.Top Then Jugador2.Top = Jugador2.Top - 2
            Case 2: If .Pos.Y > Jugador2.Top + 41 Then Jugador2.Top = Jugador2.Top + 2
        End Select
    Else
        PcError = Int((Rnd * 4) + 1)
        Select Case PcError
            Case 1: If .Pos.Y + 16 < Jugador2.Top Then Jugador2.Top = Jugador2.Top - 2
            Case 2: If .Pos.Y > Jugador2.Top + 41 Then Jugador2.Top = Jugador2.Top + 2
            Case 3: If .Pos.Y + 16 < Jugador2.Top Then Jugador2.Top = Jugador2.Top - 2
            Case 4: If .Pos.Y > Jugador2.Top + 41 Then Jugador2.Top = Jugador2.Top + 2
        End Select
    End If
 End If
    If .Pos.Y < 0 Then
        DirY = 1
        sndPlaySound App.Path + "\bank01.WAV", SND_NODEFAULT + SND_ASYNC
    End If
    If .Pos.Y > 179 Then
        DirY = 0
        sndPlaySound App.Path + "\bank01.WAV", SND_NODEFAULT + SND_ASYNC
    End If
    If .Pos.X <= Jugador1.Left + 8 And .Pos.Y + 16 > Jugador1.Top And .Pos.Y < Jugador1.Top + Jugador1.Height Then
        If .Pos.Y < Jugador1.Top + 20 Then AngY = (.Pos.Y - Jugador1.Top) / 2 Else AngY = ((Jugador1.Top + 41) - .Pos.Y) / 2
        BitBlt Jugador1.hDC, 0, 0, 8, 41, J1.hDC, 8, 0, vbSrcCopy
        sndPlaySound App.Path + "\hit01.WAV", SND_NODEFAULT + SND_ASYNC
        Jugador1.Refresh
        TLuzJ1 = 0
        DirX = 1
        TAngY = 0
    End If
    If .Pos.X + 16 >= Jugador2.Left And .Pos.Y + 16 > Jugador2.Top And .Pos.Y < Jugador2.Top + Jugador2.Height Then
        If .Pos.Y < Jugador2.Top + 20 Then AngY = (.Pos.Y - Jugador2.Top) / 2 Else AngY = ((Jugador2.Top + 41) - .Pos.Y) / 2
        BitBlt Jugador2.hDC, 0, 0, 8, 41, J2.hDC, 8, 0, vbSrcCopy
        sndPlaySound App.Path + "\hit01.WAV", SND_NODEFAULT + SND_ASYNC
        Jugador2.Refresh
        TLuzJ2 = 0
        DirX = 0
        TAngY = 0
    End If
TAngY = TAngY + 1
    If DirX = 0 Then .Pos.X = .Pos.X - 1 Else .Pos.X = .Pos.X + 1
    If DirY = 0 And TAngY >= AngY Then
        .Pos.Y = .Pos.Y - 1
        TAngY = 0
    End If
    If DirY = 1 And TAngY >= AngY Then
        .Pos.Y = .Pos.Y + 1
        TAngY = 0
    End If
    If .Pos.X < 0 Or .Pos.X > 283 Then
        Detener = True
        If .Pos.X < 0 Then PuntosJ2 = PuntosJ2 + 1 Else PuntosJ1 = PuntosJ1 + 1
        sndPlaySound App.Path + "\score.WAV", SND_NODEFAULT + SND_ASYNC
        Form_Load
        ContinuarB.Visible = True
    End If
End With
End Sub

Private Sub RestaurarFondo()
  Dim i As Integer
   For i = 0 To MaxSprites - 1
      With Sprites(i)
         If .Visible Then
            BitBlt fondo.hDC, .Pos.X, .Pos.Y, .Dim.X, .Dim.Y, SalvaFondo.hDC, .Pos.X, .Pos.Y, vbSrcCopy
         End If
       End With
   Next i
End Sub

Private Sub Demorar()
If MouseDemora < 4 Then MouseDemora = MouseDemora + 1
If TLuzJ2 < 40 Then
    TLuzJ2 = TLuzJ2 + 1
Else
    BitBlt Jugador2.hDC, 0, 0, 8, 41, J2.hDC, 0, 0, vbSrcCopy
    Jugador2.Refresh
End If
If TLuzJ1 < 40 Then
    TLuzJ1 = TLuzJ1 + 1
Else
    BitBlt Jugador1.hDC, 0, 0, 8, 41, J1.hDC, 0, 0, vbSrcCopy
    Jugador1.Refresh
End If
End Sub

Private Sub Animar()
   Dim i As Integer
   Detener = False
   Do
      RestaurarFondo
      For i = 0 To MaxSprites - 1
         If Sprites(i).Visible Then
            MoverSprite i
            AvanzarAnimacion i
            MostrarSprite i
         End If
      Next i
      Demorar
      fondo.Refresh
      DoEvents
   Loop Until Detener
End Sub

Private Sub MoverBarra()
If juego.WindowState = 0 Then
    FormDrag Me
End If
End Sub

Private Sub barra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverBarra
End Sub

Private Sub barra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Desmarcar
End Sub

Private Sub BarraD_Click()
PcDifi = BarraD.Value
End Sub

Private Sub BContinuar_Click()
If Tec1J1.Value = True Then TeclasJ1 = True
If Tec2J1.Value = True Then FlechasJ1 = True
If MouJ1.Value = True Then MouseJ1 = True
If Tec1J2.Value = True Then TeclasJ2 = True
If Tec2J2.Value = True Then FlechasJ2 = True
If MouJ2.Value = True Then MouseJ2 = True
If Pcj2.Value = True Then VsPc = True
PcDifi = PcDifi + 1
BContinuar.Visible = False
Frame1.Visible = False
AngY = 35
fondo.SetFocus
Animar
End Sub

Private Sub ContinuarB_Click()
ContinuarB.Visible = False
fondo.SetFocus
Animar
End Sub

Private Sub EndB_Click()
End
End Sub

Private Sub fondo_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape: End
End Select
If ContinuarB.Visible = False Then
If FlechasJ1 = True Then
    Select Case KeyCode
        Case vbKeyUp: If Jugador1.Top - 7 > 0 Then Jugador1.Top = Jugador1.Top - 7
        Case vbKeyDown: If Jugador1.Top + 7 < 154 Then Jugador1.Top = Jugador1.Top + 7
    End Select
End If
If FlechasJ2 = True Then
    Select Case KeyCode
        Case vbKeyUp: If Jugador2.Top - 7 > 0 Then Jugador2.Top = Jugador2.Top - 7
        Case vbKeyDown: If Jugador2.Top + 7 < 154 Then Jugador2.Top = Jugador2.Top + 7
    End Select
End If
If TeclasJ1 = True Then
    Select Case KeyCode
        Case vbKeyW: If Jugador1.Top - 7 > 0 Then Jugador1.Top = Jugador1.Top - 7
        Case vbKeyS: If Jugador1.Top + 7 < 154 Then Jugador1.Top = Jugador1.Top + 7
    End Select
End If
If TeclasJ2 = True Then
    Select Case KeyCode
        Case vbKeyW: If Jugador2.Top - 7 > 0 Then Jugador2.Top = Jugador2.Top - 7
        Case vbKeyS: If Jugador2.Top + 7 < 154 Then Jugador2.Top = Jugador2.Top + 7
    End Select
End If
End If
End Sub

Private Sub fondo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MouseJ1 = True And MouseDemora = 4 Then
    If Y < Jugador1.Top Then Jugador1.Top = Jugador1.Top - 7
    If Y > Jugador1.Top + Jugador1.Height Then Jugador1.Top = Jugador1.Top + 7
    MouseDemora = 0
End If
If MouseJ2 = True And MouseDemora = 4 Then
    If Y < Jugador2.Top Then Jugador2.Top = Jugador2.Top - 7
    If Y > Jugador2.Top + Jugador2.Height Then Jugador2.Top = Jugador2.Top + 7
    MouseDemora = 0
End If
End Sub

Private Sub Form_Load()
Jugador1.Top = 78
Jugador2.Top = 78
AngY = AngY * 3
PVel = 1
If DirX = 0 Then DirX = 1 Else DirX = 0
fondo.Left = 1
fondo.Top = barra.Height + 17
Tamaño juego, (fondo.Width + 2) * 15, (fondo.Height + barra.Height + 18) * 15
MaxB_Click
juego.Caption = Nombre
juego.Icon = Icono
barra.ToolTipText = Nombre
Nombre.ToolTipText = Nombre
Icono.ToolTipText = Nombre
   
   Sprites(Pelota).Dim.X = 16
   Sprites(Pelota).Dim.Y = 16
   Sprites(Pelota).Pos.X = 142
   Sprites(Pelota).Pos.Y = 86
   Sprites(Pelota).TC = 1
   Sprites(Pelota).Vel = 1
   Sprites(Pelota).Visible = True

BitBlt fondo.hDC, 0, 0, 300, 195, SalvaFondo.hDC, 0, 0, vbSrcCopy
BitBlt Jugador1.hDC, 0, 0, 8, 41, J1.hDC, 0, 0, vbSrcCopy
BitBlt Jugador2.hDC, 0, 0, 8, 41, J2.hDC, 0, 0, vbSrcCopy
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Desmarcar
End Sub

Private Sub MaxB_Click()
MinB.Top = 0
EndB.Top = 0
barra.Left = 0
barra.Top = 0
Icono.Left = 3
Icono.Top = 3
Nombre.Left = 6 + Icono.Width
Nombre.Top = 3
Borde.Left = 0
Borde.Top = 0
Borde.Width = juego.Width / 15
Borde.Height = juego.Height / 15
Borde2.Y1 = barra.Height
Borde2.Y2 = barra.Height
Borde2.X1 = 0
Borde2.X2 = juego.Width / 15
barra.Width = juego.Width / 15
EndB.Left = (juego.Width / 15) - (EndB.Width + 4)
MinB.Left = (juego.Width / 15) - ((EndB.Width * 2) + (4 * 2))
Desmarcar
End Sub

Private Sub Label7_Click()
asddfg = MsgBox("Este programa fue hecho por Alvaro Vidueiro en el 2002." + vbCrLf + "Se distribuye gratiutamente. No Acepte copias." + vbCrLf + "Requisitos de sistem: Windows 98 en adelante." + vbCrLf + "La j en el nombre no es una falta de ortografía.", vbOKOnly, "Air-Jokey | Acerca de...")
End Sub

Private Sub minb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If MinB.ForeColor = &HFF0000 Then
    Desmarcar
    MinB.ForeColor = &HFF&
End If
End Sub

Private Sub endB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If EndB.ForeColor = &HFF0000 Then
    Desmarcar
    EndB.ForeColor = &HFF&
End If
End Sub

Private Sub Desmarcar()
MinB.ForeColor = &HFF0000
EndB.ForeColor = &HFF0000
End Sub

Private Sub MinB_Click()
juego.WindowState = 1
End Sub

Private Sub Icono_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverBarra
End Sub

Private Sub MouJ1_Click()
If MouJ2.Value = True Then Tec1J2.Value = True
End Sub

Private Sub MouJ2_Click()
If MouJ1.Value = True Then Tec1J1.Value = True
BarraD.Enabled = False
End Sub

Private Sub Nombre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverBarra
End Sub

Private Sub Pcj2_Click()
BarraD.Enabled = True
End Sub

Private Sub Tec1J1_Click()
If Tec2J2.Value = True Or Tec1J2.Value = True Then MouJ2.Value = True
End Sub

Private Sub Tec1J2_Click()
If Tec2J1.Value = True Or Tec1J1.Value = True Then MouJ1.Value = True
BarraD.Enabled = False
End Sub

Private Sub Tec2J1_Click()
If Tec2J2.Value = True Or Tec1J2.Value = True Then MouJ2.Value = True
End Sub

Private Sub Tec2J2_Click()
If Tec2J1.Value = True Or Tec1J1.Value = True Then MouJ1.Value = True
BarraD.Enabled = False
End Sub

Private Sub Timer1_Timer()
AngY = 35
Timer1.Enabled = False
Animar
End Sub
