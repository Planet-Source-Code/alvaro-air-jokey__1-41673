VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   ScaleHeight     =   271
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   StartUpPosition =   1  'CenterOwner
   Begin VB.Line Borde2 
      BorderColor     =   &H00FF0000&
      X1              =   237
      X2              =   146
      Y1              =   27
      Y2              =   32
   End
   Begin VB.Shape Borde 
      BorderColor     =   &H00FF0000&
      Height          =   585
      Left            =   2505
      Top             =   30
      Width           =   930
   End
   Begin VB.Image Icono 
      Height          =   240
      Left            =   150
      Picture         =   "menu.frx":0000
      Stretch         =   -1  'True
      Top             =   90
      Width           =   240
   End
   Begin VB.Label Nombre 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AV Software"
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
      Top             =   90
      Width           =   1170
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
      Top             =   90
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
      Top             =   90
      Width           =   240
   End
   Begin VB.Image barra 
      Height          =   285
      Left            =   150
      Picture         =   "menu.frx":0442
      Stretch         =   -1  'True
      Top             =   75
      Width           =   4500
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MoverBarra()
If Menu.WindowState = 0 Then
    FormDrag Me
End If
End Sub

Private Sub barra_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverBarra
End Sub

Private Sub barra_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Desmarcar
End Sub

Private Sub EndB_Click()
End
End Sub

Private Sub Form_Load()
MaxB_Click
Menu.Caption = Nombre
Menu.Icon = Icono
barra.ToolTipText = Nombre
Nombre.ToolTipText = Nombre
Icono.ToolTipText = Nombre
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
End With
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
Menu.WindowState = 1
End Sub

Private Sub Icono_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverBarra
End Sub

Private Sub Nombre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoverBarra
End Sub
