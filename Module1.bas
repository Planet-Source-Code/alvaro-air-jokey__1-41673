Attribute VB_Name = "Module1"
Option Explicit

Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Global Const SND_ASYNC = 0
Global Const SND_NODEFAULT = 1

Type T_XY
   X As Integer
   Y As Integer
End Type

Type T_Sprite
   Dim As T_XY             'Dimensiones del sprite
   Pos As T_XY             'Posicion actual en la pantalla
   Inc As T_XY             'Incremento o desplazamiento cuando se mueve
   Vel As Integer          'Velocidad 0=Quieto 1=Rapido >2 Lento
   Pas As Integer          'Pasada (para controlar la velocidad)
   C   As Integer          'Cuadro actual de la animación de un sprite
   TC  As Integer          'Total de cuadros de la animación
   Visible As Boolean      '¿Está visible ahora?
End Type

Sub Tamaño(Source As Object, X As Long, Y As Long)
Source.Width = X
Source.Height = Y
End Sub

Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub



