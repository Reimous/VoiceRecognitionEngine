Imports System.Speech.Recognition
Imports System.Globalization
Imports System.Threading

Public Class Form1
#Region "SpeechEvent"
    Public Delegate Sub SpeechEventHandler(sender As Object, recognized As String)
    'Public Event SpeechEvent As SpeechEventHandler
    Public Event SpeechRecognized As  _
     EventHandler(Of SpeechRecognizedEventArgs)
    Public Event SpeechRecognitionRejected As  _
     EventHandler(Of SpeechRecognitionRejectedEventArgs)
    'Mouse event
    Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Integer, _
      ByVal dx As Integer, ByVal dy As Integer, ByVal cButtons As Integer, _
      ByVal dwExtraInfo As Integer)


    Public Const MOUSEEVENTF_LEFTDOWN As Integer = &H2
    Public Const MOUSEEVENTF_LEFTUP As Integer = &H4
    Public Const MOUSEEVENTF_MIDDLEDOWN As Integer = &H20
    Public Const MOUSEEVENTF_MIDDLEUP As Integer = &H40
    Public Const MOUSEEVENTF_RIGHTDOWN As Integer = &H8
    Public Const MOUSEEVENTF_RIGHTUP As Integer = &H10
    Public Const MOUSEEVENTF_MOVE As Integer = &H1
    Public Sub recevent(ByVal sender As System.Object, _
      ByVal e As RecognitionEventArgs)
        Dim result As String = e.Result.Text
        If e.Result.Confidence >= RequiredConfidence Then
            Select Case result
                Case "Fuego"
                    Dim caca As Integer = Keys.M
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                Case "Fuego a discreción"
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                Case "Fuego continuo"
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
                Case "Alto el fuego"
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                Case "Izquierda"
                    mouse_event(MOUSEEVENTF_MOVE, -100, 0, 0, 0)
                Case "Derecha"
                    mouse_event(MOUSEEVENTF_MOVE, 100, 0, 0, 0)
                Case "Arriba"
                    mouse_event(MOUSEEVENTF_MOVE, 0, -100, 0, 0)
                Case "Abajo"
                    mouse_event(MOUSEEVENTF_MOVE, 0, 100, 0, 0)
                Case "Mapa"
                    SendKeys.Send("{M}")
                Case "Cámara"
                    SendKeys.Send("{C}")
                Case "Gracias"
                    SendKeys.Send("{F4}")
                Case "Afirmativo"
                    SendKeys.Send("{F5}")
                Case "Negativo"
                    SendKeys.Send("{F6}")
                Case "Solicito apoyo"
                    SendKeys.Send("{F7}")
                Case "Socorro"
                    SendKeys.Send("{F8}")
                Case "Que gran batalla"
                    SendKeys.Send("{F9}")
                Case "Buena suerte a todos"
                    SendKeys.Send("{F10}")
                Case "Me cago en la puta"
                    SendKeys.Send("{F11}")
            End Select
        End If
    End Sub
    ' recognition failed event
    Public Sub recfailevent(ByVal sender As System.Object, _
      ByVal e As RecognitionEventArgs)
        'Código cuando ha sido rechazado una speech
    End Sub
#End Region
    Public RequiredConfidence As Single = 0.7F '0.9F
    Public commands As String() = New String() {"Fuego", "Fuego a discreción", "Mapa", "Cámara", "Izquierda", "Derecha", "Arriba", "Abajo", "Fuego continuo", "Gracias", "Afirmativo", "Negativo", "Solicito apoyo", _
                                               "Socorro", "Que gran batalla", "Buena suerte a todos", "Me cago en la puta"}
    Public FormNumber As Integer = 5
    Private rec As SpeechRecognitionEngine

    Public Sub StartListening()
        'AddHandler rec.SpeechRecognized, AddressOf rec_SpeechRecognized
        AddHandler rec.SpeechRecognized, AddressOf Me.recevent
        AddHandler rec.SpeechRecognitionRejected, AddressOf Me.recfailevent
        rec.RecognizeAsync(RecognizeMode.Multiple)
    End Sub
    Public Sub StopListening()
        rec.RecognizeAsyncStop()
        RemoveHandler rec.SpeechRecognized, AddressOf Me.recevent
        RemoveHandler rec.SpeechRecognitionRejected, AddressOf Me.recfailevent
        'RemoveHandler rec.SpeechRecognized, AddressOf rec_SpeechRecognized
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Thread.CurrentThread.CurrentCulture = New CultureInfo("es-ES")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("es-ES")

        rec = New SpeechRecognitionEngine()
        rec.SetInputToDefaultAudioDevice()
        Dim c As New Choices()
        For i As Integer = 0 To commands.Length - 1
            c.Add(commands(i))
        Next
        Dim gb As New GrammarBuilder(c)
        Dim g As New Grammar(gb)
        rec.LoadGrammar(g)

        Me.WindowState = FormWindowState.Minimized
        StartListening()
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        StopListening()
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            NotifyIcon1.Visible = True
            NotifyIcon1.Icon = SystemIcons.Application
            NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
            NotifyIcon1.BalloonTipTitle = "Reconocimiento de voz"
            NotifyIcon1.BalloonTipText = "Reconocimiento de voz"
            NotifyIcon1.ShowBalloonTip(50000)
            'Me.Hide()
            ShowInTaskbar = False
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        ShowInTaskbar = True
        Me.WindowState = FormWindowState.Normal
    End Sub
End Class
