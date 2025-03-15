Imports System.Windows.Forms
Imports System.IO.Ports
Imports System.Globalization
Imports System.Drawing
Imports System.Text

Public Class TrayApp
    Inherits Form

    Private WithEvents notifyIcon As New NotifyIcon()
    Private lastPorts As List(Of String) = New List(Of String)()
    Private Shadows contextMenu As New ContextMenuStrip()
    Private WithEvents updateTimer As New Timer()
    Private ReadOnly iconFalse As Icon = My.Resources.icon_false
    Private ReadOnly iconTrue As Icon = My.Resources.icon_true
    Private selectedCulture As CultureInfo

    Public Sub New(ByVal culture As CultureInfo)
        selectedCulture = culture
        ' Инициализация иконки в трее
        notifyIcon.Icon = iconFalse
        notifyIcon.Visible = True

        ' Инициализация контекстного меню
        Dim exitMenuItem As New ToolStripMenuItem(GetLocalizedString("Exit"))
        AddHandler exitMenuItem.Click, AddressOf OnExit
        contextMenu.Items.Add(exitMenuItem)

        notifyIcon.ContextMenuStrip = contextMenu

        ' Инициализация таймера для обновления списка портов каждые 2 секунд
        updateTimer.Interval = 2000
        updateTimer.Start()

        ' Первоначальное обновление списка портов
        UpdatePorts()
    End Sub

    Private Sub UpdatePorts()
        Dim ports As List(Of String) = GetAvailablePorts()

        If ports.Count = 0 Then
            notifyIcon.Icon = iconFalse
            notifyIcon.Text = GetLocalizedString("NoPorts")
        Else
            notifyIcon.Icon = iconTrue
            notifyIcon.Text = GetPortsText(ports)
        End If

        ' Проверка на появление нового порта
        If ports.Count <> lastPorts.Count OrElse Not AreListsEqual(ports, lastPorts) Then
            If ports.Count > lastPorts.Count Then
                Dim newPort As String = FindNewPort(ports, lastPorts)
                If newPort IsNot Nothing Then
                    notifyIcon.ShowBalloonTip(1500, GetLocalizedString("NewPort"), String.Format(GetLocalizedString("NewPortDetected"), newPort), ToolTipIcon.Info)
                End If
            End If
            lastPorts = ports
        End If
    End Sub

    Private Function GetPortsText(ByVal ports As List(Of String)) As String
        Dim sb As New StringBuilder()
        For Each port As String In ports
            sb.AppendLine(port)
        Next
        Return sb.ToString()
    End Function

    Private Function AreListsEqual(ByVal list1 As List(Of String), ByVal list2 As List(Of String)) As Boolean
        If list1.Count <> list2.Count Then Return False
        For i As Integer = 0 To list1.Count - 1
            If list1(i) <> list2(i) Then Return False
        Next
        Return True
    End Function

    Private Function FindNewPort(ByVal currentPorts As List(Of String), ByVal previousPorts As List(Of String)) As String
        For Each port As String In currentPorts
            If Not previousPorts.Contains(port) Then Return port
        Next
        Return Nothing
    End Function

    Private Function GetAvailablePorts() As List(Of String)
        Dim portArray As String() = SerialPort.GetPortNames()
        Dim portList As New List(Of String)()

        ' Очищаем список перед повторным использованием
        portList.Clear()

        For Each port As String In portArray
            portList.Add(port)
        Next

        Return portList
    End Function

    Private Sub OnExit(ByVal sender As Object, ByVal e As EventArgs)
        notifyIcon.Visible = False
        updateTimer.Stop() ' Останавливаем таймер перед выходом
        Application.Exit()
    End Sub

    Private Function GetLocalizedString(ByVal key As String) As String
        Select Case selectedCulture.TwoLetterISOLanguageName
            Case "ru"
                Select Case key
                    Case "Exit" : Return "Выход"
                    Case "NoPorts" : Return "Нет доступных COM-портов"
                    Case "NewPort" : Return "Новый порт"
                    Case "NewPortDetected" : Return "Обнаружен: {0}"
                End Select
            Case "uk"
                Select Case key
                    Case "Exit" : Return "Вихід"
                    Case "NoPorts" : Return "Немає доступних COM-портів"
                    Case "NewPort" : Return "Новий порт"
                    Case "NewPortDetected" : Return "Виявлений: {0}"
                End Select
            Case Else
                Select Case key
                    Case "Exit" : Return "Exit"
                    Case "NoPorts" : Return "No available COM ports"
                    Case "NewPort" : Return "New port"
                    Case "NewPortDetected" : Return "Detected: {0}"
                End Select
        End Select
        Return key
    End Function

    Protected Overrides Sub OnLoad(ByVal e As EventArgs)
        Visible = False
        ShowInTaskbar = False
        MyBase.OnLoad(e)
    End Sub

    <STAThread()> _
    Public Shared Sub Main()
        Dim culture As CultureInfo = CultureInfo.InstalledUICulture

        ' Parse command-line arguments
        Dim args As String() = Environment.GetCommandLineArgs()
        If args.Length > 1 Then
            Select Case args(1).ToLower()
                Case "-ru"
                    culture = New CultureInfo("ru-RU")
                Case "-ua"
                    culture = New CultureInfo("uk-UA")
                Case "-en"
                    culture = New CultureInfo("en-US")
            End Select
        End If

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New TrayApp(culture))
    End Sub

    ' Обработка тика таймера
    Private Sub UpdateTimer_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles updateTimer.Tick
        UpdatePorts()
    End Sub
End Class
