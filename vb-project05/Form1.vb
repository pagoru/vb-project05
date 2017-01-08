Imports System.Threading

Public Class Form1

    'Retorna un string amb el nom del mes en
    'catala en funcio del numero de mes que
    'se li pasa com a parametre
    Private Function GetCatalanMonth(num As Integer)
        Select Case num
            Case 1
                Return "Gener"
            Case 2
                Return "Febrer"
            Case 3
                Return "Març"
            Case 4
                Return "Abril"
            Case 5
                Return "Maig"
            Case 6
                Return "Juny"
            Case 7
                Return "Juliol"
            Case 8
                Return "Agost"
            Case 9
                Return "Septembre"
            Case 10
                Return "Octubre"
            Case 11
                Return "Novembre"
            Case 12
                Return "Desembre"
        End Select
        Return "Undecember"
    End Function

    'Carrega tot el contingut necessari
    'quan el programa es carrega per
    'primera vegada
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SetLocalClock()
        SetLocalDate()

        CountryComboBox.SelectedIndex = 0

        Dim thread As New Thread(New ThreadStart(AddressOf ThreadClock))
        thread.Start()

    End Sub

    'Estableix la data local representada
    'en catala mitjançant la funcio
    'anterior que retorna el mes,
    'tambe representa numericament el dia
    'i lany.
    Private Sub SetLocalDate()
        Dim dDate As DateTime = Now()
        SetTextToLabel(CurrentDate, dDate.ToString("dd") + " de " + GetCatalanMonth(Integer.Parse(dDate.ToString("MM"))) + " del " + dDate.ToString("yyyy"))
    End Sub

    'Estableix lhora local del rellotge
    'mitjançant lhora actual que te
    'el sistema en el que s'executa
    Private Sub SetLocalClock()
        SetTextToLabel(ClockSeconds, TimeOfDay.ToString("ss"))
        SetTextToLabel(ClockMinutes, TimeOfDay.ToString("mm"))
        SetTextToLabel(ClockHours, TimeOfDay.ToString("HH"))
    End Sub

    'Fil que controla el rellotge
    'i tambe lalarma mitjançant les
    'funcions de control de rellotge
    'i destabliment de dades en els
    'labels corresponents
    Private Sub ThreadClock()
        While True
            Dim countryCode As String = GetTextFromComboBox(CountryComboBox)
            Dim canCountryClock As Boolean = countryCode.Equals("Cap")

            If CheckBoxAlarm.Checked And Integer.Parse(TimeOfDay.ToString("HH")) = GetTextFromNumericUpDown(NumericAlarmHours) And Integer.Parse(TimeOfDay.ToString("mm")) = GetTextFromNumericUpDown(NumericAlarmMinutes) And TimeOfDay.ToString("ss").Equals("00") Then
                SetTextToTextBox(StopAlarmTextBox, "Alarma activada!")
            End If

            If Not canCountryClock Then
                SetTextToLabel(ClockCountrySeconds, TimeOfDay.ToString("ss"))
            End If
            SetTextToLabel(ClockSeconds, TimeOfDay.ToString("ss"))
            If TimeOfDay.ToString("ss").Equals("00") Then

                If Not canCountryClock Then
                    SetTextToLabel(ClockCountryMinutes, TimeOfDay.ToString("mm"))
                End If
                SetTextToLabel(ClockMinutes, TimeOfDay.ToString("mm"))
                If TimeOfDay.ToString("mm").Equals("00") Then

                    If Not canCountryClock Then
                        SetTextToLabel(ClockCountryHours, TimeOfDay.AddHours(Integer.Parse(countryCode.Split(" ").GetValue(0))).ToString("HH"))
                    End If
                    SetTextToLabel(ClockMinutes, TimeOfDay.ToString("HH"))
                    If TimeOfDay.ToString("HH").Equals("00") Then
                        SetLocalDate()
                    End If
                End If
            End If

            Thread.Sleep(1000)
        End While
    End Sub

    'Funcio abstracta que sutilitza
    'per fer recursivitat per
    'obtenir un valor desde un
    'NumericUpDown
    Private Delegate Function GetTextFromNumericUpDownDelegate(ByVal NumericUpDown As NumericUpDown)

    'Funcio que sutilitza
    'per fer recursivitat per
    'obtenir un valor desde un
    'NumericUpDown
    Private Function GetTextFromNumericUpDown(ByVal NumericUpDown As NumericUpDown)
        If NumericUpDown.InvokeRequired Then
            Return NumericUpDown.Invoke(New GetTextFromNumericUpDownDelegate(AddressOf GetTextFromNumericUpDown), New Object() {NumericUpDown})
        End If
        Return NumericUpDown.Value
    End Function

    'Funcio abstracta que sutilitza
    'per fer recursivitat per
    'establir un valor per a un
    'TextBox
    Private Delegate Sub SetTextToTextBoxDelegate(ByVal TextBox As TextBox, ByVal txt As String)

    'Funcio que sutilitza
    'per fer recursivitat per
    'establir un valor per a un
    'TextBox
    Private Sub SetTextToTextBox(ByVal TextBox As TextBox, ByVal Txt As String)
        If TextBox.InvokeRequired Then
            TextBox.Invoke(New SetTextToTextBoxDelegate(AddressOf SetTextToTextBox), New Object() {TextBox, Txt})
        Else
            TextBox.Text = Txt
        End If
    End Sub

    'Funcio abstracta que sutilitza
    'per fer recursivitat per
    'obtenir un valor desde un
    'ComboBox
    Private Delegate Function GetTextFromComboBoxDelegate(ByVal ComboBox As ComboBox)

    'Funcio que sutilitza
    'per fer recursivitat per
    'obtenir un valor desde un
    'ComboBox
    Private Function GetTextFromComboBox(ByVal ComboBox As ComboBox)
        If ComboBox.InvokeRequired Then
            Return ComboBox.Invoke(New GetTextFromComboBoxDelegate(AddressOf GetTextFromComboBox), New Object() {ComboBox})
        End If
        Return ComboBox.SelectedItem.ToString()
    End Function

    'Funcio abstracta que sutilitza
    'per fer recursivitat per
    'establir un valor per a un
    'Label
    Private Delegate Sub SetTextToLabelDelegate(ByVal Label As Label, ByVal txt As String)

    'Funcio que sutilitza
    'per fer recursivitat per
    'establir un valor per a un
    'Label
    Private Sub SetTextToLabel(ByVal Label As Label, ByVal Txt As String)
        If Label.InvokeRequired Then
            Label.Invoke(New SetTextToLabelDelegate(AddressOf SetTextToLabel), New Object() {Label, Txt})
        Else
            Label.Text = Txt
        End If
    End Sub

    'Event que es dispara quan
    'es prem el botó per detenir
    'lalarma, aquesta la detindra
    'mitjançant una representacio
    'visual mitjançant text
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles StopAlarmButton.Click
        SetTextToTextBox(StopAlarmTextBox, "Alarma parada")
    End Sub

    'Event que es dispara quan
    'es canvia dopcio en el
    'combobox que representa
    'el rellotge internacional
    'si detecta cap, sestablira
    'una hora 00 00 00,
    'en cas de detectar correctament
    'el pais, restara o sumara al
    'rellotge actual i fara la
    'representacio
    Private Sub CountryComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CountryComboBox.SelectedIndexChanged
        Dim countryCode As String = CountryComboBox.SelectedItem.ToString()

        If countryCode.Equals("Cap") Then
            SetTextToLabel(ClockCountryHours, "00")
            SetTextToLabel(ClockCountryMinutes, "00")
            SetTextToLabel(ClockCountrySeconds, "00")
            Return
        End If
        SetInternationalClock()
    End Sub

    'Estableix en funcio de la
    'seleccio de pais, el rellotge
    'internacional de forma correcta
    Private Sub SetInternationalClock()
        Dim countryCode As String = GetTextFromComboBox(CountryComboBox)
        SetTextToLabel(ClockCountrySeconds, TimeOfDay.ToString("ss"))
        SetTextToLabel(ClockCountryMinutes, TimeOfDay.ToString("mm"))
        SetTextToLabel(ClockCountryHours, TimeOfDay.AddHours(Integer.Parse(countryCode.Split(" ").GetValue(0))).ToString("HH"))
    End Sub
End Class
