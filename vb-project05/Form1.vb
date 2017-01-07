Imports System.Threading

Public Class Form1
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
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        SetLocalClock()
        SetLocalDate()

        CountryComboBox.SelectedIndex = 0

        Dim thread As New Thread(New ThreadStart(AddressOf ThreadClock))
        thread.Start()

    End Sub

    Private Sub SetLocalDate()
        Dim dDate As DateTime = Now()
        SetTextToLabel(CurrentDate, dDate.ToString("dd") + " de " + GetCatalanMonth(Integer.Parse(dDate.ToString("MM"))) + " del " + dDate.ToString("yyyy"))
    End Sub

    Private Sub SetLocalClock()
        SetTextToLabel(ClockSeconds, TimeOfDay.ToString("ss"))
        SetTextToLabel(ClockMinutes, TimeOfDay.ToString("mm"))
        SetTextToLabel(ClockHours, TimeOfDay.ToString("HH"))
    End Sub

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

    Private Delegate Function GetTextFromNumericUpDownDelegate(ByVal NumericUpDown As NumericUpDown)
    Private Function GetTextFromNumericUpDown(ByVal NumericUpDown As NumericUpDown)
        If NumericUpDown.InvokeRequired Then
            Return NumericUpDown.Invoke(New GetTextFromNumericUpDownDelegate(AddressOf GetTextFromNumericUpDown), New Object() {NumericUpDown})
        End If
        Return NumericUpDown.Value
    End Function

    Private Delegate Sub SetTextToTextBoxDelegate(ByVal TextBox As TextBox, ByVal txt As String)
    Private Sub SetTextToTextBox(ByVal TextBox As TextBox, ByVal Txt As String)
        If TextBox.InvokeRequired Then
            TextBox.Invoke(New SetTextToTextBoxDelegate(AddressOf SetTextToTextBox), New Object() {TextBox, Txt})
        Else
            TextBox.Text = Txt
        End If
    End Sub

    Private Delegate Function GetTextFromComboBoxDelegate(ByVal ComboBox As ComboBox)
    Private Function GetTextFromComboBox(ByVal ComboBox As ComboBox)
        If ComboBox.InvokeRequired Then
            Return ComboBox.Invoke(New GetTextFromComboBoxDelegate(AddressOf GetTextFromComboBox), New Object() {ComboBox})
        End If
        Return ComboBox.SelectedItem.ToString()
    End Function

    Private Delegate Sub SetTextToLabelDelegate(ByVal Label As Label, ByVal txt As String)
    Private Sub SetTextToLabel(ByVal Label As Label, ByVal Txt As String)
        If Label.InvokeRequired Then
            Label.Invoke(New SetTextToLabelDelegate(AddressOf SetTextToLabel), New Object() {Label, Txt})
        Else
            Label.Text = Txt
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles StopAlarmButton.Click
        SetTextToTextBox(StopAlarmTextBox, "Alarma parada")
    End Sub

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

    Private Sub SetInternationalClock()
        Dim countryCode As String = GetTextFromComboBox(CountryComboBox)
        SetTextToLabel(ClockCountrySeconds, TimeOfDay.ToString("ss"))
        SetTextToLabel(ClockCountryMinutes, TimeOfDay.ToString("mm"))
        SetTextToLabel(ClockCountryHours, TimeOfDay.AddHours(Integer.Parse(countryCode.Split(" ").GetValue(0))).ToString("HH"))
    End Sub
End Class
