<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.ClockHours = New System.Windows.Forms.Label()
        Me.ClockMinutes = New System.Windows.Forms.Label()
        Me.ClockSeconds = New System.Windows.Forms.Label()
        Me.ClockGroup = New System.Windows.Forms.GroupBox()
        Me.DateGroup = New System.Windows.Forms.GroupBox()
        Me.CurrentDate = New System.Windows.Forms.Label()
        Me.ClockCountry = New System.Windows.Forms.GroupBox()
        Me.ClockCountrySeconds = New System.Windows.Forms.Label()
        Me.CountryComboBox = New System.Windows.Forms.ComboBox()
        Me.ClockCountryHours = New System.Windows.Forms.Label()
        Me.ClockCountryMinutes = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.StopAlarmTextBox = New System.Windows.Forms.TextBox()
        Me.StopAlarmButton = New System.Windows.Forms.Button()
        Me.CheckBoxAlarm = New System.Windows.Forms.CheckBox()
        Me.NumericAlarmMinutes = New System.Windows.Forms.NumericUpDown()
        Me.NumericAlarmHours = New System.Windows.Forms.NumericUpDown()
        Me.ClockGroup.SuspendLayout()
        Me.DateGroup.SuspendLayout()
        Me.ClockCountry.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericAlarmMinutes, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericAlarmHours, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ClockHours
        '
        Me.ClockHours.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClockHours.Location = New System.Drawing.Point(6, 16)
        Me.ClockHours.Name = "ClockHours"
        Me.ClockHours.Size = New System.Drawing.Size(30, 23)
        Me.ClockHours.TabIndex = 0
        Me.ClockHours.Text = "00"
        Me.ClockHours.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ClockMinutes
        '
        Me.ClockMinutes.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClockMinutes.Location = New System.Drawing.Point(36, 16)
        Me.ClockMinutes.Name = "ClockMinutes"
        Me.ClockMinutes.Size = New System.Drawing.Size(30, 23)
        Me.ClockMinutes.TabIndex = 1
        Me.ClockMinutes.Text = "00"
        Me.ClockMinutes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ClockSeconds
        '
        Me.ClockSeconds.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClockSeconds.Location = New System.Drawing.Point(66, 16)
        Me.ClockSeconds.Name = "ClockSeconds"
        Me.ClockSeconds.Size = New System.Drawing.Size(30, 23)
        Me.ClockSeconds.TabIndex = 2
        Me.ClockSeconds.Text = "00"
        Me.ClockSeconds.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ClockGroup
        '
        Me.ClockGroup.Controls.Add(Me.ClockSeconds)
        Me.ClockGroup.Controls.Add(Me.ClockHours)
        Me.ClockGroup.Controls.Add(Me.ClockMinutes)
        Me.ClockGroup.Location = New System.Drawing.Point(12, 12)
        Me.ClockGroup.Name = "ClockGroup"
        Me.ClockGroup.Size = New System.Drawing.Size(101, 48)
        Me.ClockGroup.TabIndex = 3
        Me.ClockGroup.TabStop = False
        Me.ClockGroup.Text = "Hora local"
        '
        'DateGroup
        '
        Me.DateGroup.Controls.Add(Me.CurrentDate)
        Me.DateGroup.Location = New System.Drawing.Point(119, 12)
        Me.DateGroup.Name = "DateGroup"
        Me.DateGroup.Size = New System.Drawing.Size(230, 48)
        Me.DateGroup.TabIndex = 4
        Me.DateGroup.TabStop = False
        Me.DateGroup.Text = "Data local"
        '
        'CurrentDate
        '
        Me.CurrentDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CurrentDate.Location = New System.Drawing.Point(6, 16)
        Me.CurrentDate.Name = "CurrentDate"
        Me.CurrentDate.Size = New System.Drawing.Size(217, 23)
        Me.CurrentDate.TabIndex = 0
        Me.CurrentDate.Text = "00 de Septembre de 0000"
        Me.CurrentDate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ClockCountry
        '
        Me.ClockCountry.Controls.Add(Me.ClockCountrySeconds)
        Me.ClockCountry.Controls.Add(Me.CountryComboBox)
        Me.ClockCountry.Controls.Add(Me.ClockCountryHours)
        Me.ClockCountry.Controls.Add(Me.ClockCountryMinutes)
        Me.ClockCountry.Location = New System.Drawing.Point(12, 80)
        Me.ClockCountry.Name = "ClockCountry"
        Me.ClockCountry.Size = New System.Drawing.Size(337, 48)
        Me.ClockCountry.TabIndex = 4
        Me.ClockCountry.TabStop = False
        Me.ClockCountry.Text = "Hora internacional"
        '
        'ClockCountrySeconds
        '
        Me.ClockCountrySeconds.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClockCountrySeconds.Location = New System.Drawing.Point(66, 16)
        Me.ClockCountrySeconds.Name = "ClockCountrySeconds"
        Me.ClockCountrySeconds.Size = New System.Drawing.Size(30, 23)
        Me.ClockCountrySeconds.TabIndex = 2
        Me.ClockCountrySeconds.Text = "00"
        Me.ClockCountrySeconds.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CountryComboBox
        '
        Me.CountryComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.CountryComboBox.FormattingEnabled = True
        Me.CountryComboBox.Items.AddRange(New Object() {"Cap", "-1 hora | Regne Unit", "+1 hora | Finlandia", "+7 hores | China", "-10 hores | Alaska"})
        Me.CountryComboBox.Location = New System.Drawing.Point(117, 16)
        Me.CountryComboBox.Name = "CountryComboBox"
        Me.CountryComboBox.Size = New System.Drawing.Size(204, 21)
        Me.CountryComboBox.TabIndex = 5
        '
        'ClockCountryHours
        '
        Me.ClockCountryHours.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClockCountryHours.Location = New System.Drawing.Point(6, 16)
        Me.ClockCountryHours.Name = "ClockCountryHours"
        Me.ClockCountryHours.Size = New System.Drawing.Size(30, 23)
        Me.ClockCountryHours.TabIndex = 0
        Me.ClockCountryHours.Text = "00"
        Me.ClockCountryHours.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ClockCountryMinutes
        '
        Me.ClockCountryMinutes.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ClockCountryMinutes.Location = New System.Drawing.Point(36, 16)
        Me.ClockCountryMinutes.Name = "ClockCountryMinutes"
        Me.ClockCountryMinutes.Size = New System.Drawing.Size(30, 23)
        Me.ClockCountryMinutes.TabIndex = 1
        Me.ClockCountryMinutes.Text = "00"
        Me.ClockCountryMinutes.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.StopAlarmTextBox)
        Me.GroupBox1.Controls.Add(Me.StopAlarmButton)
        Me.GroupBox1.Controls.Add(Me.CheckBoxAlarm)
        Me.GroupBox1.Controls.Add(Me.NumericAlarmMinutes)
        Me.GroupBox1.Controls.Add(Me.NumericAlarmHours)
        Me.GroupBox1.Location = New System.Drawing.Point(373, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(174, 116)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Alarma local"
        '
        'StopAlarmTextBox
        '
        Me.StopAlarmTextBox.Location = New System.Drawing.Point(8, 58)
        Me.StopAlarmTextBox.Name = "StopAlarmTextBox"
        Me.StopAlarmTextBox.ReadOnly = True
        Me.StopAlarmTextBox.Size = New System.Drawing.Size(155, 20)
        Me.StopAlarmTextBox.TabIndex = 4
        Me.StopAlarmTextBox.Text = "Alarma parada"
        Me.StopAlarmTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'StopAlarmButton
        '
        Me.StopAlarmButton.Location = New System.Drawing.Point(8, 84)
        Me.StopAlarmButton.Name = "StopAlarmButton"
        Me.StopAlarmButton.Size = New System.Drawing.Size(155, 23)
        Me.StopAlarmButton.TabIndex = 3
        Me.StopAlarmButton.Text = "Parar alarma"
        Me.StopAlarmButton.UseVisualStyleBackColor = True
        '
        'CheckBoxAlarm
        '
        Me.CheckBoxAlarm.AutoSize = True
        Me.CheckBoxAlarm.Location = New System.Drawing.Point(8, 18)
        Me.CheckBoxAlarm.Name = "CheckBoxAlarm"
        Me.CheckBoxAlarm.Size = New System.Drawing.Size(15, 14)
        Me.CheckBoxAlarm.TabIndex = 2
        Me.CheckBoxAlarm.UseVisualStyleBackColor = True
        '
        'NumericAlarmMinutes
        '
        Me.NumericAlarmMinutes.Location = New System.Drawing.Point(98, 16)
        Me.NumericAlarmMinutes.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        Me.NumericAlarmMinutes.Name = "NumericAlarmMinutes"
        Me.NumericAlarmMinutes.Size = New System.Drawing.Size(65, 20)
        Me.NumericAlarmMinutes.TabIndex = 1
        Me.NumericAlarmMinutes.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'NumericAlarmHours
        '
        Me.NumericAlarmHours.Location = New System.Drawing.Point(27, 16)
        Me.NumericAlarmHours.Maximum = New Decimal(New Integer() {23, 0, 0, 0})
        Me.NumericAlarmHours.Name = "NumericAlarmHours"
        Me.NumericAlarmHours.Size = New System.Drawing.Size(65, 20)
        Me.NumericAlarmHours.TabIndex = 0
        Me.NumericAlarmHours.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(561, 140)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.ClockCountry)
        Me.Controls.Add(Me.DateGroup)
        Me.Controls.Add(Me.ClockGroup)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "Form1"
        Me.ShowIcon = False
        Me.Text = "Pablo González Rubio"
        Me.ClockGroup.ResumeLayout(False)
        Me.DateGroup.ResumeLayout(False)
        Me.ClockCountry.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.NumericAlarmMinutes, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericAlarmHours, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ClockHours As Label
    Friend WithEvents ClockMinutes As Label
    Friend WithEvents ClockSeconds As Label
    Friend WithEvents ClockGroup As GroupBox
    Friend WithEvents DateGroup As GroupBox
    Friend WithEvents CurrentDate As Label
    Friend WithEvents ClockCountry As GroupBox
    Friend WithEvents ClockCountrySeconds As Label
    Friend WithEvents ClockCountryHours As Label
    Friend WithEvents ClockCountryMinutes As Label
    Friend WithEvents CountryComboBox As ComboBox
    Friend WithEvents GroupBox1 As GroupBox
    Friend WithEvents NumericAlarmMinutes As NumericUpDown
    Friend WithEvents NumericAlarmHours As NumericUpDown
    Friend WithEvents CheckBoxAlarm As CheckBox
    Friend WithEvents StopAlarmTextBox As TextBox
    Friend WithEvents StopAlarmButton As Button
End Class
