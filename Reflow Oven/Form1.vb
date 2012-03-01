' Written by Ben Barenz Janary 2012

'Commands List to Oven
''A’ = Top Heater ON 
''a’ =  Top Heater OFF

''B’ = Bottom Heater ON 
''b’ =  Bottom Heater OFF

''C’ = Convection Fan ON 
''c’ =  Convection Fan OFF

''D’ = Exhaust Fan ON 
''d’ =  Exhaust Fan OFF

''T’ = Send Oven Temp (returns raw temp. sensor data in hex)
''t’ = Send Case Temp  ( returns current case temperature in hex)

''*’ = Keep alive (must be sent in less than 5 second intervals)

''s’ = Send Serial Number
''f’ = Send Firmware
''U’ = Set unconnected Mode
''N’ = Set connected Mode

'Commands from Oven 
''L’ = Lost connection indicator
'======================================================================
' Imports 
Imports System
Imports System.IO
Imports System.Xml
Imports System.IO.Ports
Imports System.Threading
Imports GemBox.Spreadsheet
'Imports MSChart20Lib
Imports System.Windows.Forms.DataVisualization.Charting
'======================================================================

Public Class Form1
    '======================================================================
    '======================================================================
    ' Declare Strings
    Private ProgramLibaryPath As String
    Private RawOvenTemp1 As String
    Private RawCaseTemp1 As String
    Private SerialNumber As String
    Private FirmwareVer As String
    Private Mode As String
    Private EditorStatus As String = "STOPPED"
    '======================================================================
    ' Declare Arrays

    '======================================================================
    ' Declare Booleans
    Private TempUnit As Boolean ' 0 = F, 1 = C
    Private ComFailure As Boolean
    Private ComFailNotified As Boolean
    Private TempAvgInit As Boolean
    Private shuttingdown As Boolean
    Private OKtoUpdate As Boolean
    Private RunUpdate As Boolean
    Private OverHeat As Boolean
    Private StopProfileEditor As Boolean
    Private ProfileEditorRunning As Boolean
    Private ProfileDesignerRunning As Boolean
    Private GraphCreatedby As Boolean ' false = Program editor, true = program designer
    Private formIsclosed As Boolean
    Private LblUpdating As Boolean
    Private LblWaitingUpdate As Boolean
    Private CheckOvenPower As Boolean
    Private Debug As Boolean
    '======================================================================
    ' Declare Integers
    Private Oven_Temp As Integer = 100
    Private Case_Temp As Integer = 70
    Private OTAvg As Integer

    Private OvenTempLow As Integer = 200 ' 93 C
    Private OvenTempmed As Integer = 450 ' 232 C
    Private OvenTempHigh As Integer = 600 ' 315 C
    Private OvenTempCritical As Integer = 700 ' 371 C

    Private CaseTempLow As Integer = 90 ' 32 C
    Private CaseTempmed As Integer = 110 ' 43 C
    Private CaseTempHigh As Integer = 130 ' 54 C
    Private CaseTempCritical As Integer = 150 ' 65 C
    Private MinOvenTemp As Integer = 100 ' 38

    Private OTArrCnt As Integer
    Private CTArrCnt As Integer

    Private AddScalar As Integer = 0
    Private MulitplyScalar As Integer = 1

    Private MaxOvenTemp As Integer = 600
    Private MaxCaseTemp As Integer = 150
    Private AvgCoefficient As Integer

    Private Elapsed_Hours As Integer = 0
    Private Elapsed_Min As Integer = 0
    Private Elapsed_Sec As Integer = 0
    Private Elapsed_inSec As Integer = 0
    Private Elapsed_RunSec As Integer = 0

    Private RecordedMaxTemp As Integer
    Private DataPointsCnt As Integer
    Private UpdateListsCnt As Integer

    Private HourMeter As Integer

    Private MinOvenTime As Integer = 1
    Private MaxOvenTime As Integer = 1200 ' 20min

    Private RXtimeout As Integer = 50
    Private ExceptionCnt As Integer
    '======================================================================
    ' Declare Integers Arrays
    Private OTArr(99) As Integer
    Private CTArr(99) As Integer
    '======================================================================
    ' Declare TriState
    Private TopHeaterStatus As TriState = TriState.UseDefault
    Private BottomHeaterStatus As TriState = TriState.UseDefault
    Private ConvFanStatus As TriState = TriState.UseDefault
    Private ExhFanStatus As TriState = TriState.UseDefault

    '======================================================================
    ' Declare custome colors 
    Dim offcolor As Color = Color.FromArgb(255, 128, 128) ' light red
    Dim oncolor As Color = Color.FromArgb(192, 255, 192) ' light green
    '======================================================================
    ' Declare others
    Private Elapsedtime As New System.Diagnostics.Stopwatch
    Private RunTime As New System.Diagnostics.Stopwatch
    Public T1 As System.Threading.Thread
    Public T2 As System.Threading.Thread
    Dim s1 As New Series
    ' End Variable declarations
    '======================================================================
    '======================================================================

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        '------------------------------------------------------------------------------------------------------------
        CheckForExistingInstance()
        '------------------------------------------------------------------------------------------------------------
        UpdatePorts() ' update serial ports
        '------------------------------------------------------------------------------------------------------------
        ' Deterime to use Celsius or Fahrenheit units
        If chk_UnitSelect.Checked = True Then
            rbtn_Celsius.Checked = True ' start using Celsius
            rbtn_Fahrenheit.Checked = False
            ChangeTempUnit("C")
        End If
        '------------------------------------------------------------------------------------------------------------
        ' setup the numerical up down selector for the averaging coefficient 
        If nud_AvgCoefficient.Value > 1 And nud_AvgCoefficient.Value < 100 Then AvgCoefficient = nud_AvgCoefficient.Value
        '------------------------------------------------------------------------------------------------------------
        If chk_AutoConnect.Checked = True Then
            SerialPortConnect()
            SerialPort1Data = "" ' used to pass data to the thread in charge of comm with controller
        End If
        '------------------------------------------------------------------------------------------------------------
        If chk_Topmost.Checked = True Then
            Me.TopMost = True
        End If
        '------------------------------------------------------------------------------------------------------------
        ' figure out if we have an (x86) or normal Program Files directory
        If Directory.Exists("C:\Program Files (x86)\") = True Then
            ProgramLibaryPath = "C:\Program Files (x86)\Reflow Oven\" ' Win 7 
        Else
            ProgramLibaryPath = "C:\Program Files\Reflow Oven\" ' Win XP 
        End If
        '------------------------------------------------------------------------------------------------------------
        TabPage2.Dispose() ' hide the profile designer tab 
        '------------------------------------------------------------------------------------------------------------
        ' give the graph a line  
        ' these values are added to the graph only and not the textboxs or saved file
        Try
            ' update the graph data 
            Chart1.Series.Clear()
            '  s1.Name = "Temp"
            s1.ChartType = SeriesChartType.Line
            s1.Points.AddXY(10, 50)
            Chart1.Series.Add(s1)

            s1.Points.AddXY(20, 100)
            Chart1.Series.Add(s1)
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

        '------------------------------------------------------------------------------------------------------------
        ' load the max temps 
        Try
            MaxCaseTemp = CInt(txt_MaxCaseTemp.Text)
            MaxOvenTemp = CInt(txt_MaxOvenTemp.Text)
        Catch ex As Exception

            txt_MaxCaseTemp.Text = "600"
            txt_MaxOvenTemp.Text = "150"
        End Try
        '------------------------------------------------------------------------------------------------------------
        '------------------------------------------------------------------------------------------------------------
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ' closing is handled by closeapp
        If formIsclosed = False Then
            CloseApp()
        End If

    End Sub

    Private Function TempConvert(ByVal temp As Integer, ByVal conv As Char)
        ' function takes temp value as integer and a char specifying C or F 
        Dim ConvertedTemp As Integer = 0

        Select Case conv
            ' figure out what it is to be converted to 
            Case "C", "c" ' convert to Celcius
                ConvertedTemp = (5 / 9) * (temp - 32)
            Case "F", "f" ' convert to fahrenheit
                ConvertedTemp = ((9 / 5) * temp) + 32
        End Select

        Return ConvertedTemp
    End Function

    Private Sub ExitToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem1.Click

        PortWrite("N")

    End Sub

    Private Sub StandbyToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles StandbyToolStripMenuItem.Click

        PortWrite("U") ' put in standby mode

    End Sub

    Private Sub ExitToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem2.Click

        If formIsclosed = False Then
            CloseApp()
        End If

    End Sub

    Private Sub AToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AToolStripMenuItem.Click

        MsgBox("Reflow Oven Console" & vbNewLine & "Version 1.20" & vbNewLine & "February 2012", MsgBoxStyle.Information)

    End Sub

    Private Sub CloseApp()


        If SerialPort1.IsOpen = True Then
            PortWrite("U") ' put oven in unconnected mode 
            delay(100)
            SerialPort1.Close()
        End If



        ' close editor if running 
        '       If BackgroundWorker1.IsBusy = True Then
        BackgroundWorker1.CancelAsync()
        '  End If

        tmr_Main.Enabled = False
        RunUpdate = False ' this should stop thread T1
        OKtoUpdate = False
        formIsclosed = True
        Me.Dispose()
        Application.Exit()

        '   Me.Close()


    End Sub

    Public Sub CheckForExistingInstance()
        'Get number of processes of you program
        Try
            If Process.GetProcessesByName _
              (Process.GetCurrentProcess.ProcessName).Length > 1 Then

                MessageBox.Show _
                 ("Another Instance of Reflow Oven is already running", _
                     "Multiple Instances Forbidden", _
                      MessageBoxButtons.OK, _
                     MessageBoxIcon.Exclamation)
                Application.Exit()
                Me.Close()
            End If
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If

        End Try
    End Sub

    Private Sub delay(ByVal delaytime)
        ' delay in ms 
        System.Threading.Thread.Sleep(delaytime)

    End Sub

#Region "Serial Port 1 Code Block"
    ' This region has all code to create a serial port instance. It also has form 1 controls. The acompanying group box must be copied as well 
    ' Serial Port 1 
    '----------------------------------------------------------------------
    ' Strings
    Private RxBuff As String
    '-----------------------------------
    ' String Arrays
    Private ports() As String
    Private Selected As String
    Private NewPorts() As String
    Private CurrentPorts() As String
    Private SerialPort1Data As String
    '-----------------------------------
    ' Integers 
    Private Highest As Integer
    Private DataTimeout As Integer
    Private FailToConnectCnt As Integer
    '-----------------------------------
    ' Booleans 
    Private NewData As Boolean
    Private ComTimeOut As Boolean
    Private serialport1_fail As Boolean
    '-----------------------------------
    ' Threads
    Public Thread_Serial_Port As System.Threading.Thread
    ' End Serial Port Declarations 
    '----------------------------------------------------------------------

    Private Sub UpdatePorts()

        ' update list box 
        lst_ComPorts.Items.Clear()

        ' populate list box with portnames 
        ports = SerialPort.GetPortNames
        Dim size As Integer
        size = SerialPort.GetPortNames.GetLength(size)

        '------------------------------------------------------------------------
        'Auto set to the hightest port number when no port name is saved 
        ' if a portname has been saved, connect to that 
        If lbl_SelectedPort.Text = Nothing Then   ' only when the box is empty
            SerialPort1.PortName = ports(size - 1) ' this always makes the highest serial port number the curren one 
            lbl_SelectedPort.Text = SerialPort1.PortName
        Else
            Try
                SerialPort1.PortName = lbl_SelectedPort.Text
            Catch ex As Exception
                If Debug = True Then
                    ExceptionCnt = ExceptionCnt + 1
                    MsgBox(ex.ToString)
                End If
            End Try

        End If
        '------------------------------------------------------------------------------------------
        Dim cnt As Integer
        While cnt < size
            lst_ComPorts.Items.Insert(cnt, ports(cnt)) ' fill in both list boxes
            cnt = cnt + 1
        End While
        '------------------------------------------------------------------------------------------
        If SerialPort1.IsOpen = False Then
            lbl_SerialPortStatus.BackColor = Color.OrangeRed ' Intialize the status box
            lbl_SerialPortStatus.Text = "Disconnected"
        Else
            lbl_SerialPortStatus.BackColor = Color.Red ' Intialize the status box
            lbl_SerialPortStatus.Text = "Connected!"
        End If
        '------------------------------------------------------------------------------------------
        btn_Connect.Enabled = True ' Enable the buttons 
        Btn_Disconnect.Enabled = True
        '------------------------------------------------------------------------------------------
    End Sub

    Private Sub SerialPortConnect()

        '-------------------------------------------------------------------------------------------------
        If FailToConnectCnt > 4 Then
            MsgBox("The port has failed to connect 4 times.")
            Exit Sub
        End If
        '-------------------------------------------------------------------------------------------------
        ' set baud rate
        Try
            SerialPort1.BaudRate = 9600

        Catch ex As Exception
            MsgBox("Error setting baud rate on port 1" & ex.ToString)
        End Try
        '-------------------------------------------------------------------------------------------------
        ' Now try to open the port 
        Try
            SerialPort1.Open()
            tmr_Main.Enabled = True
        Catch ex As Exception
            FailToConnectCnt = +FailToConnectCnt ' increment this value 
            Exit Sub
        End Try

        If SerialPort1.IsOpen = True Then
            lbl_SerialPortStatus.BackColor = Color.LightGreen
            lbl_SerialPortStatus.Text = "Connected"
        Else
            lbl_SerialPortStatus.BackColor = Color.OrangeRed
            lbl_SerialPortStatus.Text = "Disconnected"
        End If

        lbl_ConnectionStatus.Text = "Connected"
        '-------------------------------------------------------------------------------------------------
        GetUpdate() ' start updating the form 
    End Sub

    Private Sub SerialPortClose()
        ' Close the serial port 
        Try
            SerialPort1.Close()

        Catch ex As Exception
            MsgBox(ex.ToString) ' if it fails give the windows message why
        End Try


        If SerialPort1.IsOpen = False Then
            lbl_SerialPortStatus.BackColor = Color.OrangeRed
            lbl_SerialPortStatus.Text = "Disconnected"
        Else
            lbl_SerialPortStatus.BackColor = Color.LightGreen
            lbl_SerialPortStatus.Text = "Connected"
        End If

        lbl_ConnectionStatus.Text = "Disconnected"
        FailToConnectCnt = 0 ' reset the counter

    End Sub

    Private Sub SerialPort1_DataReceived(ByVal sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived

        delay(50)
        Try

            RxBuff = (SerialPort1.ReadExisting)
            '  RxBuff = String.Concat(RxBuff, SerialPort1.ReadExisting)

        Catch ex As Exception
            ' just move on if it fails 
        End Try





        If RxBuff <> Nothing And RxBuff <> "?" Then
            NewData = True      ' only flag new data if rxbuff acctually has something in it 
            RXtimeout = 50 ' reset timeout counter 
            If lbl_ConnectionStatus.ForeColor = Color.Blue Then ' flash the label color to indicate data 
                lbl_ConnectionStatus.ForeColor = Color.Orange
            Else : lbl_ConnectionStatus.ForeColor = Color.Blue
            End If
        Else
            NewData = False
        End If
        '-----------------------------------------------------------------------------
        If RxBuff = "?" Or RxBuff = "" Then
            CheckOvenPower = True
        Else
            CheckOvenPower = False
        End If
        '-----------------------------------------------------------------------------
    End Sub

    Private Function QueryController(ByVal StatText As String)
        '-----------------------------------------------------------------------------
        ' sends a character out the serial port and waits for a response 
        Dim trycnt As Integer = 15 ' how many times to try sending 
        Dim newRXdata As String = ""
        NewData = False
        RxBuff = ""
        While trycnt <> 0 And newRXdata = ""

            '-----------------------------------------------------------------------------
            Try

                SerialPort1.Write(StatText)

            Catch ex As Exception
                '       SerialPort1.Dispose()
                Try

                    ComFailure = True
                    Exit While
                Catch ex1 As Exception

                End Try

            End Try
            '-----------------------------------------------------------------------------
            Dim cnt As Integer = 50

            While NewData = False And cnt <> 0
                delay(20)          ' if it doesnt connect within a second time out 
                cnt = cnt - 1
            End While

            If NewData = False Then
                If trycnt = 1 Then
                    '        ComFailure = True 
                    '    MsgBox("would have failed")
                    ' Exit While
                Else
                    trycnt = trycnt - 1

                End If

            Else
                newRXdata = RxBuff
                RxBuff = "" ' clear rxbuff
                NewData = False
                '  Exit While
            End If
            '-----------------------------------------------------------------------------
        End While

        Return newRXdata ' return the information 
    End Function

    Private Sub PortWrite(ByVal text)
        ' put data in this variable which is read by the thread that monitors it
        SerialPort1Data = text

    End Sub

    Private Sub lst_ComPorts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lst_ComPorts.SelectedIndexChanged
        Try

            Selected = CStr(lst_ComPorts.SelectedItem)

            SerialPort1.PortName = Selected
            lbl_SelectedPort.Text = Selected
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try

    End Sub
    ' Buttons 
    Private Sub btn_Connect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Connect.Click

        SerialPortConnect() ' Connect the port 
        ComFailNotified = False
        ComFailure = False ' reset the communications failure flags
        GetUpdate()

    End Sub

    Private Sub Btn_Disconnect_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Disconnect.Click
        PortWrite("U")
        delay(100)
        SerialPortClose()
    End Sub

    Private Sub btn_UpdatePorts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        UpdatePorts()
    End Sub

#End Region

    Private Sub lbl_ExhaustFan_Click(sender As System.Object, e As System.EventArgs) Handles lbl_ExhaustFan.Click

        If ProfileEditorRunning = True Or ProfileDesignerRunning = True Then
            MsgBox("Manual control is disabled while a program is running")

        Else
            Select Case ExhFanStatus

                Case TriState.UseDefault
                    ' if the state is unknown ignore the click
                Case TriState.False
                    PortWrite("D") ' turn it ON
                Case TriState.True
                    PortWrite("d") ' turn it OFF
            End Select
        End If

    End Sub

    Private Sub lbl_ConvFan_Click(sender As System.Object, e As System.EventArgs) Handles lbl_ConvFan.Click

        If ProfileEditorRunning = True Or ProfileDesignerRunning = True Then
            MsgBox("Manual control is disabled while a program is running")

        Else

            Select Case ConvFanStatus

                Case TriState.UseDefault
                    ' if the state is unknown ignore the click
                Case TriState.False
                    PortWrite("C") ' turn it ON
                Case TriState.True
                    PortWrite("c") ' turn it OFF
            End Select

        End If

    End Sub

    Private Sub lbl_TopHeater_Click(sender As System.Object, e As System.EventArgs) Handles lbl_TopHeater.Click

        If ProfileEditorRunning = True Or ProfileDesignerRunning = True Then
            MsgBox("Manual control is disabled while a program is running")

        Else

            Select Case TopHeaterStatus

                Case TriState.UseDefault
                    ' if the state is unknown ignore the click
                Case TriState.False
                    PortWrite("A") ' turn it ON
                Case TriState.True
                    PortWrite("a") ' turn it OFF
            End Select

        End If

    End Sub

    Private Sub lbl_BottomHeater_Click(sender As System.Object, e As System.EventArgs) Handles lbl_BottomHeater.Click
        If ProfileEditorRunning = True Or ProfileDesignerRunning = True Then
            MsgBox("Manual control is disabled while a program is running")

        Else

            Select Case BottomHeaterStatus

                Case TriState.UseDefault
                    ' if the state is unknown ignore the click
                Case TriState.False
                    PortWrite("B") ' turn it ON
                Case TriState.True
                    PortWrite("b") ' turn it OFF
            End Select

        End If

    End Sub

    Private Sub lbl_OvenMode_Click(sender As System.Object, e As System.EventArgs) Handles lbl_OvenMode.Click

        If Mode = "run" Then
            PortWrite("U")
        Else
            PortWrite("N")
        End If

    End Sub

    Private Sub tmr_Main_Tick(sender As System.Object, e As System.EventArgs) Handles tmr_Main.Tick
        Try
            '------------------------------------------------------------------------------------------------------------
            If ComFailure = True And ComFailNotified = False Then
                ComFailNotified = True ' notify the failure 
                SerialPortClose()
                MsgBox("Communications with the Oven has been lost:" & vbNewLine & "  1) Check all connections to the oven" & vbNewLine &
                   "  2) Cycle power on oven" & vbNewLine & "  3) Go to settings Tab and click Connect or restart the program", MsgBoxStyle.Critical)
                '------------------------------------------------------------------------------------------------------------
                ' reset status to unknown 
                TopHeaterStatus = TriState.UseDefault
                BottomHeaterStatus = TriState.UseDefault
                ConvFanStatus = TriState.UseDefault
                ExhFanStatus = TriState.UseDefault
                StopProfileEditor = True
                '------------------------------------------------------------------------------------------------------------

            End If
            '------------------------------------------------------------------------------------------------------------
            RunUpdate = True
            '------------------------------------------------------------------------------------------------------------
            If OKtoUpdate = True Then
                UpdateStatusLabels()
                OKtoUpdate = False
            End If
            '------------------------------------------------------------------------------------------------------------
            ' calculate the new temp entered
            Try
                lbl_NewTemp.Text = CInt(lbl_CurrentOvenTemp.Text) * CInt(txt_OTMulitpleScalar.Text) + CInt(txt_OTAddScalar.Text)
                txt_OTMulitpleScalar.BackColor = System.Drawing.SystemColors.Window
                txt_OTAddScalar.BackColor = System.Drawing.SystemColors.Window
            Catch ex As Exception
                txt_OTMulitpleScalar.BackColor = offcolor
                txt_OTAddScalar.BackColor = offcolor
                txt_OTAddScalar.Text = "0"
                txt_OTMulitpleScalar.Text = "1"
            End Try
            '------------------------------------------------------------------------------------------------------------
            If OverHeat = True Then
                lbl_OverHeatWarning.Visible = True
                If lbl_OverHeatWarning.ForeColor = Color.Red Then
                    lbl_OverHeatWarning.ForeColor = Color.White
                Else
                    lbl_OverHeatWarning.ForeColor = Color.Red
                End If
            Else
                lbl_OverHeatWarning.Visible = False
            End If
            '------------------------------------------------------------------------------------------------------------
            ' update profile editor info if running 


            If ProfileEditorRunning = True Then

                If chk_ProfileEditorLogData.Checked = True Then
                    If Oven_Temp > RecordedMaxTemp Then RecordedMaxTemp = Oven_Temp
                    lbl_MaxTempReached.Text = RecordedMaxTemp
                    lbl_TotalDataPoints.Text = DataPointsCnt
                End If

                UpdateElapsedTime()
                lbl_PhaseTime.Text = Elapsed_inSec
            Else
                UpdateProfileEditor() ' this enables the run button, only run when profile editor is not running

            End If
            '------------------------------------------------------------------------------------------------------------
            If ProfileEditorRunning = True Or ProfileDesignerRunning = True Then
                ' update program run timer if either is running 
                Dim results As TimeSpan
                Dim Str1 As String
                Dim TimeSplit() As String
                results = RunTime.Elapsed
                Str1 = results.ToString
                TimeSplit = Str1.Split(".")
                lbl_RunTime.Text = TimeSplit(0)
                TimeSplit = Str1.Split(":")
                Elapsed_Hours = CInt(TimeSplit(0))
                Elapsed_Min = CInt(TimeSplit(1))
                Elapsed_Sec = CInt(TimeSplit(2))
                Elapsed_RunSec = Elapsed_Min * 60 + Elapsed_Sec ' the elapsed time in just seconds for the Run timer
                lbl_totalTime.Text = Elapsed_RunSec

            End If
            '------------------------------------------------------------------------------------------------------------
            UpdateEditorStatus()
            '------------------------------------------------------------------------------------------------------------
            ' update the program list boxes every 3 seconds ( tmr_Main = 500ms * 6 = 3s)
            If UpdateListsCnt = 0 Then
                UpdateListsCnt = 6
                UpdatePEList() ' update the Program Editor list box 
                lbl_PESaveSuccess.Visible = False ' hide some labels 
                lbl_CSVsaveStat.Visible = False
            Else
                UpdateListsCnt = UpdateListsCnt - 1

            End If
            '------------------------------------------------------------------------------------------------------------
            If RunUpdate = True And T1.IsAlive = False Then
                GetUpdate()
                '  MsgBox("reset T1")

            End If
            '------------------------------------------------------------------------------------------------------------
            ' save the hour meter text
            If IsNumeric(lbl_HourMeter.Text) = False Or lbl_HourMeter.Text = "" Then ' if the hour meter labels is not a number make it one 
                lbl_HourMeter.Text = "1"
            Else
                My.Settings.HourMeterSave = lbl_HourMeter.Text ' save this often incase the program crashes 
                My.Settings.Save()
            End If
            '------------------------------------------------------------------------------------------------------------
            ' check if oven power is off 
            If RXtimeout = 0 Or CheckOvenPower = True Then
                lbl_ConnectionStatus.Text = "Check Oven Power"
                lbl_ConnectionStatus.ForeColor = Color.Red
            Else
                lbl_ConnectionStatus.Text = "Connected"
                RXtimeout = RXtimeout - 1
            End If
            'If CheckOvenPower = True Then
            '    lbl_ConnectionStatus.Text = "Check Oven Power"
            '    lbl_ConnectionStatus.ForeColor = Color.Red
            'Else
            '    lbl_ConnectionStatus.Text = "Connected"
            'End If
            '------------------------------------------------------------------------------------------------------------
            lbl_ExceptionCnt.Text = ExceptionCnt
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception

            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If

        End Try
    End Sub

    Private Sub UpdateElapsedTime()
        ' shave off the miliseconds part of the elapsed time and show new time on label
        ' split the hours, min, sec into integer variables so they can be used else where
        Try
            '------------------------------------------------------------------------------------------------------------
            Dim results As TimeSpan
            results = Elapsedtime.Elapsed
            Dim Str1 As String
            Str1 = results.ToString
            Dim TimeSplit() As String
            TimeSplit = Str1.Split(".")
            Str1 = TimeSplit(0)
            TimeSplit = Str1.Split(":")
            Elapsed_Hours = CInt(TimeSplit(0))
            Elapsed_Min = CInt(TimeSplit(1))
            Elapsed_Sec = CInt(TimeSplit(2))
            Elapsed_inSec = Elapsed_Min * 60 + Elapsed_Sec ' the elapsed time in just seconds 
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub UpdateStatusLabels()
        ' Update the controls status's
        Try
            '----------------------------------------------------------------------------------------------------
            ' some hand shaking for use updating the max and min temp labels 
            If LblWaitingUpdate = True Then
                LblUpdating = False
                Exit Sub
            Else
                LblUpdating = True
            End If
            '----------------------------------------------------------------------------------------------------
            ' update serial number and firmware version 
            lbl_SerialNumber.Text = SerialNumber
            lbl_FirmwareVersion.Text = FirmwareVer
            '---------------------------------------------------------
            ' Update Top Heater status
            Select Case TopHeaterStatus

                Case TriState.UseDefault
                    lbl_TopHeater.BackColor = Color.Silver
                    lbl_TopHeater.Text = "UNK"
                Case TriState.False
                    lbl_TopHeater.BackColor = offcolor ' light red
                    lbl_TopHeater.Text = "OFF"
                Case TriState.True
                    lbl_TopHeater.BackColor = oncolor ' light green
                    lbl_TopHeater.Text = "ON"
            End Select
            '---------------------------------------------------------
            ' Update Bottom Heater status
            Select Case BottomHeaterStatus

                Case TriState.UseDefault
                    lbl_BottomHeater.BackColor = Color.Silver
                    lbl_BottomHeater.Text = "UNK"
                Case TriState.False
                    lbl_BottomHeater.BackColor = offcolor ' light red
                    lbl_BottomHeater.Text = "OFF"
                Case TriState.True
                    lbl_BottomHeater.BackColor = oncolor ' light green
                    lbl_BottomHeater.Text = "ON"
            End Select
            '---------------------------------------------------------
            ' Update Convection Fan Status
            Select Case ConvFanStatus

                Case TriState.UseDefault
                    lbl_ConvFan.BackColor = Color.Silver
                    lbl_ConvFan.Text = "UNK"
                Case TriState.False
                    lbl_ConvFan.BackColor = offcolor ' light red
                    lbl_ConvFan.Text = "OFF"
                Case TriState.True
                    lbl_ConvFan.BackColor = oncolor ' light green
                    lbl_ConvFan.Text = "ON"
            End Select
            '---------------------------------------------------------
            ' Update Top Heater status
            Select Case ExhFanStatus

                Case TriState.UseDefault
                    lbl_ExhaustFan.BackColor = Color.Silver
                    lbl_ExhaustFan.Text = "UNK"
                Case TriState.False
                    lbl_ExhaustFan.BackColor = offcolor ' light red
                    lbl_ExhaustFan.Text = "OFF"
                Case TriState.True
                    lbl_ExhaustFan.BackColor = oncolor ' light green
                    lbl_ExhaustFan.Text = "ON"
            End Select
            '---------------------------------------------------------
            lbl_OvenMode.Text = Mode ' update mode label 
            If Mode = "standby" Then
                rbtn_StandByMode.Checked = True
                rbtn_RunMode.Checked = False
            End If

            If Mode = "run" Then

                rbtn_StandByMode.Checked = False
                rbtn_RunMode.Checked = True

            End If
            '---------------------------------------------------------
            ' Update Oven Temp status

            Select Case Oven_Temp

                Case Is < OvenTempLow
                    lbl_OvenTemp.BackColor = oncolor ' temp is low and label is green
                    OverHeat = False

                Case Is < OvenTempmed
                    lbl_OvenTemp.BackColor = Color.FromArgb(255, 192, 128) ' light orange
                    OverHeat = False

                Case Is < OvenTempHigh
                    lbl_OvenTemp.BackColor = offcolor
                    OverHeat = False

                Case Is >= OvenTempCritical ' special case, flash it RED and white
                    If OverHeat = False Then
                        SerialPort1Data = "L" ' set overheat mode
                        OverHeat = True
                    End If

                    OverHeat = True
                    If lbl_OvenTemp.BackColor = Color.Red Then
                        lbl_OvenTemp.BackColor = Color.White
                        lbl_OverHeatWarning.Visible = True
                    Else
                        lbl_OvenTemp.BackColor = Color.Red
                        lbl_OverHeatWarning.Visible = False
                    End If
            End Select
            '---------------------------------------------------------
            ' Update Case Temp status
            Select Case Case_Temp

                Case Is < CaseTempLow
                    lbl_CaseTemp.BackColor = oncolor ' temp is low and label is green
                    OverHeat = False
                Case Is < CaseTempmed
                    lbl_CaseTemp.BackColor = Color.FromArgb(255, 192, 128) ' light orange
                    OverHeat = False
                Case Is < CaseTempHigh
                    lbl_CaseTemp.BackColor = offcolor
                    OverHeat = False
                Case Is >= CaseTempCritical ' special case, flash it RED and white
                    If OverHeat = False Then
                        SerialPort1Data = "L" ' set overheat mode
                        OverHeat = True
                    End If
                    If lbl_CaseTemp.BackColor = Color.Red Then
                        lbl_CaseTemp.BackColor = Color.White
                    Else
                        lbl_CaseTemp.BackColor = Color.Red
                    End If
            End Select
            '---------------------------------------------------------
            ' update temperature labels

            If TempUnit = False Then
                lbl_CurrentOvenTemp.Text = Oven_Temp
                lbl_OvenTemp.Text = Oven_Temp & "°F"
                lbl_CaseTemp.Text = Case_Temp & "°F"
            Else
                lbl_CurrentOvenTemp.Text = Oven_Temp
                lbl_OvenTemp.Text = Oven_Temp & "°C"
                lbl_CaseTemp.Text = Case_Temp & "°C"
            End If
            '---------------------------------------------------------
            ' Check user defined max temp values 
            If Case_Temp >= MaxCaseTemp And OverHeat = False Then
                SerialPort1Data = "L" ' set overheat mode
                OverHeat = True
            Else
                OverHeat = False
            End If

            If Oven_Temp >= MaxOvenTemp And OverHeat = False Then
                SerialPort1Data = "L" ' set overheat mode
                OverHeat = True
            Else
                OverHeat = False
            End If
            '---------------------------------------------------------
            LblUpdating = False
            '---------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub GetUpdate()
        Try
            T1 = New Thread(New ThreadStart(AddressOf GetUpdateT1)) ' define the new thread and run it 
            T1.Start()
            RunUpdate = True ' start running the update 
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try

    End Sub

    Private Sub GetUpdateT1()
        Try
            While (RunUpdate = True)

                While ComFailure = True Or SerialPort1.IsOpen = False
                    ' spin tires here
                End While
                ' this sub requests the temperature information from the oven
                '---------------------------------------------------------------------------
                ' get serial number and firmware version 
                If SerialNumber = "" Or SerialNumber = Nothing Then
                    SerialNumber = QueryController("s")
                    delay(50)
                End If

                If FirmwareVer = "" Or FirmwareVer = Nothing Then
                    FirmwareVer = QueryController("f")
                End If

                '---------------------------------------------------------------------------
                Dim str() As String
                str = Split(QueryController("S"), ",") ' get the data from serial port and split it
                If str.Length < 4 Then
                    '  GoTo skippy
                    Exit Sub
                End If
                '------------------------------------------
                Select Case str(0) ' Top Heater
                    Case "A"
                        TopHeaterStatus = TriState.True
                    Case "a"
                        TopHeaterStatus = TriState.False
                    Case Else
                        TopHeaterStatus = TriState.UseDefault
                End Select
                '------------------------------------------
                Select Case str(1) ' Bottom Heater
                    Case "B"
                        BottomHeaterStatus = TriState.True
                    Case "b"
                        BottomHeaterStatus = TriState.False
                    Case Else
                        BottomHeaterStatus = TriState.UseDefault
                End Select
                '------------------------------------------
                Select Case str(2) ' Covection Fan 
                    Case "C"
                        ConvFanStatus = TriState.True
                    Case "c"
                        ConvFanStatus = TriState.False
                    Case Else
                        ConvFanStatus = TriState.UseDefault
                End Select
                '------------------------------------------
                Select Case str(3) ' Exhaust Fan
                    Case "D"
                        ExhFanStatus = TriState.True
                    Case "d"
                        ExhFanStatus = TriState.False
                    Case Else
                        ExhFanStatus = TriState.UseDefault
                End Select
                '------------------------------------------
                Select Case str(4) ' Oven Mode
                    Case "U"
                        Mode = "standby"
                    Case "N"
                        Mode = "run"
                    Case "O"
                        Mode = "over heat"
                    Case "L"
                        Mode = "lost connection"
                    Case Else
                        Mode = "unknown"
                End Select
                '------------------------------------------
                ' get temperatures 
                Dim RawOvenTemp As String = QueryController("T")
                delay(50)
                Dim RawCaseTemp As String = QueryController("t")
                Try
                    RawOvenTemp = CLng("&h" & RawOvenTemp)
                    RawCaseTemp = CLng("&h" & RawCaseTemp)
                Catch ex As Exception
                    RawCaseTemp = "1"
                    RawOvenTemp = "1" ' give them a value 
                End Try
                '----------------------------------------------------------------
                ' initialize the temperature arrays if they have not been yet
                If TempAvgInit = False Then
                    TempAvgInit = True
                    ' initialize the 0-10 of each array to 1 
                    Dim cnt1 As Integer = 11
                    While cnt1 <> 0
                        OTArr(cnt1) = RawOvenTemp
                        CTArr(cnt1) = RawCaseTemp
                        cnt1 = cnt1 - 1
                    End While
                End If
                '----------------------------------------------------------------
                ' clean up the Oven Temp Signal signals
                If AvgCoefficient < 1 Or AvgCoefficient > 100 Then AvgCoefficient = 10 ' use the user defined avg coefficient
                Dim MaxOTAvgScalar As Integer = AvgCoefficient
                If OTArrCnt > MaxOTAvgScalar Then ' this is the max number of averages 
                    OTArrCnt = 0 ' clear it if it goes above the max
                End If
                OTArrCnt = OTArrCnt + 1
                ' GroupBox1.Text = RawOvenTemp
                OTArr(OTArrCnt) = CInt(RawOvenTemp)

                ' now average them 
                Dim i As Integer = 0
                Dim cnt As Integer = MaxOTAvgScalar

                While cnt <> 0
                    i = i + OTArr(cnt)
                    cnt = cnt - 1
                End While

                i = i / MaxOTAvgScalar ' divide by number added 
                '----------------------------------------------------------------
                ' clean up the Case Temp Signal signals
                Dim MaxctAvgScalar As Integer = 10
                If CTArrCnt > MaxctAvgScalar Then ' this is the max number of averages 
                    CTArrCnt = 0 ' clear it if it goes above the max
                End If
                CTArrCnt = CTArrCnt + 1
                CTArr(CTArrCnt) = CInt(RawCaseTemp)

                ' now average them 
                i = 0
                cnt = MaxctAvgScalar

                While cnt <> 0
                    i = i + CTArr(cnt)
                    cnt = cnt - 1
                End While

                i = i / MaxctAvgScalar ' divide by number added 
                RawCaseTemp = i
                '----------------------------------------------------------------

                Dim x As Integer = CInt(RawOvenTemp)
                Dim y As Integer

                If x < 490 Then ' scale for temp range
                    y = 102 ' less then 130 deg F

                ElseIf x < 524 Then
                    y = 122 ' less then 200deg f 
                Else
                    y = 132 ' above 200 F
                End If

                x = y + ((x - 486) * 2.11)
                '  RawOvenTemp = x
                If OTAvg <> 0 Then
                    RawOvenTemp = (x + OTAvg) / 2
                    OTAvg = RawOvenTemp
                Else
                    OTAvg = x
                    RawOvenTemp = x
                End If

                If RawOvenTemp < 100 And chk_showbelow100.Checked = False Then RawOvenTemp = 100 ' if its less then 100 dont display it

                '-----------------------------------------------------------------------------------------
                ' factor in the user configurable scalar values from the form
                Try
                    RawOvenTemp = CInt(RawOvenTemp) * MulitplyScalar + AddScalar
                Catch ex As Exception
                    If Debug = True Then
                        ExceptionCnt = ExceptionCnt + 1
                        MsgBox(ex.ToString)
                    End If
                End Try
                '-----------------------------------------------------------------------------------------
                If TempUnit = False Then
                    ' configure fahrenheit
                    Case_Temp = CInt(RawCaseTemp) ' it comes in F
                    Oven_Temp = CInt(RawOvenTemp)
                Else
                    ' configure Celsius
                    Case_Temp = (CInt(RawCaseTemp) - 32) * (5 / 9)
                    Oven_Temp = (CInt(RawOvenTemp) - 32) * (5 / 9)
                End If
                '-----------------------------------------------------------------------------------------
                OKtoUpdate = True
                '-----------------------------------------------------------------------------------------
                ' this is the delay between status updates. it also polls the write varible and transmits
                ' if it has been written 
                Dim cnt2 As Integer = 1000
                While cnt2 <> 0
                    delay(1)
                    cnt2 = cnt2 - 1

                    If SerialPort1Data <> "" Then
                        Try

                            SerialPort1.Write(SerialPort1Data)

                        Catch ex As Exception
                            If Debug = True Then
                                ExceptionCnt = ExceptionCnt + 1
                                MsgBox(ex.ToString)
                            End If
                        End Try

                        SerialPort1Data = ""
                    End If
skippy:
                End While
                '-----------------------------------------------------------------------------------------
            End While ' main while 
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub rbtn_Celsius_CheckedChanged_1(sender As System.Object, e As System.EventArgs) Handles rbtn_Celsius.CheckedChanged
        If LblWaitingUpdate = False Then

            ChangeTempUnit("C")

        End If
    End Sub

    Private Sub rbtn_Fahrenheit_CheckedChanged_1(sender As System.Object, e As System.EventArgs) Handles rbtn_Fahrenheit.CheckedChanged
        If LblWaitingUpdate = False Then

            ChangeTempUnit("F")

        End If
    End Sub

    Private Sub ChangeTempUnit(ByVal T As Char)
        Try
            '------------------------------------------------------------------------------------------------------------

            LblWaitingUpdate = True
            While LblUpdating = True ' wait here for access to the label update routine 

            End While

            Select Case T

                Case "F", "f"
                    If TempUnit = True Then
                        TempUnit = False
                        OvenTempLow = TempConvert(OvenTempLow, "F")
                        OvenTempmed = TempConvert(OvenTempmed, "F")
                        OvenTempHigh = TempConvert(OvenTempHigh, "F")
                        OvenTempCritical = TempConvert(OvenTempCritical, "F")

                        CaseTempmed = TempConvert(CaseTempmed, "F")
                        CaseTempmed = TempConvert(CaseTempmed, "F")
                        CaseTempHigh = TempConvert(CaseTempHigh, "F")
                        CaseTempCritical = TempConvert(CaseTempCritical, "F")

                        MaxOvenTemp = TempConvert(MaxOvenTemp, "F")
                        MaxCaseTemp = TempConvert(MaxCaseTemp, "F")
                        MinOvenTemp = TempConvert(MinOvenTemp, "F")
                        txt_MaxOvenTemp.Text = MaxOvenTemp
                        txt_MaxCaseTemp.Text = MaxCaseTemp

                        TempUnit = False ' set Fahrenheit units
                        lbl_TempUnit_1.Text = "°F"
                        lbl_TempUnit_2.Text = "°F"
                        lbl_TempUnit_3.Text = "°F"
                        lbl_TempUnit_4.Text = "°F"
                        lbl_MaxOvenTempUnit.Text = "°F"
                        lbl_MaxCaseTempUnit.Text = "°F"
                        '      rbtn_Celsius.Checked = False
                        rbtn_Fahrenheit.Checked = True
                    End If

                Case "C", "c"
                    If TempUnit = False Then
                        TempUnit = True
                        OvenTempLow = TempConvert(OvenTempLow, "C")
                        OvenTempmed = TempConvert(OvenTempmed, "C")
                        OvenTempHigh = TempConvert(OvenTempHigh, "C")
                        OvenTempCritical = TempConvert(OvenTempCritical, "C")

                        CaseTempmed = TempConvert(CaseTempmed, "C")
                        CaseTempmed = TempConvert(CaseTempmed, "C")
                        CaseTempHigh = TempConvert(CaseTempHigh, "C")
                        CaseTempCritical = TempConvert(CaseTempCritical, "C")

                        MaxOvenTemp = TempConvert(MaxOvenTemp, "C")
                        MinOvenTemp = TempConvert(MinOvenTemp, "C")
                        MaxCaseTemp = TempConvert(MaxCaseTemp, "C")
                        txt_MaxOvenTemp.Text = MaxOvenTemp ' update the textboxes for the user to edit 
                        txt_MaxCaseTemp.Text = MaxCaseTemp

                        lbl_TempUnit_1.Text = "°C"
                        lbl_TempUnit_2.Text = "°C"
                        lbl_TempUnit_3.Text = "°C"
                        lbl_TempUnit_4.Text = "°C"
                        lbl_MaxOvenTempUnit.Text = "°C"
                        lbl_MaxCaseTempUnit.Text = "°C"

                        rbtn_Celsius.Checked = True
                        ' convert temps to Celsius. this is done in the update thread however if its not done here the oven may 
                        ' go into overheat mode due to the case and oven temps being in F and the test temps in C. its erratic
                        ' b/c it depends when the update thread triggers. This was hard to find.
                        Case_Temp = (Case_Temp - 32) * (5 / 9)
                        Oven_Temp = (Oven_Temp - 32) * (5 / 9)
                    End If

            End Select

            LblWaitingUpdate = False ' allow the labels to be updated 
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub
    Private Sub Button1_Click_2(sender As System.Object, e As System.EventArgs)
        GetUpdate()
    End Sub

    Private Sub btn_ResetAlarm_Click(sender As System.Object, e As System.EventArgs) Handles btn_ResetAlarm.Click
        Try
            PortWrite("N")
            rbtn_RunMode.Checked = True
            rbtn_StandByMode.Checked = False
            GetUpdate()
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub rbtn_UnconnectedMode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtn_StandByMode.CheckedChanged
        Try
            PortWrite("U")
            Mode = "standby"
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub rbtn_ConnectedMode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles rbtn_RunMode.CheckedChanged
        Try
            PortWrite("N")
            Mode = "run"
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub btn_ApplyScalarValues_Click(sender As System.Object, e As System.EventArgs) Handles btn_ApplyScalarValues.Click
        Try
            '----------------------------------------------------------------------------------_
            ' check add scalar
            If IsNumeric(txt_OTAddScalar.Text) = False Then
                txt_OTAddScalar.Text = "0"
                MsgBox("add scaler must a number ")

            Else
                AddScalar = CInt(txt_OTAddScalar.Text)
            End If
            '----------------------------------------------------------------------------------_
            ' check mulitply scalar
            If IsNumeric(txt_OTMulitpleScalar.Text) = False And CInt(txt_OTMulitpleScalar.Text) < 0 And CInt(txt_OTMulitpleScalar.Text) > 10 Then
                txt_OTMulitpleScalar.Text = "1"
                MsgBox("Multiply scalar must be a number greater then 0 and less then 10")

            Else
                MulitplyScalar = CInt(txt_OTMulitpleScalar.Text)
            End If

            '----------------------------------------------------------------------------------_
            ' update averaging coeffient
            AvgCoefficient = nud_AvgCoefficient.Value ' store the averaging coefficient 
            '----------------------------------------------------------------------------------_
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub btn_MaxTempApply_Click(sender As System.Object, e As System.EventArgs) Handles btn_MaxTempApply.Click
        Try
            '------------------------------------------------------------------------------------------------------------
            If IsNumeric(txt_MaxCaseTemp.Text) = False And CInt(txt_MaxCaseTemp.Text) < 50 And CInt(txt_MaxCaseTemp.Text) > 150 Then
                txt_MaxCaseTemp.Text = "120"
                MsgBox("Max. Case Temperature must be a whole number greater than 50°F and less than 150°F")

            Else
                MaxCaseTemp = CInt(txt_MaxCaseTemp.Text)
            End If

            If IsNumeric(txt_MaxOvenTemp.Text) = False And CInt(txt_MaxOvenTemp.Text) < 100 And CInt(txt_MaxOvenTemp.Text) > 750 Then
                txt_MaxOvenTemp.Text = "600"
                MsgBox("Max. oven Temperature must be a whole number greater than 100°F and less than 750°F")

            Else
                MaxOvenTemp = CInt(txt_MaxOvenTemp.Text)
            End If

            MaxCaseTemp = CInt(txt_MaxCaseTemp.Text)
            MaxOvenTemp = CInt(txt_MaxOvenTemp.Text)
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception

        End Try
    End Sub

    Private Sub tmr_Clock_Tick(sender As System.Object, e As System.EventArgs) Handles tmr_Clock.Tick
        Try
            lbl_CurrentTime.Text = Format(DateTime.Now, "HH:mm:ss") 'FormatDateTime(Now, DateFormat.LongTime) ' update the time 
        Catch ex As Exception

        End Try

    End Sub
#Region "Profile Editor"
    Private Editor_PreheatTemp As Integer
    Private Editor_PreheatTime As Integer
    Private Editor_SoakTime As Integer
    Private Editor_SoakTemp As Integer
    Private Editor_ReflowTemp As Integer
    Private Editor_ReflowTime As Integer
    Private Editor_CoolDownTemp As Integer
    Private Editor_CoolDownTime As Integer

    Private Sub UpdateProfileEditor()
        If txt_ProfileEditor_PreheatTemp.Text <> "" And txt_ProfileEditor_PreheatTime.Text <> "" And txt_ProfileEditor_SoakTemp.Text <> "" And
            txt_ProfileEditor_SoakTime.Text <> "" And txt_ProfileEditor_ReflowTime.Text <> "" And txt_ProfileEditor_ReflowTemp.Text <> "" And
            txt_ProfileEditor_CoolDownTime.Text <> "" And txt_ProfileEditor_CoolDownTemp.Text <> "" Then

            btn_PofileEditor_Run.Enabled = True
        Else
            btn_PofileEditor_Run.Enabled = False
        End If
    End Sub
    Private Sub btn_PofileEditor_Run_Click(sender As System.Object, e As System.EventArgs) Handles btn_PofileEditor_Run.Click
        ' input validation first
        Try
            '------------------------------------------------------------------------------
            ' preheat temp
            If IsNumeric(txt_ProfileEditor_PreheatTemp.Text) = False Or CInt(txt_ProfileEditor_PreheatTemp.Text) > MaxOvenTemp Or CInt(txt_ProfileEditor_PreheatTemp.Text) < MinOvenTemp Then
                txt_ProfileEditor_PreheatTemp.BackColor = offcolor
                txt_ProfileEditor_PreheatTemp.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_PreheatTemp.BackColor = Color.White
                Editor_PreheatTemp = CInt(txt_ProfileEditor_PreheatTemp.Text)
            End If
            '------------------------------------------------------------------------------
            ' preheat time 
            If IsNumeric(txt_ProfileEditor_PreheatTime.Text) = False Or CInt(txt_ProfileEditor_PreheatTime.Text) > 1200 Then
                txt_ProfileEditor_PreheatTime.BackColor = offcolor
                txt_ProfileEditor_PreheatTime.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_PreheatTime.BackColor = Color.White
                Editor_PreheatTime = CInt(txt_ProfileEditor_PreheatTime.Text)
            End If
            '------------------------------------------------------------------------------
            ' soak temp
            If IsNumeric(txt_ProfileEditor_SoakTemp.Text) = False Or CInt(txt_ProfileEditor_SoakTemp.Text) > MaxOvenTemp Or CInt(txt_ProfileEditor_SoakTemp.Text) < MinOvenTemp Then
                txt_ProfileEditor_SoakTemp.BackColor = offcolor
                txt_ProfileEditor_SoakTemp.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_SoakTemp.BackColor = Color.White
                Editor_SoakTemp = CInt(txt_ProfileEditor_SoakTemp.Text)
            End If
            ' soak time 
            If IsNumeric(txt_ProfileEditor_SoakTime.Text) = False Or CInt(txt_ProfileEditor_SoakTime.Text) > 1200 Then
                txt_ProfileEditor_SoakTime.BackColor = offcolor
                txt_ProfileEditor_SoakTime.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_SoakTime.BackColor = Color.White
                Editor_SoakTime = CInt(txt_ProfileEditor_SoakTime.Text)
            End If
            '------------------------------------------------------------------------------
            ' reflow temp
            If IsNumeric(txt_ProfileEditor_ReflowTemp.Text) = False Or CInt(txt_ProfileEditor_ReflowTemp.Text) > MaxOvenTemp Or CInt(txt_ProfileEditor_ReflowTemp.Text) < MinOvenTemp Then
                txt_ProfileEditor_ReflowTemp.BackColor = offcolor
                txt_ProfileEditor_ReflowTemp.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_ReflowTemp.BackColor = Color.White
                Editor_ReflowTemp = CInt(txt_ProfileEditor_ReflowTemp.Text)
            End If
            ' reflow time 
            If IsNumeric(txt_ProfileEditor_ReflowTime.Text) = False Or CInt(txt_ProfileEditor_ReflowTime.Text) > 1200 Then
                txt_ProfileEditor_ReflowTime.BackColor = offcolor
                txt_ProfileEditor_ReflowTime.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_ReflowTime.BackColor = Color.White
                Editor_ReflowTime = CInt(txt_ProfileEditor_ReflowTime.Text)
            End If
            '------------------------------------------------------------------------------
            ' cooldown temp
            If IsNumeric(txt_ProfileEditor_CoolDownTemp.Text) = False Or CInt(txt_ProfileEditor_CoolDownTemp.Text) > MaxOvenTemp Or CInt(txt_ProfileEditor_CoolDownTemp.Text) < MinOvenTemp Then
                txt_ProfileEditor_CoolDownTemp.BackColor = offcolor
                txt_ProfileEditor_CoolDownTemp.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_CoolDownTemp.BackColor = Color.White
                Editor_CoolDownTemp = CInt(txt_ProfileEditor_CoolDownTemp.Text)
            End If
            ' cooldown time 
            If IsNumeric(txt_ProfileEditor_CoolDownTime.Text) = False Or CInt(txt_ProfileEditor_CoolDownTime.Text) > 1200 Then

                txt_ProfileEditor_CoolDownTime.BackColor = offcolor
                txt_ProfileEditor_CoolDownTime.Text = ""
                MsgBox("Temp must be a number between " & MinOvenTemp & " and " & MaxOvenTemp & " and time must be from 1-1200. Please correct the highlighted fields")
                Exit Sub
            Else
                txt_ProfileEditor_CoolDownTime.BackColor = Color.White
                Editor_CoolDownTime = CInt(txt_ProfileEditor_CoolDownTime.Text)
            End If
            '------------------------------------------------------------------------------
            ' make sure it is in run mode 
            Dim cnt1 As Integer = 5 ' how many time it will send the command 
            While cnt1 <> 0 And Mode <> "run"
                PortWrite("N") ' put oven in run mode 

                Dim cnt As Integer = 100
                While Mode <> "run"
                    delay(10)
                    cnt = cnt - 1

                End While

            End While

            If cnt1 = 0 Then
                MsgBox("Failed to set run mode, program aborted")
                Exit Sub
            End If

            delay(100)
            '------------------------------------------------------------------------------
            Elapsedtime.Start() ' start the timer
            RunTime.Start() ' star the program run timer 
            '------------------------------------------------------------------------------
            lbl_PhaseTime.Text = Format(DateTime.Now, "HH:mm:ss") 'FormatDateTime(Now, DateFormat.LongTime) ' update the time 
            '------------------------------------------------------------------------------
            If chk_ProfileEditorLogData.Checked = True Then
                ' enable data logging 
                tmr_DataLogger.Interval = nud_LogInterval.Value * 1000 ' get the user defined interval
                tmr_DataLogger.Enabled = True ' start logging 
                TabControl1.SelectedTab = TabPage5 ' show the logging tab page 
            End If
            '------------------------------------------------------------------------------
            btn_PofileEditor_Run.Enabled = False
            btn_PofileEditor_Stop.Enabled = True
            '     lbl_ProramStatus.Text = "Editor Running"
            '    lbl_ProramStatus.ForeColor = Color.ForestGreen
            '------------------------------------------------------------------------------
            s1.Points.Clear() ' clear out the current chart 
            rtxt_LoggedTempValues.Clear()
            rtxt_loggedTimeValues.Clear()
            gbox_DataLogControls.Enabled = False ' disable data logging controls while running
            RecordedMaxTemp = 0
            DataPointsCnt = 0
            '------------------------------------------------------------------------------
            ' if it makes it through input validation run the program
            BackgroundWorker1.RunWorkerAsync()
            ProfileEditorRunning = True ' tell the other async process to start working
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RunProfileEditor()
        delay(100)
        Try
            '------------------------------------------------------------------------------
            ' Preheat sequence
            EditorStatus = "PRE-HEAT"
            Dim waitCnt As Integer = 300
            Dim WhosTurn As Boolean ' false = Top, true = bottom 
            While Oven_Temp <= Editor_PreheatTemp And Elapsed_inSec < Editor_PreheatTime

                If chk_PEHalfPwr.Checked = True Then ' here for half power
                    '---------------------------------------------
                    If waitCnt = 0 Then
                        waitCnt = 300 ' this is how long the heater will be on 
                        If WhosTurn = False Then WhosTurn = True Else WhosTurn = False ' toggle whos turn it is
                    Else
                        waitCnt = waitCnt - 1
                    End If
                    '---------------------------------------------
                    If WhosTurn = False Then

                        If TopHeaterStatus = TriState.False Then
                            PortWrite("A") ' turn ON Top Heater
                            delay(100)
                        End If

                        If TopHeaterStatus = TriState.True Then
                            PortWrite("b") ' turn OFF Bottom Heater
                            delay(100)
                        End If


                    Else
                        If TopHeaterStatus = TriState.False Then
                            PortWrite("B") ' turn ON Bottom Heater
                            delay(100)
                        End If

                        If TopHeaterStatus = TriState.True Then
                            PortWrite("a") ' turn OFF Top Heater
                            delay(100)
                        End If

                    End If
                    '---------------------------------------------
                Else ' here for full power 

                    If TopHeaterStatus = TriState.False Then
                        PortWrite("A") ' turn ON Top Heater
                        delay(100)
                    End If
                    If TopHeaterStatus = TriState.False Then
                        PortWrite("B") ' turn ON Bottom Heater
                        delay(100)
                    End If
                End If
                '---------------------------------------------
                If chk_PreheatConvOption.Checked = True And ConvFanStatus = TriState.False Then
                    PortWrite("C") ' turn ON Convection Fan
                    delay(100)

                ElseIf chk_PreheatConvOption.Checked = False And ConvFanStatus = TriState.True Then
                    PortWrite("c") ' turn OFF Convection Fan
                    delay(100)
                End If
                '---------------------------------------------
                If BackgroundWorker1.CancellationPending = True Then
                    ' put the oven back in standby mode 
                    PortWrite("U")
                    Exit Sub
                End If
                '---------------------------------------------
                delay(100)
                '---------------------------------------------
                If Mode <> "run" Then
                    PortWrite("U")
                    EditorStatus = "STOPPED"
                    StopProfileEditor = True
                    Exit Sub
                End If
                '---------------------------------------------
            End While
            '------------------------------------------------------------------------------
            ' Soak Sequence

            EditorStatus = "SOAK"
            Elapsedtime.Restart() ' restart the timer
            UpdateElapsedTime() ' update the Elapsed_inSec Variable
            While Oven_Temp <= Editor_SoakTemp And Elapsed_inSec < Editor_SoakTime

                If TopHeaterStatus = TriState.True Then
                    PortWrite("a") ' turn OFF Top Heater
                    delay(100)
                End If

                If BottomHeaterStatus = TriState.True Then
                    PortWrite("b") ' turn OFF Bottom Heater
                    delay(100)
                End If

                If chk_SoakConvOption.Checked = True And ConvFanStatus = TriState.False Then
                    PortWrite("C") ' turn ON Convection Fan
                    delay(100)

                ElseIf chk_SoakConvOption.Checked = False And ConvFanStatus = TriState.True Then
                    PortWrite("c") ' turn OFF Convection Fan
                    delay(100)
                End If

                If BackgroundWorker1.CancellationPending = True Then
                    ' put the oven back in standby mode 
                    PortWrite("U")
                    Exit Sub
                End If
                delay(100)
                If Mode <> "run" Then
                    PortWrite("U")
                    EditorStatus = "STOPPED"
                    StopProfileEditor = True
                    Exit Sub
                End If
            End While
            '------------------------------------------------------------------------------
            ' Reflow Sequence
            EditorStatus = "REFLOW"

            Elapsedtime.Restart() ' restart the timer
            UpdateElapsedTime() ' update the Elapsed_inSec Variable
            While Oven_Temp <= Editor_ReflowTemp And Elapsed_inSec < Editor_ReflowTime

                If TopHeaterStatus = TriState.False Then
                    PortWrite("A") ' turn ON Top Heater
                    delay(100)
                End If

                If BottomHeaterStatus = TriState.False Then
                    PortWrite("B") ' turn ON Bottom Heater
                    delay(100)
                End If

                If chk_ReflowConvOption.Checked = True And ConvFanStatus = TriState.False Then
                    PortWrite("C") ' turn ON Convection Fan
                    delay(100)

                ElseIf chk_ReflowConvOption.Checked = False And ConvFanStatus = TriState.True Then
                    PortWrite("c") ' turn OFF Convection Fan
                    delay(100)
                End If

                If BackgroundWorker1.CancellationPending = True Then
                    ' put the oven back in standby mode 
                    PortWrite("U")
                    Exit Sub
                End If
                delay(100)
                If Mode <> "run" Then
                    PortWrite("U")
                    EditorStatus = "STOPPED"
                    StopProfileEditor = True
                    Exit Sub
                End If
            End While
            '------------------------------------------------------------------------------
            ' Cool down sequence
            EditorStatus = "COOLDOWN"
            Elapsedtime.Restart() ' restart the timer
            UpdateElapsedTime() ' update the Elapsed_inSec Variable
            While Oven_Temp >= Editor_CoolDownTemp And Elapsed_inSec < Editor_CoolDownTime

                If TopHeaterStatus = TriState.True Then
                    PortWrite("a") ' turn OFF Top Heater
                    delay(100)
                End If

                If BottomHeaterStatus = TriState.True Then
                    PortWrite("b") ' turn OFF Bottom Heater
                    delay(100)
                End If

                If ExhFanStatus = TriState.False Then
                    PortWrite("D") ' turn ON Exhaust Fan
                    delay(100)
                End If

                If chk_CooldownConvOption.Checked = True And ConvFanStatus = TriState.False Then
                    PortWrite("C") ' turn ON Convection Fan
                    delay(100)

                ElseIf chk_CooldownConvOption.Checked = False And ConvFanStatus = TriState.True Then
                    PortWrite("c") ' turn OFF Convection Fan
                    delay(100)
                End If

                If BackgroundWorker1.CancellationPending = True Then
                    ' put the oven back in standby mode 
                    PortWrite("U")
                    Exit Sub
                End If
                delay(100)
                If Mode <> "run" Then
                    PortWrite("U")
                    EditorStatus = "STOPPED"
                    StopProfileEditor = True
                    Exit Sub
                End If
            End While
            '------------------------------------------------------------------------------
            ' put the oven back in standby mode 
            PortWrite("U")
            delay(100)
            EditorStatus = "STOPPED"
            StopProfileEditor = True
            PortWrite("V") ' long beep to signal completion 
            '------------------------------------------------------------------------------
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception

        End Try
    End Sub

    Private Sub EndProfileEditor()
        Try
            '------------------------------------------------------------------------------
            ProfileEditorRunning = False
            ' Cancel the background process and disable the stop button
            BackgroundWorker1.CancelAsync()
            btn_PofileEditor_Stop.Enabled = False
            btn_PofileEditor_Run.Enabled = True
            Elapsedtime.Reset()
            RunTime.Reset()
            '    lbl_ProramStatus.Text = "Program STOPPED"
            '    lbl_ProramStatus.ForeColor = Color.Red
            tmr_DataLogger.Enabled = False
            gbox_DataLogControls.Enabled = True
            '------------------------------------------------------------------------------
            ' Lastly, pump a few values at the end of the graph to make it look nicer 
            ' these values are added to the graph only and not the textboxs or saved file

            ' update the graph data 
            Chart1.Series.Clear()
            '  s1.Name = "Temp"
            s1.ChartType = SeriesChartType.Line
            s1.Points.AddXY(Elapsed_RunSec + 50, Oven_Temp)
            Chart1.Series.Add(s1)
            '------------------------------------------------------------------------------
            ' udpate hour meter 
            Dim i As Double = CDbl(lbl_HourMeter.Text)
            lbl_HourMeter.Text = Math.Round((CInt(Elapsed_RunSec) / 3600) + i, 2) ' formated to 2 digits 
            '------------------------------------------------------------------------------
        Catch ex As Exception

        End Try
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        RunProfileEditor()
    End Sub

    Private Sub btn_PofileEditor_Stop_Click(sender As System.Object, e As System.EventArgs) Handles btn_PofileEditor_Stop.Click
        EndProfileEditor()
    End Sub

    Private Sub UpdateEditorStatus()
        Try
            '------------------------------------------------------------------------------------------------------------
            If StopProfileEditor = True Then
                StopProfileEditor = False
                EndProfileEditor()
            End If

            Select Case EditorStatus
                Case "PRE-HEAT"
                    lbl_PreheatIndicator.ForeColor = Color.ForestGreen
                    lbl_SoakIndicator.ForeColor = Color.Red
                    lbl_ReflowIndicator.ForeColor = Color.Red
                    lbl_CoolDownIndicator.ForeColor = Color.Red
                Case "SOAK"
                    lbl_PreheatIndicator.ForeColor = Color.Red
                    lbl_SoakIndicator.ForeColor = Color.ForestGreen
                    lbl_ReflowIndicator.ForeColor = Color.Red
                    lbl_CoolDownIndicator.ForeColor = Color.Red
                Case "REFLOW"
                    lbl_PreheatIndicator.ForeColor = Color.Red
                    lbl_SoakIndicator.ForeColor = Color.Red
                    lbl_ReflowIndicator.ForeColor = Color.ForestGreen
                    lbl_CoolDownIndicator.ForeColor = Color.Red
                Case "COOLDOWN"
                    lbl_PreheatIndicator.ForeColor = Color.Red
                    lbl_SoakIndicator.ForeColor = Color.Red
                    lbl_ReflowIndicator.ForeColor = Color.Red
                    lbl_CoolDownIndicator.ForeColor = Color.ForestGreen
                Case "STOPPED"
                    lbl_PreheatIndicator.ForeColor = Color.Red
                    lbl_SoakIndicator.ForeColor = Color.Red
                    lbl_ReflowIndicator.ForeColor = Color.Red
                    lbl_CoolDownIndicator.ForeColor = Color.Red
            End Select

            lbl_EditorStatus.Text = EditorStatus
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_PofileEditor_Save_Click(sender As System.Object, e As System.EventArgs) Handles btn_PofileEditor_Save.Click
        ' PERFORM 
        '----------------------------------------------------------------------------------------------------------------
        ' create the program's root directory if it doesnt exist 
        ' it should have been created by the installation program
        If Directory.Exists(ProgramLibaryPath & "Profile Editor\") = True Then

        Else

            Try
                Directory.CreateDirectory(ProgramLibaryPath & "Profile Editor\")
                MsgBox("Error: Could not find root directory" & vbNewLine & ProgramLibaryPath & "A new directory has been created at the above path")
            Catch ex As Exception
                MsgBox("Error: Could not find root directory" & vbNewLine & ProgramLibaryPath & "A new directory could not be created" &
                       vbNewLine & "Check the user has permission to access C: and try again")
                Exit Sub
            End Try

        End If
        '----------------------------------------------------------------------------------------------------------------
        ' load all the form data into an array to be saved to the XML file 
        Dim commands(20) As String ' 
        commands(0) = txt_ProfileEditor_PreheatTemp.Text
        commands(1) = txt_ProfileEditor_PreheatTime.Text
        commands(2) = txt_ProfileEditor_SoakTemp.Text
        commands(3) = txt_ProfileEditor_SoakTime.Text
        commands(4) = txt_ProfileEditor_ReflowTemp.Text
        commands(5) = txt_ProfileEditor_ReflowTime.Text
        commands(6) = txt_ProfileEditor_CoolDownTemp.Text
        commands(7) = txt_ProfileEditor_CoolDownTime.Text

        ' save convection fan status's
        If chk_PreheatConvOption.Checked = True Then ' save convection fan state for preheat
            commands(8) = "1"
        Else
            commands(8) = "0"
        End If

        If chk_SoakConvOption.Checked = True Then ' save convection fan state for soak
            commands(9) = "1"
        Else
            commands(9) = "0"
        End If

        If chk_ReflowConvOption.Checked = True Then ' save convection fan state for reflow
            commands(10) = "1"
        Else
            commands(10) = "0"
        End If

        If chk_CooldownConvOption.Checked = True Then ' save convection fan state for cooldown
            commands(11) = "1"
        Else
            commands(11) = "0"
        End If

        ' Save suceptor tray state 
        If chk_IsotropicTrayIndicator1.Checked = True Then ' save suceptor tray used status
            commands(12) = "1"
        Else
            commands(12) = "0"
        End If

        If chk_ProfileEditorLogData.Checked = True Then ' save run data logger state 
            commands(13) = "1"
        Else
            commands(13) = "0"
        End If

        ' save the notes box
        Dim str As String = Trim(rtxt_PENotes.Text)
        commands(14) = str

        If chk_PEHalfPwr.Checked = True Then ' save pre heat power setting 
            commands(15) = "1"
        Else
            commands(15) = "0"
        End If

        If TempUnit = False Then ' save temp units 
            commands(16) = "0"
        Else
            commands(16) = "1"
        End If
        '----------------------------------------------------------------------------------------------------------------
        Dim filename As String = Trim(txt_PofileEditor_SaveAs.Text) ' get file name from textbox
        Dim memory_stream As New MemoryStream
        Dim xml_text_writer As New XmlTextWriter(ProgramLibaryPath & "Profile Editor\" & filename & ".xml", System.Text.Encoding.UTF8)

        ' Use indentation to make the result look nice.
        xml_text_writer.Formatting = Formatting.Indented
        xml_text_writer.Indentation = 4

        ' Write the XML declaration.
        xml_text_writer.WriteStartDocument(True)

        ' Start the commands node.
        xml_text_writer.WriteStartElement("Commands")

        ' Write the commands

        Dim cnt As Integer = 0  ' save out the command data to an XML File 
        While commands(cnt) <> Nothing
            MakeCommand(xml_text_writer, commands(cnt))
            cnt = cnt + 1
        End While

        xml_text_writer.WriteEndElement()
        xml_text_writer.WriteEndDocument()
        xml_text_writer.Flush()
        xml_text_writer.Close()

        '----------------------------------------------------------------------------------------------------------------
        lbl_PESaveSuccess.Visible = True
        '----------------------------------------------------------------------------------------------------------------

    End Sub

    Private Sub btn_PofileEditor_Delete_Click(sender As System.Object, e As System.EventArgs) Handles btn_PofileEditor_Delete.Click
        Try
            '------------------------------------------------------------------------------------------------------------

            If txt_PofileEditor_SaveAs.Text = Nothing Then
                MsgBox("No file selected to delete")
            Else
                txt_PofileEditor_SaveAs.Text = Trim(txt_PofileEditor_SaveAs.Text)
                Dim filename As String = ProgramLibaryPath & "Profile Editor\" & (txt_PofileEditor_SaveAs.Text) & ".xml"
                filename.Trim() ' trim out accidental spaces by user

                Try
                    File.Delete(filename)
                    txt_PofileEditor_SaveAs.Clear()
                Catch ex As Exception
                    MsgBox("Could not find file to delete")
                End Try

            End If
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub UpdatePEList()
        If Directory.Exists(ProgramLibaryPath & "Profile Editor\") = True Then
            Try
                lst_ProfileEditor.Items.Clear()
                Dim items(999) As String ' array for the directory contents 
                items = Directory.GetFileSystemEntries(ProgramLibaryPath & "Profile Editor\")
                '------------------------------------
                ' Split out the file path so that the list box only shows the file name
                Dim items2(999) As String ' array for the file path 
                Dim items3(9) As String ' array for the file name 
                '------------------------------------
                Dim Icnt As Integer = 0
                While Icnt < items.Length
                    items2 = Split(items(Icnt), "\") ' split the file path by \ and take the file name 
                    items3 = Split(items2(4), ".") ' split the file name by . and take the file name and toss out the .txt 

                    lst_ProfileEditor.Items.Insert(Icnt, items3(0)) ' load the list with just the file name 
                    Icnt = Icnt + 1

                End While
            Catch ex As Exception
                If Debug = True Then
                    ExceptionCnt = ExceptionCnt + 1
                    MsgBox(ex.ToString)
                End If
            End Try
        End If
    End Sub

    Private Shared Sub MakeCommand(ByVal xml_text_writer As XmlTextWriter, ByVal command As String)
        Try
            '------------------------------------------------------------------------------------------------------------

            ' Start the Command Element
            xml_text_writer.WriteStartElement("Commands")

            ' Write the Command
            xml_text_writer.WriteStartElement("Command")
            xml_text_writer.WriteString(command)
            xml_text_writer.WriteEndElement()

            ' Close the Command element.
            xml_text_writer.WriteEndElement()
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception

        End Try
    End Sub


    Private Sub lst_ProfileEditor_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles lst_ProfileEditor.SelectedIndexChanged

        Dim filename As String = Trim(lst_ProfileEditor.SelectedItem)

        Try
            If File.Exists(ProgramLibaryPath & "Profile Editor\" & filename & ".xml") Then
                Dim document As XmlDocument = New XmlDocument()
                document.Load(ProgramLibaryPath & "Profile Editor\" & filename & ".xml")

                ' create XmlNodeReader for document
                Dim reader As XmlNodeReader = New XmlNodeReader(document)
                Dim commands(20) As String
                '----------------------------------

                While reader.Read
                    If reader.NodeType = XmlNodeType.Text Then
                        Dim cnt1 As Integer
                        commands(cnt1) = reader.Value
                        cnt1 = cnt1 + 1
                    End If
                End While
                '----------------------------------
                txt_ProfileEditor_PreheatTemp.Text = commands(0)
                txt_ProfileEditor_PreheatTime.Text = commands(1)
                txt_ProfileEditor_SoakTemp.Text = commands(2)
                txt_ProfileEditor_SoakTime.Text = commands(3)
                txt_ProfileEditor_ReflowTemp.Text = commands(4)
                txt_ProfileEditor_ReflowTime.Text = commands(5)
                txt_ProfileEditor_CoolDownTemp.Text = commands(6)
                txt_ProfileEditor_CoolDownTime.Text = commands(7)


                If commands(8) = "1" Then
                    chk_PreheatConvOption.Checked = True
                Else
                    chk_PreheatConvOption.Checked = False
                End If

                If commands(9) = "1" Then
                    chk_SoakConvOption.Checked = True
                Else
                    chk_SoakConvOption.Checked = False
                End If

                If commands(10) = "1" Then
                    chk_ReflowConvOption.Checked = True
                Else
                    chk_ReflowConvOption.Checked = False
                End If

                If commands(11) = "1" Then
                    chk_CooldownConvOption.Checked = True
                Else
                    chk_CooldownConvOption.Checked = False
                End If

                If commands(12) = "1" Then
                    chk_IsotropicTrayIndicator1.Checked = True
                Else
                    chk_IsotropicTrayIndicator1.Checked = False
                End If

                If commands(13) = "1" Then
                    chk_ProfileEditorLogData.Checked = True
                Else
                    chk_ProfileEditorLogData.Checked = False
                End If


                rtxt_PENotes.Text = commands(14)

                txt_PofileEditor_SaveAs.Text = filename ' show the filename 

                If commands(15) = "1" Then ' load the pre heat power setting 
                    chk_PEHalfPwr.Checked = True
                Else
                    chk_PEHalfPwr.Checked = False
                End If

                If commands(16) = "1" Then ' change the units to what the program was set in
                    ChangeTempUnit("C")
                Else
                    ChangeTempUnit("F")
                End If

            End If
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
        btn_PofileEditor_Stop.Enabled = False
    End Sub

    Private Sub tab_ProfileEditor_Enter(sender As Object, e As System.EventArgs) Handles tab_ProfileEditor.Enter
        ' update the list upon entering the tab
        UpdatePEList()
    End Sub

    Private Sub btn_PEClearValues_Click(sender As System.Object, e As System.EventArgs) Handles btn_PEClearValues.Click

        ' clear all the data for the Program Editor
        txt_ProfileEditor_PreheatTemp.Text = ""
        txt_ProfileEditor_PreheatTime.Text = ""
        txt_ProfileEditor_SoakTemp.Text = ""
        txt_ProfileEditor_SoakTime.Text = ""
        txt_ProfileEditor_ReflowTemp.Text = ""
        txt_ProfileEditor_ReflowTime.Text = ""
        txt_ProfileEditor_CoolDownTemp.Text = ""
        txt_ProfileEditor_CoolDownTime.Text = ""
        rtxt_PENotes.Clear()

        ' now the check boxes 
        chk_ProfileEditorLogData.Checked = False
        chk_IsotropicTrayIndicator1.Checked = False
        chk_CooldownConvOption.Checked = False
        chk_ReflowConvOption.Checked = False
        chk_SoakConvOption.Checked = False
        chk_PreheatConvOption.Checked = False

    End Sub

    Private Sub info_PEHalfPwr_Click(sender As System.Object, e As System.EventArgs) Handles info_PEHalfPwr.Click
        MsgBox("Half power pre heating toggles the top and bottom heater every 30 seconds so they are not ON at the same time", MsgBoxStyle.Information)
    End Sub

#Region "Profile Editor on the fly input validation"


    Private Sub txt_ProfileEditor_PreheatTemp_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_PreheatTemp.LostFocus

        If CInt(txt_ProfileEditor_PreheatTemp.Text) < MinOvenTemp Then
            Dim unit As String = "°F"
            If TempUnit = True Then unit = "°C"
            txt_ProfileEditor_PreheatTemp.Text = ""
            txt_ProfileEditor_PreheatTemp.Focus()
            MsgBox("Must be a whole number between " & MinOvenTemp & unit & " and " & MaxOvenTemp & unit & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
        End If

    End Sub


    Private Sub txt_ProfileEditor_PreheatTemp_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_PreheatTemp.TextChanged

        If txt_ProfileEditor_PreheatTemp.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_PreheatTemp.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_PreheatTemp.TextLength > 3 Or CInt(txt_ProfileEditor_PreheatTemp.Text) > MaxOvenTemp Then
                txt_ProfileEditor_PreheatTemp.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_PreheatTemp.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)

        End If

    End Sub

    Private Sub txt_ProfileEditor_PreheatTime_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_PreheatTime.LostFocus

        If CInt(txt_ProfileEditor_PreheatTime.Text) < MinOvenTime Then

            txt_ProfileEditor_PreheatTime.Text = ""
            txt_ProfileEditor_PreheatTime.Focus()
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_PreheatTime_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_PreheatTime.TextChanged

        If txt_ProfileEditor_PreheatTime.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_PreheatTime.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_PreheatTime.TextLength > 3 Or CInt(txt_ProfileEditor_PreheatTime.Text) > MaxOventime Then
                txt_ProfileEditor_PreheatTime.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_PreheatTime.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)

        End If

    End Sub

    Private Sub txt_ProfileEditor_SoakTime_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_SoakTime.LostFocus

        If CInt(txt_ProfileEditor_SoakTime.Text) < MinOvenTime Then

            txt_ProfileEditor_SoakTime.Text = ""
            txt_ProfileEditor_SoakTime.Focus()
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_SoakTime_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_SoakTime.TextChanged

        If txt_ProfileEditor_SoakTime.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_SoakTime.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_SoakTime.TextLength > 3 Or CInt(txt_ProfileEditor_SoakTime.Text) > MaxOvenTime Then
                txt_ProfileEditor_SoakTime.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_SoakTime.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)

        End If

    End Sub

    Private Sub txt_ProfileEditor_SoakTemp_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_SoakTemp.LostFocus

        If CInt(txt_ProfileEditor_SoakTemp.Text) < MinOvenTemp Then
            Dim unit As String = "°F"
            If TempUnit = True Then unit = "°C"
            txt_ProfileEditor_SoakTemp.Text = ""
            txt_ProfileEditor_SoakTemp.Focus()
            MsgBox("Must be a whole number between " & MinOvenTemp & unit & " and " & MaxOvenTemp & unit & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_SoakTemp_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_SoakTemp.TextChanged
        If txt_ProfileEditor_SoakTemp.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_SoakTemp.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_SoakTemp.TextLength > 3 Or CInt(txt_ProfileEditor_SoakTemp.Text) > MaxOvenTemp Then
                txt_ProfileEditor_SoakTemp.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_SoakTemp.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)

        End If
    End Sub

    Private Sub txt_ProfileEditor_ReflowTime_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_ReflowTime.LostFocus

        If CInt(txt_ProfileEditor_ReflowTime.Text) < MinOvenTime Then

            txt_ProfileEditor_ReflowTime.Text = ""
            txt_ProfileEditor_ReflowTime.Focus()
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_ReflowTime_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_ReflowTime.TextChanged

        If txt_ProfileEditor_ReflowTime.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_ReflowTime.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_ReflowTime.TextLength > 3 Or CInt(txt_ProfileEditor_ReflowTime.Text) > MaxOvenTime Then
                txt_ProfileEditor_ReflowTime.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_ReflowTime.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)

        End If

    End Sub

    Private Sub txt_ProfileEditor_ReflowTemp_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_ReflowTemp.LostFocus

        If CInt(txt_ProfileEditor_ReflowTemp.Text) < MinOvenTemp Then
            Dim unit As String = "°F"
            If TempUnit = True Then unit = "°C"
            txt_ProfileEditor_ReflowTemp.Text = ""
            txt_ProfileEditor_ReflowTemp.Focus()
            MsgBox("Must be a whole number between " & MinOvenTemp & unit & " and " & MaxOvenTemp & unit & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_ReflowTemp_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_ReflowTemp.TextChanged
        If txt_ProfileEditor_ReflowTemp.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_ReflowTemp.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_ReflowTemp.TextLength > 3 Or CInt(txt_ProfileEditor_ReflowTemp.Text) > MaxOvenTemp Then
                txt_ProfileEditor_ReflowTemp.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_ReflowTemp.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)

        End If
    End Sub

    Private Sub txt_ProfileEditor_CoolDownTime_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_CoolDownTime.LostFocus

        If CInt(txt_ProfileEditor_CoolDownTime.Text) < MinOvenTime Then

            txt_ProfileEditor_CoolDownTime.Text = ""
            txt_ProfileEditor_CoolDownTime.Focus()
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_CoolDownTime_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_CoolDownTime.TextChanged

        If txt_ProfileEditor_CoolDownTime.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_CoolDownTime.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_CoolDownTime.TextLength > 3 Or CInt(txt_ProfileEditor_CoolDownTime.Text) > MaxOvenTime Then
                txt_ProfileEditor_CoolDownTime.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_CoolDownTime.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTime & " and " & MaxOvenTime & vbNewLine & "Ex: 90", MsgBoxStyle.Exclamation)

        End If

    End Sub

    Private Sub txt_ProfileEditor_CoolDownTemp_LostFocus(sender As Object, e As System.EventArgs) Handles txt_ProfileEditor_CoolDownTemp.LostFocus

        If CInt(txt_ProfileEditor_CoolDownTemp.Text) < MinOvenTemp Then
            Dim unit As String = "°F"
            If TempUnit = True Then unit = "°C"
            txt_ProfileEditor_CoolDownTemp.Text = ""
            txt_ProfileEditor_CoolDownTemp.Focus()
            MsgBox("Must be a whole number between " & MinOvenTemp & unit & " and " & MaxOvenTemp & unit & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
        End If

    End Sub

    Private Sub txt_ProfileEditor_CoolDownTemp_TextChanged(sender As System.Object, e As System.EventArgs) Handles txt_ProfileEditor_CoolDownTemp.TextChanged

        If txt_ProfileEditor_CoolDownTemp.Text = "" Then
            Exit Sub
        End If

        If IsNumeric(txt_ProfileEditor_CoolDownTemp.Text) = True Then
            ' must do this one first to be sure it is a number otherwise errors ensue
            If txt_ProfileEditor_CoolDownTemp.TextLength > 3 Or CInt(txt_ProfileEditor_CoolDownTemp.Text) > MaxOvenTemp Then
                txt_ProfileEditor_CoolDownTemp.Text = ""
                MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)
            End If

        Else
            txt_ProfileEditor_CoolDownTemp.Text = ""
            MsgBox("Must be a whole number between " & MinOvenTemp & " and " & MaxOvenTemp & vbNewLine & "Ex: 125", MsgBoxStyle.Exclamation)

        End If

    End Sub
    ' End Profile editor input validation region 
    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#End Region


    ' end profile editor region
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#End Region

#Region "Data Logger Tab"
    Private Sub btn_PrintGraph_Click(sender As System.Object, e As System.EventArgs) Handles btn_PrintGraph.Click
        Dim shortFilename As String
        If txt_PofileEditor_SaveAs.Text <> "" Then
            shortFilename = "Graph_" & Trim(txt_PofileEditor_SaveAs.Text) & "_" & Replace(Replace(Replace(Date.Now, "/", "-"), ":", "."), " ", "_")
        Else
            shortFilename = "Graph_" & Replace(Replace(Replace(Date.Now, "/", "-"), ":", "."), " ", "_")

        End If

        Chart1.Printing.PrintDocument.DocumentName = shortFilename
        Chart1.Printing.Print(True) ' print with a dialog box

    End Sub

    Private Sub tmr_DataLogger_Tick(sender As System.Object, e As System.EventArgs) Handles tmr_DataLogger.Tick
        ' update the textbox data
        LogTempUpdate(Oven_Temp)
        LogTimeUpdate(Elapsed_RunSec)
        DataPointsCnt = DataPointsCnt + 1
        ' update the graph data 
        Chart1.Series.Clear()
        '  s1.Name = "Temp"
        s1.ChartType = SeriesChartType.Line
        s1.Points.AddXY(Elapsed_RunSec, Oven_Temp)
        Chart1.Series.Add(s1)
    End Sub

    Private Sub btn_ClearGraph_Click(sender As System.Object, e As System.EventArgs) Handles btn_ClearGraph.Click
        Chart1.Series.Clear()
        rtxt_LoggedTempValues.Clear()
        rtxt_loggedTimeValues.Clear()
        RecordedMaxTemp = 0
        DataPointsCnt = 0
        lbl_MaxTempReached.Text = "0"
        lbl_TotalDataPoints.Text = "0"
    End Sub

    Private Sub LogTempUpdate(ByVal StatText As String)
        ' auto scroll the textbox
        rtxt_LoggedTempValues.AppendText(StatText & vbNewLine)
        rtxt_LoggedTempValues.SelectionStart = Len(rtxt_LoggedTempValues.Text)
        rtxt_LoggedTempValues.ScrollToCaret()
        rtxt_LoggedTempValues.Select()

    End Sub

    Private Sub LogTimeUpdate(ByVal StatText As String)
        ' auto scroll the textbox
        rtxt_loggedTimeValues.AppendText(StatText & vbNewLine)
        rtxt_loggedTimeValues.SelectionStart = Len(rtxt_loggedTimeValues.Text)
        rtxt_loggedTimeValues.ScrollToCaret()
        rtxt_loggedTimeValues.Select()

    End Sub

    Private Sub btn_ExportCSV_Click(sender As System.Object, e As System.EventArgs) Handles btn_ExportCSV.Click
        Try
            '------------------------------------------------------------------------------------------------------------

            Dim cnt As Integer = 0

            If txt_CSVDirectory.Text = "" Then txt_CSVDirectory.Text = ProgramLibaryPath & "CSV Directory\"

            ' create a text file 
            If Directory.Exists(txt_CSVDirectory.Text) = True Then
                'look for the directory path in the cvs directory path textbox. if there use it if not create it
            Else
                Try
                    Directory.CreateDirectory(txt_CSVDirectory.Text)
                Catch ex As Exception
                    MsgBox("EXPORT ABORTED" & vbNewLine & "Could not create directory at " & txt_CSVDirectory.Text & vbNewLine & "Check for errors and try again", MsgBoxStyle.Critical)
                    Exit Sub
                End Try

            End If

            ' create a text file for this report under the job directory with the chassis number as a filename
            Dim shortFilename As String = ""

            If GraphCreatedby = False Then ' false = Program editor, true = program designer
                ' if the profile editor current program has a name give the graph that name otherwise just the date
                If txt_PofileEditor_SaveAs.Text <> "" Then
                    shortFilename = "Data_Points_" & Trim(txt_PofileEditor_SaveAs.Text) & "_" & Replace(Replace(Replace(Date.Now, "/", "-"), ":", "."), " ", "_")
                Else
                    shortFilename = "Data_Points_" & Replace(Replace(Replace(Date.Now, "/", "-"), ":", "."), " ", "_")

                End If

            Else
                ' if profile designer 
                'TODO write code for program designer
            End If

            Dim filename As String

            If txt_CSVDirectory.Text <> "" Then
                ' if the directory is populated then use that otherwise use the default 
                filename = Trim(txt_CSVDirectory.Text) & shortFilename & ".csv"
            Else
                txt_CSVDirectory.Text = ProgramLibaryPath
                filename = ProgramLibaryPath & shortFilename & ".csv"
            End If

            If rtxt_LoggedTempValues.Text = "" Or rtxt_loggedTimeValues.Text = "" Then ' make sure there is data in the textboxes 
                MsgBox("ERROR: No data to write to file", MsgBoxStyle.Exclamation)
                Exit Sub
            End If

            If System.IO.File.Exists(filename) = False Then

                '----------
                Dim swriter As StreamWriter

                Try
                    swriter = File.AppendText(filename)
                    ' here we will write each temp value, a "," then time value 
                    swriter.WriteLine("Time,Temp")
                    While cnt < rtxt_LoggedTempValues.Lines.Length
                        swriter.WriteLine(rtxt_loggedTimeValues.Lines(cnt) & "," & rtxt_LoggedTempValues.Lines(cnt))
                        cnt = cnt + 1

                    End While
                    swriter.Close()
                Catch ex As Exception
                    Label121.Text = "Failed to Save" ' show a failed to save
                    Label21.ForeColor = Color.Red
                    Label121.Visible = True
                    MsgBox(ex.ToString)
                Finally
                    Label121.Text = "Saved Successfully"
                    Label21.ForeColor = Color.Green
                    Label121.Visible = True ' show the saved successfully label shortly
                End Try
                '------------
            Else
                Try
                    File.Delete(filename) ' delete first, wait, then make a new one 
                    delay(100)
                    Dim swriter As StreamWriter
                    swriter = File.AppendText(filename)
                    swriter.WriteLine("Time,Temp")
                    While cnt < rtxt_LoggedTempValues.Lines.Length
                        swriter.WriteLine(rtxt_loggedTimeValues.Lines(cnt) & "," & rtxt_LoggedTempValues.Lines(cnt))
                        cnt = cnt + 1

                    End While

                    swriter.Close()
                Catch ex As Exception
                    lbl_CSVsaveStat.Text = "Failed to Save" ' show a failed to save
                    Label21.ForeColor = Color.Red
                    lbl_CSVsaveStat.Visible = True
                Finally
                    lbl_CSVsaveStat.Text = "Saved Successfully" ' show a failed to save
                    Label21.ForeColor = Color.Green
                    lbl_CSVsaveStat.Visible = True ' show the saved successfully label shortly
                End Try



            End If
            '------------------------------------------------------------------------------------------------------------
        Catch ex As Exception
            If Debug = True Then
                ExceptionCnt = ExceptionCnt + 1
                MsgBox(ex.ToString)
            End If
        End Try
    End Sub

    Private Sub btn_BrowseforCSVDir_Click_1(sender As System.Object, e As System.EventArgs) Handles btn_BrowseforCSVDir.Click
        If Directory.Exists(ProgramLibaryPath & "CSV Directory\") Then
            OpenFileDialog1.InitialDirectory = ProgramLibaryPath & "CSV Directory\"
        Else
            OpenFileDialog1.InitialDirectory = "C:\"
        End If

        OpenFileDialog1.FileName = ""
        'OpenFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        OpenFileDialog1.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "OpenFileDialog1" Then txt_CSVDirectory.Text = OpenFileDialog1.FileName
    End Sub

    ' end Data Logger region
    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
#End Region

    Private Sub chk_Topmost_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chk_Topmost.CheckedChanged
        If chk_Topmost.Checked = True Then
            Me.TopMost = True
        Else
            Me.TopMost = False
        End If
    End Sub


   
    Private Sub btn_SendFeedback_Click(sender As System.Object, e As System.EventArgs) Handles btn_SendFeedback.Click
        If txt_OtherSubjectText.Text = "" Then
            MsgBox(" You must enter a subject")
            Exit Sub
        End If

        If rtxt_FeedbackMsg.Text = "" Then
            MsgBox("You must enter a message body")
            Exit Sub
        End If
        btn_SendFeedback.Enabled = False
        lbl_Thankyou.Visible = True
        lbl_Thankyou.Text = "Sending message, please wait..."
        Me.Refresh()
        SendFeedback() ' send the message 
        lbl_Thankyou.Text = "Sent succesfully. Thank you"
        btn_SendFeedback.Enabled = True
    End Sub

    Private Sub Label60_Click(sender As System.Object, e As System.EventArgs) Handles Label60.Click
        If Debug = True Then
            Debug = False
            lbl_exceptionlabel.Visible = False
            lbl_ExceptionCnt.Visible = False
        Else
            Debug = True
            lbl_exceptionlabel.Visible = True
            lbl_ExceptionCnt.Visible = True
        End If
    End Sub
End Class ' END CODE SEGMENT
