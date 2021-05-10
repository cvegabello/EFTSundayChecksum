Imports System.Net.Mail
Imports Microsoft.Win32

Public Class Form1
    Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
    Public flagBool As Boolean
    Public appNameRecipients As String = "Recipients EFTSunday"


    Sub getHostModeNew(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt, numInt, isIcludedInt, spareInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Sub
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Sub
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Sub
            'Else
            '    MsgBox("Funciono OpenSessionChannel", MsgBoxStyle.OkOnly, "Funciono OpenSessionChannel. ")

        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Sub
            'Else
            '    MsgBox("Funciono SendReqPty", MsgBoxStyle.OkOnly, "Funciono SendReqPty. ")

        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Sub
            'Else
            '    MsgBox("Funciono SendReqShell", MsgBoxStyle.OkOnly, "Funciono SendReqShell. ")
        End If


        errCodeInt = SSHUntilMatch(ssh, "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "10", "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 10 ", MsgBoxStyle.OkOnly, "FUNCIONO.")
        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "01", "Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Sub
            'Else
            '    MsgBox("FUNCIONO: 01 ", MsgBoxStyle.OkOnly, "FUNCIONO.")


        End If

        errCodeInt = sentStringToSSH(ssh, "PEERS", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        ssh.Disconnect()


        posInt = InStr(1, cmdOutputStr, "live_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 10, 2))
        hostNumberArray(0) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "backup_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 11, 2))
        hostNumberArray(1) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "wait_for_spare:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 16, 2))
        spareInt = numInt + 1

        If ((spareInt = hostNumberArray(0)) Or (spareInt = hostNumberArray(1))) Then

            For i = 1 To 5
                isIcludedInt = Array.IndexOf(hostNumberArray, i)
                If isIcludedInt < 0 Then
                    hostNumberArray(2) = i
                    Exit For
                End If
            Next
        Else
            hostNumberArray(2) = spareInt

        End If


        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(3) = i
                Exit For
            End If
        Next

        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(4) = i
                Exit For
            End If
        Next

    End Sub



    Sub getHostMode(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal hostNumberArray() As Integer, ByRef errCodeInt As Integer, ByRef errMessageStr As String)


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum, posInt, numInt, isIcludedInt, spareInt As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Sub
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Sub
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Sub
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Sub
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Sub
        End If


        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            'MsgBox("NO funciono ChannelReadAndPoll " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono ChannelReadAndPoll. ")
            Exit Sub
        End If


        cmdOutputStr = ssh.GetReceivedText(channelNum, "ansi")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            'MsgBox("NO funciono GetReceivedText " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono GetReceivedText. ")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If


        errCodeInt = sentStringToSSH(ssh, "cd gtms/bin", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If


        errCodeInt = sentStringToSSH(ssh, "gxvision", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        errCodeInt = sentStringToSSH(ssh, "PEERS", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Sub
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        ssh.Disconnect()


        posInt = InStr(1, cmdOutputStr, "live_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 10, 2))
        hostNumberArray(0) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "backup_idn:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 11, 2))
        hostNumberArray(1) = numInt + 1

        posInt = InStr(1, cmdOutputStr, "wait_for_spare:")
        numInt = CInt(Mid(cmdOutputStr, posInt + 16, 2))
        spareInt = numInt + 1

        If ((spareInt = hostNumberArray(0)) Or (spareInt = hostNumberArray(1))) Then

            For i = 1 To 5
                isIcludedInt = Array.IndexOf(hostNumberArray, i)
                If isIcludedInt < 0 Then
                    hostNumberArray(2) = i
                    Exit For
                End If
            Next
        Else
            hostNumberArray(2) = spareInt

        End If


        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(3) = i
                Exit For
            End If
        Next

        For i = 1 To 5
            isIcludedInt = Array.IndexOf(hostNumberArray, i)
            If isIcludedInt < 0 Then
                hostNumberArray(4) = i
                Exit For
            End If
        Next



    End Sub

    Function getWeekly_eftFromSSH_NEW(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal fecha As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As String


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim fileNameLog, yearStr, monthStr, dayStr, commandStr As String
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Function
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Function
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Function
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Function
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Function
        End If


        errCodeInt = SSHUntilMatch(ssh, "Enter Selection:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
            'Else
            '    MsgBox("FUNCIONO: ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "98", "return to the menu):", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
            'Else
            '    MsgBox("FUNCIONO: 10 ", MsgBoxStyle.OkOnly, "FUNCIONO.")
        End If

        errCodeInt = sentStringToSSHUntilMatch(ssh, "sudo su - prosys", "Password:", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
            'Else
            '    MsgBox("FUNCIONO: 01 ", MsgBoxStyle.OkOnly, "FUNCIONO.")

        End If



        errCodeInt = sentStringToSSH(ssh, password, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If

        'errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        errCodeInt = sentStringToSSH(ssh, "cd elog/files", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")

        fileNameLog = "elf" & yearStr & monthStr & dayStr & ".fil"

        commandStr = "grep acct:weekly_eft " & fileNameLog

        errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        ssh.Disconnect()

        Return cmdOutputStr


    End Function




    Function getWeekly_eftFromSSH(ByVal hostName As String, ByVal userName As String, ByVal password As String, ByVal fecha As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As String


        Dim ssh As New Chilkat.Ssh()
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim fileNameLog, yearStr, monthStr, dayStr, commandStr As String
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        port = 22

        success = ssh.Connect(hostName, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Function
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userName, password)
        If (success <> True) Then
            errCodeInt = -3
            errMessageStr = "Error related to the authentication: " & ssh.LastErrorText

            'MsgBox("No Funciono Autenticacion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Autenticacion")
            Exit Function
            'Else
            '    MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If

        channelNum = ssh.OpenSessionChannel()
        If (channelNum < 0) Then
            errCodeInt = -4
            errMessageStr = "Error related to OpenSessionChannel: " & ssh.LastErrorText

            'MsgBox("No Funciono Apertura Canal. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Apertura del canal")
            Exit Function
        End If

        success = ssh.SendReqPty(channelNum, termType, widthInChars, heightInChars, pixWidth, pixHeight)
        If (success <> True) Then
            errCodeInt = -5
            errMessageStr = "Error related to SendReqPty: " & ssh.LastErrorText
            'MsgBox("NO funciono Sendreqpty " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono Sendreqpty. ")
            Exit Function
        End If

        success = ssh.SendReqShell(channelNum)
        If (success <> True) Then
            errCodeInt = -6
            errMessageStr = "Error related to SendReqShell: " & ssh.LastErrorText
            'MsgBox("NO funciono SendReqShell " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono SendReqShell. ")
            Exit Function
        End If


        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            errCodeInt = -7
            errMessageStr = "Error related to ChannelReadAndPoll: " & ssh.LastErrorText
            'MsgBox("NO funciono ChannelReadAndPoll " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono ChannelReadAndPoll. ")
            Exit Function
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            errCodeInt = -8
            errMessageStr = "Error related to GetReceivedText: " & ssh.LastErrorText
            'MsgBox("NO funciono GetReceivedText " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono GetReceivedText. ")
            Exit Function
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        'errCodeInt = sentStringToSSH(ssh, "df", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        errCodeInt = sentStringToSSH(ssh, "cd elog/files", channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")

        fileNameLog = "elf" & yearStr & monthStr & dayStr & ".fil"

        commandStr = "grep acct:weekly_eft " & fileNameLog

        errCodeInt = sentStringToSSH(ssh, commandStr, channelNum, pollTimeoutMs, cmdOutputStr, msgError)

        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            Exit Function
        End If


        ssh.Disconnect()

        Return cmdOutputStr


    End Function

    Function sentStringToSSH(ByVal ssh As Chilkat.Ssh, ByVal strText As String, ByVal channelNum As Integer, ByVal pollTimeoutMs As Integer, ByRef cmdOutputStr As String, ByRef msgError As String) As Integer
        Dim success As Boolean
        Dim n As Integer

        success = ssh.ChannelSendString(channelNum, strText & vbCrLf, "utf-8")
        If (success <> True) Then
            msgError = "ChannelSendString Error: " + ssh.LastErrorText
            Return -1
        End If

        n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        If (n < 0) Then
            msgError = "ChannelReadAndPoll Error: " + ssh.LastErrorText
            Return -2
        End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            msgError = "GetReceivedText Error: " + ssh.LastErrorText
            Return -3
        Else
            Return 0
        End If

    End Function

    Private Sub okBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles okBtn.Click

        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim file As System.IO.StreamWriter
        Dim infoStr As String = ""
        Dim fecha As Date
        Dim posInt As Integer
        Dim monthStr, dayStr, yearStr, infoCutStr, strPathRemoteFile, fileNameStr, strFilePathLocal As String
        Dim hostNumberArray(5), dayInt, resCode As Integer
        Dim conStrinHostStr As String
        Dim substrings(), fechaStr, recipientsStr, errorStr As String



        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")

        dayInt = Weekday(fecha)

        If dayInt <> 1 Then

            MsgBox("The date selected must be SUNDAY", MsgBoxStyle.OkOnly, "Error")

        Else
            Me.okBtn.Visible = False
            Me.cancelBtn.Enabled = False
            Button1.Visible = False
            Me.Cursor = Cursors.WaitCursor


            Me.Height = 212
            Button1.Text = "View/Edit email recipients ..."
            flagBool = False

            monthStr = CStr(fecha.Month).PadLeft(2, "0")
            dayStr = CStr(fecha.Day).PadLeft(2, "0")
            yearStr = CStr(fecha.Year).PadLeft(4, "0")

            getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCodeInt, errMessageStr)

            If errCodeInt <> 0 Then
                MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                Me.Dispose()
            End If

            Me.RichTextBox1.Text = "Getting weekly_eft_checksum from Primary..." & vbCrLf
            Me.Refresh()


            Select Case hostNumberArray(0)
                Case 1
                    infoStr = getWeekly_eftFromSSH_NEW(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 2

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select

            infoCutStr = Trim(Mid(infoStr, 40))
            posInt = InStr(infoCutStr, "End of task")
            infoCutStr = Trim(Mid(infoCutStr, 1, posInt + 34))
            fileNameStr = "weekly_host_eft_checksum_" & monthStr & dayStr & yearStr & ".txt"

            file = My.Computer.FileSystem.OpenTextFileWriter(fileNameStr, False)
            file.WriteLine(infoCutStr)
            file.Close()


            strPathRemoteFile = "files/weekly_eft_checksum_logFiles"
            strFilePathLocal = System.Windows.Forms.Application.StartupPath
            Me.RichTextBox1.Text = "Uploading to SFTP..." & vbCrLf
            Me.Refresh()


            errCodeInt = upLoadFileSFTP("10.1.5.182", "22", "xfer", "Welcome1", strPathRemoteFile & "/" & fileNameStr, strFilePathLocal & "\" & fileNameStr, errMessageStr)


            Select Case errCodeInt

                Case 1
                    MsgBox("SFTP Component error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub
                Case 2

                    MsgBox("Conection error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 3
                    MsgBox("Authenticate error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 4
                    MsgBox("SFTP Initialize error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub


                Case 5
                    MsgBox("Upload file error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 0

                    fechaStr = Format(DateTimePicker1.Value, "dddd, MMM d yyyy")
                    recipientsStr = Trim(emailTxt.Text)
                    resCode = sent_Email("156.24.14.132", recipientsStr, strFilePathLocal & "\" & fileNameStr, "Weekly EFT Checksum Log: " & fechaStr, "Attached is log file with the weekly EFT checksum for: " & fechaStr, errorStr)

                    If resCode = 0 Then
                        Me.RichTextBox1.Text = "File sent by email." & vbCrLf
                        Me.Refresh()
                        Sleep(3000)
                        Me.Cursor = Cursors.Default
                        Me.RichTextBox1.Text = "All Done."
                        Me.Refresh()
                        Sleep(3000)
                        Me.cancelBtn.Enabled = True

                    Else
                        MessageBox.Show("Error sending the file by email. " & errorStr, "Error sending the file", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

            End Select
            Me.Dispose()


        End If




        'file = My.Computer.FileSystem.OpenTextFileWriter("weekly_eft_checksum_" & monthStr & dayStr & yearStr & ".txt", False)
        'file.WriteLine(infoCutStr)
        'file.Close()

        'Me.Cursor = Cursors.Default
        'Me.RichTextBox1.Text = "All Done."
        'Me.Refresh()
        'Me.cancelBtn.Enabled = True



    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'DateTimePicker1.Value = Now.AddDays(-4)
        DateTimePicker1.Value = Now
        ProgressBar1.Maximum = 100
        okBtn.Visible = False
        cancelBtn.Visible = False
        Button1.Visible = False
        Timer1.Enabled = True

        emailTxt.Text = GetSettingRecipientsEFTSunday(appNameRecipients, "StringRecEFTSunday", "").ToString()

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim errCodeInt As Integer = 0
        Dim errMessageStr As String = ""
        Dim file As System.IO.StreamWriter
        Dim infoStr As String = ""
        Dim fecha As Date
        Dim posInt As Integer
        Dim monthStr, dayStr, yearStr, infoCutStr, strPathRemoteFile, fileNameStr, strFilePathLocal As String
        Dim hostNumberArray(5) As Integer
        Dim conStrinHostStr As String
        Dim substrings() As String



        'substrings(0) -> username
        'substrings(1) -> IP ESTE1
        'substrings(2) -> Password ESTE1
        'substrings(3) -> IP ESTE2
        'substrings(4) -> Password ESTE2
        'substrings(5) -> IP ESTE3
        'substrings(6) -> Password ESTE3
        'substrings(7) -> IP ESTE4
        'substrings(8) -> Password ESTE4
        'substrings(9) -> IP ESTE5
        'substrings(10) -> Password ESTE5

        conStrinHostStr = GetSettingConfigHost(appName, "conStringHost", "").ToString()
        substrings = conStrinHostStr.Split("|")


        fecha = Format(DateTimePicker1.Value, "MM/dd/yyyy")
        monthStr = CStr(fecha.Month).PadLeft(2, "0")
        dayStr = CStr(fecha.Day).PadLeft(2, "0")
        yearStr = CStr(fecha.Year).PadLeft(4, "0")


        Static count As Integer = 0
        count = count + 10
        If count <= 100 Then
            ProgressBar1.Value = count
        Else
            Timer1.Enabled = False
            stopBtn.Visible = False
            ProgressBar1.Visible = False
            okBtn.Visible = False
            cancelBtn.Visible = False
            'getHostMode("10.1.5.12", "prosys", "Numb3r1j0b", hostNumberArray, errCodeInt, errMessageStr)

            getHostModeNew(substrings(3), substrings(0), substrings(4), hostNumberArray, errCodeInt, errMessageStr)

            If errCodeInt <> 0 Then
                MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                Me.Dispose()
            End If

            Me.RichTextBox1.Text = "Getting weekly_eft_checksum from Primary..." & vbCrLf
            Me.Refresh()


            Select Case hostNumberArray(0)
                Case 1
                    infoStr = getWeekly_eftFromSSH_NEW(substrings(1), substrings(0), substrings(2), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 2

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(3), substrings(0), substrings(4), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 3

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(5), substrings(0), substrings(6), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 4

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(7), substrings(0), substrings(8), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
                Case 5

                    infoStr = getWeekly_eftFromSSH_NEW(substrings(9), substrings(0), substrings(10), fecha, errCodeInt, errMessageStr)
                    If errCodeInt <> 0 Then
                        MsgBox(errMessageStr, MsgBoxStyle.OkOnly, "Error")
                    End If
            End Select

            infoCutStr = Trim(Mid(infoStr, 40))
            posInt = InStr(infoCutStr, "End of task")
            infoCutStr = Trim(Mid(infoCutStr, 1, posInt + 34))
            fileNameStr = "weekly_host_eft_checksum_" & monthStr & dayStr & yearStr & ".txt"

            file = My.Computer.FileSystem.OpenTextFileWriter(fileNameStr, False)
            file.WriteLine(infoCutStr)
            file.Close()



            strPathRemoteFile = "files/weekly_eft_checksum_logFiles"
            strFilePathLocal = System.Windows.Forms.Application.StartupPath
            Me.RichTextBox1.Text = "Uploading to SFTP..." & vbCrLf
            Me.Refresh()


            errCodeInt = upLoadFileSFTP("10.1.5.182", "22", "xfer", "Welcome1", strPathRemoteFile & "/" & fileNameStr, strFilePathLocal & "\" & fileNameStr, errMessageStr)


            Select Case errCodeInt

                Case 1
                    MsgBox("SFTP Component error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub
                Case 2

                    MsgBox("Conection error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 3
                    MsgBox("Authenticate error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 4
                    MsgBox("SFTP Initialize error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub


                Case 5
                    MsgBox("Upload file error. " + errMessageStr, MsgBoxStyle.OkOnly, "Error.")
                    Exit Sub

                Case 0

                    Me.Cursor = Cursors.Default
                    Me.RichTextBox1.Text = "All Done."
                    Me.Refresh()


            End Select

            Me.Dispose()


        End If
    End Sub

    Private Sub stopBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stopBtn.Click
        Timer1.Enabled = False
        stopBtn.Visible = False
        ProgressBar1.Visible = False

        okBtn.Visible = True
        cancelBtn.Visible = True
        Button1.Visible = True
    End Sub

    Private Sub cancelBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cancelBtn.Click
        Me.Close()
    End Sub


    Function sent_Email(ByVal smtpHostStr As String, ByVal toStr As String, ByVal attachStr As String, ByVal subjectStr As String, ByVal bodyStr As String, ByRef messErrorStr As String) As Integer
        Dim SMTPserver As New SmtpClient
        Dim mail As New MailMessage
        Dim oAttch As Attachment = New Attachment(attachStr)


        Try
            SMTPserver.Host = smtpHostStr
            mail = New MailMessage
            mail.From = New MailAddress("do.not.reply@gtech-noreply.com")
            mail.To.Add(toStr) 'The Man you want to send the message to him
            mail.Subject = subjectStr
            mail.Body = bodyStr
            mail.Attachments.Add(oAttch)
            SMTPserver.Send(mail)
            Return 0
            'MessageBox.Show("Done!", "Message Sent", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)


        Catch ex As Exception
            messErrorStr = ex.Message
            Return -1

        End Try




    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If flagBool = False Then
            Me.Height = 350
            Button1.Text = "Hide"
            flagBool = True
        Else
            Me.Height = 212
            Button1.Text = "View/Edit email recipients ..."
            flagBool = False

        End If

    End Sub


    Public Sub SalvarSettingRecipientsEFTSunday(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Dim mesError As String
        Try
            Key = Registry.LocalMachine.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Key.SetValue(Keyname, Value)
        Catch ex As Exception
            mesError = ex.Message
            Return
        End Try
    End Sub

    Public Function GetSettingRecipientsEFTSunday(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.LocalMachine.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SalvarSettingRecipientsEFTSunday(appNameRecipients, "StringRecEFTSunday", emailTxt.Text)
    End Sub
End Class
