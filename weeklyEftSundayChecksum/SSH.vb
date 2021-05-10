Module SSH
    Function getJackpotFromSSH(ByVal IPHostStr As String, ByVal userNameHostStr As String, ByVal passwordHostStr As String, ByVal productSysCodeInt As Integer, ByVal toKnowDate As Date, ByRef errCodeInt As Integer, ByRef errMessageStr As String) As Double


        Dim ssh As New Chilkat.Ssh()
        Dim elfFileStr As String
        Dim port As Integer
        Dim channelNum As Integer
        Dim termType As String = "xterm"
        Dim widthInChars As Integer = 80
        Dim heightInChars As Integer = 24
        Dim pixWidth As Integer = 0
        Dim pixHeight As Integer = 0
        Dim cmdOutputStr As String = ""
        Dim msgError As String = ""
        Dim n As Integer
        Dim pollTimeoutMs As Integer = 2000
        Dim posInic As Integer
        Dim jackpotDbl As Double
        Dim yearStr, monthStr, daystr As String

        Dim success As Boolean = ssh.UnlockComponent("Anything for 30-day trial")

        If (success <> True) Then
            errCodeInt = -1
            errMessageStr = "Error related to component license: " & ssh.LastErrorText
            'MsgBox("Se presento un error #1. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "Error")
            Exit Function
        End If

        'hostname = "10.1.5.103" 'NYB2BAPP1
        'hostname = "10.1.5.10" 'NYESTE1
        'hostname = "10.1.5.12" 'NYESTE3
        port = 22

        success = ssh.Connect(IPHostStr, port)
        If (success <> True) Then
            errCodeInt = -2
            errMessageStr = "Error related to the conection: " & ssh.LastErrorText

            'MsgBox("No Funciono Conneccion. " + ssh.LastErrorText, MsgBoxStyle.OkOnly, "NO funciono coneccion")
            Exit Function
            'Else
            '    MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
        End If


        success = ssh.AuthenticatePw(userNameHostStr, passwordHostStr)
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


        errCodeInt = sentStringToSSH(ssh, "cd elog/files", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        If errCodeInt <> 0 Then
            errCodeInt = -9
            errMessageStr = ssh.LastErrorText
            'MsgBox(msgError, MsgBoxStyle.OkOnly, "Error")
            Exit Function
            'Else
            '    RichTextBox1.Text = cmdOutputStr
        End If

        yearStr = Trim(Str(Year(toKnowDate)))
        monthStr = Trim(Str(Month(toKnowDate))).PadLeft(2, "0")
        daystr = Trim(Str(DateAndTime.Day(toKnowDate))).PadLeft(2, "0")
        elfFileStr = "elf" & yearStr & monthStr & daystr & ".fil"

        Select Case productSysCodeInt

            Case 8
                errCodeInt = sentStringToSSH(ssh, "grep jackpot " & elfFileStr & " | grep loto", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
            Case 12
                errCodeInt = sentStringToSSH(ssh, "grep jackpot " & elfFileStr & " | grep bigg", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
            Case 15
                errCodeInt = sentStringToSSH(ssh, "grep jackpot " & elfFileStr & " | grep pwrb", channelNum, pollTimeoutMs, cmdOutputStr, msgError)
        End Select

        If errCodeInt <> 0 Then
            errCodeInt = -10
            errMessageStr = msgError
            Exit Function
        Else
            ssh.Disconnect()
            'posInic = InStr(1, cmdOutputStr, "$", CompareMethod.Text)
            posInic = InStrRev(cmdOutputStr, "$")
            'RichTextBox1.Text = cmdOutputStr & "===" & posInic
            If posInic = 0 Then
                jackpotDbl = 0
            Else
                'posFin = InStr(1, cmdOutputStr, ".", CompareMethod.Text)
                jackpotDbl = CDbl(Mid(cmdOutputStr, posInic + 1, 10))

            End If

            Return jackpotDbl

        End If

    End Function

    Function sentStringToSSH(ByVal ssh As Chilkat.Ssh, ByVal strText As String, ByVal channelNum As Integer, ByVal pollTimeoutMs As Integer, ByRef cmdOutputStr As String, ByRef msgError As String) As Integer
        Dim success As Boolean
        Dim n As Integer

        success = ssh.ChannelSendString(channelNum, strText & vbLf, "utf-8")
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
    Function SSHUntilMatch(ByVal ssh As Chilkat.Ssh, ByVal matchPattern As String, ByVal channelNum As Integer, ByVal pollTimeoutMs As Integer, ByRef cmdOutputStr As String, ByRef msgError As String) As Integer
        Dim success As Boolean


        'ssh.ReadTimeoutMs = 10000
        success = ssh.ChannelReceiveUntilMatch(channelNum, matchPattern, "utf-8", False)
        If (success <> True) Then
            msgError = "ChannelReceiveUntilMatch Error: " + ssh.LastErrorText
            Return -2
        End If


        'n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        'If (n < 0) Then
        '    msgError = "ChannelReadAndPoll Error: " + ssh.LastErrorText
        '    Return -2
        'End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            msgError = "GetReceivedText Error: " + ssh.LastErrorText
            Return -3
        Else
            Return 0
        End If

    End Function



    Function sentStringToSSHUntilMatch(ByVal ssh As Chilkat.Ssh, ByVal strText As String, ByVal matchPattern As String, ByVal channelNum As Integer, ByVal pollTimeoutMs As Integer, ByRef cmdOutputStr As String, ByRef msgError As String) As Integer
        Dim success As Boolean
        Dim n As Integer

        success = ssh.ChannelSendString(channelNum, strText & vbLf, "utf-8")
        If (success <> True) Then
            msgError = "ChannelSendString Error: " + ssh.LastErrorText
            Return -1
        End If


        'ssh.ReadTimeoutMs = 10000
        success = ssh.ChannelReceiveUntilMatch(channelNum, matchPattern, "utf-8", False)
        If (success <> True) Then
            msgError = "ChannelReceiveUntilMatch Error: " + ssh.LastErrorText
            Return -2
        End If


        'n = ssh.ChannelReadAndPoll(channelNum, pollTimeoutMs)
        'If (n < 0) Then
        '    msgError = "ChannelReadAndPoll Error: " + ssh.LastErrorText
        '    Return -2
        'End If

        cmdOutputStr = ssh.GetReceivedText(channelNum, "utf-8")
        If (ssh.LastMethodSuccess <> True) Then
            msgError = "GetReceivedText Error: " + ssh.LastErrorText
            Return -3
        Else
            Return 0
        End If

    End Function



End Module
