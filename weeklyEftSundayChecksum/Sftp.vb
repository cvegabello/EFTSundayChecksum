
Module Sftp

    Function downloadFileSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal remoteFile As String, ByVal localFile As String, ByRef strError As String) As Integer

        Dim sftp As New Chilkat.SFtp()
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        Dim handle As String

        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 1
            'MsgBox("Error", MsgBoxStyle.OkOnly, "Se presento un error #1. " + sftp.LastErrorText)
            'Console.WriteLine(sftp.LastErrorText)

        End If


        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.
        'Dim port As Integer
        'Dim hostname As String

        'hostname = "10.2.5.14"
        'port = 22
        success = sftp.Connect(hostName, port)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 2

            'MsgBox("NO funciono coneccion", MsgBoxStyle.OkOnly, "No Funciono Conneccion. " + sftp.LastErrorText)
            'Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
            'Console.WriteLine("Exito")
        End If

        'success = sftp.AuthenticatePw("prosys", "Numb3r1j0b")
        'MsgBox(username + " " + password, MsgBoxStyle.OkOnly, "Debug Carlos")
        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 3


            'MsgBox("NO funciono Autenticacion", MsgBoxStyle.OkOnly, "No Funciono Autenticacion. " + sftp.LastErrorText)
            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 4

            'MsgBox("NO funciono Inicializacion", MsgBoxStyle.OkOnly, "No Funciono Inicializacion. " + sftp.LastErrorText)
            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else

            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")

        End If


        handle = sftp.OpenFile(remoteFile, "readOnly", "openExisting")
        If (handle = vbNullString) Then

            strError = sftp.LastErrorText
            Return 5

            'MsgBox("NO funciono abrir el archivo", MsgBoxStyle.OkOnly, "No Funciono abrir el archivo. " + sftp.LastErrorText)
            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("Funciono abrir el archivo", MsgBoxStyle.OkOnly, "Funciono abrir el archivo. ")

        End If


        '  Download the file:
        success = sftp.DownloadFile(handle, localFile)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 6

            'MsgBox("No Funciono bajar el archivo", MsgBoxStyle.OkOnly, "No Funciono bajar el archivo. " + sftp.LastErrorText)
            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("Funciono bajar el archivo", MsgBoxStyle.OkOnly, "Funciono bajar el archivo. ")

        End If


        '  Close the file.
        success = sftp.CloseHandle(handle)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 7

            'MsgBox("No Funciono cerrar el archivo", MsgBoxStyle.OkOnly, "No Funciono cerrar el archivo. " + sftp.LastErrorText)

            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("Funciono cerrar el archivo", MsgBoxStyle.OkOnly, "Funciono cerrar el archivo. ")
        End If
        Return 0

        'MsgBox("Funciono todo", MsgBoxStyle.OkOnly, "EXITO TOTAL")


    End Function

    Function upLoadFileSFTP(ByVal hostName As String, ByVal port As Integer, ByVal username As String, ByVal password As String, ByVal remoteFile As String, ByVal localFile As String, ByRef strError As String) As Integer

        Dim sftp As New Chilkat.SFtp()
        Dim success As Boolean = sftp.UnlockComponent("Anything for 30-day trial")

        'Dim handle As String

        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 1
            'MsgBox("Error", MsgBoxStyle.OkOnly, "Se presento un error #1. " + sftp.LastErrorText)
            'Console.WriteLine(sftp.LastErrorText)

        End If


        '  Set some timeouts, in milliseconds:
        sftp.ConnectTimeoutMs = 5000
        sftp.IdleTimeoutMs = 10000

        '  Connect to the SSH server.
        '  The standard SSH port = 22
        '  The hostname may be a hostname or IP address.
        'Dim port As Integer
        'Dim hostname As String

        'hostname = "10.2.5.14"
        'port = 22
        success = sftp.Connect(hostName, port)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 2

            'MsgBox("NO funciono coneccion", MsgBoxStyle.OkOnly, "No Funciono Conneccion. " + sftp.LastErrorText)
            'Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("EXITO Coneccion", MsgBoxStyle.OkOnly, "Funciono Conneccion")
            'Console.WriteLine("Exito")
        End If

        'success = sftp.AuthenticatePw("prosys", "Numb3r1j0b")
        'MsgBox(username + " " + password, MsgBoxStyle.OkOnly, "Debug Carlos")
        success = sftp.AuthenticatePw(username, password)
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 3


            'MsgBox("NO funciono Autenticacion", MsgBoxStyle.OkOnly, "No Funciono Autenticacion. " + sftp.LastErrorText)
            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else
            'MsgBox("Funciono Autenticacion", MsgBoxStyle.OkOnly, "Funciono Autenticacion. ")
        End If


        '  After authenticating, the SFTP subsystem must be initialized:
        success = sftp.InitializeSftp()
        If (success <> True) Then

            strError = sftp.LastErrorText
            Return 4

            'MsgBox("NO funciono Inicializacion", MsgBoxStyle.OkOnly, "No Funciono Inicializacion. " + sftp.LastErrorText)
            ''Console.WriteLine(sftp.LastErrorText)
            'Exit Function
        Else

            'MsgBox("Funciono Inicializacion", MsgBoxStyle.OkOnly, "Funciono Inicializacion. ")

        End If


        success = sftp.UploadFileByName(remoteFile, localFile)
        If (success <> True) Then
            strError = sftp.LastErrorText
            Return 5
        End If




        'handle = sftp.OpenFile(remoteFile, "readOnly", "openExisting")
        'If (handle = vbNullString) Then

        '    strError = sftp.LastErrorText
        '    Return 5

        '    'MsgBox("NO funciono abrir el archivo", MsgBoxStyle.OkOnly, "No Funciono abrir el archivo. " + sftp.LastErrorText)
        '    ''Console.WriteLine(sftp.LastErrorText)
        '    'Exit Function
        'Else
        '    'MsgBox("Funciono abrir el archivo", MsgBoxStyle.OkOnly, "Funciono abrir el archivo. ")

        'End If


        ''  Download the file:
        'success = sftp.DownloadFile(handle, localFile)
        'If (success <> True) Then

        '    strError = sftp.LastErrorText
        '    Return 6

        '    'MsgBox("No Funciono bajar el archivo", MsgBoxStyle.OkOnly, "No Funciono bajar el archivo. " + sftp.LastErrorText)
        '    ''Console.WriteLine(sftp.LastErrorText)
        '    'Exit Function
        'Else
        '    'MsgBox("Funciono bajar el archivo", MsgBoxStyle.OkOnly, "Funciono bajar el archivo. ")

        'End If


        '  Close the file.
        'success = sftp.CloseHandle(handle)
        'If (success <> True) Then

        '    strError = sftp.LastErrorText
        '    Return 7

        '    'MsgBox("No Funciono cerrar el archivo", MsgBoxStyle.OkOnly, "No Funciono cerrar el archivo. " + sftp.LastErrorText)

        '    ''Console.WriteLine(sftp.LastErrorText)
        '    'Exit Function
        'Else
        '    'MsgBox("Funciono cerrar el archivo", MsgBoxStyle.OkOnly, "Funciono cerrar el archivo. ")
        'End If
        Return 0

        'MsgBox("Funciono todo", MsgBoxStyle.OkOnly, "EXITO TOTAL")


    End Function





End Module
