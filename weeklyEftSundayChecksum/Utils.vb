
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Net
Imports Microsoft.Win32


Module Utils
    Public appName As String = "SETTINGS LOCAL APPS"

    Public Function GetSetting(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Return Key.GetValue(Keyname, DefVal)


        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSetting(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME)
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Sub SalvarSettingConfiServerDB(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigServerDB(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string database")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function

    Public Sub SalvarSettingConfiHost(ByVal APP_NAME As String, ByVal Keyname As String, ByVal Value As String)
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string host")
            Key.SetValue(Keyname, Value)
        Catch
            Return
        End Try
    End Sub

    Public Function GetSettingConfigHost(ByVal APP_NAME As String, ByVal Keyname As String, Optional ByVal DefVal As String = "") As String
        Dim Key As RegistryKey
        Try
            Key = Registry.CurrentUser.CreateSubKey("SOFTWARE").CreateSubKey("IGT").CreateSubKey(APP_NAME).CreateSubKey("Conection string host")
            Return Key.GetValue(Keyname, DefVal)
        Catch
            Return DefVal
        End Try
    End Function



End Module
