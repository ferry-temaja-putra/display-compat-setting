Imports Microsoft.Win32
Imports System.Reflection
Imports System.Management

Module Module1

    Private Const DPISettingValue = " GDIDPISCALING DPIUNAWARE"
    Private Const CompatibilitySettingKeyName = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
    Private ReadOnly AppFileName As String = Assembly.GetExecutingAssembly.Location
    Private Const Windows10Caption = "Windows 10"
    Private Const FallCreatorsUpdateBuildNumber = 16299
    Private Const HIGHDPIAWARESetting = " HIGHDPIAWARE"
    Private Const DPIUNAWARESetting = " DPIUNAWARE"

    Sub Main()

        If Not IsOverrideHighDPIScalingSupported() Then
            Console.ReadLine()
            Return
        End If

        If IsHighDPISettingConfigured() Then
            Console.ReadLine()
            Return
        End If

        ConfigureDPISetting()

        Console.ReadLine()

    End Sub

    Private Function IsOverrideHighDPIScalingSupported() As Boolean

        Console.WriteLine("Checking OS version...")

        Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")

        Dim osCaption = String.Empty
        Dim osVersion = String.Empty
        Dim osBuildNumber = String.Empty

        For Each result In searcher.Get
            osCaption = CType(result("Caption"), String)
            osBuildNumber = CType(result("BuildNumber"), String)
            osVersion = CType(result("Version"), String)
            Exit For
        Next

        Console.WriteLine("OS: {0} {1} {2}", osCaption, osBuildNumber, osVersion)
        Console.WriteLine()

        Dim isSupported =
            osCaption.Contains(Windows10Caption) AndAlso
            CType(osBuildNumber, Integer) >= FallCreatorsUpdateBuildNumber

        If Not isSupported Then
            Console.WriteLine("Current OS does not support overriding high dpi scaling setting.")
        End If

        Return isSupported

    End Function

    Private Function IsHighDPISettingConfigured() As Boolean

        Console.WriteLine("Checking High DPI Scaling setting...")

        Dim currentValue = CType(Registry.GetValue(CompatibilitySettingKeyName, AppFileName, String.Empty), String)

        If String.IsNullOrEmpty(currentValue) OrElse Not currentValue.Contains(DPISettingValue) Then
            Console.WriteLine("High DPI Scaling is not configured.")
            Return False
        End If

        Console.WriteLine("High DPI Scaling is already configured.")
        Return True

    End Function

    Private Sub ConfigureDPISetting()

        Console.WriteLine("Configuring High DPI Scaling...")

        Dim value = CType(Registry.GetValue(CompatibilitySettingKeyName, AppFileName, String.Empty), String)

        If String.IsNullOrEmpty(value) Then
            value = "~"
        End If

        ' Remove DPI related setting
        value = value.Replace(HIGHDPIAWARESetting, String.Empty)
        value = value.Replace(DPIUNAWARESetting, String.Empty)

        value &= DPISettingValue

        Registry.SetValue(CompatibilitySettingKeyName, AppFileName, value)

        Console.WriteLine("High DPI Scaling is configured.")

    End Sub

End Module
