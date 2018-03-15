Imports Microsoft.Win32
Imports System.Reflection
Imports System.Management

Module Module1

    Private Const DPISettingValue = "~ GDIDPISCALING DPIUNAWARE"
    Private Const CompatibilitySettingKeyName = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
    Private ReadOnly AppFileName = Assembly.GetExecutingAssembly.Location
    Private Const WindowsCaption = "Windows 10"
    Private Const WindowsVersion = ""

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

    Private Function IsOverrideHighDPIScalingSupported()

        Console.WriteLine("Checking OS version...")

        Dim searcher As New ManagementObjectSearcher("SELECT Caption, Version FROM Win32_OperatingSystem")

        For Each result In searcher.Get
            Console.WriteLine("OS Version: {0} {1}", result("Caption"), result("Version"))
        Next

        Console.WriteLine()

        Return True
    End Function

    Private Function IsHighDPISettingConfigured()

        Console.WriteLine("Checking High DPI Scaling setting...")

        Dim currentValue = CType(Registry.GetValue(CompatibilitySettingKeyName, AppFileName, String.Empty), String)

        If String.IsNullOrEmpty(currentValue) OrElse currentValue <> DPISettingValue Then
            Console.WriteLine("High DPI Scaling is not configured.")
            Return False
        End If

        Console.WriteLine("High DPI Scaling is already configured.")
        Return True

    End Function

    Private Sub ConfigureDPISetting()

        Console.WriteLine("Configuring High DPI Scaling...")

        Registry.SetValue(CompatibilitySettingKeyName, AppFileName, DPISettingValue)

        Console.WriteLine("High DPI Scaling is configured.")

    End Sub

End Module
