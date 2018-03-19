Imports Microsoft.Win32
Imports System.Reflection
Imports System.Management
Imports System.Drawing
Imports System.Runtime.InteropServices

Module Module1

    Private Const DPISettingValue = " GDIDPISCALING DPIUNAWARE"
    Private Const CompatibilitySettingKeyName = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"
    Private ReadOnly AppFileName As String = Assembly.GetExecutingAssembly.Location
    Private Const Windows10Caption = "Windows 10"
    Private Const FallCreatorsUpdateBuildNumber = 16299
    Private Const HIGHDPIAWARESetting = " HIGHDPIAWARE"
    Private Const DPIUNAWARESetting = " DPIUNAWARE"
    Private Const DefaultDpi As Single = 96

    <DllImport("gdi32.dll")>
    Private Function GetDeviceCaps(ByVal hdc As IntPtr, ByVal nIndex As Integer) As Integer
    End Function

    Public Enum DeviceCap
        VERTRES = 10
        DESKTOPVERTRES = 117
    End Enum

    Sub Main()

        If Not IsDisplayScaledUp() Then
            Console.WriteLine("Current display is not scaled up. No need to override the display scaling.")
            Console.ReadLine()
            Return
        End If

        If Not IsOverrideHighDPIScalingSupported() Then
            Console.WriteLine("Current OS does not support overriding high dpi scaling setting.")
            Console.ReadLine()
            Return
        End If

        If IsHighDPISettingConfigured() Then
            Console.WriteLine("Override High DPI Scaling is already configured.")
            Console.ReadLine()
            Return
        End If

        ConfigureDPISetting()

        Console.WriteLine("Override High DPI Scaling is configured.")
        Console.ReadLine()

    End Sub

    Private Function IsDisplayScaledUp() As Boolean

        Console.WriteLine("Detecting current scaling...")

        Dim scaleFactor = GetScalingFactor()

        Console.WriteLine("Current display scale: {0}.", scaleFactor)

        Console.WriteLine()

        Return scaleFactor > 1

    End Function

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

        Console.WriteLine("Current OS: {0} {1} {2}.", osCaption, osBuildNumber, osVersion)
        Console.WriteLine()

        Dim isSupported =
            osCaption.Contains(Windows10Caption) AndAlso
            CType(osBuildNumber, Integer) >= FallCreatorsUpdateBuildNumber

        Return isSupported

    End Function

    Private Function IsHighDPISettingConfigured() As Boolean

        Console.WriteLine("Checking High DPI Scaling setting...")

        Dim currentValue = CType(Registry.GetValue(CompatibilitySettingKeyName, AppFileName, String.Empty), String)

        If String.IsNullOrEmpty(currentValue) OrElse Not currentValue.Contains(DPISettingValue) Then
            Console.WriteLine("High DPI Scaling is not configured.")
            Return False
        End If

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

    End Sub

    Private Function GetScalingFactor() As Single

        Using graphObj = Graphics.FromHwnd(IntPtr.Zero)

            Dim desktop As IntPtr = graphObj.GetHdc()

            Dim LogicalScreenHeight As Integer = GetDeviceCaps(desktop, CInt(DeviceCap.VERTRES))
            Dim PhysicalScreenHeight As Integer = GetDeviceCaps(desktop, CInt(DeviceCap.DESKTOPVERTRES))

            Return CSng(PhysicalScreenHeight) / CSng(LogicalScreenHeight)
        End Using

    End Function

End Module
