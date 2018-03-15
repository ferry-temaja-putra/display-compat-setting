Imports Microsoft.Win32
Imports System.Reflection

Module Module1

    Private Const Value = "~ GDIDPISCALING DPIUNAWARE"
    Private Const KeyName = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\Layers"

    Sub Main()

        Console.WriteLine("Checking compatibility setting...")

        Dim valueName = Assembly.GetExecutingAssembly.Location
        Dim currentValue = CType(Registry.GetValue(KeyName, valueName, String.Empty), String)

        If String.IsNullOrEmpty(currentValue) Then

            Console.WriteLine("Compatibility setting is not set. Configuring...")
            Registry.SetValue(KeyName, valueName, Value)

            Console.WriteLine("Compatibility setting is configured.")
            Console.ReadLine()
            Return

        End If

        Console.WriteLine("Compatibility settings is already configured.")
        Console.ReadLine()

    End Sub

End Module
