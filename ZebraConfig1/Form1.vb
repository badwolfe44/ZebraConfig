Imports System.Data.SqlTypes
Imports System.Drawing.Printing
Imports System.IO
Imports System.Net.Sockets
Imports System.Runtime.InteropServices
Imports Zebra.Sdk
Imports System.Management
Imports System.Windows.Forms
Imports System.Data.Common
Imports Zebra.Sdk.Comm
Imports Zebra.Sdk.Printer
Imports Zebra.Sdk.Printer.Discovery
Imports Lextm.SharpSnmpLib
Imports System.Text.RegularExpressions
Imports System.Text
Imports File = System.IO.File
Imports System.Globalization

Public Class Form1
    Private zplString As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SetDefaultText(TextBox2, "Old IP Adress")
        SetDefaultText(TextBox1, "Length")
        SetDefaultText(TextBox3, "Width")
        ComboBox1.SelectedIndex = 0
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        ' Allow only numeric characters (0-9) and the Backspace key
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        ' Allow only numeric characters (0-9) and the Backspace key
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnPrintZPL_Click(sender As Object, e As EventArgs)
        Try
            ' Step 1: Define ZPL Content
            Dim zplContent As String = "^XA^MTd^ML" + TextBox2.Text + "^PW" + TextBox3.Text + "^JUS^XZ" ' Update with your ZPL content

            ' Construct File Path
            Dim documentsPath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            Dim zebraConfigPath As String = Path.Combine(documentsPath, "zebraconfig")
            Directory.CreateDirectory(zebraConfigPath) ' Creates the directory if it does not exist
            Dim filePath As String = Path.Combine(zebraConfigPath, "zebra_config.zpl")

            ' Write the File using 'Using' statement
            Using writer As StreamWriter = New StreamWriter(filePath)
                writer.Write(zplContent)
            End Using  ' Automatically closes the file after writing

            ' Find Printer with 'ZPL' in its name
            Dim printerName As String = FindZPLPrinter()
            If String.IsNullOrEmpty(printerName) Then
                MessageBox.Show("ZPL Printer not found.")
                Return
            End If

            ' Send ZPL to Printer
            If RawPrinterHelper.SendFileToPrinter(printerName, filePath) Then
                MessageBox.Show("File sent to printer successfully.")
            Else
                MessageBox.Show("Failed to send file to printer.")
            End If

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Function FindZPLPrinter() As String
        For Each printer As String In PrinterSettings.InstalledPrinters
            If printer.Contains("ZPL") Then
                Return printer
            End If
        Next
        Return Nothing
    End Function

    Private Sub PrintPageHandler(ByVal sender As Object, ByVal e As PrintPageEventArgs)
        Dim font As New Font("Courier New", 10)
        e.Graphics.DrawString(zplString, font, Brushes.Black, 0, 0)
    End Sub

    Private Sub SetDefaultText(tb As TextBox, defaultText As String)
        tb.Text = defaultText
        tb.ForeColor = Color.Gray
        AddHandler tb.GotFocus, Sub(sender As Object, e As EventArgs)
                                    If tb.Text = defaultText Then
                                        tb.Text = ""
                                        tb.ForeColor = Color.Black
                                    End If
                                End Sub

        AddHandler tb.LostFocus, Sub(sender As Object, e As EventArgs)
                                     If String.IsNullOrWhiteSpace(tb.Text) Then
                                         tb.Text = defaultText
                                         tb.ForeColor = Color.Gray
                                     End If
                                 End Sub
    End Sub

    ' Event handler for your button click
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Discover the connected Zebra printers
        For Each printer As DiscoveredUsbPrinter In UsbDiscoverer.GetZebraUsbPrinters(New ZebraPrinterFilter())
            Console.WriteLine(printer)
        Next

        ' Get the connection string of the first discovered Zebra printer
        Dim connectionString As String = GetUsbPortOfZplPrinter()
        If String.IsNullOrEmpty(connectionString) Then
            MessageBox.Show("ZPL Printer not found or not connected via USB.")
            Return
        End If

        Dim connection As UsbConnection = Nothing
        Try
            ' Establish a USB connection to the printer
            connection = New UsbConnection(connectionString)
            connection.Open()

            Dim generic As ZebraPrinter = ZebraPrinterFactory.GetInstance(connection)
            Dim linkos As ZebraPrinterLinkOs = ZebraPrinterFactory.CreateLinkOsPrinter(generic)
            Dim settings As HashSet(Of String) = linkos.GetAvailableSettings()
            Dim intValue As Integer
            Dim intValueSg As String = linkos.GetSettingValue("ezpl.label_length_max")
            Dim intValueSs As String = TextBox1.Text
            Dim outNumber As Single
            Dim outNumber2 As Single
            Dim divisor As Single = 203
            Dim method As String
            Dim result As Single = outNumber * divisor

            If linkos IsNot Nothing Then

                If Single.TryParse(intValueSs, outNumber) Then
                    ' Convert stringValue to a Single (float)
                    If Single.TryParse(intValueSs, 203.0) Then
                        ' Check if divisor is not zero to avoid division by zero
                        If divisor <> 0 Then
                            ' Perform the division
                            Dim result1 As Single = outNumber / divisor

                            ' Format the result with two decimal places
                            intValueSs = result1.ToString("F2")

                            If ComboBox1.SelectedItem = "Direct" Then
                                method = "direct thermal"
                            ElseIf ComboBox1.SelectedItem = "Thermal" Then
                                method = "thermal trans"
                            End If

                            linkos.SetSetting("ezpl.label_length_max", intValueSs)
                            linkos.SetSetting("ezpl.print_width", TextBox3.Text)
                            linkos.SetSetting("ezpl.print_method", method)

                            If linkos IsNot Nothing Then

                                If Single.TryParse(intValueSs, outNumber) Then
                                    ' Convert stringValue to a Single (float)
                                    If Single.TryParse(intValueSs, 203.0) Then
                                        ' Check if divisor is not zero to avoid division by zero
                                        If divisor <> 0 Then
                                            ' Perform the division
                                            result = outNumber * divisor

                                            ' Format the result with two decimal places
                                            intValueSs = result.ToString("F2")

                                            Console.WriteLine(intValueSs)
                                            Console.WriteLine(linkos.GetSettingValue("ezpl.print_width"))
                                        Else
                                            Console.WriteLine("Cannot divide by zero.")
                                        End If
                                    Else
                                        Console.WriteLine("Invalid divisor value.")
                                    End If
                                Else
                                    Console.WriteLine("Invalid input value.")
                                End If

                            End If

                            MessageBox.Show("Printer settings updated successfully. " + vbCrLf + "Width: " + linkos.GetAllSettingValues("ezpl.print_width") + vbCrLf + "Length: " + intValueSs + vbCrLf + "Print Method: " + linkos.GetAllSettingValues("ezpl.print_method"))

                        Else
                            Console.WriteLine("Cannot divide by zero.")
                        End If
                    Else
                        Console.WriteLine("Invalid divisor value.")
                    End If
                Else
                    Console.WriteLine("Invalid input value.")
                End If

            End If

        Catch ex As ConnectionException
            MessageBox.Show("Error communicating with the printer: " & ex.Message)
        Catch ex As ZebraPrinterLanguageUnknownException
            MessageBox.Show("Could not determine printer language: " & ex.Message)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Ensure the connection is closed
            If connection IsNot Nothing Then
                Try
                    connection.Close()
                Catch ex As Exception
                    MessageBox.Show("An error occurred while closing the connection: " & ex.Message)
                End Try
            End If
        End Try



    End Sub

    ' Function to get the connection string of the first discovered Zebra printer
    Private Function GetUsbPortOfZplPrinter() As String
        Try
            Dim discoveredPrinters As IEnumerable(Of DiscoveredUsbPrinter) = UsbDiscoverer.GetZebraUsbPrinters()
            If discoveredPrinters.Any() Then
                ' Assuming 'Address' or a similar property provides the necessary connection detail
                Return discoveredPrinters.First().Address
            Else
                MessageBox.Show("No Zebra USB printers found.")
            End If
        Catch ex As Exception
            MessageBox.Show("An error occurred while searching for printers: " & ex.Message)
        End Try
        Return String.Empty
    End Function


    Private Function GetDeviceNameByUsbPort(usbPort As String) As String
        Try
            Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_USBControllerDevice")
                For Each usbDevice As ManagementObject In searcher.Get()
                    Dim dependent As String = usbDevice("Dependent").ToString()
                    If dependent.Contains(usbPort) Then
                        Return dependent ' or extract and return the device name part
                    End If
                Next
            End Using
        Catch ex As Exception
            MessageBox.Show("An error occurred while searching for USB devices: " & ex.Message)
        End Try
        Return String.Empty
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Get the value to search for from the TextBox
        Dim printerHostnames As String = TextBox2.Text

        ' Split the input into an array of lines
        Dim ips As String() = printerHostnames.Replace(Chr(13), "").Split(Chr(10))
        Dim port As Integer = 9100  ' Standard port for Zebra network printers

        ' Connect to the Zebra USB printer
        Dim connectionString As String = GetUsbPortOfZplPrinter()
        Dim usbConnection As UsbConnection = Nothing

        Try
            ' Establish a USB connection to the printer
            usbConnection = New UsbConnection(connectionString)
            usbConnection.Open()

            ' Create a ZebraPrinterLinkOs instance for the USB printer
            Dim linkosUsbPrinter As ZebraPrinterLinkOs = ZebraPrinterFactory.CreateLinkOsPrinter(ZebraPrinterFactory.GetInstance(usbConnection))

            For Each ip As String In ips
                Try
                    Dim connection As TcpConnection = New TcpConnection(ip, port)
                    connection.Open()

                    Dim command As String = "^XA^HH^XZ"
                    connection.Write(System.Text.Encoding.UTF8.GetBytes(command))

                    ' Increase the delay if necessary
                    Threading.Thread.Sleep(5000)

                    Dim responseBytes As Byte() = connection.Read(10012)
                    Dim response As String = System.Text.Encoding.ASCII.GetString(responseBytes)
                    Console.WriteLine(response)
                    Dim responseList As String() = response.Split(vbNewLine)

                    For Each setting As String In responseList
                        If setting.Contains("PRINT WIDTH") Then
                            Console.WriteLine(setting)
                        End If
                    Next

                    Dim printWidthMatch As Match = Regex.Match(response, "(\d+)\s+PRINT  WIDTH")
                    Dim maxLengthMatch As Match = Regex.Match(response, "\d+(\.\d+)?\s*MM\s+MAXIMUM LENGTH")
                    Dim method As String

                    Console.WriteLine(printWidthMatch.ToString)

                    If printWidthMatch.Success AndAlso maxLengthMatch.Success Then
                        ' Extract only the numeric part for "MAXIMUM LENGTH"
                        Dim maxLengthValue As String = Regex.Match(maxLengthMatch.Value, "(\d+(\.\d+)?)\s*MM").Groups(1).Value

                        ' Convert millimeters to inches as a Single with two decimal places
                        Dim convertedValueInches As Single = CSng(Double.Parse(maxLengthValue) / 25.4)
                        convertedValueInches = CSng(Math.Round(convertedValueInches, 15))

                        ' Convert to string for display
                        Dim convertedValueString As String = convertedValueInches.ToString()
                        convertedValueInches = convertedValueInches * 203
                        Console.WriteLine(convertedValueString + "shfd")

                        If response.Contains("DIRECT-THERMAL") Then
                            method = "direct thermal"
                        ElseIf response.Contains("THERMAL-TRANS") Then
                            method = "thermal trans"
                        Else
                            method = "direct thermal"
                        End If
                        Console.WriteLine($"Print Width: {printWidthMatch.Groups(1).Value} dots")
                        Console.WriteLine($"Maximum Length: {convertedValueString} inches")

                        ' Send data directly to the USB Zebra printer using Link-OS
                        If linkosUsbPrinter IsNot Nothing Then
                            linkosUsbPrinter.SetSetting("ezpl.label_length_max", convertedValueString)
                            linkosUsbPrinter.SetSetting("ezpl.print_width", printWidthMatch.Groups(1).Value)
                            linkosUsbPrinter.SetSetting("ezpl.print_method", method)

                            ' Display success message
                            MessageBox.Show("Printer settings updated successfully. " +
                                            vbCrLf + "Width: " + linkosUsbPrinter.GetAllSettingValues("ezpl.print_width") +
                                            vbCrLf + "Length: " + convertedValueInches.ToString +
                                            vbCrLf + "Print Method: " + linkosUsbPrinter.GetAllSettingValues("ezpl.print_method"))
                        End If
                    Else
                        ' Handle case where information is not successfully extracted
                        MessageBox.Show($"Error extracting information from printer {ip}")
                    End If

                Catch ex As ConnectionException
                    ' Handle connection errors
                    MessageBox.Show($"Error communicating with printer {ip}: {ex.Message}")
                Catch ex As Exception
                    ' Handle other errors
                    MessageBox.Show($"An error occurred: {ex.Message}")
                End Try
            Next

        Catch ex As ConnectionException
            MessageBox.Show("Error communicating with the USB printer: " & ex.Message)
        Catch ex As ZebraPrinterLanguageUnknownException
            MessageBox.Show("Could not determine USB printer language: " & ex.Message)
        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message)
        Finally
            ' Ensure the USB connection is closed
            If usbConnection IsNot Nothing Then
                Try
                    usbConnection.Close()
                Catch ex As Exception
                    MessageBox.Show("An error occurred while closing the USB connection: " & ex.Message)
                End Try
            End If
        End Try
    End Sub




    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    ' Get the value to search for from the TextBox
    '    Dim searchValue As String = TextBox2.Text.Trim()

    '    ' Check if the search value is empty
    '    If String.IsNullOrEmpty(searchValue) Then
    '        MessageBox.Show("Please enter a search value.")
    '        Return
    '    End If

    '    Try
    '        ' Read all lines from the CSV file
    '        Dim csvFilePath As String = "printer_info1.csv"
    '        Dim csvLines As String() = File.ReadAllLines(csvFilePath)

    '        ' Iterate through each line to find the matching IP address
    '        For Each line As String In csvLines
    '            Dim values As String() = line.Split(","c)

    '            ' Check if the IP address in the first column matches the search value
    '            If values.Length >= 1 AndAlso values(0).Trim() = searchValue Then
    '                ' Found a matching row
    '                ' Extract values from the second and third columns
    '                Dim printWidthValue As String = If(values.Length >= 2, values(1).Trim(), "N/A")
    '                Dim maxLengthValue As String = If(values.Length >= 3, values(2).Trim(), "N/A")
    '                Dim printMethod As String = If(values.Length >= 4, values(3).Trim(), "N/A")

    '                ' Display the values or use them as needed
    '                MessageBox.Show($"Print Width: {printWidthValue}, Maximum Length: {maxLengthValue}")

    '                ' Connect to the Zebra printer
    '                Dim connectionString As String = GetUsbPortOfZplPrinter()
    '                Dim connection As UsbConnection = Nothing

    '                Try
    '                    ' Establish a USB connection to the printer
    '                    connection = New UsbConnection(connectionString)
    '                    connection.Open()

    '                    ' Create a ZebraPrinterLinkOs instance
    '                    Dim linkos As ZebraPrinterLinkOs = ZebraPrinterFactory.CreateLinkOsPrinter(ZebraPrinterFactory.GetInstance(connection))
    '                    Dim intValue As Integer
    '                    Dim intValueSg As String = maxLengthValue
    '                    Dim intValueSs As String = maxLengthValue
    '                    Dim outNumber As Single
    '                    Dim divisor As Single = 203
    '                    Dim method As String
    '                    Dim result As Single = outNumber * divisor

    '                    If linkos IsNot Nothing Then

    '                        If Single.TryParse(intValueSs, outNumber) Then
    '                            ' Convert stringValue to a Single (float)
    '                            If Single.TryParse(intValueSs, 203.0) Then
    '                                ' Check if divisor is not zero to avoid division by zero
    '                                If divisor <> 0 Then
    '                                    ' Perform the division
    '                                    Dim result1 As Single = outNumber / divisor

    '                                    ' Format the result with two decimal places
    '                                    intValueSs = result1.ToString("F2")

    '                                    If ComboBox1.SelectedItem = "Direct" Then
    '                                        method = "direct thermal"
    '                                    ElseIf ComboBox1.SelectedItem = "Thermal" Then
    '                                        method = "thermal trans"
    '                                    End If

    '                                    linkos.SetSetting("ezpl.label_length_max", intValueSs)
    '                                    linkos.SetSetting("ezpl.print_width", printWidthValue)
    '                                    linkos.SetSetting("ezpl.print_method", method)

    '                                    If linkos IsNot Nothing Then

    '                                        If Single.TryParse(intValueSs, outNumber) Then
    '                                            ' Convert stringValue to a Single (float)
    '                                            If Single.TryParse(intValueSs, 203.0) Then
    '                                                ' Check if divisor is not zero to avoid division by zero
    '                                                If divisor <> 0 Then
    '                                                    ' Perform the division
    '                                                    result = outNumber * divisor

    '                                                    ' Format the result with two decimal places
    '                                                    intValueSs = result.ToString("F2")

    '                                                    Console.WriteLine(intValueSs)
    '                                                    Console.WriteLine(linkos.GetSettingValue("ezpl.print_width"))
    '                                                Else
    '                                                    Console.WriteLine("Cannot divide by zero.")
    '                                                End If
    '                                            Else
    '                                                Console.WriteLine("Invalid divisor value.")
    '                                            End If
    '                                        Else
    '                                            Console.WriteLine("Invalid input value.")
    '                                        End If

    '                                    End If

    '                                    MessageBox.Show("Printer settings updated successfully. " + vbCrLf + "Width: " + linkos.GetAllSettingValues("ezpl.print_width") + vbCrLf + "Length: " + intValueSs + vbCrLf + "Print Method: " + linkos.GetAllSettingValues("ezpl.print_method"))

    '                                Else
    '                                    Console.WriteLine("Cannot divide by zero.")
    '                                End If
    '                            Else
    '                                Console.WriteLine("Invalid divisor value.")
    '                            End If
    '                        Else
    '                            Console.WriteLine("Invalid input value.")
    '                        End If

    '                    End If
    '                Catch ex As ConnectionException
    '                    MessageBox.Show("Error communicating with the printer: " & ex.Message)
    '                Catch ex As ZebraPrinterLanguageUnknownException
    '                    MessageBox.Show("Could not determine printer language: " & ex.Message)
    '                Catch ex As Exception
    '                    MessageBox.Show("An error occurred: " & ex.Message)
    '                Finally
    '                    ' Ensure the connection is closed
    '                    If connection IsNot Nothing Then
    '                        Try
    '                            connection.Close()
    '                        Catch ex As Exception
    '                            MessageBox.Show("An error occurred while closing the connection: " & ex.Message)
    '                        End Try
    '                    End If
    '                End Try

    '                Return ' Exit the loop after finding the first match
    '            End If
    '        Next

    '        ' If no match is found
    '        MessageBox.Show($"No match found for IP: {searchValue}")
    '    Catch ex As Exception
    '        MessageBox.Show($"An error occurred: {ex.Message}")
    '    End Try
    'End Sub



    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Dim printerHostnames As String = TextBox2.Text

    '    ' Split the input into an array of lines
    '    Dim ips As String() = printerHostnames.Replace(Chr(13), "").Split(Chr(10))
    '    Dim port As Integer = 9100  ' Standard port for Zebra network printers

    '    ' Create a StringBuilder to store CSV content
    '    Dim csvContent As New StringBuilder()

    '    For Each ip As String In ips
    '        Try
    '            Dim connection As TcpConnection = New TcpConnection(ip, port)
    '            connection.Open()

    '            Dim command As String = "^XA^HH^XZ"
    '            connection.Write(System.Text.Encoding.UTF8.GetBytes(command))

    '            ' Increase the delay if necessary
    '            Threading.Thread.Sleep(500)

    '            Dim responseBytes As Byte() = connection.Read(1024)
    '            Dim response As String = System.Text.Encoding.ASCII.GetString(responseBytes)
    '            'Console.WriteLine(response)
    '            Dim printWidthMatch As Match = Regex.Match(response, "(\d+)\s+PRINT WIDTH")
    '            Dim maxLengthMatch As Match = Regex.Match(response, "\d+(\.\d+)?\s*MM\s+MAXIMUM LENGTH")
    '            Dim method As String

    '            If printWidthMatch.Success AndAlso maxLengthMatch.Success Then
    '                ' Extract only the numeric part for "MAXIMUM LENGTH"
    '                Dim maxLengthValue As String = Regex.Match(maxLengthMatch.Value, "(\d+(\.\d+)?)\s*MM").Groups(1).Value

    '                ' Convert millimeters to inches, then multiply by 203 and round to the nearest whole number
    '                Dim convertedValue As Integer = CInt(Math.Round(Double.Parse(maxLengthValue) * 0.0393701 * 203))
    '                If response.Contains("DIRECT-THERMAL") Then
    '                    method = "DIRECT-THERMAL"
    '                ElseIf response.Contains("THERMAL-TRANS") Then
    '                    method = "THERMAL-TRANS"
    '                Else
    '                    method = "DIRECT-THERMAL"
    '                End If
    '                Console.WriteLine($"Print Width: {printWidthMatch.Groups(1).Value} dots")
    '                Console.WriteLine($"Maximum Length: {convertedValue} dots")

    '                ' Append to CSV content with new line
    '                csvContent.AppendLine($"{ip},{printWidthMatch.Groups(1).Value},{convertedValue},{method}")
    '            Else
    '                ' Handle case where information is not successfully extracted
    '                csvContent.AppendLine($"{ip},Error,Error")
    '            End If

    '        Catch ex As ConnectionException
    '            ' Handle connection errors
    '            csvContent.AppendLine($"{ip},Error,Error")
    '        Catch ex As Exception
    '            ' Handle other errors
    '            csvContent.AppendLine($"{ip},Error,Error")
    '        End Try
    '    Next

    '    ' Save CSV content to a file
    '    Dim csvFilePath As String = "printer_info1.csv"
    '    File.WriteAllText(csvFilePath, csvContent.ToString())

    '    ' Display success message
    '    MessageBox.Show($"Printer information saved to {csvFilePath}")
    'End Sub



    'Private Async Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    ' Read IPs from RichTextBox and create a list
    '    Dim ipList As String = RichTextBox1.Text

    '    Dim ips As String() = ipList.Replace(Chr(13), "").Split(Chr(10))

    '    ' Create a StringBuilder to store CSV content
    '    Dim csvContent As New StringBuilder()

    '    ' Iterate through each IP asynchronously
    '    For Each printerIp As String In ipList
    '        Dim connection As TcpConnection = New TcpConnection(printerIp, 9100)

    '        Try
    '            ' Open connection
    '            connection.Open()

    '            ' Send command to retrieve printer information
    '            Dim command As String = "^XA^HH^XZ"
    '            connection.Write(System.Text.Encoding.UTF8.GetBytes(command))

    '            ' Delay if necessary
    '            Threading.Thread.Sleep(500)

    '            Dim responseBytes As Byte() = connection.Read(1024)
    '            Dim response As String = System.Text.Encoding.ASCII.GetString(responseBytes)
    '            Dim responseLines As List(Of String) = response.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries).ToList()
    '            Console.WriteLine(response)
    '            Dim printWidthPattern As String = "(\d+)\s+PRINT WIDTH"
    '            Dim maxLengthPattern As String = "\d+(\.\d+)?\s*MM\s+MAXIMUM LENGTH"

    '            Dim printWidthMatch As Match = Regex.Match(response, printWidthPattern)
    '            Dim maxLengthMatch As Match = Regex.Match(response, maxLengthPattern)

    '            If printWidthMatch.Success AndAlso maxLengthMatch.Success Then
    '                ' Convert max length to inches, multiply by 203, and round to the nearest whole number
    '                Dim maxLengthValue As Integer = CInt(Math.Round(Double.Parse(maxLengthMatch.Groups(1).Value) * 0.0393701 * 203))
    '                Console.WriteLine(maxLengthValue)

    '                ' Append to CSV content
    '                'csvContent.AppendLine($"{printerIp},{printWidthMatch.Groups(1).Value},{maxLengthValue}")
    '            Else
    '                ' Handle case where information is not successfully extracted
    '                'csvContent.AppendLine($"{printerIp},Error,Error")
    '            End If

    '        Catch ex As ConnectionException
    '            ' Handle connection errors
    '            csvContent.AppendLine($"{printerIp},Error,Error")
    '        Catch ex As Exception
    '            ' Handle other errors
    '            csvContent.AppendLine($"{printerIp},Error,Error")
    '        Finally
    '            ' Ensure the connection is closed
    '            If connection IsNot Nothing Then
    '                Try
    '                    connection.Close()
    '                Catch ex As Exception
    '                    ' Handle error while closing the connection
    '                End Try
    '            End If
    '        End Try
    '    Next

    '    ' Save CSV content to a file
    '    Dim csvFilePath As String = "printer_info.csv"
    '    File.WriteAllText(csvFilePath, csvContent.ToString())

    '    ' Display success message
    '    MessageBox.Show($"Printer information saved to {csvFilePath}")
    'End Sub
End Class
