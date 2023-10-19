Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.FileIO
Imports System.Drawing.Imaging

Module Module1

    Sub Main(args As String())
        If args.Length = 0 Then
            Console.WriteLine("Please provide a file path.")
            Return
        End If

        Dim filePath As String = args(0)

        If Not System.IO.File.Exists(filePath) Then
            Console.WriteLine("File does not exist.")
            Return
        End If

        Dim imageFormat As ImageFormat = GetImageFormat(filePath)

        If imageFormat Is Nothing Then
            Console.WriteLine("File is not an image.")
            Return
        End If

        Dim image As Image = Image.FromFile(filePath)
        Dim width As Integer = image.Width
        Dim height As Integer = image.Height

        ' Read aspect ratio from INI file
        Dim iniFilePath As String = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Dimension-Calc-CL.ini")
        Dim aspectRatio As Double = 16 / 9
        If System.IO.File.Exists(iniFilePath) Then
            Using parser As New TextFieldParser(iniFilePath)
                parser.TextFieldType = FieldType.Delimited
                parser.SetDelimiters("=")
                While Not parser.EndOfData
                    Dim fields() As String = parser.ReadFields()
                    If fields(0) = "AspectRatio" Then
                        Dim ratioParts() As String = fields(1).Split(":"c)
                        aspectRatio = Double.Parse(ratioParts(0)) / Double.Parse(ratioParts(1))
                        Exit While
                    End If
                End While
            End Using
        End If

        ' Calculate missing dimension
        If height > width Then
            Dim calculatedWidth As Integer = width
            Dim calculatedHeight As Integer = CInt(width / aspectRatio)

            ' Copy the result to clipboard
            Clipboard.SetText(calculatedWidth.ToString() & " " & calculatedHeight.ToString())
        Else
            Dim calculatedWidth As Integer = CInt(height * aspectRatio)
            Dim calculatedHeight As Integer = height

            ' Copy the result to clipboard
            Clipboard.SetText(calculatedWidth.ToString() & " " & calculatedHeight.ToString())
        End If
    End Sub

    Private Function GetImageFormat(filePath As String) As ImageFormat
        Try
            Using image As Image = Image.FromFile(filePath)
                Return image.RawFormat
            End Using
        Catch ex As OutOfMemoryException
            Return Nothing
        End Try
    End Function

End Module
