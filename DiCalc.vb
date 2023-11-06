Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
' Imports System.Runtime.InteropServices
' Imports Microsoft.Win32
' =================================================================
 ' Custom NumericUpDown to support Up and Down Buttons
Public Class CustomNumericUpDown
    Inherits NumericUpDown

    Private calculationAction As Action

    Public Sub New(calculationAction As Action)
        MyBase.New()
        Me.calculationAction = calculationAction
    End Sub

    Public Overrides Sub UpButton()
        MyBase.UpButton()
        ' Execute the calculation action
        calculationAction()
    End Sub

    Public Overrides Sub DownButton()
        MyBase.DownButton()
        ' Execute the calculation action
        calculationAction()
    End Sub
End Class
' =================================================================
Public Class WindowsApi
    ''' <summary>
    ''' Extracts an icon from a specified exe, dll or icon file. Call DestroyIcon on handles once finished.
    ''' </summary>
    ''' <param name="lpszFile">The file to extract the icon from</param>
    ''' <param name="nIconIndex">Specifies the zero-based index of the icon to extract. If this value is –1 and phiconLarge and phiconSmall are both NULL, the function returns the total number of icons in the specified file</param>
    ''' <param name="phiconLarge">Handle to the large version of the icon requested</param>
    ''' <param name="phiconSmall">handle to the small version of the icon requested</param>
    ''' <param name="nIcons">Number of icons to extract</param>
    <System.Runtime.InteropServices.DllImportAttribute("shell32.dll", EntryPoint:="ExtractIconExW", CallingConvention:=System.Runtime.InteropServices.CallingConvention.StdCall)> _
    Public Shared Function ExtractIconExW(<System.Runtime.InteropServices.InAttribute(), System.Runtime.InteropServices.MarshalAsAttribute(System.Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpszFile As String, ByVal nIconIndex As Integer, ByRef phiconLarge As System.IntPtr, ByRef phiconSmall As System.IntPtr, ByVal nIcons As UInteger) As UInteger
    End Function
End Class
' =================================================================
Public Class Form1
    Inherits Form
    ' Create TextBoxes
	' Dim aspectRatio1 As New CustomNumericUpDown(Sub()
                                                       ' Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
                                                           ' dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
                                                   ' End Sub)
	' Dim aspectRatio2 As New CustomNumericUpDown(Sub()
                                                       ' Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
                                                           ' dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
                                                   ' End Sub)
	Dim dimensionW As New CustomNumericUpDown(Sub()
                                                       Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
                                                       dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
                                                       Clipboard.SetText(Math.Round(dimensionW.Value / aspectRatio))
                                                   End Sub)
	Dim dimensionH As New CustomNumericUpDown(Sub()
                                                       Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
                                                       dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
                                                       Clipboard.SetText(Math.Round(dimensionH.Value * aspectRatio))
                                                   End Sub)
    Private aspectRatio1 As New NumericUpDown()
    Private aspectRatio2 As New NumericUpDown()
    ' Private dimensionW As New NumericUpDown()
    ' Private dimensionH As New NumericUpDown()
    ' Create Buttons
    ' Private calculateButton As New Button()
	Private closeButton As New Button()
    Private minimizeButton As New Button()
	Private openWebsiteButton1 As New Button()
	Private toolTip1 As New ToolTip()
    ' Create Labels
    Private AspectRatioLabel As New Label()
    Private ratioWidthLabel As New Label()
    Private ratioHeightLabel As New Label()
    Private pixelsWidthLabel As New Label()
    Private pixelsHeightLabel As New Label()

    Public Sub New()
        ' Set form properties
        ' Enable drag and drop
        Me.AllowDrop = True

		Me.FormBorderStyle = FormBorderStyle.FixedDialog
		' Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
		' Me.showintaskbar = False
		Me.ControlBox = False
		Me.Size = New Size(220, 190)
		' Me.Text = "=)"
		Me.Text = "DiCalc - amymor OgomnamO"
        Me.BackColor = Color.FromArgb(60,75,66)
        ' Me.BackColor = Color.FromArgb(60,66,75)
        Me.ForeColor = Color.FromArgb(235,235,235) ' Set text color (for Labels and Buttons if not set already)
        ' Me.ForeColor = Color.White
        Me.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath)
' =================================================================
        ' Initialize Labels
        AspectRatioLabel.Text = "Aspect ratio"
        AspectRatioLabel.Location = New Point(18, 0)
        AspectRatioLabel.Size = New Size(100, 30)
        AspectRatioLabel.Font = New Font("Segoe UI", 12) ' Set font size
        AspectRatioLabel.ForeColor = Color.FromArgb(238,190,123) ' Set text color
		
        ' ratioWidthLabel.Text = "Ratio width"
        ' ratioWidthLabel.Location = New Point(140, 10)
        ' ratioWidthLabel.Size = New Size(100, 25)
        ' ratioWidthLabel.Font = New Font("Segoe UI", 12) ' Set font size
        ' ratioHeightLabel.Text = "Ratio height"
        ' ratioHeightLabel.Location = New Point(140, 40)
        ' ratioHeightLabel.Size = New Size(100, 25)
        ' ratioHeightLabel.Font = New Font("Segoe UI", 12) ' Set font size
		
        ' pixelsWidthLabel.Text = "W"
        ' pixelsWidthLabel.Location = New Point(5, 82)
        ' pixelsWidthLabel.Size = New Size(100, 25)
        ' pixelsWidthLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        ' pixelsWidthLabel.ForeColor = Color.FromArgb(193,141,227)

        ' pixelsHeightLabel.Text = "H"
        ' pixelsHeightLabel.Location = New Point(5, 118)
        ' pixelsHeightLabel.Size = New Size(100, 25)
        ' pixelsHeightLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        ' pixelsHeightLabel.ForeColor = Color.FromArgb(151,215,151)
' =================================================================
        ' Initialize TextBoxes
        aspectRatio1.Location = New Point(10, 25)
        aspectRatio1.Size = New Size(59, 20)
        aspectRatio1.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        aspectRatio1.BackColor = Color.FromArgb(54,69,60) ' Set background color
        ' aspectRatio1.BackColor = Color.FromArgb(54,60,69) ' Set background color
        aspectRatio1.ForeColor = Color.FromArgb(235,235,235) ' Set text color
		
        aspectRatio2.Location = New Point(71, 25)
        aspectRatio2.Size = New Size(59, 20)
        aspectRatio2.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        aspectRatio2.BackColor = Color.FromArgb(54,69,60) ' Set background color
        aspectRatio2.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        
		dimensionW.Location = New Point(30, 80)
        dimensionW.Size = New Size(105, 20)
        dimensionW.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        dimensionW.BackColor = Color.FromArgb(54,69,60) ' Set background color
        dimensionW.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        
		dimensionH.Location = New Point(30, 115)
        dimensionH.Size = New Size(105, 20)
        dimensionH.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        dimensionH.BackColor = Color.FromArgb(54,69,60) ' Set background color
        dimensionH.ForeColor = Color.FromArgb(235,235,235) ' Set text color
		
		' TextBoxes BorderStyle
        aspectRatio1.BorderStyle = BorderStyle.FixedSingle
        aspectRatio2.BorderStyle = BorderStyle.FixedSingle
        dimensionW.BorderStyle = BorderStyle.FixedSingle
        dimensionH.BorderStyle = BorderStyle.FixedSingle
		
        ' NumericUpDown colors
        ' aspectRatio1.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' aspectRatio1.Controls("UpDownButtons").ForeColor = Color.White
        ' aspectRatio2.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' aspectRatio2.Controls("UpDownButtons").ForeColor = Color.White
        ' dimensionW.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' dimensionW.Controls("UpDownButtons").ForeColor = Color.White
        ' dimensionH.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' dimensionH.Controls("UpDownButtons").ForeColor = Color.White
		
        ' maximum value of the data type
        dimensionW.Maximum = 9999999
        dimensionH.Maximum = 9999999
' =================================================================
        ' Initialize Buttons
        ' calculateButton.Text = "Calculate"
        ' calculateButton.Location = New Point(5, 160)
        ' calculateButton.Size = New Size(130, 60)
        ' calculateButton.Font = New Font("Segoe UI", 18) ' Set font size
        ' calculateButton.ForeColor = Color.FromArgb(100,205,100) ' Set text color
        ' calculateButton.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        ' calculateButton.BackColor = Color.FromArgb(50,80,50) ' Set background color
        ' calculateButton.FlatStyle = FlatStyle.Flat
        ' calculateButton.FlatAppearance.BorderColor = Color.FromArgb(25,60,25)
        ' calculateButton.FlatAppearance.BorderSize = 1
		
        closeButton.Text = "❌"
		closeButton.Location =  New Point(182, 0)
        closeButton.Size = New Size(32, 32)
        closeButton.Font = New Font("Segoe UI", 16) ' Set font size
        closeButton.BackColor = Color.Brown
        closeButton.FlatStyle = FlatStyle.Flat
        closeButton.FlatAppearance.BorderColor = Color.FromArgb(50,50,50)
        closeButton.FlatAppearance.BorderSize = 1
		
        minimizeButton.Text = "—"
        minimizeButton.Location = New Point(151, 0)
        minimizeButton.Size = New Size(32, 32)
        minimizeButton.Font = New Font("Segoe UI", 16, FontStyle.Bold) ' Set font size
        ' minimizeButton.ForeColor = Color.FromArgb(100,205,100) ' Set text color
        ' minimizeButton.BackColor = Color.FromArgb(54,60,69) ' Set background color
        minimizeButton.BackColor = Color.SteelBlue ' Set background color
        minimizeButton.FlatStyle = FlatStyle.Flat
        minimizeButton.FlatAppearance.BorderColor = Color.FromArgb(50,50,50)
        minimizeButton.FlatAppearance.BorderSize = 1

        ' openWebsiteButton1.Text = "ℹ"
        ' openWebsiteButton1.Text = "v1.5"
        AddHandler openWebsiteButton1.Paint, Sub(sender As Object, e As PaintEventArgs)
                                                             ' Calculate the bounds of the text
                                                             Dim bounds As New Rectangle(New Point(0, 0), openWebsiteButton1.Size)

                                                             ' Draw the text on the button
                                                             TextRenderer.DrawText(e.Graphics, "v1.5", openWebsiteButton1.Font, bounds, openWebsiteButton1.ForeColor, TextFormatFlags.SingleLine)
                                                             End Sub
        openWebsiteButton1.Location = New Point(182, 142)
        openWebsiteButton1.Size = New Size(32, 18)
        ' openWebsiteButton1.Font = New Font("Segoe UI black", 24, FontStyle.Bold) ' Set font size
        openWebsiteButton1.Font = New Font("Tahoma", 10) ' Set font size
        ' openWebsiteButton1.Padding = New Padding(0, 0, 0, 0)
        ' openWebsiteButton1.TextAlign = ContentAlignment.MiddleCenter
        openWebsiteButton1.ForeColor = Color.Aquamarine
        ' openWebsiteButton1.BackColor = Color.FromArgb(60,66,75)
        openWebsiteButton1.FlatStyle = FlatStyle.Flat
        openWebsiteButton1.FlatAppearance.BorderColor = Color.FromArgb(100,100,100)
        openWebsiteButton1.FlatAppearance.BorderSize = 1

' =================================================================
        ' Add TextBoxes, Labels, and Buttons to Form
        Me.Controls.Add(aspectRatio1)
        Me.Controls.Add(aspectRatio2)
        Me.Controls.Add(dimensionW)
        Me.Controls.Add(dimensionH)
        Me.Controls.Add(AspectRatioLabel)
        ' Me.Controls.Add(ratioWidthLabel)
        ' Me.Controls.Add(ratioHeightLabel)
        ' Me.Controls.Add(pixelsWidthLabel)
        ' Me.Controls.Add(pixelsHeightLabel)
        ' Me.Controls.Add(calculateButton)
		Me.Controls.Add(closeButton)
		Me.Controls.Add(minimizeButton)
		Me.Controls.Add(openWebsiteButton1)
		' Me.Controls.Add(openWebsiteButton2)

' =================================================================
		Dim CopyButtonW As New Button
		Dim CopyButtonH As New Button
		Dim CopyButtonALL As New Button
		
        CopyButtonW.Text = "W"
        CopyButtonW.Location = New Point(6, 80)
        CopyButtonW.Size = New Size(25, 32)
        CopyButtonW.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        CopyButtonW.ForeColor = Color.FromArgb(193,141,227)
        CopyButtonW.BackColor = Color.FromArgb(60,66,75)
        CopyButtonW.FlatStyle = FlatStyle.Flat
        CopyButtonW.FlatAppearance.BorderColor = Color.FromArgb(100,100,100)
        CopyButtonW.FlatAppearance.BorderSize = 1
		
        CopyButtonH.Text = "H"
        CopyButtonH.Location = New Point(6, 115)
        CopyButtonH.Size = New Size(25, 32)
        CopyButtonH.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        CopyButtonH.ForeColor = Color.FromArgb(151,215,151)
        CopyButtonH.BackColor = Color.FromArgb(60,66,75)
        CopyButtonH.FlatStyle = FlatStyle.Flat
        CopyButtonH.FlatAppearance.BorderColor = Color.FromArgb(100,100,100)
        CopyButtonH.FlatAppearance.BorderSize = 1

        ' CopyButtonALL.Text = "📋"
        CopyButtonALL.Location = New Point(140, 93)
        CopyButtonALL.Size = New Size(35, 41)
        ' CopyButtonALL.Font = New Font("Segoe UI", 10) ' Set font size
        ' CopyButtonALL.ForeColor = Color.FromArgb(151,215,151)
        CopyButtonALL.BackColor = Color.FromArgb(60,66,75)
        CopyButtonALL.FlatStyle = FlatStyle.Flat
        CopyButtonALL.FlatAppearance.BorderColor = Color.FromArgb(100,100,100)
        CopyButtonALL.FlatAppearance.BorderSize = 1

        Me.Controls.Add(CopyButtonW)
        Me.Controls.Add(CopyButtonH)
        Me.Controls.Add(CopyButtonALL)

        ' Copy buttons
        AddHandler CopyButtonW.Click, Sub()
                                        Clipboard.SetText(dimensionW.Value)
                                    End Sub
        AddHandler CopyButtonH.Click, Sub()
                                        Clipboard.SetText(dimensionH.Value)
                                    End Sub
        AddHandler CopyButtonALL.Click, Sub()
                                        Clipboard.SetText(dimensionW.Value & " " & dimensionH.Value)
                                    End Sub
' --------------------------------------------------------------------
        ' Create an array to hold the handle of the extracted icon
        Dim phiconLarge(0) As IntPtr
        Dim phiconSmall(0) As IntPtr
        ' Extract the icon from shell32.dll
        Dim iconCopyButtonALL1 As UInteger = WindowsApi.ExtractIconExW("C:\Windows\System32\mmcndmgr.dll", -30518, phiconLarge(0), phiconSmall(0), 1)
        Dim iconCopyButtonALL2 As Icon = Icon.FromHandle(phiconLarge(0))
        Dim iconCopyButtonALL3 As Bitmap = iconCopyButtonALL2.ToBitmap()
        ' Set the Icon object as the button's image
        CopyButtonALL.Image = iconCopyButtonALL3
' --------------------------------------------------------------------
' =================================================================
		Dim DeleteButton As New Button

        ' DeleteButton.Text = "⬇️"
        ' DeleteButton.Text = "⟳"
        ' DeleteButton.Text = "🧹"
        DeleteButton.Location = New Point(174, 93)
        DeleteButton.Size = New Size(35, 41)
        DeleteButton.Font = New Font("Segoe UI", 24, FontStyle.Bold) ' Set font size
        ' DeleteButton.ForeColor = Color.FromArgb(151,215,151)
        DeleteButton.BackColor = Color.FromArgb(60,66,75)
        DeleteButton.FlatStyle = FlatStyle.Flat
        DeleteButton.FlatAppearance.BorderColor = Color.FromArgb(100,100,100)
        DeleteButton.FlatAppearance.BorderSize = 1

        Me.Controls.Add(DeleteButton)

        ' Delete Button
        AddHandler DeleteButton.Click, Sub()
                                        dimensionW.value = 0
                                        dimensionH.value = 0
                                        dimensionW.text = ""
                                        dimensionH.text = ""
                                    End Sub
' --------------------------------------------------------------------
        ' Create an array to hold the handle of the extracted icon
        ' Extract the icon from wmploc.dll
        Dim iconDeleteButton1 As UInteger = WindowsApi.ExtractIconExW("C:\Windows\System32\wmploc.dll", -29607, phiconLarge(0), phiconSmall(0), 1)
        Dim iconDeleteButton2 As Icon = Icon.FromHandle(phiconLarge(0))
        Dim iconDeleteButton3 As Bitmap = iconDeleteButton2.ToBitmap()
        ' Set the Icon object as the button's image
        DeleteButton.Image = iconDeleteButton3
' --------------------------------------------------------------------
' =================================================================
        ' Set 16:9 as default
        ' aspectRatio1.text = 16
        ' aspectRatio2.text = 9
        aspectRatio1.Value = 16
        aspectRatio2.Value = 9
		dimensionW.text = ""
		dimensionH.text = ""
		' dimensionW.Value = 0
		' dimensionH.Value = 0

		' focus to the TextBox at startup
		' dimensionW.Select()
		dimensionH.Select()
		
		toolTip1.SetToolTip(pixelsWidthLabel, "Pixels width")
		toolTip1.SetToolTip(pixelsHeightLabel, "Pixels height")
		toolTip1.SetToolTip(closeButton, "Close")
		toolTip1.SetToolTip(minimizeButton, "Minimize")
		toolTip1.SetToolTip(openWebsiteButton1, "Click to open the GitHub webpage and check manually for updates")
		toolTip1.SetToolTip(CopyButtonW, "Copy Width to the clipboard")
		toolTip1.SetToolTip(CopyButtonH, "Copy Height to the clipboard")
		toolTip1.SetToolTip(CopyButtonALL, "Copy Width and Height to the clipboard")
		toolTip1.SetToolTip(DeleteButton, "Clear Width and Height")

' Add event handler =============================================
        ' AddHandler calculateButton.Click, AddressOf calculateButton_Click
		
        ' closeButton
        AddHandler closeButton.Click, AddressOf closeButton_Click
        ' minimizeButton
        AddHandler minimizeButton.Click, AddressOf minimizeButton_Click
        ' openWebsiteButton
		AddHandler openWebsiteButton1.Click, AddressOf openWebsiteButton1_Click

		' Add ValueChanged event handlers to NumericUpDown controls
        AddHandler aspectRatio1.ValueChanged, AddressOf aspectRatio1_ValueChanged
        AddHandler aspectRatio2.ValueChanged, AddressOf aspectRatio2_ValueChanged
		
		' Add TextChanged  event handlers to NumericUpDown controls (to make field empty while using KeyUp)
        AddHandler dimensionW.TextChanged , AddressOf calculateButton_TextChangedW
        AddHandler dimensionH.TextChanged , AddressOf calculateButton_TextChangedH

        ' Add KeyUp event handlers to NumericUpDown controls
        AddHandler dimensionW.KeyUp, AddressOf calculateButton_KeyUpW
        AddHandler dimensionH.KeyUp, AddressOf calculateButton_KeyUpH
		
        ' Add KeyDown event handlers to NumericUpDown controls (to Enable Ctrl+A in NumericUpDown field)
        AddHandler dimensionW.KeyDown, AddressOf dimension_KeyDownW
        AddHandler dimensionH.KeyDown, AddressOf dimension_KeyDownH
	
        ' Add KeyPress event handlers to NumericUpDown controls
        ' AddHandler dimensionW.KeyPress, AddressOf calculateButton_ClickWP
        ' AddHandler dimensionH.KeyPress, AddressOf calculateButton_ClickHP

        ' Add Click event handlers to NumericUpDown controls
        ' AddHandler aspectRatio1.Click , AddressOf calculateButton_Click1
        ' AddHandler aspectRatio2.Click , AddressOf calculateButton_Click2
        ' AddHandler dimensionW.Click , AddressOf calculateButton_Click2
        ' AddHandler dimensionH.Click , AddressOf calculateButton_Click1
		
        ' Add MouseUp event handlers to NumericUpDown controls (for Right-click > Paste)
        AddHandler dimensionW.MouseUp , AddressOf dimension_MouseUpW
        AddHandler dimensionH.MouseUp , AddressOf dimension_MouseUpH
		
        ' Add Enter event handlers to NumericUpDown controls
        ' AddHandler aspectRatio1.Enter , AddressOf calculateButton_Click1
        ' AddHandler aspectRatio2.Enter , AddressOf calculateButton_Click2
        ' AddHandler dimensionW.Enter , AddressOf calculateButton_EnterW
        ' AddHandler dimensionH.Enter , AddressOf calculateButton_EnterH
		
        ' Add GotFocus (focus) event handlers to NumericUpDown controls
        ' AddHandler aspectRatio1.GotFocus, AddressOf calculateButton_Click1
        ' AddHandler aspectRatio2.GotFocus, AddressOf calculateButton_Click2
        ' AddHandler dimensionW.GotFocus, AddressOf calculateButton_ClickW
        ' AddHandler dimensionH.GotFocus, AddressOf calculateButton_ClickH
		
        ' Add Leave (loss focus) event handlers to NumericUpDown controls
        ' AddHandler aspectRatio1.Leave, AddressOf calculateButton_Click1
        ' AddHandler aspectRatio2.Leave, AddressOf calculateButton_Click2
        ' AddHandler dimensionW.Leave, AddressOf calculateButton_ClickW
        ' AddHandler dimensionH.Leave, AddressOf calculateButton_ClickH
    End Sub

' Main Calc =================================================================
' Tying to use ValueChanged instead of CustomNumericUpDown------------------------------------------------------------
' Private isProgrammaticChange As Boolean = False

' Private Sub calculateButton_ClickW(ByVal sender As Object, ByVal e As EventArgs) Handles dimensionW.TextChanged
        ' isProgrammaticChange = True
        ' PerformCalculationW(dimensionW, dimensionH, aspectRatio1, aspectRatio2)
' End Sub

' Private Sub calculateButton_ClickH(ByVal sender As Object, ByVal e As EventArgs) Handles dimensionH.TextChanged
        ' isProgrammaticChange = True
        ' PerformCalculationH(dimensionW, dimensionH, aspectRatio1, aspectRatio2)
' End Sub

' Private Sub PerformCalculationW(ByVal dimensionW As NumericUpDown, ByVal dimensionH As NumericUpDown, ByVal aspectRatio1 As NumericUpDown, ByVal aspectRatio2 As NumericUpDown)
        ' Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        ' dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
        ' Clipboard.SetText(Math.Round(dimensionW.Value / aspectRatio))
    ' isProgrammaticChange = False
' End Sub

' Private Sub PerformCalculationH(ByVal dimensionW As NumericUpDown, ByVal dimensionH As NumericUpDown, ByVal aspectRatio1 As NumericUpDown, ByVal aspectRatio2 As NumericUpDown)
        ' Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        ' dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
        ' Clipboard.SetText(Math.Round(dimensionH.Value * aspectRatio))
    ' isProgrammaticChange = False
' End Sub

' ----------------------------------------------------------------------------
 
    Private Sub aspectRatio1_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
		 Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
	End Sub

    Private Sub aspectRatio2_ValueChanged(ByVal sender As Object, ByVal e As EventArgs) 
        Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
    End Sub

    ' Private Sub calculateButton_ClickW(ByVal sender As Object, ByVal e As EventArgs) 
        ' Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        ' dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
        ' Clipboard.SetText(Math.Round(dimensionW.Value / aspectRatio))
    ' End Sub

    ' Private Sub calculateButton_ClickH(ByVal sender As Object, ByVal e As EventArgs) 
        ' Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        ' dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
        ' Clipboard.SetText(Math.Round(dimensionH.Value * aspectRatio))
    ' End Sub

    Private Sub calculateButton_KeyUpW(ByVal sender As Object, ByVal e As KeyEventArgs)
	    If Not (Char.IsDigit(ChrW(e.KeyCode)) OrElse e.KeyCode = Keys.Enter OrElse (e.KeyCode = Keys.V AndAlso e.Control) OrElse e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Delete OrElse e.KeyCode >= Keys.NumPad0 AndAlso e.KeyCode <= Keys.NumPad9) Then
            Return
		Else
            Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
            dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
            Clipboard.SetText(Math.Round(dimensionW.Value / aspectRatio))
        End If
    End Sub

    Private Sub calculateButton_KeyUpH(ByVal sender As Object, ByVal e As KeyEventArgs)
	    If Not (Char.IsDigit(ChrW(e.KeyCode)) OrElse e.KeyCode = Keys.Enter OrElse (e.KeyCode = Keys.V AndAlso e.Control) OrElse e.KeyCode = Keys.Back OrElse e.KeyCode = Keys.Delete OrElse e.KeyCode >= Keys.NumPad0 AndAlso e.KeyCode <= Keys.NumPad9) Then
            Return
		Else
            Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
            dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
            Clipboard.SetText(Math.Round(dimensionH.Value * aspectRatio))
        End If
    End Sub

' to make field empty while using KeyUp =================================================
    Private Sub calculateButton_TextChangedW(ByVal sender As Object, ByVal e As EventArgs)
        if  dimensionW.Text  = ""  Then
		dimensionH.Value = 0
		dimensionH.text = ""
    End If
	End Sub
	
    Private Sub calculateButton_TextChangedH(ByVal sender As Object, ByVal e As EventArgs)
        if  dimensionH.Text  = "" Then
		dimensionW.Value = 0
		dimensionW.text = ""
    End If
	End Sub
' Right-click > Paste =======================================================
    Private Sub dimension_MouseUpW(ByVal sender As Object, ByVal e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
            dimensionH.Text = Math.Round(dimensionW.Value / aspectRatio)
            Clipboard.SetText(Math.Round(dimensionW.Value / aspectRatio))
        End If
    End Sub
	
Private Sub dimension_MouseUpH(ByVal sender As Object, ByVal e As MouseEventArgs)
    If e.Button = MouseButtons.Right Then
        Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value
        dimensionW.Text = Math.Round(dimensionH.Value * aspectRatio)
        Clipboard.SetText(Math.Round(dimensionH.Value * aspectRatio))
    End If
End Sub
' Ctrl+A '=================================================================
    Private Sub dimension_KeyDownW(ByVal sender As Object, ByVal e As KeyEventArgs)
	    If e.Control AndAlso e.KeyCode = Keys.A Then
       dimensionW.Select(0, dimensionW.Text.Length)
	    End If
    End Sub

    Private Sub dimension_KeyDownH(ByVal sender As Object, ByVal e As KeyEventArgs)
	    If e.Control AndAlso e.KeyCode = Keys.A Then
       dimensionH.Select(0, dimensionH.Text.Length)
	    End If
    End Sub

' =================================================================
    ' closeButton
    Private Sub closeButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub
    ' minimizeButton
    Private Sub minimizeButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub
	' openWebsiteButton
	Private Sub openWebsiteButton1_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("https://github.com/amymor/Dimension-Calc")
    End Sub
' =================================================================
	' Drag & Drop1
    Private Sub Form_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub
	' Drag & Drop2
    Private Sub Form_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs) Handles Me.DragDrop
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
		' Fix
		dimensionW.text = 0
		dimensionH.text = 0
		
        For Each file As String In files
            Dim image As Image = Image.FromFile(file)
            Dim width As Integer = image.Width
            Dim height As Integer = image.Height
            Dim Image_aspectRatio As Double = width / height
            Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value

            If Image_aspectRatio < aspectRatio Then
                dimensionW.Value = width
                dimensionH.Value = CInt(width / aspectRatio)
            Else
                dimensionW.Value = CInt(height * aspectRatio)
                dimensionH.Value = height
            End If
            ' Copy the result to clipboard
            Clipboard.SetText(dimensionW.Value & " " & dimensionH.Value)
        Next
    End Sub

' =================================================================

    ' Entry point of the application
    Public Shared Sub Main()
        Application.Run(New Form1())
    End Sub
End Class
