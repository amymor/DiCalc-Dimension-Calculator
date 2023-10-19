Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class Form1
    Inherits Form
    ' Create TextBoxes
    Private aspectRatio1 As New NumericUpDown()
    Private aspectRatio2 As New NumericUpDown()
    Private dimension1 As New NumericUpDown()
    Private dimension2 As New NumericUpDown()
    ' Create Buttons
    Private calculateButton As New Button()
	Private closeButton As New Button()
    Private minimizeButton As New Button()
	Private openWebsiteButton1 As New Button()
	Private openWebsiteButton2 As New Button()
	Private toolTip1 As New ToolTip()
    ' Create Labels
    Private AspectRatio As New Label()
    Private ratioWidthLabel As New Label()
    Private ratioHeightLabel As New Label()
    Private pixelsWidthLabel As New Label()
    Private pixelsHeightLabel As New Label()
	
    Public Sub New()
        ' Set form properties
        ' Enable drag and drop
        Me.AllowDrop = True
        ' Add event handlers to Form
        AddHandler Me.DragEnter, AddressOf Form_DragEnter
        AddHandler Me.DragDrop, AddressOf Form_DragDrop
		
		Me.FormBorderStyle = FormBorderStyle.FixedDialog
		' Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
		' Me.showintaskbar = False
		Me.ControlBox = False
		Me.Size = New Size(220, 190)
		' Me.Text = "=)"
		Me.Text = "DCalc - amymor(OgomnamO)"
        Me.BackColor = Color.FromArgb(60,75,66)
        ' Me.BackColor = Color.FromArgb(60,66,75)
        Me.ForeColor = Color.FromArgb(235,235,235) ' Set text color (for Labels and Buttons if not set already)
        ' Me.ForeColor = Color.White
 '=================================================================
        ' Initialize Labels
        AspectRatio.Text = "Aspect ratio"
        AspectRatio.Location = New Point(15, 0)
        AspectRatio.Size = New Size(120, 30)
        AspectRatio.Font = New Font("Segoe UI", 14) ' Set font size
        AspectRatio.ForeColor = Color.FromArgb(238,190,123) ' Set text color
		
        ratioWidthLabel.Text = "Ratio width"
        ratioWidthLabel.Location = New Point(140, 10)
        ratioWidthLabel.Size = New Size(100, 25)
        ratioWidthLabel.Font = New Font("Segoe UI", 12) ' Set font size
        ratioHeightLabel.Text = "Ratio height"
        ratioHeightLabel.Location = New Point(140, 40)
        ratioHeightLabel.Size = New Size(100, 25)
        ratioHeightLabel.Font = New Font("Segoe UI", 12) ' Set font size
		
        pixelsWidthLabel.Text = "W"
        pixelsWidthLabel.Location = New Point(5, 82)
        pixelsWidthLabel.Size = New Size(100, 25)
        pixelsWidthLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        pixelsWidthLabel.ForeColor = Color.FromArgb(193,141,227)

        pixelsHeightLabel.Text = "H"
        pixelsHeightLabel.Location = New Point(5, 118)
        pixelsHeightLabel.Size = New Size(100, 25)
        pixelsHeightLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        pixelsHeightLabel.ForeColor = Color.FromArgb(151,215,151)
 '=================================================================
        ' Initialize TextBoxes
        aspectRatio1.Location = New Point(15, 25)
        aspectRatio1.Size = New Size(59, 20)
        aspectRatio1.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        aspectRatio1.BackColor = Color.FromArgb(54,69,60) ' Set background color
        ' aspectRatio1.BackColor = Color.FromArgb(54,60,69) ' Set background color
        aspectRatio1.ForeColor = Color.FromArgb(235,235,235) ' Set text color
		
        aspectRatio2.Location = New Point(76, 25)
        aspectRatio2.Size = New Size(59, 20)
        aspectRatio2.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        aspectRatio2.BackColor = Color.FromArgb(54,69,60) ' Set background color
        aspectRatio2.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        
		dimension1.Location = New Point(30, 80)
        dimension1.Size = New Size(105, 20)
        dimension1.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        dimension1.BackColor = Color.FromArgb(54,69,60) ' Set background color
        dimension1.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        
		dimension2.Location = New Point(30, 115)
        dimension2.Size = New Size(105, 20)
        dimension2.Font = New Font("Segoe UI", 14, FontStyle.Bold) ' Set font size
        dimension2.BackColor = Color.FromArgb(54,69,60) ' Set background color
        dimension2.ForeColor = Color.FromArgb(235,235,235) ' Set text color
		
		' TextBoxes BorderStyle
        aspectRatio1.BorderStyle = BorderStyle.FixedSingle
        aspectRatio2.BorderStyle = BorderStyle.FixedSingle
        dimension1.BorderStyle = BorderStyle.FixedSingle
        dimension2.BorderStyle = BorderStyle.FixedSingle
		
        ' NumericUpDown colors
        ' aspectRatio1.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' aspectRatio1.Controls("UpDownButtons").ForeColor = Color.White
        ' aspectRatio2.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' aspectRatio2.Controls("UpDownButtons").ForeColor = Color.White
        ' dimension1.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' dimension1.Controls("UpDownButtons").ForeColor = Color.White
        ' dimension2.Controls("UpDownButtons").BackColor = Color.DarkGray
        ' dimension2.Controls("UpDownButtons").ForeColor = Color.White
		
        ' maximum value of the data type
        dimension1.Maximum = 9999999
        dimension2.Maximum = 9999999
 '=================================================================
        ' Initialize Buttons
        calculateButton.Text = "Calculate"
        calculateButton.Location = New Point(5, 160)
        calculateButton.Size = New Size(130, 60)
        calculateButton.Font = New Font("Segoe UI", 18) ' Set font size
        ' calculateButton.ForeColor = Color.FromArgb(100,205,100) ' Set text color
        calculateButton.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        calculateButton.BackColor = Color.FromArgb(50,80,50) ' Set background color
        calculateButton.FlatStyle = FlatStyle.Flat
        calculateButton.FlatAppearance.BorderColor = Color.FromArgb(25,60,25)
        calculateButton.FlatAppearance.BorderSize = 1
		
        closeButton.Text = "❌"
		closeButton.Location =  New Point(182, 0)
        closeButton.Size = New Size(32, 32)
        closeButton.Font = New Font("Segoe UI", 16) ' Set font size
        closeButton.BackColor = Color.Brown
        closeButton.FlatStyle = FlatStyle.Flat
        closeButton.FlatAppearance.BorderColor = Color.FromArgb(50,50,50)
        closeButton.FlatAppearance.BorderSize = 1
		
        minimizeButton.Text = "—"
        minimizeButton.Location = New Point(150, 0)
        minimizeButton.Size = New Size(32, 32)
        minimizeButton.Font = New Font("Segoe UI", 16, FontStyle.Bold) ' Set font size
        ' minimizeButton.ForeColor = Color.FromArgb(100,205,100) ' Set text color
        ' minimizeButton.BackColor = Color.FromArgb(54,60,69) ' Set background color
        minimizeButton.BackColor = Color.SteelBlue ' Set background color
        minimizeButton.FlatStyle = FlatStyle.Flat
        minimizeButton.FlatAppearance.BorderColor = Color.FromArgb(50,50,50)
        minimizeButton.FlatAppearance.BorderSize = 1

        openWebsiteButton1.Text = "ℹ"
        openWebsiteButton1.Location = New Point(182, 128)
        openWebsiteButton1.Size = New Size(32, 32)
        openWebsiteButton1.Font = New Font("Segoe UI black", 24, FontStyle.Bold) ' Set font size
        openWebsiteButton1.ForeColor = Color.Aquamarine
        openWebsiteButton1.BackColor = Color.FromArgb(60,66,75)
        openWebsiteButton1.FlatStyle = FlatStyle.Flat
        openWebsiteButton1.FlatAppearance.BorderColor = Color.FromArgb(50,50,50)
        openWebsiteButton1.FlatAppearance.BorderSize = 1
		
        openWebsiteButton2.Text = "@"
        openWebsiteButton2.Location = New Point(148, 128)
        openWebsiteButton2.Size = New Size(32, 32)
        openWebsiteButton2.Font = New Font("Segoe UI", 16) ' Set font size
        openWebsiteButton2.ForeColor = Color.Aquamarine
        openWebsiteButton2.BackColor = Color.FromArgb(60,66,75)
        openWebsiteButton2.FlatStyle = FlatStyle.Flat
        openWebsiteButton2.FlatAppearance.BorderColor = Color.FromArgb(50,50,50)
        openWebsiteButton2.FlatAppearance.BorderSize = 1
 '=================================================================
        ' Add TextBoxes, Labels, and Buttons to Form
        Me.Controls.Add(aspectRatio1)
        Me.Controls.Add(aspectRatio2)
        Me.Controls.Add(dimension1)
        Me.Controls.Add(dimension2)
        Me.Controls.Add(AspectRatio)
        ' Me.Controls.Add(ratioWidthLabel)
        ' Me.Controls.Add(ratioHeightLabel)
        Me.Controls.Add(pixelsWidthLabel)
        Me.Controls.Add(pixelsHeightLabel)
        ' Me.Controls.Add(calculateButton)
		Me.Controls.Add(closeButton)
		Me.Controls.Add(minimizeButton)
		Me.Controls.Add(openWebsiteButton1)
		Me.Controls.Add(openWebsiteButton2)

        ' Set 16:9 as default
        aspectRatio1.text = 16
        aspectRatio2.text = 9
		dimension1.text = ""
		dimension2.text = ""
		' dimension1.text = 0
		' dimension2.text = 0

		' focus to the dimension1 TextBox at startup
		dimension1.Select()
		
		toolTip1.SetToolTip(pixelsWidthLabel, "Pixels width")
		toolTip1.SetToolTip(pixelsHeightLabel, "Pixels height")
		toolTip1.SetToolTip(closeButton, "Close")
		toolTip1.SetToolTip(minimizeButton, "Minimize")
		toolTip1.SetToolTip(openWebsiteButton1, "Click to open the GitHub webpage and check manually for updates")
		toolTip1.SetToolTip(openWebsiteButton2, "Click to open the Telegram channel(@PC_Mods_ir)")

        ' Add event handler to Button
        ' AddHandler calculateButton.Click, AddressOf calculateButton_Click
        ' minimizeButton
        AddHandler minimizeButton.Click, AddressOf minimizeButton_Click
        ' closeButton
        AddHandler closeButton.Click, AddressOf closeButton_Click
        ' openWebsiteButton
		AddHandler openWebsiteButton1.Click, AddressOf openWebsiteButton1_Click
		AddHandler openWebsiteButton2.Click, AddressOf openWebsiteButton2_Click
		' Add ValueChanged event handlers to NumericUpDown controls
        AddHandler aspectRatio1.ValueChanged, AddressOf calculateButton_Click
        AddHandler aspectRatio2.ValueChanged, AddressOf calculateButton_Click
        AddHandler dimension1.ValueChanged, AddressOf calculateButton_Click
        AddHandler dimension2.ValueChanged, AddressOf calculateButton_Click

    End Sub

    Private Sub calculateButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Try to parse the text from the TextBoxes to Doubles
        ' Dim aspectRatio1Value As Double
        ' Dim aspectRatio2Value As Double
        ' Dim dimension1Value As Double
        ' Dim dimension2Value As Double

        ' Check if the TextBoxes are empty and set the corresponding Double variables to zero if they are
        ' If Not Double.TryParse(aspectRatio1.Text, aspectRatio1Value) Then aspectRatio1Value = 0
        ' If Not Double.TryParse(aspectRatio2.Text, aspectRatio2Value) Then aspectRatio2Value = 0
        ' If Not Double.TryParse(dimension1.Text, dimension1Value) Then dimension1Value = 0
        ' If Not Double.TryParse(dimension2.Text, dimension2Value) Then dimension2Value = 0

        ' Calculate missing dimension
        Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value

        If dimension1.Text = "" Then
            dimension1.Text = Math.Round(dimension2.Value * aspectRatio).ToString()
        Else
            dimension2.Text = Math.Round(dimension1.Value / aspectRatio).ToString()
        End If
    End Sub

    ' closeButton
    Private Sub closeButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

    ' minimizeButton
    Private Sub minimizeButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub

	Private Sub openWebsiteButton1_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("https://github.com/amymor/Dimension-Calc")
    End Sub

	Private Sub openWebsiteButton2_Click(ByVal sender As Object, ByVal e As EventArgs)
        Process.Start("https://t.me/PC_Mods_ir")
    End Sub
	
	
    Private Sub Form_DragEnter(ByVal sender As Object, ByVal e As DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub


    Private Sub Form_DragDrop(ByVal sender As Object, ByVal e As DragEventArgs)
        Dim files() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())
		' Fix
		dimension1.text = 0
		dimension2.text = 0
		
        For Each file As String In files
            Dim image As Image = Image.FromFile(file)
            Dim width As Integer = image.Width
            Dim height As Integer = image.Height

            ' Calculate missing dimension
            Dim aspectRatio As Double = aspectRatio1.Value / aspectRatio2.Value

            If height > width Then
                dimension1.Value = width
                dimension2.Value = CInt(width / aspectRatio)
            Else
                dimension1.Value = CInt(height * aspectRatio)
                dimension2.Value = height
            End If

            ' Copy the result to clipboard
            Clipboard.SetText(dimension1.Value.ToString() & " " & dimension2.Value.ToString())
        Next
    End Sub


    ' Entry point of the application
    <STAThread>
    Public Shared Sub Main()
        Application.Run(New Form1())
    End Sub
End Class
