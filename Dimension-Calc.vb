Imports System
Imports System.Diagnostics
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms

Public Class Form1
    Inherits Form
    ' Create TextBoxes
    Private aspectRatio1 As New TextBox()
    Private aspectRatio2 As New TextBox()
    Private dimension1 As New TextBox()
    Private dimension2 As New TextBox()
    ' Create Buttons
    Private calculateButton As New Button()
    Private minimizeButton As New Button()
	Private closeButton As New Button()
	Private openWebsiteButton1 As New Button()
	Private openWebsiteButton2 As New Button()
    ' Create Labels
    Private AppLabel As New Label()
    Private ratioWidthLabel As New Label()
    Private ratioHeightLabel As New Label()
    Private pixelsWidthLabel As New Label()
    Private pixelsHeightLabel As New Label()
	
    Public Sub New()
        ' Set form properties
		Me.FormBorderStyle = FormBorderStyle.FixedDialog
		' Me.FormBorderStyle = FormBorderStyle.FixedToolWindow
		' Me.showintaskbar = False
		Me.ControlBox = False
		Me.Size = New Size(300, 280)
		' Me.Text = "=)"
		Me.Text = "Dimension Calc by amymor(OgomnamO)"
        Me.BackColor = Color.FromArgb(60,75,66)
        ' Me.BackColor = Color.FromArgb(60,66,75)
        Me.ForeColor = Color.FromArgb(235,235,235) ' Set text color (for Labels and Buttons if not set already)
        ' Me.ForeColor = Color.White

        ' Initialize TextBoxes
        aspectRatio1.Location = New Point(20, 30)
        aspectRatio1.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        aspectRatio1.BackColor = Color.FromArgb(54,69,60) ' Set background color
        ' aspectRatio1.BackColor = Color.FromArgb(54,60,69) ' Set background color
        aspectRatio1.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        aspectRatio2.Location = New Point(20, 58)
        aspectRatio2.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        aspectRatio2.BackColor = Color.FromArgb(54,69,60) ' Set background color
        aspectRatio2.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        dimension1.Location = New Point(20, 110)
        dimension1.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        dimension1.BackColor = Color.FromArgb(54,69,60) ' Set background color
        dimension1.ForeColor = Color.FromArgb(235,235,235) ' Set text color
        dimension2.Location = New Point(20, 138)
        dimension2.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        dimension2.BackColor = Color.FromArgb(54,69,60) ' Set background color
        dimension2.ForeColor = Color.FromArgb(235,235,235) ' Set text color
		' TextBoxes BorderStyle
        aspectRatio1.BorderStyle = BorderStyle.FixedSingle
        aspectRatio2.BorderStyle = BorderStyle.FixedSingle
        dimension1.BorderStyle = BorderStyle.FixedSingle
        dimension2.BorderStyle = BorderStyle.FixedSingle

        ' Initialize Labels
        AppLabel.Size = New Size(260, 30)
        AppLabel.Location = New Point(5, 5)
        AppLabel.Font = New Font("Segoe UI", 10) ' Set font size
        AppLabel.ForeColor = Color.FromArgb(170,225,170) ' Set text color
        AppLabel.Text = "Dimension Calc by amymor(OgomnamO)"
        ratioWidthLabel.Location = New Point(130, 30)
        ratioWidthLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        ratioWidthLabel.Size = New Size(140, 25)
        ratioWidthLabel.Text = "Ratio width"
        ratioHeightLabel.Location = New Point(130, 60)
        ratioHeightLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        ratioHeightLabel.Size = New Size(140, 25)
        ratioHeightLabel.Text = "Ratio height"
        pixelsWidthLabel.Location = New Point(130, 110)
        pixelsWidthLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        pixelsWidthLabel.Size = New Size(140, 25)
        pixelsWidthLabel.Text = "Pixels width"
        pixelsHeightLabel.Location = New Point(130, 140)
        pixelsHeightLabel.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        pixelsHeightLabel.Size = New Size(140, 25)
        pixelsHeightLabel.Text = "Pixels height"

        ' Initialize Buttons
        calculateButton.Location = New Point(10, 180)
        calculateButton.Size = New Size(270, 60)
        calculateButton.Text = "Calculate"
        calculateButton.Font = New Font("Segoe UI", 24, FontStyle.Bold) ' Set font size
        calculateButton.ForeColor = Color.FromArgb(100,205,100) ' Set text color
        ' calculateButton.BackColor = Color.FromArgb(54,60,69) ' Set background color
        calculateButton.FlatStyle = FlatStyle.Flat
        ' calculateButton.FlatAppearance.BorderColor = Color.Red
        ' calculateButton.FlatAppearance.BorderSize = 1
		
        minimizeButton.Location = New Point(270, 30)
        minimizeButton.Size = New Size(24, 24)
        minimizeButton.Text = "—"
        minimizeButton.Font = New Font("Segoe UI", 12, FontStyle.Bold) ' Set font size
        ' minimizeButton.ForeColor = Color.FromArgb(100,205,100) ' Set text color
        ' minimizeButton.BackColor = Color.FromArgb(54,60,69) ' Set background color
        minimizeButton.BackColor = Color.DodgerBlue ' Set background color
        minimizeButton.FlatStyle = FlatStyle.Flat
        minimizeButton.FlatAppearance.BorderColor = Color.DimGray
        minimizeButton.FlatAppearance.BorderSize = 1
		
		' closeButton.Location = New Point(Me.Width - 30, 0)
		closeButton.Location =  New Point(270, 0)
        closeButton.Size = New Size(24, 24)
        closeButton.Text = "❌"
        closeButton.Font = New Font("Segoe UI", 12) ' Set font size
        closeButton.BackColor = Color.Red
        closeButton.FlatStyle = FlatStyle.Flat
        closeButton.FlatAppearance.BorderColor = Color.black
        closeButton.FlatAppearance.BorderSize = 1

        openWebsiteButton1.Location = New Point(10, 190)
        openWebsiteButton1.Text = "Open Website"
        openWebsiteButton1.Font = New Font("Arial", 12) ' Set font size
		
        openWebsiteButton2.Location = New Point(10, 190)
        openWebsiteButton2.Text = "Open Website"
        openWebsiteButton2.Font = New Font("Arial", 12) ' Set font size

        ' Add TextBoxes, Labels, and Button to Form
        Me.Controls.Add(aspectRatio1)
        Me.Controls.Add(aspectRatio2)
        Me.Controls.Add(dimension1)
        Me.Controls.Add(dimension2)
        ' Me.Controls.Add(AppLabel)
        Me.Controls.Add(ratioWidthLabel)
        Me.Controls.Add(ratioHeightLabel)
        Me.Controls.Add(pixelsWidthLabel)
        Me.Controls.Add(pixelsHeightLabel)
        Me.Controls.Add(calculateButton)
		Me.Controls.Add(minimizeButton)
		Me.Controls.Add(closeButton)

        ' Set 16:9 as default
        aspectRatio1.text = 16
        aspectRatio2.text = 9

        ' Add event handler to Button
        AddHandler calculateButton.Click, AddressOf calculateButton_Click

    ' minimizeButton
        AddHandler minimizeButton.Click, AddressOf minimizeButton_Click

    ' closeButton
        AddHandler closeButton.Click, AddressOf closeButton_Click
    End Sub

    Private Sub calculateButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Try to parse the text from the TextBoxes to Doubles
        Dim aspectRatio1Value As Double
        Dim aspectRatio2Value As Double
        Dim dimension1Value As Double
        Dim dimension2Value As Double

        ' Check if the TextBoxes are empty and set the corresponding Double variables to zero if they are
        If Not Double.TryParse(aspectRatio1.Text, aspectRatio1Value) Then aspectRatio1Value = 0
        If Not Double.TryParse(aspectRatio2.Text, aspectRatio2Value) Then aspectRatio2Value = 0
        If Not Double.TryParse(dimension1.Text, dimension1Value) Then dimension1Value = 0
        If Not Double.TryParse(dimension2.Text, dimension2Value) Then dimension2Value = 0

        ' Calculate missing dimension
        Dim aspectRatio As Double = aspectRatio1Value / aspectRatio2Value

        If dimension1.Text = "" Then
            dimension1.Text = (dimension2Value * aspectRatio).ToString()
        Else
            dimension2.Text = (dimension1Value / aspectRatio).ToString()
        End If
    End Sub

    ' minimizeButton
    Private Sub minimizeButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.WindowState = FormWindowState.Minimized
    End Sub

    ' closeButton
    Private Sub closeButton_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.Close()
    End Sub

	
    ' Entry point of the application
    <STAThread>
    Public Shared Sub Main()
        Application.Run(New Form1())
    End Sub
End Class
