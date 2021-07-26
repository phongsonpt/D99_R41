Imports System.ComponentModel
Imports System.Drawing.Drawing2D
' -----------------------------------------------------------------------------
' Copyright © 2008 Timothy C. Alvord
'
' This source code is provided AS IS, with no warranty expressed or implied,
' including without limitation, warranties of merchantability or 
' fitness for a particular purpose or any warranty of title or 
' non-infringement. This disclaimer must be passed on whenever the Software 
' is distributed in either source form or as a derivative works.
' It may be used for both commercial and non-commercial applications.
'
' -----------------------------------------------------------------------------
' Notes:
' Properties Added:
'    GradientOptions        
'       Gradient            - True, False
'       Orientation         - BackwardDiagonal, ForwardDiagonal,
'                             Horizontal, Vertical
'       StartColor          - Start Gradient Color
'       EndColor            - End Gradient Color
'    ButtonBorderColor      - Color around the border of the button
'    ButtonBackColor        - Color of the button background
'    ButtonRadius           - Corner Radius for button
'    ButtonImage            - Image for Background
'    ButtonImageLayout      - Center, None, Stretch, Tile, Zoom
'    TextAlign              - TopLeft, TopCenter, TopRight,
'                             MiddleLeft, MiddleCenter, MiddleRight,
'                             BottomLeft, BottomCenter, BottomRight
'    FlatStyle              - Flat:         The button appears flat
'                             Popup:        The button appears flat until the
'                                           mouse pointer moves over it, at which
'                                           point it appears three-dimensional
'                             Standard:     The button appears three-dimensional
'                             System:       The appearance of the button is
'                                           determined by the operating system
'                             Traditional:  The button appears to go up and down
'                                           when clicked
'    FlatAppearance         - 
' --------------------------------------------------------------------------------
' Version:
' --------------------------------------------------------------------------------
' 1.0 2007-07-10 TCA
'     Initial Release
' --------------------------------------------------------------------------------

<ToolboxBitmap(GetType(CustomButton), "CustomButton.bmp")> _
Public Class CustomButton : Inherits Control

    Private m_AutoSize As Boolean
    Private m_AutoSizeMode As AutoSizeMode
    Private m_OriginalSize As Size
    Private m_ButtonBorderColor As Color
    Private m_ButtonBackColor As Color
    Private m_ButtonRadius As Integer
    Private m_ButtonImage As System.Drawing.Image = Nothing
    Private m_ButtonImageLayout As ImageLayout
    Private m_DialogResult As Windows.Forms.DialogResult
    Private m_TextAlign As ContentAlignment
    Private m_FlatStyle As EnumFlatStyle
    Private m_FlatAppearance As FlatButtonAppearance

    Private m_GradientOptions As New GradientOptions

    Public Enum EnumFlatStyle
        Flat = 0
        Popup = 1
        Standard = 2
        Traditional = 3
    End Enum

    Sub New()
        ' ---------------------------------------------------------------------
        ' New Button
        ' ---------------------------------------------------------------------
        '
        '
        ' Setup defaults
        ButtonBorderColor = Color.RosyBrown ' RoyalBlue
        ButtonBackColor = Color.Transparent
        ButtonRadius = 0
        ButtonImage = Nothing
        FlatStyle = EnumFlatStyle.Standard
        TextAlign = ContentAlignment.MiddleCenter
        AutoSize = False
        AutoSizeMode = Windows.Forms.AutoSizeMode.GrowOnly
        Width = 85
        Height = 23
    End Sub

    <Category("Appearance"), _
    Description("Expand to specifiy the Gradient Options of the button."), _
    DesignerSerializationVisibility(DesignerSerializationVisibility.Content)> _
    Public Property Gradient() As GradientOptions
        Get
            Gradient = m_GradientOptions
        End Get
        Set(ByVal Value As GradientOptions)
            m_GradientOptions = Value
        End Set
    End Property

    <CategoryAttribute("Layout"), DefaultValue(False), _
    System.ComponentModel.Browsable(True), EditorBrowsable(EditorBrowsableState.Always), _
    DesignerSerializationVisibility(DesignerSerializationVisibility.Visible), _
    DescriptionAttribute("Specifies whether the button will automatically size itself to fit its contents.")> _
    Public Overrides Property AutoSize() As Boolean
        Get
            AutoSize = m_AutoSize
        End Get
        Set(ByVal value As Boolean)
            m_AutoSize = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Layout"), DefaultValue(AutoSizeMode.GrowOnly), _
    DescriptionAttribute("Specifies whether the button will automatically size itself to fit its contents.")> _
    Public Property AutoSizeMode() As AutoSizeMode
        Get
            AutoSizeMode = m_AutoSizeMode
        End Get
        Set(ByVal value As AutoSizeMode)
            m_AutoSizeMode = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Button"), DefaultValueAttribute(""), _
    DescriptionAttribute("Color around the border of the button.")> _
    Public Property ButtonBorderColor() As Color
        Get
            ButtonBorderColor = m_ButtonBorderColor
        End Get
        Set(ByVal value As Color)
            m_ButtonBorderColor = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Button"), DefaultValueAttribute(""), _
    DescriptionAttribute("Color around the border of the button.")> _
    Public Property ButtonBackColor() As Color
        Get
            ButtonBackColor = m_ButtonBackColor
        End Get
        Set(ByVal value As Color)
            m_ButtonBackColor = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Button"), DefaultValueAttribute(""), _
    DescriptionAttribute("Corner Radius for button.")> _
    Public Property ButtonRadius() As Integer
        Get
            ButtonRadius = m_ButtonRadius
        End Get
        Set(ByVal value As Integer)
            m_ButtonRadius = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Button"), DefaultValueAttribute(""), _
    DescriptionAttribute("Image for Button.")> _
    Public Property ButtonImage() As System.Drawing.Image
        Get
            ButtonImage = m_ButtonImage
        End Get
        Set(ByVal value As System.Drawing.Image)
            m_ButtonImage = value
            m_FlatStyle = EnumFlatStyle.Traditional
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Button"), DefaultValueAttribute(ImageLayout.Center), _
    DescriptionAttribute("Center, None, Stretch, Tile, Zoom.")> _
    Public Property ButtonImageLayout() As ImageLayout
        Get
            ButtonImageLayout = m_ButtonImageLayout
        End Get
        Set(ByVal value As ImageLayout)
            m_ButtonImageLayout = value
            Me.Invalidate()
        End Set
    End Property

    <CategoryAttribute("Behavior"), DefaultValueAttribute(Windows.Forms.DialogResult.None), _
    DescriptionAttribute("The dialog-result produced in a modal form by clicking the button.")> _
    Public Property DialogResult() As Windows.Forms.DialogResult
        Get
            DialogResult = m_DialogResult
        End Get
        Set(ByVal value As Windows.Forms.DialogResult)
            m_DialogResult = value
            Me.Invalidate()
        End Set
    End Property

    <DescriptionAttribute("Text Alignment for Button Text."), DefaultValueAttribute(ContentAlignment.MiddleCenter)> _
    Public Property TextAlign() As ContentAlignment
        Get
            TextAlign = m_TextAlign
        End Get
        Set(ByVal value As ContentAlignment)
            m_TextAlign = value
            Me.Invalidate()
        End Set
    End Property

    <DescriptionAttribute("Specifies the appearance of the Button."), DefaultValueAttribute(EnumFlatStyle.Traditional)> _
    Public Property FlatStyle() As EnumFlatStyle
        Get
            FlatStyle = m_FlatStyle
        End Get
        Set(ByVal value As EnumFlatStyle)
            m_FlatStyle = value
            Me.Invalidate()
        End Set
    End Property

    '<DescriptionAttribute("Specifies the appearance of the Button."), DefaultValueAttribute("")> _
    'Public Property FlatAppearance() As FlatButtonAppearance
    '    Get
    '        m_FlatAppearance = New Button().FlatAppearance
    '        FlatAppearance = m_FlatAppearance
    '    End Get
    '    Set(ByVal value As FlatButtonAppearance)
    '        'm_FlatAppearance.BorderColor = value.BorderColor
    '        'm_FlatAppearance.BorderSize = value.BorderSize
    '        'm_FlatAppearance.CheckedBackColor = value.CheckedBackColor
    '        'm_FlatAppearance.MouseDownBackColor = value.MouseDownBackColor
    '        'm_FlatAppearance.MouseOverBackColor = value.MouseOverBackColor
    '        'Me.Invalidate()
    '    End Set
    'End Property

    Protected Overrides Sub OnCreateControl()
        MyBase.OnCreateControl()
        m_OriginalSize = Me.Size
    End Sub

    Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
        MyBase.OnResize(e)
        If m_AutoSize <> True Then
            m_OriginalSize = Me.Size
        End If
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        Dim brushText As SolidBrush
        Dim iXPos, iYPos As Integer
        Dim DrawRect As New Rectangle
        Dim LayoutSize As New SizeF
        Dim Format As New StringFormat

        If m_AutoSize = True Then
            LayoutSize = e.Graphics.MeasureString(Me.Text, Me.Font)
            Me.Width = CInt(LayoutSize.Width + 10)
            If m_AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink Then
                Me.Height = CInt(LayoutSize.Height + 10)
            End If
        Else
            Me.Size = m_OriginalSize
            LayoutSize = e.Graphics.MeasureString(Me.Text, Me.Font, Me.Size)
        End If

        If m_ButtonImage IsNot Nothing Then
            DrawRect = New Rectangle(0, 0, m_ButtonImage.Width, m_ButtonImage.Height)
            e.Graphics.DrawImage(m_ButtonImage, DrawRect)
            Select Case m_ButtonImageLayout
                Case ImageLayout.Center
                    Dim brushBackground As SolidBrush

                    brushBackground = New SolidBrush(m_ButtonBackColor)
                    e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                    brushBackground.Dispose()
                    DrawRect = New Rectangle(CInt((Me.Width - m_ButtonImage.Width) / 2), CInt((Me.Height - m_ButtonImage.Height) / 2), m_ButtonImage.Width, m_ButtonImage.Height)
                    e.Graphics.DrawImage(m_ButtonImage, DrawRect)
                Case ImageLayout.None
                    Dim brushBackground As SolidBrush

                    brushBackground = New SolidBrush(m_ButtonBackColor)
                    e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                    brushBackground.Dispose()
                    e.Graphics.DrawImage(m_ButtonImage, New Rectangle(0, 0, m_ButtonImage.Width, m_ButtonImage.Height))
                Case ImageLayout.Stretch
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    e.Graphics.DrawImage(m_ButtonImage, DrawRect)
                Case ImageLayout.Tile
                    Dim iLeft, iTop As Int16

                    iTop = 0
                    While iTop < Me.Height
                        iLeft = 0
                        While iLeft < Me.Width
                            e.Graphics.DrawImage(m_ButtonImage, New Rectangle(iLeft, iTop, m_ButtonImage.Width, m_ButtonImage.Height))
                            iLeft = CShort(iLeft + m_ButtonImage.Width)
                        End While
                        iTop = CShort(iTop + m_ButtonImage.Height)
                    End While
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                Case ImageLayout.Zoom
                    Dim dImageWidthRatio, dImageHeightRatio As Double
                    Dim brushBackground As SolidBrush

                    brushBackground = New SolidBrush(m_ButtonBackColor)
                    e.Graphics.FillRectangle(brushBackground, DrawRect)
                    brushBackground.Dispose()

                    dImageWidthRatio = Me.Width / m_ButtonImage.Width
                    dImageHeightRatio = Me.Height / m_ButtonImage.Height
                    If dImageWidthRatio < dImageHeightRatio Then
                        e.Graphics.DrawImage(m_ButtonImage, New Rectangle(CInt((Me.Width - m_ButtonImage.Width * dImageWidthRatio) / 2), CInt((Me.Height - m_ButtonImage.Height * dImageWidthRatio) / 2), CInt(m_ButtonImage.Width * dImageWidthRatio), CInt(m_ButtonImage.Height * dImageWidthRatio)))
                    Else
                        e.Graphics.DrawImage(m_ButtonImage, New Rectangle(CInt((Me.Width - m_ButtonImage.Width * dImageHeightRatio) / 2), CInt((Me.Height - m_ButtonImage.Height * dImageHeightRatio) / 2), CInt(m_ButtonImage.Width * dImageHeightRatio), CInt(m_ButtonImage.Height * dImageHeightRatio)))
                    End If
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
            End Select
        Else
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias
            Select Case m_FlatStyle
                Case EnumFlatStyle.Flat
                    Dim BorderPen As New Pen(m_ButtonBorderColor)

                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            e.Graphics.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            e.Graphics.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Popup
                    Dim BorderPen As New Pen(Color.FromArgb(165, Me.BackColor.R, Me.BackColor.G, Me.BackColor.B))

                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            e.Graphics.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            e.Graphics.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Standard
                    Dim BorderPen As New Pen(m_ButtonBorderColor)
                    Dim pathRoundedCorners As New GraphicsPath

                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_ButtonBackColor = Color.Transparent Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush
                        Dim EndGradientColor As Color

                        EndGradientColor = Color.FromArgb(255, 171, 189, 211)
                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), Color.White, EndGradientColor, LinearGradientMode.Vertical)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    End If
                    e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                    pathRoundedCorners = RoundedCorners(New Rectangle(1, 1, Me.Width - 3, Me.Height - 3), 3)
                    BorderPen = New Pen(Color.FromArgb(255, 45, 80, 130))
                    e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                    BorderPen = New Pen(Color.White)
                    BorderPen.Width = 2
                    e.Graphics.DrawRectangle(BorderPen, New Rectangle(3, 3, Me.Width - 6, Me.Height - 6))
                    pathRoundedCorners.Dispose()
                Case EnumFlatStyle.Traditional
                    Dim BorderPen As New Pen(m_ButtonBorderColor)

                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            e.Graphics.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            e.Graphics.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            e.Graphics.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            e.Graphics.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            e.Graphics.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()

            End Select
        End If
        If Me.Enabled Then
            brushText = New SolidBrush(Me.ForeColor)
        Else
            brushText = New SolidBrush(Color.Gray)
        End If

        Select Case m_TextAlign
            Case ContentAlignment.TopLeft
                Format.LineAlignment = StringAlignment.Near
                Format.Alignment = StringAlignment.Near
                iXPos = DrawRect.X + 2
                iYPos = DrawRect.Y + 2
            Case ContentAlignment.TopCenter
                Format.LineAlignment = StringAlignment.Near
                Format.Alignment = StringAlignment.Center
                iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                iYPos = DrawRect.Y + 2
            Case ContentAlignment.TopRight
                Format.LineAlignment = StringAlignment.Near
                Format.Alignment = StringAlignment.Far
                iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                iYPos = DrawRect.Y + 2
            Case ContentAlignment.MiddleLeft
                Format.LineAlignment = StringAlignment.Center
                Format.Alignment = StringAlignment.Near
                iXPos = DrawRect.X + 2
                iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
            Case ContentAlignment.MiddleCenter
                Format.LineAlignment = StringAlignment.Center
                Format.Alignment = StringAlignment.Center
                iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
            Case ContentAlignment.MiddleRight
                Format.LineAlignment = StringAlignment.Center
                Format.Alignment = StringAlignment.Far
                iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
            Case ContentAlignment.BottomLeft
                Format.LineAlignment = StringAlignment.Far
                Format.Alignment = StringAlignment.Near
                iXPos = DrawRect.X + 2
                iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
            Case ContentAlignment.BottomCenter
                Format.LineAlignment = StringAlignment.Far
                Format.Alignment = StringAlignment.Center
                iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
            Case ContentAlignment.BottomRight
                Format.LineAlignment = StringAlignment.Far
                Format.Alignment = StringAlignment.Far
                iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
        End Select
        e.Graphics.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos, iYPos, LayoutSize.Width, LayoutSize.Height), Format)
        brushText.Dispose()

    End Sub

    Private Sub CustomButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Click
        If Me.DialogResult <> Windows.Forms.DialogResult.None Then
            Dim ParentForm As Form = FindForm()

            ' Clear any remaining events
            Me.Events.Dispose()

            ParentForm.DialogResult = Me.DialogResult
            ParentForm.close()
        End If
    End Sub

    Public Sub DoMouseDown(ByVal e As MouseEventArgs)
        OnMouseDown(e)
    End Sub

    Public Sub DoMouseUp(ByVal e As MouseEventArgs)
        OnMouseUp(e)
    End Sub

    Public Sub CustomButton_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        Dim g As Graphics = Me.CreateGraphics()
        Dim brushText As SolidBrush
        Dim iXPos, iYPos As Integer
        Dim DrawRect As New Rectangle
        Dim BorderPen As New Pen(m_ButtonBorderColor)
        Dim LayoutSize As New SizeF
        Dim Format As New StringFormat

        If m_AutoSize = True Then
            LayoutSize = g.MeasureString(Me.Text, Me.Font)
            Me.Width = CInt(LayoutSize.Width + 10)
            If m_AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink Then
                Me.Height = CInt(LayoutSize.Height + 10)
            End If
        Else
            Me.Size = m_OriginalSize
            LayoutSize = g.MeasureString(Me.Text, Me.Font, Me.Size)
        End If

        g.SmoothingMode = SmoothingMode.AntiAlias
        Select Case m_FlatStyle
            Case EnumFlatStyle.Flat
                DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                Dim brushBackground As SolidBrush
                brushBackground = New SolidBrush(Color.FromArgb(127, Me.ForeColor.R, Me.ForeColor.G, Me.ForeColor.B))
                If m_GradientOptions.Background = True Then

                    If m_ButtonRadius > 0 Then
                        Dim pathRoundedCorners As New GraphicsPath

                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        g.DrawPath(BorderPen, pathRoundedCorners)
                        pathRoundedCorners.Dispose()
                    Else
                        g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                        g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                    End If
                    brushBackground.Dispose()
                Else
                    If m_ButtonRadius > 0 Then
                        Dim pathRoundedCorners As New GraphicsPath

                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        g.DrawPath(BorderPen, pathRoundedCorners)
                        pathRoundedCorners.Dispose()
                    Else
                        g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                        g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                    End If
                    brushBackground.Dispose()
                End If
                BorderPen.Dispose()
                If Me.Enabled Then
                    brushText = New SolidBrush(Me.ForeColor)
                Else
                    brushText = New SolidBrush(Color.Gray)
                End If

                Select Case m_TextAlign
                    Case ContentAlignment.TopLeft
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.TopCenter
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.TopRight
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.MiddleLeft
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.MiddleCenter
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.MiddleRight
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.BottomLeft
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                    Case ContentAlignment.BottomCenter
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                    Case ContentAlignment.BottomRight
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                End Select
                g.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos, iYPos, LayoutSize.Width, LayoutSize.Height), Format)
                brushText.Dispose()
            Case EnumFlatStyle.Popup
                DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                If m_GradientOptions.Background = True Then
                    Dim brushBackground As Drawing2D.LinearGradientBrush

                    brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                    If m_ButtonRadius > 0 Then
                        Dim pathRoundedCorners As New GraphicsPath

                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        g.DrawPath(BorderPen, pathRoundedCorners)
                        pathRoundedCorners.Dispose()
                    Else
                        g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                        g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                    End If
                    brushBackground.Dispose()
                Else
                    Dim brushBackground As SolidBrush

                    brushBackground = New SolidBrush(m_ButtonBackColor)
                    If m_ButtonRadius > 0 Then
                        Dim pathRoundedCorners As New GraphicsPath

                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        g.DrawPath(BorderPen, pathRoundedCorners)
                        pathRoundedCorners.Dispose()
                    Else
                        g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                        g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                    End If
                    brushBackground.Dispose()
                End If
                BorderPen.Dispose()
                If Me.Enabled Then
                    brushText = New SolidBrush(Me.ForeColor)
                Else
                    brushText = New SolidBrush(Color.Gray)
                End If
                Select Case m_TextAlign
                    Case ContentAlignment.TopLeft
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.TopCenter
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.TopRight
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.MiddleLeft
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.MiddleCenter
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.MiddleRight
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.BottomLeft
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                    Case ContentAlignment.BottomCenter
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                    Case ContentAlignment.BottomRight
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                End Select
                g.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos + 1, iYPos + 1, LayoutSize.Width, LayoutSize.Height), Format)
                brushText.Dispose()
            Case EnumFlatStyle.Standard
                Dim pathRoundedCorners As New GraphicsPath

                BorderPen = New Pen(m_ButtonBorderColor)
                DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                If m_ButtonBackColor = Color.Transparent Then
                    Dim brushBackground As Drawing2D.LinearGradientBrush
                    Dim EndGradientColor As Color

                    EndGradientColor = Color.FromArgb(255, 171, 189, 211)
                    brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), EndGradientColor, Color.White, LinearGradientMode.Vertical)
                    pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                    g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                    brushBackground.Dispose()
                Else
                    Dim brushBackground As SolidBrush

                    brushBackground = New SolidBrush(m_ButtonBackColor)
                    pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                    g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                    brushBackground.Dispose()
                End If
                g.DrawPath(BorderPen, pathRoundedCorners)
                pathRoundedCorners = RoundedCorners(New Rectangle(1, 1, Me.Width - 3, Me.Height - 3), 3)
                BorderPen = New Pen(Color.FromArgb(255, 45, 80, 130))
                g.DrawPath(BorderPen, pathRoundedCorners)
                BorderPen = New Pen(Color.White)
                BorderPen.Width = 2
                g.DrawRectangle(BorderPen, New Rectangle(3, 3, Me.Width - 6, Me.Height - 6))
                pathRoundedCorners.Dispose()
                BorderPen.Dispose()
                If Me.Enabled Then
                    brushText = New SolidBrush(Me.ForeColor)
                Else
                    brushText = New SolidBrush(Color.Gray)
                End If
                Select Case m_TextAlign
                    Case ContentAlignment.TopLeft
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.TopCenter
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.TopRight
                        Format.LineAlignment = StringAlignment.Near
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = DrawRect.Y + 2
                    Case ContentAlignment.MiddleLeft
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.MiddleCenter
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.MiddleRight
                        Format.LineAlignment = StringAlignment.Center
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                    Case ContentAlignment.BottomLeft
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Near
                        iXPos = DrawRect.X + 2
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                    Case ContentAlignment.BottomCenter
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Center
                        iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                    Case ContentAlignment.BottomRight
                        Format.LineAlignment = StringAlignment.Far
                        Format.Alignment = StringAlignment.Far
                        iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                        iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                End Select
                g.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos, iYPos, LayoutSize.Width, LayoutSize.Height), Format)
                brushText.Dispose()
            Case EnumFlatStyle.Traditional
                Me.Left = Me.Left + 1
                Me.Top = Me.Top + 1
        End Select
    End Sub

    Public Sub CustomButton_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        Try

            Dim g As Graphics = Me.CreateGraphics()
            Dim brushText As SolidBrush
            Dim iXPos, iYPos As Integer
            Dim DrawRect As New Rectangle
            Dim BorderPen As New Pen(m_ButtonBorderColor)
            Dim LayoutSize As New SizeF
            Dim Format As New StringFormat

            If m_AutoSize = True Then
                LayoutSize = g.MeasureString(Me.Text, Me.Font)
                Me.Width = CInt(LayoutSize.Width + 10)
                If m_AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink Then
                    Me.Height = CInt(LayoutSize.Height + 10)
                End If
            Else
                Me.Size = m_OriginalSize
                LayoutSize = g.MeasureString(Me.Text, Me.Font, Me.Size)
            End If

            g.SmoothingMode = SmoothingMode.AntiAlias
            Select Case m_FlatStyle
                Case EnumFlatStyle.Flat
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Popup
                    BorderPen = New Pen(Color.FromArgb(165, Me.BackColor.R, Me.BackColor.G, Me.BackColor.B))

                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Standard
                    Dim pathRoundedCorners As New GraphicsPath

                    BorderPen = New Pen(m_ButtonBorderColor)
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_ButtonBackColor = Color.Transparent Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush
                        Dim EndGradientColor As Color

                        EndGradientColor = Color.FromArgb(255, 171, 189, 211)
                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), Color.White, EndGradientColor, LinearGradientMode.Vertical)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    End If
                    g.DrawPath(BorderPen, pathRoundedCorners)
                    pathRoundedCorners = RoundedCorners(New Rectangle(1, 1, Me.Width - 3, Me.Height - 3), 3)
                    BorderPen = New Pen(Color.FromArgb(255, 45, 80, 130))
                    g.DrawPath(BorderPen, pathRoundedCorners)
                    BorderPen = New Pen(Color.FromArgb(255, 239, 207, 134))
                    BorderPen.Width = 2
                    g.DrawRectangle(BorderPen, New Rectangle(3, 3, Me.Width - 6, Me.Height - 6))
                    pathRoundedCorners.Dispose()
                    BorderPen.Dispose()
                Case EnumFlatStyle.Traditional
                    Me.Left = Me.Left - 1
                    Me.Top = Me.Top - 1
                    Exit Sub
            End Select

            If Me.Enabled Then
                brushText = New SolidBrush(Me.ForeColor)
            Else
                brushText = New SolidBrush(Color.Gray)
            End If

            Select Case m_TextAlign
                Case ContentAlignment.TopLeft
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.TopCenter
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.TopRight
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.MiddleLeft
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.MiddleCenter
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.MiddleRight
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.BottomLeft
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                Case ContentAlignment.BottomCenter
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                Case ContentAlignment.BottomRight
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
            End Select
            g.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos, iYPos, LayoutSize.Width, LayoutSize.Height), Format)
            brushText.Dispose()
        Catch ex As Exception
            
        End Try
    End Sub

    Private Function RoundedCorners(ByVal rect As Rectangle, ByVal Radius As Integer) As GraphicsPath
        Dim path As New GraphicsPath
        Dim rectRoundBounds As New Rectangle

        path.StartFigure()
        path.AddLine(rect.X + Radius, rect.Y, rect.X + rect.Width - 2 * Radius, rect.Y)
        path.AddArc(rect.X + rect.Width - 2 * Radius, rect.Y, 2 * Radius, 2 * Radius, 270, 90)
        path.AddLine(rect.X + rect.Width, rect.Y + Radius, rect.X + rect.Width, rect.Y + rect.Height - 2 * Radius)
        path.AddArc(rect.X + rect.Width - 2 * Radius, rect.Y + rect.Height - 2 * Radius, 2 * Radius, 2 * Radius, 0, 90)
        path.AddLine(rect.X + rect.Width - 2 * Radius, rect.Y + rect.Height, rect.X + Radius, rect.Y + rect.Height)
        path.AddArc(rect.X, rect.Y + rect.Height - 2 * Radius, 2 * Radius, 2 * Radius, 90, 90)
        path.AddLine(rect.X, rect.Y + rect.Height - 2 * Radius, rect.X, rect.Y + Radius)
        path.AddArc(rect.X, rect.Y, 2 * Radius, 2 * Radius, 180, 90)
        path.CloseFigure()

        RoundedCorners = path
    End Function

    Private Sub CustomButton_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseEnter
        Dim g As Graphics = Me.CreateGraphics()
        Dim brushText As SolidBrush
        Dim iXPos, iYPos As Integer
        Dim DrawRect As New Rectangle
        Dim BorderPen As New Pen(m_ButtonBorderColor)
        Dim LayoutSize As New SizeF
        Dim Format As New StringFormat

        If m_AutoSize = True Then
            LayoutSize = g.MeasureString(Me.Text, Me.Font)
            Me.Width = CInt(LayoutSize.Width + 10)
            If m_AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink Then
                Me.Height = CInt(LayoutSize.Height + 10)
            End If
        Else
            Me.Size = m_OriginalSize
            LayoutSize = g.MeasureString(Me.Text, Me.Font, Me.Size)
        End If

        If m_ButtonImage Is Nothing Then
            g.SmoothingMode = SmoothingMode.AntiAlias
            Select Case m_FlatStyle
                Case EnumFlatStyle.Flat
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(Color.FromArgb(30, Color.Black.R, Color.Black.G, Color.Black.B))
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Popup
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            BorderPen = New Pen(Color.FromArgb(127, m_ButtonBorderColor.R, m_ButtonBorderColor.G, m_ButtonBorderColor.B))
                            g.DrawLine(BorderPen, 0, 0, 0, Me.Height)
                            g.DrawLine(BorderPen, 1, 0, Me.Width, 0)
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Standard
                    Dim pathRoundedCorners As New GraphicsPath

                    BorderPen = New Pen(m_ButtonBorderColor)
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_ButtonBackColor = Color.Transparent Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush
                        Dim EndGradientColor As Color

                        EndGradientColor = Color.FromArgb(255, 171, 189, 211)
                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), Color.White, EndGradientColor, LinearGradientMode.Vertical)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    End If
                    g.DrawPath(BorderPen, pathRoundedCorners)
                    pathRoundedCorners = RoundedCorners(New Rectangle(1, 1, Me.Width - 3, Me.Height - 3), 3)
                    BorderPen = New Pen(Color.FromArgb(255, 45, 80, 130))
                    g.DrawPath(BorderPen, pathRoundedCorners)
                    BorderPen = New Pen(Color.FromArgb(255, 239, 207, 134))
                    BorderPen.Width = 2
                    g.DrawRectangle(BorderPen, New Rectangle(3, 3, Me.Width - 6, Me.Height - 6))
                    pathRoundedCorners.Dispose()
                    BorderPen.Dispose()
                Case EnumFlatStyle.Traditional
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
            End Select
            If Me.Enabled Then
                brushText = New SolidBrush(Me.ForeColor)
            Else
                brushText = New SolidBrush(Color.Gray)
            End If
            Select Case m_TextAlign
                Case ContentAlignment.TopLeft
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.TopCenter
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.TopRight
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.MiddleLeft
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.MiddleCenter
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.MiddleRight
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.BottomLeft
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                Case ContentAlignment.BottomCenter
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                Case ContentAlignment.BottomRight
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
            End Select
            g.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos, iYPos, LayoutSize.Width, LayoutSize.Height), Format)
            brushText.Dispose()
        End If
    End Sub

    Private Sub CustomButton_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.MouseLeave
        Dim g As Graphics = Me.CreateGraphics()
        Dim brushText As SolidBrush
        Dim iXPos, iYPos As Integer
        Dim DrawRect As New Rectangle
        Dim BorderPen As New Pen(m_ButtonBorderColor)
        Dim LayoutSize As New SizeF
        Dim Format As New StringFormat

        If m_AutoSize = True Then
            LayoutSize = g.MeasureString(Me.Text, Me.Font)
            Me.Width = CInt(LayoutSize.Width + 10)
            If m_AutoSizeMode = Windows.Forms.AutoSizeMode.GrowAndShrink Then
                Me.Height = CInt(LayoutSize.Height + 10)
            End If
        Else
            Me.Size = m_OriginalSize
            LayoutSize = g.MeasureString(Me.Text, Me.Font, Me.Size)
        End If

        If m_ButtonImage Is Nothing Then
            g.SmoothingMode = SmoothingMode.AntiAlias
            Select Case m_FlatStyle
                Case EnumFlatStyle.Flat
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    BorderPen = New Pen(m_ButtonBorderColor)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Popup
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    BorderPen = New Pen(Color.FromArgb(165, Me.BackColor.R, Me.BackColor.G, Me.BackColor.B))
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
                Case EnumFlatStyle.Standard
                    Dim pathRoundedCorners As New GraphicsPath

                    BorderPen = New Pen(m_ButtonBorderColor)
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_ButtonBackColor = Color.Transparent Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush
                        Dim EndGradientColor As Color

                        EndGradientColor = Color.FromArgb(255, 171, 189, 211)
                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), Color.White, EndGradientColor, LinearGradientMode.Vertical)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), 3)
                        g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                        brushBackground.Dispose()
                    End If
                    g.DrawPath(BorderPen, pathRoundedCorners)
                    pathRoundedCorners = RoundedCorners(New Rectangle(1, 1, Me.Width - 3, Me.Height - 3), 3)
                    BorderPen = New Pen(Color.FromArgb(255, 45, 80, 130))
                    g.DrawPath(BorderPen, pathRoundedCorners)
                    BorderPen = New Pen(Color.White)
                    BorderPen.Width = 2
                    g.DrawRectangle(BorderPen, New Rectangle(3, 3, Me.Width - 6, Me.Height - 6))
                    pathRoundedCorners.Dispose()
                    BorderPen.Dispose()
                Case EnumFlatStyle.Traditional
                    DrawRect = New Rectangle(0, 0, Me.Width, Me.Height)
                    If m_GradientOptions.Background = True Then
                        Dim brushBackground As Drawing2D.LinearGradientBrush

                        brushBackground = New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), m_GradientOptions.StartColor, m_GradientOptions.EndColor, m_GradientOptions.Orientation)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    Else
                        Dim brushBackground As SolidBrush

                        brushBackground = New SolidBrush(m_ButtonBackColor)
                        If m_ButtonRadius > 0 Then
                            Dim pathRoundedCorners As New GraphicsPath

                            pathRoundedCorners = RoundedCorners(New Rectangle(0, 0, Me.Width - 1, Me.Height - 1), m_ButtonRadius)
                            g.FillRegion(brushBackground, New Region(pathRoundedCorners))
                            g.DrawPath(BorderPen, pathRoundedCorners)
                            pathRoundedCorners.Dispose()
                        Else
                            g.FillRectangle(brushBackground, New Rectangle(0, 0, Me.Width, Me.Height))
                            g.DrawRectangle(BorderPen, New Rectangle(0, 0, Me.Width - 1, Me.Height - 1))
                        End If
                        brushBackground.Dispose()
                    End If
                    BorderPen.Dispose()
            End Select
            If Me.Enabled Then
                brushText = New SolidBrush(Me.ForeColor)
            Else
                brushText = New SolidBrush(Color.Gray)
            End If
            Select Case m_TextAlign
                Case ContentAlignment.TopLeft
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.TopCenter
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.TopRight
                    Format.LineAlignment = StringAlignment.Near
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = DrawRect.Y + 2
                Case ContentAlignment.MiddleLeft
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.MiddleCenter
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.MiddleRight
                    Format.LineAlignment = StringAlignment.Center
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = CInt(DrawRect.Y + (DrawRect.Height - LayoutSize.Height) / 2.0)
                Case ContentAlignment.BottomLeft
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Near
                    iXPos = DrawRect.X + 2
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                Case ContentAlignment.BottomCenter
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Center
                    iXPos = CInt(DrawRect.X + (DrawRect.Width - LayoutSize.Width) / 2.0)
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
                Case ContentAlignment.BottomRight
                    Format.LineAlignment = StringAlignment.Far
                    Format.Alignment = StringAlignment.Far
                    iXPos = CInt(DrawRect.X + DrawRect.Width - LayoutSize.Width - 2)
                    iYPos = CInt(DrawRect.Y + DrawRect.Height - LayoutSize.Height - 4.0)
            End Select
            g.DrawString(CStr(Me.Text), Me.Font, brushText, New RectangleF(iXPos, iYPos, LayoutSize.Width, LayoutSize.Height), Format)
            brushText.Dispose()
        End If
    End Sub

    Private Sub CustomButton_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.TextChanged
        Me.Invalidate()
    End Sub
End Class
