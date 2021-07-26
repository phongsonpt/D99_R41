Imports System.ComponentModel
Imports System.Drawing.Drawing2D

< _
    Serializable(), _
    TypeConverter(GetType(GradientOptionsConverter)) _
> _
Public Class GradientOptions
    Private Const DefaultBackground As Boolean = False
    Private Const DefaultOrientation As LinearGradientMode = LinearGradientMode.Horizontal
    Private Shared ReadOnly DefaultStartColor As Color = Color.Black
    Private Shared ReadOnly DefaultEndColor As Color = Color.Blue

    Private m_Background As Boolean
    Private m_Orientation As Drawing2D.LinearGradientMode = DefaultOrientation
    Private m_StartColor As Color = DefaultStartColor
    Private m_EndColor As Color = DefaultEndColor

    Public Sub New()
        m_Background = DefaultBackground
        m_EndColor = DefaultEndColor
        m_Orientation = DefaultOrientation
        m_StartColor = DefaultStartColor
    End Sub

    Public Sub New(ByVal useGradient As Boolean, ByVal startColor As Color, ByVal endColor As Color, ByVal orientation As LinearGradientMode)
        Me.New()
        m_Background = useGradient
        m_StartColor = startColor
        m_EndColor = endColor
        m_Orientation = orientation
    End Sub

    < _
        Category("Gradient"), _
        Description("Turns ON/OFF Gradient."), _
        DefaultValue(DefaultBackground) _
    > _
    Public Property Background() As Boolean
        Get
            Background = m_Background
        End Get
        Set(ByVal value As Boolean)
            m_Background = value
        End Set
    End Property

    < _
        Category("Gradient"), _
        DefaultValue(DefaultOrientation), _
        DescriptionAttribute("Select Gradient Orientation.") _
    > _
    Public Property Orientation() As Drawing2D.LinearGradientMode
        Get
            Orientation = m_Orientation
        End Get
        Set(ByVal value As Drawing2D.LinearGradientMode)
            m_Orientation = value
        End Set
    End Property

    < _
        CategoryAttribute("Gradient"), _
        DescriptionAttribute("Starting Gradient Color.") _
    > _
    Public Property StartColor() As Color
        Get
            StartColor = m_StartColor
        End Get
        Set(ByVal value As Color)
            m_StartColor = value
        End Set
    End Property

    Protected Function ShouldSerializeStartColor() As Boolean
        Return m_StartColor <> DefaultStartColor
    End Function

    Protected Function ResetStartColor() As Boolean
        m_StartColor = DefaultStartColor
    End Function

    < _
        CategoryAttribute("Gradient"), _
        DescriptionAttribute("Ending Gradient Color.") _
    > _
    Public Property EndColor() As Color
        Get
            EndColor = m_EndColor
        End Get
        Set(ByVal value As Color)
            m_EndColor = value
        End Set
    End Property

    Protected Function ShouldSerializeEndColor() As Boolean
        Return m_EndColor <> DefaultEndColor
    End Function

    Protected Function ResetEndColor() As Boolean
        m_EndColor = DefaultEndColor
    End Function

    Public Overrides Function ToString() As String
        Dim sStartColor As String, sEndColor As String

        If m_StartColor.IsNamedColor Then
            sStartColor = m_StartColor.Name
        Else
            sStartColor = String.Format("{0}, {1}, {2}, {3}", New String() {m_StartColor.A.ToString, m_StartColor.R.ToString, m_StartColor.G.ToString, m_StartColor.B.ToString})
        End If
        If m_EndColor.IsNamedColor Then
            sEndColor = m_EndColor.Name
        Else
            sEndColor = String.Format("{0}, {1}, {2}, {3}", New String() {m_EndColor.A.ToString, m_EndColor.R.ToString, m_EndColor.G.ToString, m_EndColor.B.ToString})
        End If
        Return String.Format("{0}; {1}; {2}; {3}", New String() {m_Background.ToString, sStartColor, sEndColor, m_Orientation.ToString})
    End Function
End Class

Public Class GradientOptionsConverter
    Inherits ExpandableObjectConverter

    'Returns whether this converter can convert an object 
    'of one type to the type of this converter.
    Public Overloads Overrides Function CanConvertFrom _
        (ByVal context As System.ComponentModel.ITypeDescriptorContext, _
        ByVal sourceType As System.Type) As Boolean

        If sourceType Is GetType(String) Then
            Return True
        End If
        Return MyBase.CanConvertFrom(context, sourceType)
    End Function

    'Converts the given value to the type of this converter.
    Public Overloads Overrides Function ConvertFrom _
        (ByVal context As System.ComponentModel.ITypeDescriptorContext, _
        ByVal culture As System.Globalization.CultureInfo, ByVal value As Object) As Object

        If TypeOf value Is String Then
            Try
                Dim strValue As String = value.ToString
                Dim arrValues() As String
                Dim objGradientOptions As GradientOptions = New GradientOptions

                arrValues = strValue.Split(";"c)
                If Not IsNothing(arrValues) Then
                    If Not IsNothing(arrValues(0)) Then
                        objGradientOptions.Background = Convert.ToBoolean(arrValues(0))
                    End If
                    If Not IsNothing(arrValues(1)) Then
                        Dim arrColorValues() As String

                        arrColorValues = arrValues(1).Split(","c)
                        Select Case arrColorValues.Length
                            Case 1
                                Dim ccv As ColorConverter = New System.Drawing.ColorConverter()
                                objGradientOptions.StartColor = CType(ccv.ConvertFromString(arrValues(1)), Color)
                            Case 3  ' RGB
                                objGradientOptions.StartColor = Color.FromArgb(255, L3Int(arrColorValues(0)), L3Int(arrColorValues(1)), L3Int(arrColorValues(2)))
                            Case 4  ' ARGB
                                objGradientOptions.StartColor = Color.FromArgb(L3Int(arrColorValues(0)), L3Int(arrColorValues(1)), L3Int(arrColorValues(2)), L3Int(arrColorValues(3)))
                        End Select
                    End If
                    If Not IsNothing(arrValues(2)) Then
                        Dim arrColorValues() As String

                        arrColorValues = arrValues(2).Split(","c)
                        Select Case arrColorValues.Length
                            Case 1
                                Dim ccv As ColorConverter = New System.Drawing.ColorConverter()
                                objGradientOptions.EndColor = CType(ccv.ConvertFromString(arrValues(2)), Color)
                            Case 3  ' RGB
                                objGradientOptions.EndColor = Color.FromArgb(255, L3Int(arrColorValues(0)), L3Int(arrColorValues(1)), L3Int(arrColorValues(2)))
                            Case 4  ' ARGB
                                objGradientOptions.EndColor = Color.FromArgb(L3Int(arrColorValues(0)), L3Int(arrColorValues(1)), L3Int(arrColorValues(2)), L3Int(arrColorValues(3)))
                        End Select
                    End If
                    If Not IsNothing(Trim(arrValues(3))) Then
                        Dim EnumValues As List(Of String) = New List(Of String)([Enum].GetNames(GetType(LinearGradientMode)))

                        arrValues(3) = Trim(arrValues(3))
                        objGradientOptions.Orientation = CType([Enum].Parse(GetType(LinearGradientMode), arrValues(3), True), LinearGradientMode)
                    End If
                End If
                Return objGradientOptions
            Catch ex As Exception
                Throw New ArgumentException("'" & value.ToString & "' cannot be converted to GradientOptions." & ex.Message)
            End Try
        End If

        Return MyBase.ConvertFrom(context, culture, value)
    End Function
End Class