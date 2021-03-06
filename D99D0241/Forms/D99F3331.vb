Imports System
Public Class D99F3331
    Const sOperator As String = "(+-*/^am\mod abs int ln log sqrt sin cos tg cotg"
    Const sPriority As String = "0112239 22 4 4 4 4 4 4 4 4 4"
    Dim oprStack, rpnStack As Collection

    Dim sDivideError As String = "Cannot divide by zero"
    Dim sMathError As String = "Math Error"

    Private _outputResult As String = "0"
    Public ReadOnly Property outputResult() As String
        Get
            Return _outputResult
        End Get
    End Property

    'Cho phép nhập số âm
    Private _allowNegative As Boolean = False
    Public WriteOnly Property AllowNegative() As Boolean 
        Set(ByVal Value As Boolean )
            _allowNegative = Value
        End Set
    End Property

    'Nçi giÀ trÜ tÁi vÜ trÛ con trà ¢ang ¢÷ng
    Private Sub Display(ByVal sValue As String)
        Dim bm As Integer = txtFormular.SelectionStart + Len(sValue)
        If txtFormular.SelectionStart = 0 Then 'vÜ trÛ ¢Çu
            txtFormular.Text = sValue & txtFormular.Text
        ElseIf txtFormular.SelectionStart = Len(txtFormular.Text) Then 'vÜ trÛ cuçi
            txtFormular.Text = txtFormular.Text & sValue
        Else
            txtFormular.Text = Strings.Left(txtFormular.Text, txtFormular.SelectionStart) & sValue & Strings.Right(txtFormular.Text, Len(txtFormular.Text) - txtFormular.SelectionStart)
        End If
        ReturnResult()
        txtFormular.SelectionStart = bm
    End Sub

    Private Sub D99F3331_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        If lblOutputWindow.Text.Contains("-") And Not _allowNegative Then
            D99C0008.MsgL3(rL3("So_khong_hop_le"))
            _outputResult = "0"
            Exit Sub
        End If
        Select Case lblOutputWindow.Text
            Case sMathError, sDivideError, "am"
                _outputResult = "0"
            Case Else
                _outputResult = lblOutputWindow.Text
        End Select
    End Sub


    Private Sub D99F3331_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If geLanguage = EnumLanguage.English Then Me.Text = "Calculator"
        Me.Text = rL3("May_tinhF")
    End Sub

    Private Sub txtFormular_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFormular.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnEqual_Click(sender, Nothing)
        End If

    End Sub

    Private Sub txtFormular_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtFormular.KeyPress
        e.Handled = CheckKeyPress(e.KeyChar, EnumKey.Custom, "0123456789.%+-*/()")
    End Sub

    Private Function Expr2Balan(ByVal sExpression As String) As Boolean
        On Error GoTo hell
        Expr2Balan = False
        oprStack = New Collection
        rpnStack = New Collection

        sExpression = Trim$(sExpression)
        sExpression = "(" & Replace(sExpression, " ", "") & ")"

        Dim sChar As String, i As Integer, OperPre As Integer, OperCurr As Integer

        i = 0

        Do While i < sExpression.Length
            sChar = sExpression.Substring(i, 1)

            Select Case sChar

                Case "("
                    oprStack.Add("(")

                Case ")"
                    If oprStack.Count = 0 Then Exit Function
                    Do While oprStack(oprStack.Count).ToString <> "("
                        rpnStack.Add(oprStack(oprStack.Count))
                        oprStack.Remove(oprStack.Count)
                    Loop

                    oprStack.Remove(oprStack.Count)

                Case ","

                    Do While oprStack(oprStack.Count).ToString <> "("
                        rpnStack.Add(oprStack(oprStack.Count))
                        oprStack.Remove(oprStack.Count)
                    Loop
                Case "%"
                    Dim dNumber As Object
                    dNumber = CalcOpr(Number(rpnStack.Item(rpnStack.Count)), "/", 100)
                    rpnStack.Remove(rpnStack.Count)
                    rpnStack.Add(dNumber)
                Case "+", "-", "*", "/", "^", "\"
                    'Update 20/08/2009
                    'ChÆn nÕu nhËp phÏp toÀn cuçi cîng thØ bà qua ' -1 vØ câ ) cuoi cung
                    If i = sExpression.Length - 1 Then GoTo 3
                    '*************
                    If Not sExpression.Substring(i - 1, 1) Like "[0-9)%]" Then
                        If sChar = "-" Then
                            sChar = "am"
                            If oprStack(oprStack.Count).ToString = "am" Then
                                oprStack.Remove(oprStack.Count)
                                GoTo 3
                            End If
                        Else
                            GoTo 1
                        End If
                    End If

2:                  OperCurr = L3Int(Mid$(sPriority, InStr(sOperator, sChar), 1))
                    OperPre = L3Int(Mid$(sPriority, InStr(sOperator, oprStack(oprStack.Count).ToString), 1))

                    Do While OperCurr <= OperPre
                        rpnStack.Add(oprStack(oprStack.Count))
                        oprStack.Remove(oprStack.Count)
                        OperPre = L3Int(Mid$(sPriority, InStr(sOperator, oprStack(oprStack.Count).ToString), 1))
                    Loop
                    oprStack.Add(sChar)
1:

                Case "0" To "9", "."
                    sChar = ""
                    'LÊy tÊt c¶ sç VD: 1000 th£m vªo stack
                    Do
                        sChar = sChar & sExpression.Substring(i, 1)
                        i = i + 1
                    Loop Until Not (sExpression.Substring(i, 1) Like "[0-9.,]")

                    i = i - 1
                    rpnStack.Add(CDbl(sChar))

                Case "a" To "z"
                    sChar = ""

                    Do
                        sChar = sChar & Mid$(sExpression, i, 1)
                        i = i + 1
                    Loop Until Not (Mid$(sExpression, i, 1) Like "[a-z]")

                    i = i - 1

                    ' If InStr(sOperator, sChar) Then
                    If sOperator.Contains(sChar) Then
                        If sChar = "mod" Then GoTo 2
                        oprStack.Add(CStr(sChar))
                    Else
                        MsgBox("Function " & sChar & " is not defined")
                        Exit Function
                    End If

            End Select
3:
            i = i + 1
        Loop
        Expr2Balan = True
        Exit Function
hell:
        Expr2Balan = False
        lblOutputWindow.Text = sMathError '"Math Error"

    End Function

    Private Function CalcBalan() As String
        On Error GoTo hell
        Dim i As Integer = 1
        Dim schar As String
        Do While rpnStack.Count > 1 'TH nhap vao 1 so va 1 toan tu
            i = i + 1
            schar = rpnStack(i).ToString
            'If InStr(sOperator, sChar) Then
            If sOperator.Contains(schar) Then
                Select Case schar

                    Case "am"
                        rpnStack.Add(-Number(rpnStack(i - 1)), , , i)
                    Case "+", "*", "-", "\", "^", "mod", "log" '"/",
                        rpnStack.Add(CalcOpr(Number(rpnStack(i - 2)), schar, Number(rpnStack(i - 1))), , , i)
                        rpnStack.Remove(i - 2) : i = i - 1
                    Case "/"
                        If Number(rpnStack(i - 1)) = 0 Then
                            CalcBalan = sDivideError '"Cannot divide by zero"
                            Exit Function
                        End If
                        rpnStack.Add(CalcOpr(Number(rpnStack(i - 2)), schar, Number(rpnStack(i - 1))), , , i)
                        rpnStack.Remove(i - 2) : i = i - 1
                    Case Else
                        rpnStack.Add(CalcOpr(Number(rpnStack(i - 1)), schar), , , i)
                End Select

                rpnStack.Remove(i)
                rpnStack.Remove(i - 1)

                i = 0
            End If

        Loop
        CalcBalan = rpnStack(1).ToString
        Exit Function
hell:
        CalcBalan = sMathError '"Math Error"
    End Function

    Private Function CalcOpr(ByVal Number1 As Double, ByVal Opr As String, Optional ByVal Number2 As Double = 0) As Object
        On Error GoTo hell
        CalcOpr = 0
        Select Case Opr

            Case "+" : CalcOpr = Number1 + Number2

            Case "-" : CalcOpr = Number1 - Number2

            Case "*" : CalcOpr = Number1 * Number2

            Case "/" : CalcOpr = Number1 / Number2

            Case "\" : CalcOpr = CLng(Number1) \ CLng(Number2)

            Case "^" : CalcOpr = Number1 ^ Number2

            Case "mod" : CalcOpr = Number1 Mod Number2

                'Case "sin" : CalcOpr = Sin(Number1)

                'Case "cos" : CalcOpr = Cos(Number1)

                'Case "tg" : CalcOpr = Tan(Number1)

                'Case "abs" : CalcOpr = Abs(Number1)

                'Case "int" : CalcOpr = Int(Number1)

                'Case "abs" : CalcOpr = Abs(Number1)

                'Case "sqrt" : CalcOpr = Sqr(Number1)

                'Case "ln" : CalcOpr = Log(Number1)

                'Case "log" : CalcOpr = Log(Number2) / Log(Number1)
        End Select
        Return CalcOpr
hell:
        CalcOpr = sMathError '"Math Error"

    End Function

    Private Function CalcExpression(ByVal s As String) As String
        On Error GoTo hell
        Dim bSuccess As Boolean
        bSuccess = Expr2Balan(s)
        If bSuccess Then CalcExpression = CalcBalan : Exit Function
hell:
        CalcExpression = sMathError '"Math Error"
    End Function

    Private Sub txtFormular_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtFormular.KeyUp
        If e.KeyCode = Keys.Decimal Then Exit Sub
        ReturnResult()
    End Sub

    Private Sub ReturnResult()
        Try
            If txtFormular.Text = "" Then
                lblOutputWindow.Text = ""
                Exit Sub
            End If

            lblOutputWindow.Text = CalcExpression(txtFormular.Text)
            If lblOutputWindow.Text = sMathError Or lblOutputWindow.Text = "am" Or lblOutputWindow.Text = sDivideError Then Exit Sub
            Dim sValue As String
            sValue = FormatDecimal(lblOutputWindow.Text)
            If sValue <> "" Then lblOutputWindow.Text = sValue

            If txtFormular.Text.Length >= 34 Then
                txtFormular.Font = New Font("Arial", 11.0!)
            Else
                txtFormular.Font = New Font("Arial", 14.0!)
            End If
            Exit Sub

        Catch ex As Exception
            lblOutputWindow.Text = sMathError
        End Try
    End Sub


    Private Function FormatDecimal(ByVal sValue As String) As String
        If sValue = "" Or sValue = sMathError Or sValue = "am" Then Return sMathError
        If sValue.Contains("E+") Then Return sMathError
        Dim arr() As String = Microsoft.VisualBasic.Split(sValue, ".")
        Dim j As Integer = 0
        If arr.Length > 1 Then j = arr(1).Length
        FormatDecimal = Format(Number(sValue), "#,##0" & InsertZero(j))
    End Function

    Private Sub btn0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn0.Click, btn00.Click, btn1.Click, btn2.Click, btn3.Click, btn4.Click, btn5.Click, btn6.Click, btn7.Click, btn8.Click, btn9.Click, btnDivide.Click, btnMulti.Click, btnOpen.Click, btnPercent.Click, btnPlus.Click, btnPoint.Click, btnSub.Click, btnClose.Click
        Display(CType(sender, CustomButton).Text)
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtFormular.Text = ""
        lblOutputWindow.Text = ""
    End Sub

    Private Sub btnBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack.Click
        If txtFormular.Text = "" Or txtFormular.SelectionStart = 0 Then txtFormular.Focus() : Exit Sub
        Dim iSelstart As Integer
        iSelstart = txtFormular.SelectionStart
        txtFormular.Text = Strings.Left(txtFormular.Text, txtFormular.SelectionStart - 1) & Strings.Right(txtFormular.Text, Len(txtFormular.Text) - txtFormular.SelectionStart)
        ReturnResult()
        txtFormular.SelectionStart = iSelstart - 1
    End Sub

    Private Sub btnDel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDel.Click
        If txtFormular.Text = "" Or txtFormular.SelectionStart = Len(txtFormular.Text) Then txtFormular.Focus() : Exit Sub
        Dim iSelstart As Integer
        iSelstart = txtFormular.SelectionStart
        txtFormular.Text = Strings.Left(txtFormular.Text, txtFormular.SelectionStart) & Strings.Right(txtFormular.Text, Len(txtFormular.Text) - txtFormular.SelectionStart - 1)
        ReturnResult()
        txtFormular.SelectionStart = iSelstart
    End Sub

    Private Sub btnEqual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEqual.Click
        Me.Close()
    End Sub
End Class