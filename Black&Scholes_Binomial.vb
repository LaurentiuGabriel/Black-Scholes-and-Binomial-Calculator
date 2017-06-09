Public Class Form1
    Public Input As String                               ''declares the variables for the basic calculator
    Public Operator1 As String
    Public Operator2 As String
    Public Operation As Char
    Public Result As String

    Public Function Norm(ByVal z As Double) As Double                   ''declares the prerequisites of the normal distribution function used by the Black & Scholes formula
        Dim normsdistval As Double
        normsdistval = 1 / (Math.Sqrt(2 * Math.PI)) * Math.Exp(-Math.Pow(z, 2) / 2)
        Return normsdistval
    End Function

    Public Function NormDist(ByVal x As Double) As Double                 ''holds the alghorithm of the normal distribution function
        Const b0 As Double = 0.2316419
        Const b1 As Double = 0.31938153
        Const b2 As Double = -0.356563782
        Const b3 As Double = 1.781477937
        Const b4 As Double = -1.821255978
        Dim b5 As Double
        b5 = 1.330274429
        Dim t As Double
        t = 1 / (1 + b0 * x)
        Dim sigma As Double
        sigma = 1 - Norm(x) * (b1 * t + b2 * Math.Pow(t, 2) + b3 * Math.Pow(t, 3) + b4 * Math.Pow(t, 4) + b5 * Math.Pow(t, 5))
        Return sigma

    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click          ''Code for the 1 button.
        Display.Text = ""
        Input += "1"
        Display.Text += Input
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click         ''Code for the 2 button.
        Display.Text = ""
        Input += "2"
        Display.Text += Input
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Display.Text = ""
        Input += "3"
        Display.Text = Input
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Display.Text = ""
        Input += "4"
        Display.Text += Input
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Display.Text = ""
        Input += "5"
        Display.Text += Input
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Display.Text = ""
        Input += "6"
        Display.Text += Input
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Display.Text = ""
        Input += "7"
        Display.Text += Input
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Display.Text = ""
        Input += "8"
        Display.Text += Input

    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Display.Text = ""
        Input += "9"
        Display.Text += Input
    End Sub
    Private Sub Separator(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Display.Text = ""
        Input += "."
        Display.Text += Input
    End Sub
    Private Sub Button0(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Display.Text = ""
        Input += "0"
        Display.Text += Input
    End Sub

    Private Sub Addition(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click                ''Handles what to do when the plus button is clicked.

        Display.Text = ""
        Operator1 = Input
        Operation = "+"
        Input = String.Empty

    End Sub
    Private Sub Cancel(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click                   ''''Handles what to do when the cancel (C) button is clicked.
        Display.Text = ""
        Input = String.Empty
        Operator1 = String.Empty
        Operator2 = String.Empty

    End Sub

    Private Sub Equals(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click                         ''''Handles what to do when the equal button is clicked.
        Operator2 = Input
        Dim Num1 As Double
        Dim Num2 As Double
        Num1 = Double.Parse(Display.Text)
        Num2 = Double.Parse(Display.Text)
        Double.TryParse(Operator1, Num1)
        Double.TryParse(Operator2, Num2)
        If Operation = "+" Then
            Result = Num1 + Num2
            Display.Text = Result.ToString
        ElseIf Operation = "-" Then
            Result = Num1 - Num2
            Display.Text = Result.ToString
        ElseIf Operation = "*" Then
            Result = Num1 * Num2
            Display.Text = Result.ToString
        ElseIf Operation = "/" Then
            If Num2 <> 0 Then
                Result = Num1 / Num2
                Display.Text = Result.ToString
            Else : Display.Text = "Division to zero?"
            End If
        ElseIf Operation = "^" Then
            If Num1 > 0 Then
                Result = Num1 ^ Num2
                Display.Text = Result.ToString
            End If
            If Num2 = 0 Then
                Result = 1
                Display.Text = Result.ToString

            End If
            If Num1 < 0 And Num2 Mod 2 = 0 Then
                Result = Num1 ^ Num2
                Display.Text = "+" & Result.ToString
            End If
            If Num1 < 0 And Num2 Mod 2 <> 0 Then
                Result = Num1 ^ Num2
                Display.Text = "-" & Result.ToString
            End If
            If Num1 <> 0 And Num2 < 0 Then
                Result = (1 / Num1) ^ Num2
            End If
        End If
    End Sub

    Private Sub Substraction(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click                         ''''Handles what to do when the minus button is clicked.
        Display.Text = ""
        Operator1 = Input
        Operation = "-"
        Input = String.Empty

    End Sub

    Private Sub Division(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click                           ''Handles what to do when the division button is clicked.
        Display.Text = ""
        Operator1 = Input
        Operation = "/"
        Input = String.Empty
    End Sub

    Private Sub Multiplication(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click                    ''Handles what to do when the multiplication button is clicked.
        Display.Text = ""
        Operator1 = Input
        Operation = "*"
        Input = String.Empty
    End Sub

    Private Sub Factorial(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click                           ''Holds the algorithm for performing factorial operations.
        Dim i As Double
        Dim Num1 As Double
        Dim Fact As Double
        Fact = 1
        Num1 = Input
        Num1 = Double.Parse(Display.Text)
        Double.TryParse(Num1, Input)
        If Num1 > 0 Then
            For i = 1 To Num1
                Fact = Fact * i
            Next
            Display.Text = Fact.ToString
        End If
    End Sub

    Private Sub RaisingToPower(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click                           ''Handles what to do when the raise-to-the-power button is clicked.
        Display.Text = ""
        Operator1 = Input
        Operation = "^"
        Input = String.Empty
    End Sub

    Private Sub CalculateBS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CalculateBS.Click
        Dim SpotPrice As Double                 ''Declares the required variables for calculating the price
        Dim StrikePrice As Double
        Dim InterestRate As Double
        Dim DividendY As Double
        Dim Maturity As Double
        Dim Volatility As Double
        Dim D1 As Double
        Dim D2 As Double
        Dim Price As Double

        If RadioButton1.Checked = True Then                     ''Implements the pricing algorithm for European call options

            If System.Double.TryParse(Spot.Text, SpotPrice) Then          ''Throws an error message if the input value is not a number
                SpotPrice = Double.Parse(Spot.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Strike.Text, StrikePrice) Then
                StrikePrice = Double.Parse(Strike.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Rate.Text, InterestRate) Then
                InterestRate = Double.Parse(Rate.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Dividend.Text, DividendY) Then
                DividendY = Double.Parse(Dividend.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(TextBox2.Text, Maturity) Then
                Maturity = Double.Parse(TextBox2.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Vol.Text, Volatility) Then
                Volatility = Double.Parse(Vol.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            InterestRate = InterestRate / 100
            DividendY = DividendY / 100
            Volatility = Volatility / 100

            D1 = ((Math.Log(SpotPrice / StrikePrice) + (InterestRate - DividendY + ((Volatility * Volatility) / 2)) * Maturity) / Volatility * Math.Sqrt(Maturity))
            D2 = D1 - Volatility * Math.Sqrt(Maturity)

            Price = SpotPrice * Math.Exp(DividendY * Maturity) * NormDist(D1) - StrikePrice * Math.Exp(-InterestRate * Maturity) * NormDist(D2)
            Display.Text = Price.ToString

        End If

        If RadioButton2.Checked = True Then                           ''Implements the pricing algorithm for European put options

            If System.Double.TryParse(Spot.Text, SpotPrice) Then
                SpotPrice = Double.Parse(Spot.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Strike.Text, StrikePrice) Then
                StrikePrice = Double.Parse(Strike.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Rate.Text, InterestRate) Then
                InterestRate = Double.Parse(Rate.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Dividend.Text, DividendY) Then
                DividendY = Double.Parse(Dividend.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(TextBox2.Text, Maturity) Then
                Maturity = Double.Parse(TextBox2.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            If System.Double.TryParse(Vol.Text, Volatility) Then
                Volatility = Double.Parse(Vol.Text)
            Else : MessageBox.Show("You must specify a number only")
            End If
            InterestRate = InterestRate / 100
            DividendY = DividendY / 100
            Volatility = Volatility / 100

            D1 = ((Math.Log(SpotPrice / StrikePrice) + (InterestRate - DividendY + ((Volatility * Volatility) / 2)) * Maturity) / Volatility * Math.Sqrt(Maturity))
            D2 = D1 - Volatility * Math.Sqrt(Maturity)

            Price = SpotPrice * Math.Exp(-DividendY * Maturity) * NormDist(D1) - StrikePrice * Math.Exp(-InterestRate * Maturity) * NormDist(D2)
            Dim PutPrice As Double
            PutPrice = (Price + StrikePrice * Math.Exp(-InterestRate * Maturity)) - SpotPrice * Math.Exp(-DividendY * Maturity)
            Display.Text = PutPrice.ToString

        End If
        If RadioButton1.Checked = False And RadioButton2.Checked = False Then     ''Handles the possibility when non of the radio buttons are selected.
            MessageBox.Show("Select at least one type of option")
        End If
    End Sub


    Function binoCoeff(ByVal n, ByVal j)                 ''Calculate the binomial coefficient
        Dim i As Integer
        Dim b As Double

        b = 1
        For i = 0 To j - 1
            b = b * (n - i) / (j - i)
        Next i
        binoCoeff = b

    End Function


    Function Bi_Call_Eur(ByVal s As Double, ByVal x As Double, ByVal t As Double, ByVal r As Double, ByVal sd As Double, ByVal n As Double)    ''Implements the Binomial model for call options
        Dim sdd As Single
        Dim j As Integer
        Dim rr As Single
        Dim q As Single
        Dim u As Single
        Dim d As Single
        Dim bicomp As Single
        Dim sumbi As Single
        Dim nj As Double
        '' Dim firstBicomp As Single

        rr = Math.Exp(r * (t / n)) - 1
        sdd = sd * Math.Sqrt(t / n)
        u = Math.Exp(rr + sdd)
        d = Math.Exp(rr - sdd)
        q = (1 + rr - d) / (u - d)

        For j = 0 To n
            nj = binoCoeff(n, j)
            bicomp = nj * (q ^ j) * ((1 - q) ^ (n - j)) * (s * (u ^ j) * (d ^ (n - j)) - x)
            If bicomp < 0 Then bicomp = 0
            sumbi = sumbi + bicomp
        Next j

        Bi_Call_Eur = sumbi / ((1 + rr) ^ n)

    End Function

    Public Function Bi_Put_Eur(ByVal s As Double, ByVal x As Double, ByVal t As Double, ByVal r As Double, ByVal sd As Double, ByVal n As Double)    ''Implements the Binomial model for put options
        Dim sdd As Single
        Dim j As Integer
        Dim rr As Single
        Dim q As Single
        Dim u As Single
        Dim d As Single
        Dim bicomp As Single
        Dim sumbi As Single
        Dim nj As Double
        '' Dim firstBicomp As Single

        rr = Math.Exp(r * (t / n)) - 1
        sdd = sd * Math.Sqrt(t / n)
        u = Math.Exp(rr + sdd)
        d = Math.Exp(rr - sdd)
        q = (1 + rr - d) / (u - d)

        For j = 0 To n
            nj = binoCoeff(n, j)
            bicomp = nj * (q ^ j) * ((1 - q) ^ (n - j)) * (x - (s * (u ^ j) * (d ^ (n - j))))
            If bicomp < 0 Then bicomp = 0
            sumbi = sumbi + bicomp
        Next j

        Bi_Put_Eur = sumbi / ((1 + rr) ^ n)

    End Function



    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click       ''Code for the 'Calculate' button
        Dim s, x, t, r, sd, n As Double
        Dim BinomialPrice As Double

        If System.Double.TryParse(TextBox3.Text, s) Then
            s = Double.Parse(TextBox3.Text)
        Else : MessageBox.Show("You must specify a number only")
        End If
        If System.Double.TryParse(TextBox4.Text, x) Then
            x = Double.Parse(TextBox4.Text)
        Else : MessageBox.Show("You must specify a number only")
        End If

        If System.Double.TryParse(TextBox5.Text, t) Then
            t = Double.Parse(TextBox5.Text)
        Else : MessageBox.Show("You must specify a number only")
        End If

        If System.Double.TryParse(TextBox6.Text, sd) Then
            sd = Double.Parse(TextBox6.Text)
        Else : MessageBox.Show("You must specify a number only")
        End If

        If System.Double.TryParse(TextBox7.Text, r) Then
            r = Double.Parse(TextBox7.Text)
        Else : MessageBox.Show("You must specify a number only")
        End If


        If System.Double.TryParse(TextBox8.Text, n) Then
            n = Double.Parse(TextBox8.Text)
        Else : MessageBox.Show("You must specify a number only")
        End If
        If RadioButton3.Checked = True Then
            BinomialPrice = Bi_Call_Eur(s, x, t, r, sd, n)
            Display.Text = BinomialPrice.ToString


        End If

        If RadioButton4.Checked = True Then
            BinomialPrice = Bi_Put_Eur(s, x, t, r, sd, n)
            Display.Text = BinomialPrice.ToString
        End If

        If RadioButton3.Checked = False And RadioButton4.Checked = False Then
            MessageBox.Show("Select at least one type of option")
        End If
    End Sub
End Class
