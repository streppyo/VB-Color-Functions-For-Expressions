' Module ColorExp
    ' TODO: Documentation for TextColor functions
    ' TODO: Convert Color Names (KnownColors) to hex
    Public Function ColorMix(ByVal Value As Double, ByVal MaxValue As Double, ByVal MinValue As Double, ByVal ColStrMax As String, Optional ByVal ColStrMin As String = "#FFFFFF") As String
        Dim ColR1, ColG1, ColB1 As Integer
        Dim ColR2, ColG2, ColB2 As Integer
        Value = Math.Max(Math.Min(MaxValue, Value), MinValue)
        ColR1 = Convert.ToInt32(Left(Right(ColStrMax, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStrMax, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStrMax, 2), 16)
        ColR2 = Convert.ToInt32(Left(Right(ColStrMin, 6), 2), 16)
        ColG2 = Convert.ToInt32(Left(Right(ColStrMin, 4), 2), 16)
        ColB2 = Convert.ToInt32(Right(ColStrMin, 2), 16)
        Dim MaxMinDiff As Double
        MaxMinDiff = Math.Abs(MaxValue - MinValue)
        Dim NewColR, NewColG, NewColB As Integer
        Dim strColor As String
        NewColR = ColR1 + CInt(Math.Round((MaxValue - Value) * ((ColR2 - ColR1) / MaxMinDiff)))
        NewColG = ColG1 + CInt(Math.Round((MaxValue - Value) * ((ColG2 - ColG1) / MaxMinDiff)))
        NewColB = ColB1 + CInt(Math.Round((MaxValue - Value) * ((ColB2 - ColB1) / MaxMinDiff)))
        strColor = "#" & NewColR.ToString("X2") & NewColG.ToString("X2") & NewColB.ToString("X2")
        Return strColor
    End Function


    Public Function ColorMix2(ByVal Value As Double, ByVal MaxValue As Double, ByVal MinValue As Double, ByVal ColStrMax As String, ByVal ColStrMin As String, Optional ByVal ColStrMid As String = "#FFFFFF") As String
        Dim ColR1, ColG1, ColB1 As Integer
        Dim ColR2, ColG2, ColB2 As Integer
        Dim ColR3, ColG3, ColB3 As Integer
        Value = Math.Max(Math.Min(MaxValue, Value), MinValue)
        ColR1 = Convert.ToInt32(Left(Right(ColStrMax, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStrMax, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStrMax, 2), 16)
        ColR2 = Convert.ToInt32(Left(Right(ColStrMin, 6), 2), 16)
        ColG2 = Convert.ToInt32(Left(Right(ColStrMin, 4), 2), 16)
        ColB2 = Convert.ToInt32(Right(ColStrMin, 2), 16)
        ColR3 = Convert.ToInt32(Left(Right(ColStrMid, 6), 2), 16)
        ColG3 = Convert.ToInt32(Left(Right(ColStrMid, 4), 2), 16)
        ColB3 = Convert.ToInt32(Right(ColStrMid, 2), 16)
        Dim MaxMinDiff As Double
        MaxMinDiff = Math.Abs(MaxValue - MinValue) / 2
        Dim NewColR, NewColG, NewColB As Integer
        Dim strColor As String
        If Value > (MaxValue + MinValue) / 2 Then
            NewColR = ColR1 + CInt(Math.Round((MaxValue - Value) * ((ColR3 - ColR1) / MaxMinDiff)))
            NewColG = ColG1 + CInt(Math.Round((MaxValue - Value) * ((ColG3 - ColG1) / MaxMinDiff)))
            NewColB = ColB1 + CInt(Math.Round((MaxValue - Value) * ((ColB3 - ColB1) / MaxMinDiff)))
        Else
            NewColR = ColR2 - CInt(Math.Round((MinValue - Value) * ((ColR3 - ColR2) / MaxMinDiff)))
            NewColG = ColG2 - CInt(Math.Round((MinValue - Value) * ((ColG3 - ColG2) / MaxMinDiff)))
            NewColB = ColB2 - CInt(Math.Round((MinValue - Value) * ((ColB3 - ColB2) / MaxMinDiff)))
        End If
        strColor = "#" & NewColR.ToString("X2") & NewColG.ToString("X2") & NewColB.ToString("X2")
        Return strColor
    End Function

    Public Function RGB(ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double) As String
        Dim RVal, GVal, BVal As Integer
        RVal = Math.Round(Math.Max(Math.Min(Red, 255), 0))
        GVal = Math.Round(Math.Max(Math.Min(Green, 255), 0))
        BVal = Math.Round(Math.Max(Math.Min(Blue, 255), 0))
        RGB = "#" & RVal.ToString("X2") & GVal.ToString("X2") & BVal.ToString("X2")
        Return RGB
    End Function

    Public Function HSV(ByVal Hue As Double, ByVal Saturation As Double, ByVal Value As Double) As String
        Dim HVal As Double, SVal As Double, VVal As Double
        Dim f As Double, p As Double, q As Double, t As Double
        Dim QVal As Double
        HVal = Math.Max(Math.Min(Hue - Int(Hue), 1), 0) * 360
        SVal = Math.Max(Math.Min(Saturation, 1), 0)
        VVal = Math.Max(Math.Min(Value, 1), 0)
        QVal = Math.Floor(HVal / 60)
        f = HVal / 60 - QVal
        p = VVal * (1 - SVal)
        q = VVal * (1 - f * SVal)
        t = VVal * (1 - (1 - f) * SVal)
        Dim RVal As Integer, GVal As Integer, BVal As Integer
        Select Case QVal
            Case 0
                RVal = 255 * VVal
                GVal = 255 * t
                BVal = 255 * p
            Case 1
                RVal = 255 * q
                GVal = 255 * VVal
                BVal = 255 * p
            Case 2
                RVal = 255 * p
                GVal = 255 * VVal
                BVal = 255 * t
            Case 3
                RVal = 255 * p
                GVal = 255 * q
                BVal = 255 * VVal
            Case 4
                RVal = 255 * t
                GVal = 255 * p
                BVal = 255 * VVal
            Case 5
                RVal = 255 * VVal
                GVal = 255 * p
                BVal = 255 * q
        End Select
        HSV = "#" & RVal.ToString("X2") & GVal.ToString("X2") & BVal.ToString("X2")
        Return HSV
    End Function

    Public Function HSV360(ByVal Hue As Double, ByVal Saturation As Double, ByVal Value As Double) As String
        Dim HVal As Double, SVal As Double, VVal As Double
        Dim f As Double, p As Double, q As Double, t As Double
        Dim QVal As Double
        HVal = Math.Max(Math.Min(Hue - Int(Hue / 360) * 360, 360), 0)
        SVal = Math.Max(Math.Min(Saturation, 100), 0) / 100
        VVal = Math.Max(Math.Min(Value, 100), 0) / 100
        QVal = Math.Floor(HVal / 60)
        f = HVal / 60 - QVal
        p = VVal * (1 - SVal)
        q = VVal * (1 - f * SVal)
        t = VVal * (1 - (1 - f) * SVal)
        Dim RVal As Integer, GVal As Integer, BVal As Integer
        Select Case QVal
            Case 0
                RVal = 255 * VVal
                GVal = 255 * t
                BVal = 255 * p
            Case 1
                RVal = 255 * q
                GVal = 255 * VVal
                BVal = 255 * p
            Case 2
                RVal = 255 * p
                GVal = 255 * VVal
                BVal = 255 * t
            Case 3
                RVal = 255 * p
                GVal = 255 * q
                BVal = 255 * VVal
            Case 4
                RVal = 255 * t
                GVal = 255 * p
                BVal = 255 * VVal
            Case 5
                RVal = 255 * VVal
                GVal = 255 * p
                BVal = 255 * q
        End Select
        HSV360 = "#" & RVal.ToString("X2") & GVal.ToString("X2") & BVal.ToString("X2")
        Return HSV360
    End Function

    Public Function HSL(ByVal Hue As Double, ByVal Saturation As Double, ByVal Luminosity As Double) As String
        If Math.Max(Math.Min(Saturation, 1), 0) = 0 Then
            Dim LVal As Integer
            LVal = 255 * Math.Max(Math.Min(Luminosity, 1), 0)
            HSL = "#" & LVal.ToString("X2") & LVal.ToString("X2") & LVal.ToString("X2")
            Return HSL
        Else
            Dim l As Double, c As Double, x As Double, s As Double, h As Double, m As Double
            h = Math.Max(Math.Min(Hue, 1), 0) * 360
            s = Math.Max(Math.Min(Saturation, 1), 0)
            l = Math.Max(Math.Min(Luminosity, 1), 0)
            c = (1 - Math.Abs(2 * l - 1)) * s
            x = c * (1 - Math.Abs(((h / 60) Mod 2) - 1))
            m = l - c / 2
            Dim tCArray(3) As Double
            Dim quad As Integer
            quad = Math.Floor(h / 60) + 1
            Select Case quad
                Case 1
                    tCArray(1) = c + m
                    tCArray(2) = x + m
                    tCArray(3) = m
                Case 2
                    tCArray(1) = x + m
                    tCArray(2) = c + m
                    tCArray(3) = m
                Case 3
                    tCArray(1) = m
                    tCArray(2) = c + m
                    tCArray(3) = x + m
                Case 4
                    tCArray(1) = m
                    tCArray(2) = x + m
                    tCArray(3) = c + m
                Case 5
                    tCArray(1) = x + m
                    tCArray(2) = m
                    tCArray(3) = c + m
                Case 6
                    tCArray(1) = c + m
                    tCArray(2) = m
                    tCArray(3) = x + m
            End Select
            Dim RVal As Integer, GVal As Integer, BVal As Integer
            RVal = Math.Round(Math.Max(Math.Min(tCArray(1) * 255, 255), 0))
            GVal = Math.Round(Math.Max(Math.Min(tCArray(2) * 255, 255), 0))
            BVal = Math.Round(Math.Max(Math.Min(tCArray(3) * 255, 255), 0))
            HSL = "#" & RVal.ToString("X2") & GVal.ToString("X2") & BVal.ToString("X2")
            Return HSL
        End If
    End Function

    Public Function HSL360(ByVal Hue As Double, ByVal Saturation As Double, ByVal Luminosity As Double) As String
        If Math.Max(Math.Min(Saturation, 100), 0) = 0 Then
            Dim LVal As Integer
            LVal = 255 * Math.Max(Math.Min(Luminosity / 100, 1), 0)
            HSL360 = "#" & LVal.ToString("X2") & LVal.ToString("X2") & LVal.ToString("X2")
            Return HSL360
        Else
            Dim l As Double, c As Double, x As Double, s As Double, h As Double, m As Double
            h = Math.Max(Math.Min(Hue Mod 360, 360), 0)
            s = Math.Max(Math.Min(Saturation / 100, 1), 0)
            l = Math.Max(Math.Min(Luminosity / 100, 1), 0)
            c = (1 - Math.Abs(2 * l - 1)) * s
            x = c * (1 - Math.Abs(((h / 60) Mod 2) - 1))
            m = l - c / 2
            Dim tCArray(3) As Double
            Dim quad As Integer
            quad = Math.Floor(h / 60) + 1
            Select Case quad
                Case 1
                    tCArray(1) = c + m
                    tCArray(2) = x + m
                    tCArray(3) = m
                Case 2
                    tCArray(1) = x + m
                    tCArray(2) = c + m
                    tCArray(3) = m
                Case 3
                    tCArray(1) = m
                    tCArray(2) = c + m
                    tCArray(3) = x + m
                Case 4
                    tCArray(1) = m
                    tCArray(2) = x + m
                    tCArray(3) = c + m
                Case 5
                    tCArray(1) = x + m
                    tCArray(2) = m
                    tCArray(3) = c + m
                Case 6
                    tCArray(1) = c + m
                    tCArray(2) = m
                    tCArray(3) = x + m
            End Select

            Dim RVal As Integer, GVal As Integer, BVal As Integer
            RVal = Math.Round(Math.Max(Math.Min(tCArray(1) * 255, 255), 0))
            GVal = Math.Round(Math.Max(Math.Min(tCArray(2) * 255, 255), 0))
            BVal = Math.Round(Math.Max(Math.Min(tCArray(3) * 255, 255), 0))
            HSL360 = "#" & RVal.ToString("X2") & GVal.ToString("X2") & BVal.ToString("X2")
            Return HSL360
        End If
    End Function

    Public Function RandomColor() As String
        Dim r As Double, g As Double, b As Double
        r = Rnd()
        g = Rnd()
        b = Rnd()
        Dim RVal As Integer, GVal As Integer, BVal As Integer
        RVal = Math.Round(Math.Max(Math.Min(r * 255, 255), 0))
        GVal = Math.Round(Math.Max(Math.Min(g * 255, 255), 0))
        BVal = Math.Round(Math.Max(Math.Min(b * 255, 255), 0))
        RandomColor = "#" & RVal.ToString("X2") & GVal.ToString("X2") & BVal.ToString("X2")
        Return RandomColor
    End Function

    Public Function RandomColorHSL() As String
        Dim r As Double, g As Double, b As Double
        r = Rnd()
        g = Rnd()
        b = Rnd()
        RandomColorHSL = HSL(r, g, b)
        Return RandomColorHSL
    End Function

    Public Function RandomColorHSV() As String
        Dim r As Double, g As Double, b As Double
        r = Rnd()
        g = Rnd()
        b = Rnd()
        RandomColorHSV = HSV(r, g, b)
        Return RandomColorHSV
    End Function

    Public Function TextColorBG601(ByVal ColStr As String) As String
        Dim ColR1, ColG1, ColB1 As Integer
        ColR1 = Convert.ToInt32(Left(Right(ColStr, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStr, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStr, 2), 16)
        If (ColR1 * 0.299 + ColG1 * 0.587 + ColB1 * 0.114) < 186 Then
            TextColorBG601 = "#FFFFFF"
            Return TextColorBG601
        Else
            TextColorBG601 = "#000000"
            Return TextColorBG601
        End If
    End Function

    Public Function TextColorBG709(ByVal ColStr As String) As String
        Dim ColR1, ColG1, ColB1 As Integer
        ColR1 = Convert.ToInt32(Left(Right(ColStr, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStr, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStr, 2), 16)
        If (ColR1 * 0.2126 + ColG1 * 0.7152 + ColB1 * 0.0722) < 186 Then
            TextColorBG709 = "#FFFFFF"
            Return TextColorBG709
        Else
            TextColorBG709 = "#000000"
            Return TextColorBG709
        End If
    End Function

    Public Function TextColorBG240(ByVal ColStr As String) As String
        Dim ColR1, ColG1, ColB1 As Integer
        ColR1 = Convert.ToInt32(Left(Right(ColStr, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStr, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStr, 2), 16)
        If (ColR1 * 0.212 + ColG1 * 0.701 + ColB1 * 0.087) < 186 Then
            TextColorBG240 = "#FFFFFF"
            Return TextColorBG240
        Else
            TextColorBG240 = "#000000"
            Return TextColorBG240
        End If
    End Function

    Public Function TextColorBG601W(ByVal ColStr As String) As String
        Dim ColR1, ColG1, ColB1 As Integer
        ColR1 = Convert.ToInt32(Left(Right(ColStr, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStr, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStr, 2), 16)
        If (ColR1 * 0.299 + ColG1 * 0.587 + ColB1 * 0.114) < 150 Then
            TextColorBG601W = "#FFFFFF"
            Return TextColorBG601W
        Else
            TextColorBG601W = "#000000"
            Return TextColorBG601W
        End If
    End Function

    Public Function TextColorBG709W(ByVal ColStr As String) As String
        Dim ColR1, ColG1, ColB1 As Integer
        ColR1 = Convert.ToInt32(Left(Right(ColStr, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStr, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStr, 2), 16)
        If (ColR1 * 0.2126 + ColG1 * 0.7152 + ColB1 * 0.0722) < 150 Then
            TextColorBG709W = "#FFFFFF"
            Return TextColorBG709W
        Else
            TextColorBG709W = "#000000"
            Return TextColorBG709W
        End If
    End Function

    Public Function TextColorBG240W(ByVal ColStr As String) As String
        Dim ColR1, ColG1, ColB1 As Integer
        ColR1 = Convert.ToInt32(Left(Right(ColStr, 6), 2), 16)
        ColG1 = Convert.ToInt32(Left(Right(ColStr, 4), 2), 16)
        ColB1 = Convert.ToInt32(Right(ColStr, 2), 16)
        If (ColR1 * 0.212 + ColG1 * 0.701 + ColB1 * 0.087) < 186 Then
            TextColorBG240W = "#FFFFFF"
            Return TextColorBG240W
        Else
            TextColorBG240W = "#000000"
            Return TextColorBG240W
        End If
    End Function

'End Module
