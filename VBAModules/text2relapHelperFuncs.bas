Attribute VB_Name = "text2relapHelperFuncs"
Option Explicit

Function brackedExpression(ParamArray var() As Variant)
' Action: Returns a formatted string inside brackets and separated with semicolon
'         with values
'
    Dim IsXYData As Boolean
    Dim OutputString As String
    Dim i As Integer, j As Integer
    Dim CurrVal As Variant

    On Error GoTo ErrorHandler
    If UBound(var) = 1 Then
        'If UBound(var(0)) <> UBound(var(1)) Then
        IsXYData = True
        If VarType(var(0)) = vbArray + vbVariant Then
            ' Do nothing
        Else
            IsXYData = False
        End If
        'Debug.Print UBound(var(0))
    Else
        IsXYData = False
    End If
    
    OutputString = "["
    
    If IsXYData = False Then
        For i = LBound(var) To UBound(var)
            If TypeName(var(i)) = "Range" Then
                CurrVal = var(i).Text
            Else
                CurrVal = var(i)
            End If
            OutputString = OutputString & CurrVal & IIf(i < UBound(var), ";", "")
        Next i
    Else
        For i = 1 To var(0).Rows.Count
            OutputString = OutputString & var(0)(i).Text & ";" & var(1)(i).Text & IIf(i < var(0).Rows.Count, ";", "")
        Next i
    End If
    
    OutputString = OutputString + "]"
    brackedExpression = OutputString
    Exit Function
ErrorHandler:
        brackedExpression = "<ERROR>"
End Function


Function At2K(A1, A2, At, Adef)
' Funktion som returnerar K-värdet för abrupt area
'
' K2 = (1 - e/(ec*et))^2
'
' e  = A2/A1
' ec = Ac/At = 0.62 + 0.38*et^3
' et = At/A1
'
' A1 = uppströms area
' A2 = nedströms area
' At = geometrisk area
'
'
    Dim et As Double, e As Double, ec As Double
    Dim K2 As Double, Kdef As Double
    et = At / A1
    e = A2 / A1
    ec = 0.62 + 0.38 * et ^ 3
    
    K2 = (1 - e / (ec * et)) ^ 2
    
    Kdef = K2 * (Adef / A2) ^ 2
    At2K = Kdef
    
    
End Function

Function K2At(k, Adef, A1, A2)
' Action: Calculates At (A_throat) from a given K-value
'
    Dim At_guess As Double
    Dim Check As Boolean
    Dim K_guess As Double
    Dim Amax As Double, Amin As Double
    Dim loopCnt As Integer
    
    If A1 < A2 Then Amax = A1 Else Amax = A2
    Amin = 0#
    
    At_guess = Amax / 2
    
    Debug.Print "Första gissning " + CStr(At_guess)
    
    K_guess = At2K(A1, A2, At_guess, Adef)
    
    Check = True
    
    loopCnt = 0
    
    Do    ' Outer loop.
        loopCnt = loopCnt + 1
        Debug.Print "Loop " + CStr(loopCnt)
        Debug.Print "   At_guess = A1 = " + CStr(At_guess / A1)
        
        If loopCnt > 20 Then Check = False
        
        Debug.Print "K_guess = " + CStr(K_guess)
        If Abs((k - K_guess) / k) < 0.001 Then
            Debug.Print "Gissning OK"
            Check = False
        Else
            If K_guess > k Then    ' Om gissat K är större - Förstora At_guess
                Debug.Print "Förstora At_guess"
                Amin = At_guess
                At_guess = (Amax + At_guess) * 0.5
                K_guess = At2K(A1, A2, At_guess, Adef)
            Else     ' Om gissat K är för litet - Förminska At_guess
                Debug.Print "Förminska At_guess"
                Amax = At_guess
                At_guess = (Amin + At_guess) * 0.5
                K_guess = At2K(A1, A2, At_guess, Adef)
            End If
        End If
        
    Loop Until Check = False    ' Exit outer loop immediately
    
    K2At = At_guess

End Function

Function getK1(mflow, dp, rho, Adef)
' Funktion som returnerar K-värde givet flöde, tryckfall, densitet och referensarea
getK1 = dp * 2 * Adef ^ 2 * rho / mflow ^ 2
End Function

Function K2Cv(k, Adef)
K2Cv = Sqr(2 * Adef ^ 2 / (991.091 * k)) / 0.0000007598055
End Function

Function Cv2K(Cv, Adef)
Cv2K = 2 * Adef ^ 2 / (991.091 * (0.0000007598055 * Cv) ^ 2)
End Function


Function getActualDx(L, Dx)
' Action: Return true dx using the same alogoritm as text2relap
    Dim actualDx As Double, largeDx As Double, smallDx As Double
    Dim diffLargeDx As Double, diffSmallDx As Double
    
    If L <= Dx Then
        actualDx = L
    Else
        largeDx = L / Int(((L + 0.00001) / Dx))
        smallDx = L / Int(((L + 0.00001) / Dx) + 1)
    
        diffLargeDx = Abs(Dx - largeDx)
        diffSmallDx = Abs(Dx - smallDx)
    
        If diffSmallDx <= 0.5 * diffLargeDx Then
            actualDx = smallDx
        Else
            actualDx = largeDx
        End If
    End If
    
    getActualDx = actualDx
End Function




Function void2xGas(voidg, Pabs, t)
' Funktion som returnerar massandelen ånga+gas som ska anges
'  för att erhålla en viss volymsfraktion (void) gas+ånga när ebt=104 används
'
' INPUT
'       voidg       Volymsandel gas+vattenånga (0-1)
'       Pabs        Totaltryck
'       T           Temperatur på blandning. psat_T(T) >= Pabs
'
    
    Dim pps As Double, ppa As Double
    Dim rho_l As Double, rho_s As Double, rho_a As Double
    Dim mass_l As Double, mass_sa As Double, x As Double
    
    pps = psat_T(t)     ' Partialtryck ånga = mättnadstryck för aktuell temp
    ppa = Pabs - pps    ' Partialtryck gas = Totaltryck - pps
    
    If ppa < 0 Then
        void2xGas = "Fel, negativt partialtryck för gas. T < T_sat(Pabs) = " & Tsat_p(Pabs) & " degC"
    End If
    
    rho_l = rhoL_T(t)
    rho_s = pps * 100000# / (461.5 * (273.15 + t))
    rho_a = ppa * 100000# / (296.8 * (273.15 + t))
    
    mass_l = rho_l * (1 - voidg)
    mass_sa = (rho_s + rho_a) * voidg
    
    x = mass_sa / (mass_sa + mass_l)
    
    void2xGas = x

End Function



Function void2x(voidg, Pabs)
' Funktion som returnerar massandelen ånga som ska anges
'  för att erhålla en viss volymsfraktion (void) ånga när ebt=102 används
'
' INPUT
'       voidg       Volymsandel ånga (0-1)
'       Pabs        Totaltryck
'
'
    Dim rho_s As Double, rho_l As Double, x As Double

    rho_s = rhoV_p(Pabs)      ' Densitet på ångan
    rho_l = rhoL_p(Pabs)      ' Densitet på vattnet
    
    x = (voidg / rho_l) / (1 / rho_s + voidg * (1 / rho_l - 1 / rho_s))
    
    void2x = x
End Function




Sub LastRow()

With ActiveSheet

MsgBox .Range(.Cells(1, 1), .Cells(1, 22)).End(xlDown).Row
End With
End Sub
