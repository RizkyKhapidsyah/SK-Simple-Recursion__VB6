Attribute VB_Name = "RecursiveAlgorithms"
Option Explicit



Public Function Factorial(Number As Integer) As Long

    'Get the factorial of of a number (N!)
    'eg. 5! = 4*3*2*1
    
    If Number < 1 Then
        Factorial = 1
    Else
        Factorial = Number * Factorial(Number - 1)
    End If
    
End Function
Public Function Fibonacci(Number As Integer)

    If Number <= 1 Then
        Fibonacci = Number
    Else
        Fibonacci = Fibonacci(Number - 1) + Fibonacci(Number - 2)
    End If
    
End Function

Public Function GreatestCommonDivisor(A As Integer, B As Integer) As Integer

    If B Mod A = 0 Then
        GreatestCommonDivisor = A
    Else
        GreatestCommonDivisor = GreatestCommonDivisor(B Mod A, A)
    End If
    
End Function
