Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()>
Public Class CalculatorLogicTests

    Private _calculator As CalculatorLogic

    <TestInitialize()>
    Public Sub Setup()
        _calculator = New CalculatorLogic()
    End Sub

    <TestMethod()>
    Public Sub TestNorm()
        Dim inputValue As Double = 1 
        Dim expectedValue As Double = 2.05

        Dim result = _calculator.Norm(inputValue)
        
        Assert.AreEqual(expectedValue, result)
    End Sub

    <TestMethod()>
    Public Sub TestNormDist()
        Dim inputValue As Double = 1 
        Dim expectedValue As Double = 1.02333

        Dim result = _calculator.NormDist(inputValue)
        
        Assert.AreEqual(expectedValue, result)
    End Sub

    <TestMethod()>
    Public Sub TestBinoCoeff()
        Dim n As Integer = 5 
        Dim j As Integer = 3 
        Dim expectedValue As Double = 5.3332

        Dim result = _calculator.BinoCoeff(n, j)
        
        Assert.AreEqual(expectedValue, result)
    End Sub


End Class
