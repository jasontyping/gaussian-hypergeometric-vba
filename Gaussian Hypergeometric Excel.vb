' How to use
'
' Open the Visual Basic window in Excel. Create a new module in your spreadsheet and paste this code in
' Then you can use either "AscendingFactorial" or "GaussianHypergeometricFunction" in the formulas
' in your spreadsheet.

Option Explicit

Public Function AscendingFactorial(ByVal x As Double, ByVal n As Double) As Double
    ' The ascending factorial can be defined in terms of the Gamma function
    AscendingFactorial = WorksheetFunction.Gamma(x + n) / WorksheetFunction.Gamma(x)
    
End Function

Public Function GaussianHypergeometricFunction(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal z As Double) As Double
    Dim Summation As Double
    Dim k As Double
    Dim ThisSumIteration As Double
    Dim MinimumSumIncrement As Double
    Dim Precision As Integer
    Dim Denominator As Double
    
    ' Maximum that Excel can handle is about 15 decimals of precision
    Precision = 15
    MinimumSumIncrement = 1 / (10 ^ Precision)
       
    k = 0
    ThisSumIteration = 0
    Summation = 0
    
    ' Iterate and sum until the iteration converges on an increment smaller or
    ' equal to the minimum sum increment above
    '
    ' Or, since some inputs reach sufficiently high iterations that values go beyond,
    ' what Excel is capable of and overflow, simply exit the loop on error
    On Error GoTo SummationComplete
    
    Do
        Denominator = AscendingFactorial(c, k) * WorksheetFunction.Fact(k)
        
        ThisSumIteration = (AscendingFactorial(a, k) * AscendingFactorial(b, k) * (z ^ k)) / Denominator
        
        Summation = Summation + ThisSumIteration
        
        k = k + 1
    Loop While ThisSumIteration > MinimumSumIncrement
    
SummationComplete:
    
    Summation = Math.Round(Summation, Precision)
    
    GaussianHypergeometricFunction = Summation
    
End Function
