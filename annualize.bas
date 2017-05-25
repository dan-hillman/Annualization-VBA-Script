Function DaysInYear(dateToAnalyze As Long) As Long
 
    DaysInYear = 365 - (Year(dateToAnalyze) Mod 4 = 0) + (Year(dateToAnalyze) Mod 100 = 0) - (Year(dateToAnalyze) Mod 400 = 0)
   
End Function
Function Annualize(valueToAnnualize As Double, dateToAnnualizeBy As Long) As Double
 
    beginningOfYear = DateSerial(Year(dateToAnnualizeBy) - 1, 12, 31)
    endOfYear = DateSerial(Year(dateToAnnualizeBy), 12, 31)
    YTDDaysElapsed = DateDiff("d", beginningOfYear, dateToAnnualizeBy)
    annualizationCoefficient = (1 / YTDDaysElapsed) * DaysInYear(dateToAnnualizeBy)
    Annualize = valueToAnnualize * annualizationCoefficient
 
End Function
