Option Explicit
Option Base 1


Function TIEREDRATEFEE(Volumes As Variant, MinimumFee As Double, Thresholds As Variant, Rates As Variant) As Variant
    
    'DIM VARIABLES
    
    Dim TempRates() As Double, TempThresholdDifferences() As Double, TempConditions() As Integer
    Dim IndividualVolume As Range
    Dim Fees() As Double, Fee As Double
    Dim VolumeCount As Integer, RateCount As Integer, ThresholdCount As Integer
    Dim VolumeCounter As Integer, LoopCounter As Integer
    
    If TypeName(Volumes) = "Range" Then VolumeCount = Volumes.Rows.Count Else VolumeCount = UBound(Volumes, 1)
    If TypeName(Rates) = "Range" Then RateCount = Rates.Rows.Count Else RateCount = UBound(Rates, 1)
    If TypeName(Thresholds) = "Range" Then ThresholdCount = Thresholds.Rows.Count Else ThresholdCount = UBound(Thresholds, 1)
    
    ReDim Fees(1 To VolumeCount, 0 To 0)
    ReDim TempRates(RateCount)
    ReDim TempThresholdDifferences(ThresholdCount)
    ReDim TempConditions(ThresholdCount)
    
    Dim i As Variant
    Dim k As Integer
    
    
    'ERROR CHECKS
    
    If TypeName(Volumes) = "Range" Then
        If Volumes.Columns.Count > 1 Then
            MsgBox "Range 'Volumes' should be on a single column.", vbCritical, "Incorrect Range!"
            Exit Function
        End If
    End If
    
    If TypeName(Rates) = "Range" Then
        If Rates.Columns.Count > 1 Then
            MsgBox "Range 'Rates' should be on a single column.", vbCritical, "Incorrect Range!"
            Exit Function
        End If
    End If
    
    If TypeName(Thresholds) = "Range" Then
        If Thresholds.Columns.Count > 1 Then
            MsgBox "Range 'Thresholds' should be on a single column.", vbCritical, "Incorrect Range!"
            Exit Function
        End If
    End If
    
    If RateCount <> ThresholdCount Then
        MsgBox "Ranges 'Thresholds' and 'Rates' should have the same size.", vbCritical, "Range size mismatch!"
        Exit Function
    End If
    
    
    'GET RATES
    
    LoopCounter = 0
    
    For Each i In Rates
    
        LoopCounter = LoopCounter + 1
        TempRates(LoopCounter) = i

    Next i

    
    'GET RATE DIFFERENTIALS

    For k = RateCount To 2 Step -1
    
        TempRates(k) = TempRates(k) - TempRates(k - 1)
        
    Next k
    
    
    'CALCULATE
    
    VolumeCounter = 1

    For Each IndividualVolume In Volumes.Cells
        
        LoopCounter = 0
        
        For Each i In Thresholds
        
            LoopCounter = LoopCounter + 1
            
            TempThresholdDifferences(LoopCounter) = IndividualVolume.Value - i

            If TempThresholdDifferences(LoopCounter) > 0 Then
                TempConditions(LoopCounter) = 1
            Else
                TempConditions(LoopCounter) = 0
            End If
            
        Next i
        
        Fee = 0
        
        For k = 1 To RateCount
            Fee = Fee + TempRates(k) * TempThresholdDifferences(k) * TempConditions(k)
        Next k

        If Fee < MinimumFee Then Fee = MinimumFee
        
        Fees(VolumeCounter, 0) = Fee
        VolumeCounter = VolumeCounter + 1

    Next IndividualVolume
    
    TIEREDRATEFEE = Fees
    
End Function
