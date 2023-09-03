Attribute VB_Name = "Module1"
Option Explicit
Private Declare Sub sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)

Sub LinkGraphAdd()

    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlXYScatterLinesNoMarkers
        .SetSourceData Range("B2:C6")
    End With
    
    With ActiveSheet.ChartObjects(1)
        .Top = Range("J2").Top
        .Left = Range("J2").Left
        .Height = Range("J2:J13").Height
        .Width = Range("J2:P2").Width
        .Chart.HasTitle = False
        .Chart.HasLegend = False
        .Chart.Axes(xlValue, 1).HasTitle = False
        .Chart.Axes(xlValue, 1).MaximumScale = 180
        .Chart.Axes(xlValue, 1).MinimumScale = -180
        .Chart.Axes(xlCategory, 1).MaximumScale = 240
        .Chart.Axes(xlCategory, 1).MinimumScale = -240
        .Chart.HasAxis(xlValue, 1) = False
        .Chart.HasAxis(xlCategory, 1) = False
    End With
    
    ActiveSheet.ChartObjects(1).Activate
    
    With ActiveChart
        .Axes(xlCategory).HasMajorGridlines = False
        .Axes(xlValue).HasMajorGridlines = False
    End With
    
End Sub


Sub ChartRefresh()

    Dim a As Long   '��A��`
    For a = 1 To 720 Step 2   '��A��2�����₷�i�덷�����炷����720���܂œ������j
        Range("G2") = a   '��A��G2�ɕ\��
        ActiveSheet.ChartObjects(1).Chart.Refresh   '�O���t���X�V
        Cells(Round(Range("G6")) + 8, 2) = Cells(Round(Range("G6")) + 8, 2) + 1   '�Ή������D�ɂP���L�^����
        sleep 2   '�ҋ@����
        DoEvents   '������OS�ɖ߂�
        DoEvents   '�O���t�ĕ`��
    Next a
End Sub

Sub AngleChartAdd()
    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlLineMarkers
        .SetSourceData Range("B15:B165")
    End With
    
    With ActiveSheet.ChartObjects(2)
        .Top = Range("J16").Top
        .Left = Range("J16").Left
        .Height = Range("J16:J27").Height
        .Width = Range("J16:P16").Width
        .Chart.HasTitle = False
        .Chart.HasLegend = False
        .Chart.Axes(xlValue, 1).HasTitle = False
    End With
    
End Sub

Sub Dreset()
    Range("B15:B188") = 0
End Sub

Sub sumchartadd()
    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlLineMarkers
        .SetSourceData Range("C15:C165")
    End With
    
    With ActiveSheet.ChartObjects(3)
        .Top = Range("J30").Top
        .Left = Range("J30").Left
        .Height = Range("J30:J41").Height
        .Width = Range("J30:P30").Width
        .Chart.HasTitle = False
        .Chart.HasLegend = False
        .Chart.Axes(xlValue, 1).HasTitle = False
    End With
End Sub
