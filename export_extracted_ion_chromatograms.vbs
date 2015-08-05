'the path where the resulting xy files should be exported to
exportPath = "C:\Users\AnalysisUser\Desktop\Export\"

'the list of masses for the extracted mass traces
massesToExport= Array(48,60,77,112)

Dim currentAnalysis

For Each currentAnalysis in Application.Analyses
    currentAnalysis.Chromatograms.Clear

    For Each mass in massesToExport
        'MsgBox mass (show a message box with the current mass)
        Set EIC = CreateObject("DataAnalysis.EICChromatogramDefinition")
        EIC.range = mass
        currentAnalysis.Chromatograms.AddChromatogram EIC
    Next

    Set TIC = CreateObject("DataAnalysis.TICChromatogramDefinition")
    currentAnalysis.Chromatograms.AddChromatogram TIC

    For Each currentChromatogram in currentAnalysis.Chromatograms
        currentChromatogram.Export exportPath+currentAnalysis.name+"_"+currentChromatogram.name+".xy", daXY
    Next
Next
