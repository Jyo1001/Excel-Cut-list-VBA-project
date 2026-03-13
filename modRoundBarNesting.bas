Attribute VB_Name = "modRoundBarNesting"
Option Explicit

Private Const WS_INPUT As String = "INPUT_PARTS"
Private Const WS_MAT As String = "MATERIAL_SUMMARY"
Private Const WS_PURCHASE As String = "PURCHASE_PLAN"
Private Const WS_BASIC As String = "NEST_BASIC"
Private Const WS_FINAL As String = "NEST_FINAL"
Private Const WS_CUTLIST As String = "CUTLIST"

Private Const TBL_PARTS As String = "tblParts"
Private Const TBL_MAT As String = "tblMat"
Private Const TBL_PURCHASE As String = "tblPurchase"
Private Const TBL_FINAL As String = "tblNestFinal"
Private Const TBL_CUTLIST As String = "tblCutlist"

Private Const BTN_PRINT_NAME As String = "btnPrintCutlist"
Private Const TOL As Double = 0.000001#

'=========================================================
' LOGGING
'=========================================================
Public Sub LogInfo(ByVal PROC As String, ByVal msg As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), "| INFO  |", PROC, "|", msg
End Sub

Public Sub LogError(ByVal PROC As String, ByVal errNum As Long, ByVal errDesc As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), "| ERROR |", PROC, "|", errNum, "|", errDesc
End Sub

'=========================================================
' PUBLIC ENTRY POINTS
'=========================================================
Public Sub EnsureInputValidations()
    Const PROC As String = "EnsureInputValidations"
    On Error GoTo EH

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim maxRows As Long
    Dim headerRow As Long
    Dim firstDataRow As Long
    Dim idxDiameter As Long
    Dim idxQty As Long
    Dim idxMaterial As Long

    Set ws = GetWorksheet(WS_INPUT)
    Set lo = GetTable(WS_INPUT, TBL_PARTS)

    If ws Is Nothing Then
        LogInfo PROC, "Worksheet not found: " & WS_INPUT
        Exit Sub
    End If
    If lo Is Nothing Then
        LogInfo PROC, "Table not found: " & TBL_PARTS
        Exit Sub
    End If

    idxDiameter = GetListColumnIndex(lo, "Diameter_in", PROC)
    idxQty = GetListColumnIndex(lo, "Qty", PROC)
    idxMaterial = GetListColumnIndex(lo, "Material", PROC)

    maxRows = CLng(GetNamedRangeNumber("MaxPartRows", 500))
    If maxRows < 1 Then maxRows = 500

    headerRow = lo.HeaderRowRange.Row
    firstDataRow = headerRow + 1

    LogInfo PROC, "Applying validations to input ranges. Max rows=" & CStr(maxRows)

    ApplyDecimalValidation ws.Range(ws.Cells(firstDataRow, lo.Range.Column + idxDiameter - 1), _
                                    ws.Cells(firstDataRow + maxRows - 1, lo.Range.Column + idxDiameter - 1)), _
                           0.000001#, 1000000#

    ApplyWholeValidation ws.Range(ws.Cells(firstDataRow, lo.Range.Column + idxQty - 1), _
                                  ws.Cells(firstDataRow + maxRows - 1, lo.Range.Column + idxQty - 1)), _
                         1, 1000000

    ApplyListValidationAllowFreeText ws.Range(ws.Cells(firstDataRow, lo.Range.Column + idxMaterial - 1), _
                                              ws.Cells(firstDataRow + maxRows - 1, lo.Range.Column + idxMaterial - 1)), _
                                     "=Materials"

    LogInfo PROC, "Validation complete"
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Public Sub EnsurePrintButton()
    Const PROC As String = "EnsurePrintButton"
    On Error GoTo EH

    Dim ws As Worksheet
    Dim shp As Shape
    Dim leftPos As Double
    Dim topPos As Double
    Dim btnW As Double
    Dim btnH As Double

    Set ws = GetWorksheet(WS_CUTLIST)
    If ws Is Nothing Then
        LogInfo PROC, "Worksheet not found: " & WS_CUTLIST
        Exit Sub
    End If

    btnW = 120
    btnH = 28
    leftPos = ws.Range("J1").Left
    topPos = ws.Range("J1").Top + 2

    On Error Resume Next
    Set shp = ws.Shapes(BTN_PRINT_NAME)
    On Error GoTo EH

    If shp Is Nothing Then
        LogInfo PROC, "Creating print button"
        Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnW, btnH)
        shp.Name = BTN_PRINT_NAME
    Else
        LogInfo PROC, "Updating existing print button"
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = btnW
        shp.Height = btnH
    End If

    If Not shp Is Nothing Then
        shp.OnAction = "PrintCutlist"
        shp.TextFrame.Characters.Text = "Print Cutlist"
        shp.Fill.ForeColor.RGB = RGB(225, 235, 245)
        shp.Line.ForeColor.RGB = RGB(80, 80, 80)
        shp.TextFrame.HorizontalAlignment = xlHAlignCenter
        shp.TextFrame.VerticalAlignment = xlVAlignCenter
        shp.Placement = xlFreeFloating
    End If

    LogInfo PROC, "Print button ensured"
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Public Sub ResetDemoData()
    Const PROC As String = "ResetDemoData"
    On Error GoTo EH

    LogInfo PROC, "Reset started"

    ResetInputPartsTable
    ResetTableToSingleBlankRow GetTable(WS_MAT, TBL_MAT)
    ResetTableToSingleBlankRow GetTable(WS_PURCHASE, TBL_PURCHASE)
    ResetTableToSingleBlankRow GetTable(WS_FINAL, TBL_FINAL)
    ResetTableToSingleBlankRow GetTable(WS_CUTLIST, TBL_CUTLIST)
    ClearPlainOutputSheet GetWorksheet(WS_BASIC)

    SeedDefaultNamedValues
    SeedMaterialsListDefault
    EnsureInputValidations
    EnsurePrintButton

    SetNamedRangeValue "LastFinalRun", vbNullString
    SetNamedRangeValue "LastCutlistGen", vbNullString

    LogInfo PROC, "Reset complete"
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Public Sub GenerateDemoData(Optional ByVal Seed As Long = 1, Optional ByVal PartRowCount As Long = 100, Optional ByVal ComboCount As Long = 20)
    Const PROC As String = "GenerateDemoData"
    On Error GoTo EH

    Dim stageMsg As String
    Dim lo As ListObject
    Dim dataArr() As Variant
    Dim comboMaterials() As String
    Dim comboDiams() As Double
    Dim i As Long
    Dim comboIdx As Long
    Dim baseMaterials As Variant
    Dim baseDiams As Variant
    Dim actualComboCount As Long
    Dim r As Double
    Dim qtyVal As Long
    Dim finishVal As Double
    Dim extraVal As Double
    Dim kerfVal As Variant
    Dim faceVal As Variant
    Dim matLo As ListObject

    Dim idxJob As Long
    Dim idxPartNo As Long
    Dim idxMaterial As Long
    Dim idxDiameter As Long
    Dim idxFinish As Long
    Dim idxExtra As Long
    Dim idxQty As Long
    Dim idxKerf As Long
    Dim idxFace As Long

    LogInfo PROC, "Generate demo data start. Seed=" & CStr(Seed) & ", PartRowCount=" & CStr(PartRowCount) & ", ComboCount=" & CStr(ComboCount)

    If PartRowCount < 1 Then PartRowCount = 100
    If ComboCount < 1 Then ComboCount = 20

    stageMsg = "Reset workbook state"
    ResetDemoData

    baseMaterials = Array("1018", "12L14", "4140", "4340", "8620", "17-4PH", "304", "316", "O1", "D2", _
                          "H13", "A2", "Brass360", "Al6061", "Al7075", "Ti-6Al-4V", "Monel400", "Inconel625", "UHMW", "Acetal")
    baseDiams = Array(0.5, 0.625, 0.75, 0.875, 1#, 1.125, 1.25, 1.375, 1.5, 1.625, _
                      1.75, 1.875, 2#, 2.125, 2.25, 2.5, 2.75, 3#, 3.5, 4#)

    actualComboCount = ComboCount
    If actualComboCount > 20 Then actualComboCount = 20

    stageMsg = "Build combo arrays"
    ReDim comboMaterials(1 To actualComboCount)
    ReDim comboDiams(1 To actualComboCount)

    For i = 1 To actualComboCount
        comboMaterials(i) = CStr(baseMaterials(i - 1))
        comboDiams(i) = CDbl(baseDiams(i - 1))
    Next i

    stageMsg = "Seed materials list"
    SeedMaterialsListFromArrays comboMaterials

    stageMsg = "Get input parts table"
    Set lo = GetTable(WS_INPUT, TBL_PARTS)
    If lo Is Nothing Then Err.Raise 9001, PROC, "Input parts table not found."

    idxJob = GetListColumnIndex(lo, "Job", PROC)
    idxPartNo = GetListColumnIndex(lo, "PartNo", PROC)
    idxMaterial = GetListColumnIndex(lo, "Material", PROC)
    idxDiameter = GetListColumnIndex(lo, "Diameter_in", PROC)
    idxFinish = GetListColumnIndex(lo, "FinishLen_in", PROC)
    idxExtra = GetListColumnIndex(lo, "ExtraStock_in", PROC)
    idxQty = GetListColumnIndex(lo, "Qty", PROC)
    idxKerf = GetListColumnIndex(lo, "Kerf_in", PROC)
    idxFace = GetListColumnIndex(lo, "FaceAllow_in", PROC)

    stageMsg = "Resize input table rows"
    EnsureTableRows lo, PartRowCount

    stageMsg = "Prepare demo data array"
    Randomize Seed
    ReDim dataArr(1 To PartRowCount, 1 To 9)

    For i = 1 To PartRowCount
        comboIdx = Int(actualComboCount * Rnd) + 1

        qtyVal = Int(6 * Rnd) + 1
        finishVal = RoundToPlaces(1.5 + (46.5 * Rnd), 3)
        extraVal = RoundToPlaces(0.125 + (0.875 * Rnd), 3)

        r = Rnd
        If r < 0.3 Then
            kerfVal = vbNullString
        Else
            kerfVal = RoundToPlaces(0.06 + (0.06 * Rnd), 3)
        End If

        r = Rnd
        If r < 0.3 Then
            faceVal = vbNullString
        Else
            faceVal = RoundToPlaces(0.05 + (0.2 * Rnd), 3)
        End If

        dataArr(i, 1) = "JOB-" & Format$(Seed, "000")
        dataArr(i, 2) = "RB-" & Format$(i, "0000")
        dataArr(i, 3) = comboMaterials(comboIdx)
        dataArr(i, 4) = comboDiams(comboIdx)
        dataArr(i, 5) = finishVal
        dataArr(i, 6) = extraVal
        dataArr(i, 7) = qtyVal
        dataArr(i, 8) = kerfVal
        dataArr(i, 9) = faceVal
    Next i

    stageMsg = "Write demo data to tblParts"
    lo.DataBodyRange.Columns(idxJob).Value = GetColumnFrom2DArray(dataArr, 1)
    lo.DataBodyRange.Columns(idxPartNo).Value = GetColumnFrom2DArray(dataArr, 2)
    lo.DataBodyRange.Columns(idxMaterial).Value = GetColumnFrom2DArray(dataArr, 3)
    lo.DataBodyRange.Columns(idxDiameter).Value = GetColumnFrom2DArray(dataArr, 4)
    lo.DataBodyRange.Columns(idxFinish).Value = GetColumnFrom2DArray(dataArr, 5)
    lo.DataBodyRange.Columns(idxExtra).Value = GetColumnFrom2DArray(dataArr, 6)
    lo.DataBodyRange.Columns(idxQty).Value = GetColumnFrom2DArray(dataArr, 7)
    lo.DataBodyRange.Columns(idxKerf).Value = GetColumnFrom2DArray(dataArr, 8)
    lo.DataBodyRange.Columns(idxFace).Value = GetColumnFrom2DArray(dataArr, 9)

    stageMsg = "Apply calculated columns and calculate"
    ApplyTblPartsCalculatedColumns lo
    Application.CalculateFull

    stageMsg = "Refresh all after writing parts"
    RefreshAll

    stageMsg = "Apply demo material summary defaults"
    Set matLo = GetTable(WS_MAT, TBL_MAT)
    If matLo Is Nothing Then Err.Raise 9002, PROC, "Material summary table not found."
    ApplyDemoMaterialSummaryDefaults matLo

    stageMsg = "Refresh all after material defaults"
    RefreshAll

    stageMsg = "Ensure validations and print button"
    EnsureInputValidations
    EnsurePrintButton

    LogInfo PROC, "Generate demo data complete"
    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

Public Sub RefreshAll()
    Const PROC As String = "RefreshAll"
    On Error GoTo EH

    Dim stageMsg As String
    Dim partsLo As ListObject
    Dim matLo As ListObject
    Dim settingsDict As Object
    Dim totalsDict As Object
    Dim stockKeys As Collection
    Dim keyCount As Long
    Dim i As Long
    Dim matArr() As Variant
    Dim sk As String
    Dim partsInfo As Variant
    Dim setInfo As Variant
    Dim totalReq As Double

    LogInfo PROC, "RefreshAll start"

    stageMsg = "Get core tables"
    Set partsLo = GetTable(WS_INPUT, TBL_PARTS)
    Set matLo = GetTable(WS_MAT, TBL_MAT)

    If partsLo Is Nothing Then Err.Raise 9101, PROC, "tblParts not found."
    If matLo Is Nothing Then Err.Raise 9102, PROC, "tblMat not found."

    stageMsg = "Apply calculated columns"
    ApplyTblPartsCalculatedColumns partsLo
    Application.CalculateFull

    stageMsg = "Capture existing stock settings"
    Set settingsDict = CaptureExistingStockSettings()

    stageMsg = "Build parts totals dict"
    Set totalsDict = BuildPartsTotalsDict()

    stageMsg = "Sort stock keys"
    Set stockKeys = GetSortedKeyCollection(totalsDict)
    keyCount = stockKeys.Count

    If keyCount = 0 Then
        stageMsg = "Clear tables for no stock keys"
        ResetTableToSingleBlankRow matLo
        ResetTableToSingleBlankRow GetTable(WS_PURCHASE, TBL_PURCHASE)
        SeedMaterialsListFromParts
        EnsureInputValidations
        EnsurePrintButton
        LogInfo PROC, "No stock keys found. Cleared summary tables."
        Exit Sub
    End If

    stageMsg = "Build material summary array"
    ReDim matArr(1 To keyCount, 1 To 9)

    For i = 1 To keyCount
        sk = CStr(stockKeys(i))
        partsInfo = ParseStockKey(sk)
        setInfo = GetSettingInfo(settingsDict, sk)

        totalReq = NzDouble(totalsDict(sk), 0#)

        matArr(i, 1) = sk
        matArr(i, 2) = CStr(partsInfo(0))
        matArr(i, 3) = CDbl(partsInfo(1))
        matArr(i, 4) = RoundToPlaces(totalReq, 6)
        matArr(i, 5) = RoundToPlaces(totalReq / 12#, 6)
        matArr(i, 6) = NzDouble(setInfo(0), 144#)
        If NzDouble(setInfo(1), 0#) > 0# Then
            matArr(i, 7) = NzDouble(setInfo(1), 0#)
        Else
            matArr(i, 7) = vbNullString
        End If
        matArr(i, 8) = CBool(setInfo(2))
        matArr(i, 9) = CStr(setInfo(3))
    Next i

    stageMsg = "Load material summary table"
    LoadArrayIntoTable matLo, matArr, keyCount, 9

    stageMsg = "Rebuild purchase plan rows"
    RebuildPurchasePlanRows totalsDict, settingsDict

    stageMsg = "Reseed materials list and validations"
    SeedMaterialsListFromParts
    EnsureInputValidations
    EnsurePrintButton

    LogInfo PROC, "RefreshAll complete. StockKeys=" & CStr(keyCount)
    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

Public Sub RunBasicNesting()
    Const PROC As String = "RunBasicNesting"
    On Error GoTo EH

    Dim stageMsg As String
    Dim pieces() As Variant
    Dim pieceCount As Long
    Dim settings As Object
    Dim outArr() As Variant
    Dim outRow As Long
    Dim i As Long
    Dim currSk As String
    Dim sk As String
    Dim pieceID As String
    Dim cutLen As Double
    Dim buy1 As Double
    Dim buy2 As Double
    Dim prefer1 As Boolean
    Dim currentBarNo As Long
    Dim currentBuyLen As Double
    Dim remaining As Double
    Dim lastBarOutRow As Long
    Dim ws As Worksheet

    LogInfo PROC, "RunBasicNesting start"

    stageMsg = "RefreshAll"
    RefreshAll

    stageMsg = "BuildPieceArray"
    BuildPieceArray pieces, pieceCount

    stageMsg = "Get NEST_BASIC sheet"
    Set ws = GetWorksheet(WS_BASIC)
    If ws Is Nothing Then Err.Raise 9201, PROC, WS_BASIC & " not found."

    stageMsg = "Clear and write headers"
    ClearPlainOutputSheet ws
    WriteBasicHeaders ws

    If pieceCount = 0 Then
        LogInfo PROC, "No pieces found. Exiting."
        Exit Sub
    End If

    stageMsg = "Build settings dict"
    Set settings = BuildSettingsDict()

    stageMsg = "Dimension output array"
    ReDim outArr(1 To pieceCount, 1 To 10)

    currSk = vbNullString
    currentBarNo = 0
    currentBuyLen = 0#
    remaining = 0#
    outRow = 0
    lastBarOutRow = 0

    stageMsg = "Pack BASIC nesting rows"
    For i = 1 To pieceCount
        sk = CStr(pieces(i, 1))
        cutLen = NzDouble(pieces(i, 2), 0#)
        pieceID = CStr(pieces(i, 3))

        GetSettingsForKey settings, sk, buy1, buy2, prefer1

        If currSk <> sk Then
            If currSk <> vbNullString And lastBarOutRow > 0 Then
                FinalizeBarRecord outArr, lastBarOutRow, remaining
            End If
            currSk = sk
            currentBarNo = 0
            currentBuyLen = 0#
            remaining = 0#
            lastBarOutRow = 0
        End If

        If currentBarNo = 0 Then
            currentBarNo = 1
            currentBuyLen = ChooseNewBarLength(cutLen, buy1, buy2, prefer1)
            remaining = currentBuyLen
            LogInfo PROC, "Opened first BASIC bar for " & sk & " buy=" & FormatNumberSafe(currentBuyLen)
        ElseIf cutLen > remaining + TOL Then
            If lastBarOutRow > 0 Then
                FinalizeBarRecord outArr, lastBarOutRow, remaining
            End If
            currentBarNo = currentBarNo + 1
            currentBuyLen = ChooseNewBarLength(cutLen, buy1, buy2, prefer1)
            remaining = currentBuyLen
            LogInfo PROC, "Opened next BASIC bar for " & sk & " bar#=" & CStr(currentBarNo) & " buy=" & FormatNumberSafe(currentBuyLen)
        End If

        outRow = outRow + 1
        outArr(outRow, 1) = sk
        outArr(outRow, 2) = currentBuyLen
        outArr(outRow, 3) = pieceID
        outArr(outRow, 4) = cutLen
        outArr(outRow, 5) = currentBarNo
        outArr(outRow, 6) = RoundToPlaces(remaining, 6)
        remaining = RoundToPlaces(remaining - cutLen, 6)
        outArr(outRow, 7) = remaining
        outArr(outRow, 8) = False
        outArr(outRow, 9) = vbNullString
        outArr(outRow, 10) = vbNullString
        lastBarOutRow = outRow
    Next i

    If lastBarOutRow > 0 Then
        FinalizeBarRecord outArr, lastBarOutRow, remaining
    End If

    stageMsg = "Write BASIC output sheet"
    WriteArrayToRange ws.Range("A2"), outArr, outRow, 10

    LogInfo PROC, "RunBasicNesting complete. Output rows=" & CStr(outRow)
    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

Public Sub RunFinalNesting()
    Const PROC As String = "RunFinalNesting"
    On Error GoTo EH

    Dim stageMsg As String
    Dim pieces() As Variant
    Dim pieceCount As Long
    Dim settings As Object
    Dim totalsDict As Object
    Dim basicCounts As Object
    Dim outArr() As Variant
    Dim outRow As Long
    Dim i As Long
    Dim sk As String
    Dim currSk As String
    Dim pieceID As String
    Dim cutLen As Double
    Dim buy1 As Double
    Dim buy2 As Double
    Dim prefer1 As Boolean
    Dim chosenBar As Long
    Dim bestSlack As Double
    Dim slack As Double
    Dim barCount As Long
    Dim barRemain() As Double
    Dim barBuy() As Double
    Dim barLastRow() As Long
    Dim b As Long
    Dim finalLo As ListObject
    Dim barMetrics As Object

    LogInfo PROC, "RunFinalNesting start"

    stageMsg = "RefreshAll"
    RefreshAll

    stageMsg = "RunBasicNesting"
    RunBasicNesting

    stageMsg = "Build piece array"
    BuildPieceArray pieces, pieceCount

    stageMsg = "Get final nesting table"
    Set finalLo = GetTable(WS_FINAL, TBL_FINAL)
    If finalLo Is Nothing Then Err.Raise 9301, PROC, "tblNestFinal not found."

    If pieceCount = 0 Then
        ResetTableToSingleBlankRow finalLo
        SetNamedRangeValue "LastFinalRun", vbNullString
        LogInfo PROC, "No pieces found. Exiting."
        Exit Sub
    End If

    stageMsg = "Build dictionaries"
    Set settings = BuildSettingsDict()
    Set totalsDict = BuildPartsTotalsDict()
    Set basicCounts = BuildBasicBarCountDict()
    Set barMetrics = CreateObject("Scripting.Dictionary")
    barMetrics.CompareMode = 1

    ReDim outArr(1 To pieceCount, 1 To 10)

    currSk = vbNullString
    outRow = 0
    barCount = 0

    stageMsg = "Pack final nesting bars"
    For i = 1 To pieceCount
        sk = CStr(pieces(i, 1))
        cutLen = NzDouble(pieces(i, 2), 0#)
        pieceID = CStr(pieces(i, 3))

        GetSettingsForKey settings, sk, buy1, buy2, prefer1

        If currSk <> sk Then
            If currSk <> vbNullString Then
                FinalizeAllBarsForStockKey outArr, barLastRow, barRemain, barCount, currSk, barMetrics
            End If
            currSk = sk
            barCount = 0
            Erase barRemain
            Erase barBuy
            Erase barLastRow
            LogInfo PROC, "Starting FINAL stock key " & currSk
        End If

        chosenBar = 0
        bestSlack = 1E+99

        For b = 1 To barCount
            If cutLen <= barRemain(b) + TOL Then
                slack = barRemain(b) - cutLen
                If (slack < bestSlack - TOL) Or ((Abs(slack - bestSlack) <= TOL) And (b < chosenBar Or chosenBar = 0)) Then
                    chosenBar = b
                    bestSlack = slack
                End If
            End If
        Next b

        If chosenBar = 0 Then
            barCount = barCount + 1

            If barCount = 1 Then
                ReDim barRemain(1 To 1)
                ReDim barBuy(1 To 1)
                ReDim barLastRow(1 To 1)
            Else
                ReDim Preserve barRemain(1 To barCount)
                ReDim Preserve barBuy(1 To barCount)
                ReDim Preserve barLastRow(1 To barCount)
            End If

            barBuy(barCount) = ChooseNewBarLength(cutLen, buy1, buy2, prefer1)
            barRemain(barCount) = barBuy(barCount)
            barLastRow(barCount) = 0
            chosenBar = barCount

            LogInfo PROC, "Opened FINAL bar. StockKey=" & sk & ", Bar#=" & CStr(chosenBar) & ", BuyLen=" & FormatNumberSafe(barBuy(chosenBar))
        End If

        outRow = outRow + 1
        outArr(outRow, 1) = sk
        outArr(outRow, 2) = barBuy(chosenBar)
        outArr(outRow, 3) = pieceID
        outArr(outRow, 4) = cutLen
        outArr(outRow, 5) = chosenBar
        outArr(outRow, 6) = RoundToPlaces(barRemain(chosenBar), 6)
        barRemain(chosenBar) = RoundToPlaces(barRemain(chosenBar) - cutLen, 6)
        outArr(outRow, 7) = barRemain(chosenBar)
        outArr(outRow, 8) = False
        outArr(outRow, 9) = vbNullString
        outArr(outRow, 10) = vbNullString

        barLastRow(chosenBar) = outRow
    Next i

    If currSk <> vbNullString Then
        FinalizeAllBarsForStockKey outArr, barLastRow, barRemain, barCount, currSk, barMetrics
    End If

    stageMsg = "Load final table and update purchase plan"
    LoadArrayIntoTable finalLo, outArr, outRow, 10
    SetNamedRangeValue "LastFinalRun", CDbl(Now)

    RebuildPurchasePlanRows totalsDict, settings
    UpdatePurchasePlanFromMetrics basicCounts, totalsDict, barMetrics
    BuildCutlistFromFinal

    LogInfo PROC, "RunFinalNesting complete. Output rows=" & CStr(outRow)
    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

Public Sub BuildCutlistCurrent()
    Const PROC As String = "BuildCutlistCurrent"
    On Error GoTo EH

    LogInfo PROC, "Build-only cutlist request"
    BuildCutlistFromFinal
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Public Sub BuildCutlistFromFinal()
    Const PROC As String = "BuildCutlistFromFinal"
    On Error GoTo EH

    Dim stageMsg As String
    Dim finalLo As ListObject
    Dim cutLo As ListObject
    Dim rowCount As Long
    Dim dict As Object
    Dim orderArr() As String
    Dim keyCount As Long
    Dim i As Long
    Dim k As String
    Dim stockKey As String
    Dim barNo As Long
    Dim pieceID As String
    Dim cutLen As Double
    Dim isBarEnd As Boolean
    Dim leftover As Double
    Dim remStatus As String
    Dim rec As Variant

    Dim idxStockKey As Long
    Dim idxBuyLen As Long
    Dim idxPieceID As Long
    Dim idxCutLen As Long
    Dim idxBarNo As Long
    Dim idxIsBarEnd As Long
    Dim idxLeftover As Long
    Dim idxRemnant As Long

    Dim outArr() As Variant

    LogInfo PROC, "BuildCutlistFromFinal start"

    stageMsg = "Ensure final nesting exists"
    If Not FinalNestingHasRows() Then
        LogInfo PROC, "NEST_FINAL blank or not current enough for cutlist. Running RunFinalNesting."
        RunFinalNesting
    End If

    stageMsg = "Get final/cutlist tables"
    Set finalLo = GetTable(WS_FINAL, TBL_FINAL)
    Set cutLo = GetTable(WS_CUTLIST, TBL_CUTLIST)

    If finalLo Is Nothing Then Err.Raise 9401, PROC, "tblNestFinal not found."
    If cutLo Is Nothing Then Err.Raise 9402, PROC, "tblCutlist not found."

    If IsTableEffectivelyBlank(finalLo) Then
        ResetTableToSingleBlankRow cutLo
        SetNamedRangeValue "LastCutlistGen", vbNullString
        EnsurePrintButton
        LogInfo PROC, "No final nesting rows found after refresh."
        Exit Sub
    End If

    stageMsg = "Get final table column indexes"
    idxStockKey = GetListColumnIndex(finalLo, "StockKey", PROC)
    idxBuyLen = GetListColumnIndex(finalLo, "BuyLen_in", PROC)
    idxPieceID = GetListColumnIndex(finalLo, "PieceID", PROC)
    idxCutLen = GetListColumnIndex(finalLo, "CutLen_in", PROC)
    idxBarNo = GetListColumnIndex(finalLo, "Bar#", PROC)
    idxIsBarEnd = GetListColumnIndex(finalLo, "IsBarEnd", PROC)
    idxLeftover = GetListColumnIndex(finalLo, "LeftoverAtEnd_in", PROC)
    idxRemnant = GetListColumnIndex(finalLo, "RemnantStatus", PROC)

    rowCount = finalLo.ListRows.Count

    stageMsg = "Build cutlist grouping dictionary"
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    keyCount = 0

    For i = 1 To rowCount
        stockKey = Trim$(CStr(finalLo.DataBodyRange.Cells(i, idxStockKey).Value))
        If stockKey <> vbNullString Then
            barNo = CLng(NzDouble(finalLo.DataBodyRange.Cells(i, idxBarNo).Value, 0))
            k = stockKey & Chr$(30) & CStr(barNo)

            If Not dict.Exists(k) Then
                keyCount = keyCount + 1

                If keyCount = 1 Then
                    ReDim orderArr(1 To 1)
                Else
                    ReDim Preserve orderArr(1 To keyCount)
                End If

                orderArr(keyCount) = k
                dict.Add k, Array(stockKey, barNo, NzDouble(finalLo.DataBodyRange.Cells(i, idxBuyLen).Value, 0#), vbNullString, 0#, 0#, vbNullString)
            End If

            rec = dict(k)
            pieceID = Trim$(CStr(finalLo.DataBodyRange.Cells(i, idxPieceID).Value))
            cutLen = NzDouble(finalLo.DataBodyRange.Cells(i, idxCutLen).Value, 0#)
            isBarEnd = NzBool(finalLo.DataBodyRange.Cells(i, idxIsBarEnd).Value, False)
            leftover = NzDouble(finalLo.DataBodyRange.Cells(i, idxLeftover).Value, 0#)
            remStatus = Trim$(CStr(finalLo.DataBodyRange.Cells(i, idxRemnant).Value))

            If pieceID <> vbNullString Then
                If CStr(rec(3)) = vbNullString Then
                    rec(3) = pieceID
                Else
                    rec(3) = CStr(rec(3)) & ", " & pieceID
                End If
            End If

            rec(4) = NzDouble(rec(4), 0#) + cutLen

            If isBarEnd Then
                rec(5) = leftover
                rec(6) = remStatus
            End If

            dict(k) = rec
        End If
    Next i

    If keyCount = 0 Then
        ResetTableToSingleBlankRow cutLo
        SetNamedRangeValue "LastCutlistGen", vbNullString
        EnsurePrintButton
        LogInfo PROC, "No cutlist rows built."
        Exit Sub
    End If

    stageMsg = "Write cutlist output array"
    ReDim outArr(1 To keyCount, 1 To 8)

    For i = 1 To keyCount
        rec = dict(orderArr(i))
        outArr(i, 1) = rec(0)
        outArr(i, 2) = rec(1)
        outArr(i, 3) = rec(2)
        outArr(i, 4) = rec(3)
        outArr(i, 5) = RoundToPlaces(NzDouble(rec(4), 0#), 6)
        outArr(i, 6) = RoundToPlaces(NzDouble(rec(5), 0#), 6)
        outArr(i, 7) = rec(6)
        If NzDouble(rec(2), 0#) > 0# Then
            outArr(i, 8) = RoundToPlaces(NzDouble(rec(4), 0#) / NzDouble(rec(2), 1#), 6)
        Else
            outArr(i, 8) = 0#
        End If
    Next i

    stageMsg = "Load cutlist table"
    LoadArrayIntoTable cutLo, outArr, keyCount, 8
    SetNamedRangeValue "LastCutlistGen", CDbl(Now)
    EnsurePrintButton

    LogInfo PROC, "BuildCutlistFromFinal complete. Rows=" & CStr(keyCount)
    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

Public Sub PrintCutlist()
    Const PROC As String = "PrintCutlist"
    On Error GoTo EH

    Dim stageMsg As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lastRow As Long
    Dim i As Long
    Dim prevKey As String
    Dim thisKey As String

    LogInfo PROC, "PrintCutlist start"

    stageMsg = "Ensure cutlist current"
    BuildCutlistFromFinal

    Set ws = GetWorksheet(WS_CUTLIST)
    Set lo = GetTable(WS_CUTLIST, TBL_CUTLIST)
    If ws Is Nothing Then Err.Raise 9501, PROC, WS_CUTLIST & " not found."
    If lo Is Nothing Then Err.Raise 9502, PROC, TBL_CUTLIST & " not found."

    If IsTableEffectivelyBlank(lo) Then
        LogInfo PROC, "No cutlist rows available to print."
        Exit Sub
    End If

    stageMsg = "Ensure button and reset page breaks"
    EnsurePrintButton
    ResetSheetPageBreaks ws

    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then
        LogInfo PROC, "Nothing to print."
        Exit Sub
    End If

    stageMsg = "Insert page breaks by StockKey"
    prevKey = vbNullString
    For i = 2 To lastRow
        thisKey = Trim$(CStr(ws.Cells(i, 1).Value))
        If i > 2 Then
            If thisKey <> vbNullString And thisKey <> prevKey Then
                ws.HPageBreaks.Add Before:=ws.Rows(i)
            End If
        End If
        prevKey = thisKey
    Next i

    stageMsg = "Apply page setup"
    On Error Resume Next
    With ws.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.3)
        .BottomMargin = Application.InchesToPoints(0.3)
        .HeaderMargin = Application.InchesToPoints(0.15)
        .FooterMargin = Application.InchesToPoints(0.15)
        .PrintTitleRows = "$1:$1"
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 8)).Address
    End With
    If Err.Number <> 0 Then
        LogInfo PROC, "Page setup warning: " & Err.Description
        Err.Clear
    End If
    On Error GoTo EH

    stageMsg = "Open print preview"
    On Error Resume Next
    ws.PrintPreview
    If Err.Number <> 0 Then
        LogInfo PROC, "Print preview warning: " & Err.Description
        Err.Clear
        ws.Activate
        MsgBox "Cutlist built successfully, but Excel could not open Print Preview. Check your default printer / printer driver and try again.", vbExclamation, "Print Preview Warning"
        Exit Sub
    End If
    On Error GoTo EH

    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

'=========================================================
' RESET / SETUP HELPERS
'=========================================================
Private Sub SeedDefaultNamedValues()
    SetNamedRangeValue "DefaultKerf_in", 0.125
    SetNamedRangeValue "DefaultFaceAllow_in", 0.125
    SetNamedRangeValue "MinRemnantKeep_in", 12#
    SetNamedRangeValue "MaxPartRows", 500
    SetNamedRangeValue "MaxPieces", 20000
End Sub

Private Sub ResetInputPartsTable()
    Const PROC As String = "ResetInputPartsTable"
    On Error GoTo EH

    Dim lo As ListObject
    Dim entryCols As Variant
    Dim i As Long
    Dim idx As Long

    Set lo = GetTable(WS_INPUT, TBL_PARTS)
    If lo Is Nothing Then Exit Sub

    EnsureTableRows lo, 1

    entryCols = Array("Job", "PartNo", "Material", "Diameter_in", "FinishLen_in", "ExtraStock_in", "Qty", "Kerf_in", "FaceAllow_in")

    For i = LBound(entryCols) To UBound(entryCols)
        idx = GetListColumnIndex(lo, CStr(entryCols(i)), PROC)
        lo.DataBodyRange.Columns(idx).ClearContents
    Next i

    ApplyTblPartsCalculatedColumns lo
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Sub SeedMaterialsListDefault()
    Dim arr(1 To 20) As String
    arr(1) = "1018"
    arr(2) = "12L14"
    arr(3) = "4140"
    arr(4) = "4340"
    arr(5) = "8620"
    arr(6) = "17-4PH"
    arr(7) = "304"
    arr(8) = "316"
    arr(9) = "O1"
    arr(10) = "D2"
    arr(11) = "H13"
    arr(12) = "A2"
    arr(13) = "Brass360"
    arr(14) = "Al6061"
    arr(15) = "Al7075"
    arr(16) = "Ti-6Al-4V"
    arr(17) = "Monel400"
    arr(18) = "Inconel625"
    arr(19) = "UHMW"
    arr(20) = "Acetal"
    SeedMaterialsListFromArrays arr
End Sub

Private Sub SeedMaterialsListFromParts()
    Const PROC As String = "SeedMaterialsListFromParts"
    On Error GoTo EH

    Dim lo As ListObject
    Dim dict As Object
    Dim keysCol As Collection
    Dim i As Long
    Dim keyCount As Long
    Dim ws As Worksheet
    Dim idxMaterial As Long
    Dim rowCount As Long
    Dim matVal As String
    Dim outArr() As Variant

    Set lo = GetTable(WS_INPUT, TBL_PARTS)
    Set ws = GetWorksheet(WS_INPUT)
    If lo Is Nothing Or ws Is Nothing Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    idxMaterial = GetListColumnIndex(lo, "Material", PROC)

    If Not IsTableEffectivelyBlank(lo) Then
        rowCount = lo.ListRows.Count
        For i = 1 To rowCount
            matVal = Trim$(CStr(lo.DataBodyRange.Cells(i, idxMaterial).Value))
            If matVal <> vbNullString Then
                dict(matVal) = True
            End If
        Next i
    End If

    If dict.Count = 0 Then
        SeedMaterialsListDefault
        Exit Sub
    End If

    Set keysCol = GetSortedKeyCollection(dict)
    keyCount = keysCol.Count

    ws.Range("G2").Value = "Materials"
    ws.Range("G3:G102").ClearContents

    If keyCount > 0 Then
        ReDim outArr(1 To keyCount, 1 To 1)
        For i = 1 To keyCount
            outArr(i, 1) = CStr(keysCol(i))
        Next i
        ws.Range("G3").Resize(keyCount, 1).Value = outArr
    End If

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Sub SeedMaterialsListFromArrays(ByRef materialArr() As String)
    Dim ws As Worksheet
    Dim n As Long

    Set ws = GetWorksheet(WS_INPUT)
    If ws Is Nothing Then Exit Sub

    n = UBound(materialArr) - LBound(materialArr) + 1
    WriteMaterialsRange ws, materialArr, n
End Sub

Private Sub WriteMaterialsRange(ByVal ws As Worksheet, ByRef sortedKeys() As String, ByVal keyCount As Long)
    Dim outArr() As Variant
    Dim i As Long

    ws.Range("G2").Value = "Materials"
    ws.Range("G3:G102").ClearContents

    If keyCount > 0 Then
        ReDim outArr(1 To keyCount, 1 To 1)
        For i = 1 To keyCount
            outArr(i, 1) = sortedKeys(i)
        Next i
        ws.Range("G3").Resize(keyCount, 1).Value = outArr
    End If
End Sub

Private Sub ApplyDemoMaterialSummaryDefaults(ByVal matLo As ListObject)
    Const PROC As String = "ApplyDemoMaterialSummaryDefaults"
    On Error GoTo EH

    Dim rowCount As Long
    Dim i As Long
    Dim idxStockKey As Long
    Dim idxBuy1 As Long
    Dim idxBuy2 As Long
    Dim idxPrefer As Long

    If matLo Is Nothing Then Err.Raise 9010, PROC, "Material summary table is Nothing."
    If matLo.DataBodyRange Is Nothing Then Exit Sub

    idxStockKey = GetListColumnIndex(matLo, "StockKey", PROC)
    idxBuy1 = GetListColumnIndex(matLo, "BuyLen1_in", PROC)
    idxBuy2 = GetListColumnIndex(matLo, "BuyLen2_in", PROC)
    idxPrefer = GetListColumnIndex(matLo, "PreferBuyLen1", PROC)

    rowCount = matLo.ListRows.Count
    LogInfo PROC, "Applying demo buy lengths. Rows=" & CStr(rowCount)

    If rowCount < 1 Then Exit Sub

    For i = 1 To rowCount
        If Trim$(CStr(matLo.DataBodyRange.Cells(i, idxStockKey).Value)) <> vbNullString Then
            matLo.DataBodyRange.Cells(i, idxBuy1).Value = 144#
            If (i Mod 2) = 0 Then
                matLo.DataBodyRange.Cells(i, idxBuy2).Value = 96#
                matLo.DataBodyRange.Cells(i, idxPrefer).Value = False
            Else
                matLo.DataBodyRange.Cells(i, idxBuy2).Value = vbNullString
                matLo.DataBodyRange.Cells(i, idxPrefer).Value = True
            End If
        End If
    Next i

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

'=========================================================
' NESTING HELPERS
'=========================================================
Private Sub BuildPieceArray(ByRef pieces() As Variant, ByRef pieceCount As Long)
    Const PROC As String = "BuildPieceArray"
    On Error GoTo EH

    Dim stageMsg As String
    Dim lo As ListObject
    Dim i As Long
    Dim qtyVal As Long
    Dim j As Long
    Dim sk As String
    Dim cutLen As Double
    Dim partNo As String
    Dim idxStockKey As Long
    Dim idxReqCut As Long
    Dim idxQty As Long
    Dim idxPartNo As Long
    Dim rowCount As Long

    stageMsg = "Get tblParts"
    Set lo = GetTable(WS_INPUT, TBL_PARTS)
    If lo Is Nothing Then Err.Raise 9601, PROC, "tblParts not found."

    pieceCount = 0
    If IsTableEffectivelyBlank(lo) Then Exit Sub

    stageMsg = "Get required column indexes"
    idxStockKey = GetListColumnIndex(lo, "StockKey", PROC)
    idxReqCut = GetListColumnIndex(lo, "RequiredCutLen_in", PROC)
    idxQty = GetListColumnIndex(lo, "Qty", PROC)
    idxPartNo = GetListColumnIndex(lo, "PartNo", PROC)

    rowCount = lo.ListRows.Count

    stageMsg = "Count total pieces"
    For i = 1 To rowCount
        sk = Trim$(CStr(lo.DataBodyRange.Cells(i, idxStockKey).Value))
        cutLen = NzDouble(lo.DataBodyRange.Cells(i, idxReqCut).Value, 0#)
        qtyVal = CLng(NzDouble(lo.DataBodyRange.Cells(i, idxQty).Value, 0))
        If sk <> vbNullString And cutLen > TOL And qtyVal > 0 Then
            pieceCount = pieceCount + qtyVal
        End If
    Next i

    If pieceCount = 0 Then Exit Sub

    stageMsg = "Allocate piece array"
    ReDim pieces(1 To pieceCount, 1 To 3)
    pieceCount = 0

    stageMsg = "Populate piece array"
    For i = 1 To rowCount
        sk = Trim$(CStr(lo.DataBodyRange.Cells(i, idxStockKey).Value))
        cutLen = NzDouble(lo.DataBodyRange.Cells(i, idxReqCut).Value, 0#)
        qtyVal = CLng(NzDouble(lo.DataBodyRange.Cells(i, idxQty).Value, 0))
        partNo = Trim$(CStr(lo.DataBodyRange.Cells(i, idxPartNo).Value))

        If sk <> vbNullString And cutLen > TOL And qtyVal > 0 Then
            For j = 1 To qtyVal
                pieceCount = pieceCount + 1
                pieces(pieceCount, 1) = sk
                pieces(pieceCount, 2) = RoundToPlaces(cutLen, 6)
                pieces(pieceCount, 3) = partNo & "-" & Format$(j, "000")
            Next j
        End If
    Next i

    stageMsg = "Sort piece array"
    SortPieces pieces, pieceCount

    LogInfo PROC, "Built piece array. Count=" & CStr(pieceCount)
    Exit Sub

EH:
    LogError PROC, Err.Number, "Stage=" & stageMsg & " | " & Err.Description
    Err.Raise Err.Number, PROC, "Stage=" & stageMsg & " | " & Err.Description
End Sub

Private Sub SortPieces(ByRef pieces() As Variant, ByVal pieceCount As Long)
    Const PROC As String = "SortPieces"
    On Error GoTo EH

    Dim i As Long
    Dim j As Long
    Dim tmp1 As Variant
    Dim tmp2 As Variant
    Dim tmp3 As Variant

    If pieceCount <= 1 Then Exit Sub

    For i = 2 To pieceCount
        tmp1 = pieces(i, 1)
        tmp2 = pieces(i, 2)
        tmp3 = pieces(i, 3)
        j = i - 1

        Do While j >= 1
            If ComparePieceRow(CStr(pieces(j, 1)), NzDouble(pieces(j, 2), 0#), CStr(pieces(j, 3)), _
                               CStr(tmp1), NzDouble(tmp2, 0#), CStr(tmp3)) > 0 Then
                pieces(j + 1, 1) = pieces(j, 1)
                pieces(j + 1, 2) = pieces(j, 2)
                pieces(j + 1, 3) = pieces(j, 3)
                j = j - 1
            Else
                Exit Do
            End If
        Loop

        pieces(j + 1, 1) = tmp1
        pieces(j + 1, 2) = tmp2
        pieces(j + 1, 3) = tmp3
    Next i

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Function ComparePieceRow(ByVal aStock As String, ByVal aLen As Double, ByVal aID As String, _
                                 ByVal bStock As String, ByVal bLen As Double, ByVal bID As String) As Long
    If StrComp(aStock, bStock, vbTextCompare) < 0 Then
        ComparePieceRow = -1
    ElseIf StrComp(aStock, bStock, vbTextCompare) > 0 Then
        ComparePieceRow = 1
    ElseIf aLen > bLen + TOL Then
        ComparePieceRow = -1
    ElseIf aLen < bLen - TOL Then
        ComparePieceRow = 1
    Else
        ComparePieceRow = StrComp(aID, bID, vbTextCompare)
    End If
End Function

Private Sub GetSettingsForKey(ByVal settings As Object, ByVal stockKey As String, ByRef buy1 As Double, ByRef buy2 As Double, ByRef prefer1 As Boolean)
    Dim v As Variant

    buy1 = 144#
    buy2 = 0#
    prefer1 = True

    If settings Is Nothing Then Exit Sub
    If Not settings.Exists(stockKey) Then Exit Sub

    v = settings(stockKey)
    buy1 = NzDouble(v(0), 144#)
    buy2 = NzDouble(v(1), 0#)
    prefer1 = NzBool(v(2), True)

    If buy1 <= 0# Then buy1 = 144#
End Sub

Private Function ChooseNewBarLength(ByVal cutLen As Double, ByVal buy1 As Double, ByVal buy2 As Double, ByVal prefer1 As Boolean) As Double
    If buy1 <= 0# Then buy1 = 144#

    If prefer1 Then
        ChooseNewBarLength = buy1
    ElseIf buy2 > 0# Then
        If cutLen <= buy2 + TOL Then
            ChooseNewBarLength = buy2
        Else
            ChooseNewBarLength = buy1
        End If
    Else
        ChooseNewBarLength = buy1
    End If

    If ChooseNewBarLength < cutLen Then
        ChooseNewBarLength = cutLen
    End If
End Function

Private Sub FinalizeBarRecord(ByRef outArr() As Variant, ByVal rowIndex As Long, ByVal leftover As Double)
    If rowIndex < 1 Then Exit Sub
    outArr(rowIndex, 8) = True
    outArr(rowIndex, 9) = RoundToPlaces(leftover, 6)
    If leftover >= GetNamedRangeNumber("MinRemnantKeep_in", 12#) - TOL Then
        outArr(rowIndex, 10) = "KEEP"
    Else
        outArr(rowIndex, 10) = "SCRAP"
    End If
End Sub

Private Sub FinalizeAllBarsForStockKey(ByRef outArr() As Variant, ByRef barLastRow() As Long, ByRef barRemain() As Double, _
                                       ByVal barCount As Long, ByVal stockKey As String, ByVal barMetrics As Object)
    Dim b As Long
    Dim purchased As Double
    Dim scrapLen As Double
    Dim remnantLen As Double
    Dim status As String

    purchased = 0#
    scrapLen = 0#
    remnantLen = 0#

    If barCount <= 0 Then Exit Sub

    For b = 1 To barCount
        If barLastRow(b) > 0 Then
            FinalizeBarRecord outArr, barLastRow(b), barRemain(b)
            purchased = purchased + NzDouble(outArr(barLastRow(b), 2), 0#)
            status = CStr(outArr(barLastRow(b), 10))
            If status = "KEEP" Then
                remnantLen = remnantLen + NzDouble(outArr(barLastRow(b), 9), 0#)
            Else
                scrapLen = scrapLen + NzDouble(outArr(barLastRow(b), 9), 0#)
            End If
        End If
    Next b

    If Not barMetrics Is Nothing Then
        barMetrics(stockKey) = Array(barCount, purchased, scrapLen, remnantLen)
    End If
End Sub

'=========================================================
' PURCHASE PLAN / SUMMARY HELPERS
'=========================================================
Private Sub RebuildPurchasePlanRows(ByVal totalsDict As Object, ByVal settings As Object)
    Const PROC As String = "RebuildPurchasePlanRows"
    On Error GoTo EH

    Dim lo As ListObject
    Dim stockKeys As Collection
    Dim keyCount As Long
    Dim arr() As Variant
    Dim i As Long
    Dim sk As String
    Dim buy1 As Double
    Dim buy2 As Double
    Dim prefer1 As Boolean

    Set lo = GetTable(WS_PURCHASE, TBL_PURCHASE)
    If lo Is Nothing Then Err.Raise 9700, PROC, "tblPurchase not found."

    Set stockKeys = GetSortedKeyCollection(totalsDict)
    keyCount = stockKeys.Count

    If keyCount = 0 Then
        ResetTableToSingleBlankRow lo
        Exit Sub
    End If

    ReDim arr(1 To keyCount, 1 To 13)

    For i = 1 To keyCount
        sk = CStr(stockKeys(i))
        GetSettingsForKey settings, sk, buy1, buy2, prefer1

        arr(i, 1) = sk
        arr(i, 2) = buy1
        If buy2 > 0# Then
            arr(i, 3) = buy2
        Else
            arr(i, 3) = vbNullString
        End If
        arr(i, 4) = vbNullString
        arr(i, 5) = vbNullString
        arr(i, 6) = vbNullString
        arr(i, 7) = RoundToPlaces(NzDouble(totalsDict(sk), 0#), 6)
        arr(i, 8) = vbNullString
        arr(i, 9) = vbNullString
        arr(i, 10) = vbNullString
        arr(i, 11) = vbNullString
        arr(i, 12) = vbNullString
        arr(i, 13) = vbNullString
    Next i

    LoadArrayIntoTable lo, arr, keyCount, 13
    LogInfo PROC, "Rebuilt purchase plan rows. Count=" & CStr(keyCount)
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Sub UpdatePurchasePlanFromMetrics(ByVal basicCounts As Object, ByVal totalsDict As Object, ByVal finalMetrics As Object)
    Const PROC As String = "UpdatePurchasePlanFromMetrics"
    On Error GoTo EH

    Dim lo As ListObject
    Dim rowCount As Long
    Dim i As Long
    Dim sk As String
    Dim basicBars As Long
    Dim finalBars As Long
    Dim purchased As Double
    Dim required As Double
    Dim dropLen As Double
    Dim wastePct As Double
    Dim scrapLen As Double
    Dim remnantLen As Double
    Dim scrapPct As Double
    Dim remPct As Double
    Dim rec As Variant

    Dim idxStockKey As Long
    Dim idxBarsBasic As Long
    Dim idxBarsFinal As Long
    Dim idxPurchased As Long
    Dim idxReq As Long
    Dim idxDrop As Long
    Dim idxWaste As Long
    Dim idxScrapLen As Long
    Dim idxRemnantLen As Long
    Dim idxScrapPct As Long
    Dim idxRemnantPct As Long

    Set lo = GetTable(WS_PURCHASE, TBL_PURCHASE)
    If lo Is Nothing Then Err.Raise 9701, PROC, "tblPurchase not found."

    If IsTableEffectivelyBlank(lo) Then
        LogInfo PROC, "Purchase table blank. Rebuild should have happened earlier."
        Exit Sub
    End If

    idxStockKey = GetListColumnIndex(lo, "StockKey", PROC)
    idxBarsBasic = GetListColumnIndex(lo, "BarsNeeded_BASIC", PROC)
    idxBarsFinal = GetListColumnIndex(lo, "BarsNeeded_FINAL", PROC)
    idxPurchased = GetListColumnIndex(lo, "PurchasedLen_FINAL", PROC)
    idxReq = GetListColumnIndex(lo, "TotalReq_in", PROC)
    idxDrop = GetListColumnIndex(lo, "TotalDrop_FINAL", PROC)
    idxWaste = GetListColumnIndex(lo, "Waste%_FINAL", PROC)
    idxScrapLen = GetListColumnIndex(lo, "ScrapLen_FINAL", PROC)
    idxRemnantLen = GetListColumnIndex(lo, "RemnantLen_FINAL", PROC)
    idxScrapPct = GetListColumnIndex(lo, "Scrap%_FINAL", PROC)
    idxRemnantPct = GetListColumnIndex(lo, "Remnant%_FINAL", PROC)

    rowCount = lo.ListRows.Count

    For i = 1 To rowCount
        sk = Trim$(CStr(lo.DataBodyRange.Cells(i, idxStockKey).Value))
        If sk <> vbNullString Then
            basicBars = 0
            If Not basicCounts Is Nothing Then
                If basicCounts.Exists(sk) Then basicBars = CLng(basicCounts(sk))
            End If

            finalBars = 0
            purchased = 0#
            scrapLen = 0#
            remnantLen = 0#

            If Not finalMetrics Is Nothing Then
                If finalMetrics.Exists(sk) Then
                    rec = finalMetrics(sk)
                    finalBars = CLng(rec(0))
                    purchased = NzDouble(rec(1), 0#)
                    scrapLen = NzDouble(rec(2), 0#)
                    remnantLen = NzDouble(rec(3), 0#)
                End If
            End If

            required = 0#
            If Not totalsDict Is Nothing Then
                If totalsDict.Exists(sk) Then required = NzDouble(totalsDict(sk), 0#)
            End If

            dropLen = purchased - required
            If purchased > TOL Then
                wastePct = dropLen / purchased
                scrapPct = scrapLen / purchased
                remPct = remnantLen / purchased
            Else
                wastePct = 0#
                scrapPct = 0#
                remPct = 0#
            End If

            lo.DataBodyRange.Cells(i, idxBarsBasic).Value = basicBars
            lo.DataBodyRange.Cells(i, idxBarsFinal).Value = finalBars
            lo.DataBodyRange.Cells(i, idxPurchased).Value = RoundToPlaces(purchased, 6)
            lo.DataBodyRange.Cells(i, idxReq).Value = RoundToPlaces(required, 6)
            lo.DataBodyRange.Cells(i, idxDrop).Value = RoundToPlaces(dropLen, 6)
            lo.DataBodyRange.Cells(i, idxWaste).Value = RoundToPlaces(wastePct, 6)
            lo.DataBodyRange.Cells(i, idxScrapLen).Value = RoundToPlaces(scrapLen, 6)
            lo.DataBodyRange.Cells(i, idxRemnantLen).Value = RoundToPlaces(remnantLen, 6)
            lo.DataBodyRange.Cells(i, idxScrapPct).Value = RoundToPlaces(scrapPct, 6)
            lo.DataBodyRange.Cells(i, idxRemnantPct).Value = RoundToPlaces(remPct, 6)

            LogInfo PROC, "StockKey=" & sk & _
                         ", barsBasic=" & CStr(basicBars) & _
                         ", barsFinal=" & CStr(finalBars) & _
                         ", purchased=" & FormatNumberSafe(purchased) & _
                         ", required=" & FormatNumberSafe(required) & _
                         ", drop=" & FormatNumberSafe(dropLen) & _
                         ", scrap=" & FormatNumberSafe(scrapLen) & _
                         ", remnant=" & FormatNumberSafe(remnantLen)
        End If
    Next i

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Function CaptureExistingStockSettings() As Object
    Const PROC As String = "CaptureExistingStockSettings"
    On Error GoTo EH

    Dim dict As Object
    Dim lo As ListObject
    Dim i As Long
    Dim rowCount As Long
    Dim sk As String

    Dim idxStockKey As Long
    Dim idxBuy1 As Long
    Dim idxBuy2 As Long
    Dim idxPrefer As Long
    Dim idxNotes As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set lo = GetTable(WS_MAT, TBL_MAT)
    If lo Is Nothing Then
        Set CaptureExistingStockSettings = dict
        Exit Function
    End If

    If IsTableEffectivelyBlank(lo) Then
        Set CaptureExistingStockSettings = dict
        Exit Function
    End If

    idxStockKey = GetListColumnIndex(lo, "StockKey", PROC)
    idxBuy1 = GetListColumnIndex(lo, "BuyLen1_in", PROC)
    idxBuy2 = GetListColumnIndex(lo, "BuyLen2_in", PROC)
    idxPrefer = GetListColumnIndex(lo, "PreferBuyLen1", PROC)
    idxNotes = GetListColumnIndex(lo, "Notes", PROC)

    rowCount = lo.ListRows.Count
    For i = 1 To rowCount
        sk = Trim$(CStr(lo.DataBodyRange.Cells(i, idxStockKey).Value))
        If sk <> vbNullString Then
            dict(sk) = Array( _
                NzDouble(lo.DataBodyRange.Cells(i, idxBuy1).Value, 144#), _
                NzDouble(lo.DataBodyRange.Cells(i, idxBuy2).Value, 0#), _
                NzBool(lo.DataBodyRange.Cells(i, idxPrefer).Value, True), _
                CStr(lo.DataBodyRange.Cells(i, idxNotes).Value) _
            )
        End If
    Next i

    Set CaptureExistingStockSettings = dict
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
    Set CaptureExistingStockSettings = CreateObject("Scripting.Dictionary")
End Function

Private Function BuildSettingsDict() As Object
    Const PROC As String = "BuildSettingsDict"
    On Error GoTo EH

    Dim dict As Object
    Dim lo As ListObject
    Dim i As Long
    Dim rowCount As Long
    Dim sk As String

    Dim idxStockKey As Long
    Dim idxBuy1 As Long
    Dim idxBuy2 As Long
    Dim idxPrefer As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set lo = GetTable(WS_MAT, TBL_MAT)
    If lo Is Nothing Then
        Set BuildSettingsDict = dict
        Exit Function
    End If

    If IsTableEffectivelyBlank(lo) Then
        Set BuildSettingsDict = dict
        Exit Function
    End If

    idxStockKey = GetListColumnIndex(lo, "StockKey", PROC)
    idxBuy1 = GetListColumnIndex(lo, "BuyLen1_in", PROC)
    idxBuy2 = GetListColumnIndex(lo, "BuyLen2_in", PROC)
    idxPrefer = GetListColumnIndex(lo, "PreferBuyLen1", PROC)

    rowCount = lo.ListRows.Count
    For i = 1 To rowCount
        sk = Trim$(CStr(lo.DataBodyRange.Cells(i, idxStockKey).Value))
        If sk <> vbNullString Then
            dict(sk) = Array( _
                NzDouble(lo.DataBodyRange.Cells(i, idxBuy1).Value, 144#), _
                NzDouble(lo.DataBodyRange.Cells(i, idxBuy2).Value, 0#), _
                NzBool(lo.DataBodyRange.Cells(i, idxPrefer).Value, True) _
            )
        End If
    Next i

    Set BuildSettingsDict = dict
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
    Set BuildSettingsDict = CreateObject("Scripting.Dictionary")
End Function

Private Function BuildPartsTotalsDict() As Object
    Const PROC As String = "BuildPartsTotalsDict"
    On Error GoTo EH

    Dim dict As Object
    Dim lo As ListObject
    Dim i As Long
    Dim rowCount As Long
    Dim sk As String
    Dim totalReq As Double

    Dim idxStockKey As Long
    Dim idxTotalReq As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set lo = GetTable(WS_INPUT, TBL_PARTS)
    If lo Is Nothing Then
        Set BuildPartsTotalsDict = dict
        Exit Function
    End If

    If IsTableEffectivelyBlank(lo) Then
        Set BuildPartsTotalsDict = dict
        Exit Function
    End If

    idxStockKey = GetListColumnIndex(lo, "StockKey", PROC)
    idxTotalReq = GetListColumnIndex(lo, "TotalLenReq_in", PROC)

    rowCount = lo.ListRows.Count
    For i = 1 To rowCount
        sk = Trim$(CStr(lo.DataBodyRange.Cells(i, idxStockKey).Value))
        totalReq = NzDouble(lo.DataBodyRange.Cells(i, idxTotalReq).Value, 0#)

        If sk <> vbNullString And totalReq > TOL Then
            If dict.Exists(sk) Then
                dict(sk) = NzDouble(dict(sk), 0#) + totalReq
            Else
                dict.Add sk, totalReq
            End If
        End If
    Next i

    Set BuildPartsTotalsDict = dict
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
    Set BuildPartsTotalsDict = CreateObject("Scripting.Dictionary")
End Function

Private Function BuildBasicBarCountDict() As Object
    Const PROC As String = "BuildBasicBarCountDict"
    On Error GoTo EH

    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim sk As String
    Dim barNo As String
    Dim combo As String
    Dim seen As Object

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    Set ws = GetWorksheet(WS_BASIC)
    If ws Is Nothing Then
        Set BuildBasicBarCountDict = dict
        Exit Function
    End If

    lastRow = LastUsedRow(ws)
    If lastRow < 2 Then
        Set BuildBasicBarCountDict = dict
        Exit Function
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    For i = 2 To lastRow
        sk = Trim$(CStr(ws.Cells(i, 1).Value))
        If sk <> vbNullString Then
            barNo = Trim$(CStr(ws.Cells(i, 5).Value))
            combo = sk & Chr$(30) & barNo
            If Not seen.Exists(combo) Then
                seen(combo) = True
                If dict.Exists(sk) Then
                    dict(sk) = CLng(dict(sk)) + 1
                Else
                    dict(sk) = 1
                End If
            End If
        End If
    Next i

    Set BuildBasicBarCountDict = dict
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
    Set BuildBasicBarCountDict = CreateObject("Scripting.Dictionary")
End Function

Private Function GetSettingInfo(ByVal dict As Object, ByVal stockKey As String) As Variant
    If Not dict Is Nothing Then
        If dict.Exists(stockKey) Then
            GetSettingInfo = dict(stockKey)
            Exit Function
        End If
    End If
    GetSettingInfo = Array(144#, 0#, True, vbNullString)
End Function

Private Function ParseStockKey(ByVal stockKey As String) As Variant
    Dim p As Long
    Dim mat As String
    Dim diaVal As Double

    p = InStr(1, stockKey, "|", vbTextCompare)
    If p > 0 Then
        mat = Trim$(Left$(stockKey, p - 1))
        diaVal = NzDouble(Mid$(stockKey, p + 1), 0#)
    Else
        mat = Trim$(stockKey)
        diaVal = 0#
    End If

    ParseStockKey = Array(mat, diaVal)
End Function

'=========================================================
' TABLE / RANGE HELPERS
'=========================================================
Private Function GetWorksheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function GetTable(ByVal sheetName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If Not ws Is Nothing Then Set GetTable = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetListColumnIndex(ByVal lo As ListObject, ByVal headerName As String, ByVal PROC As String) As Long
    On Error GoTo EH

    If lo Is Nothing Then Err.Raise 11001, PROC, "ListObject is Nothing while requesting column: " & headerName
    GetListColumnIndex = lo.ListColumns(headerName).Index
    Exit Function

EH:
    Err.Raise 11009, PROC, "Required column not found in table '" & lo.Name & "': " & headerName
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    On Error GoTo SAFE_EXIT
    LastUsedRow = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                                LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                                MatchCase:=False).Row
    Exit Function
SAFE_EXIT:
    LastUsedRow = 1
End Function

Private Function IsTableEffectivelyBlank(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then
        IsTableEffectivelyBlank = True
        Exit Function
    End If

    If lo.DataBodyRange Is Nothing Then
        IsTableEffectivelyBlank = True
        Exit Function
    End If

    If lo.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange.Rows(1)) = 0 Then
            IsTableEffectivelyBlank = True
            Exit Function
        End If
    End If

    IsTableEffectivelyBlank = False
End Function

Private Function ResizeTableDataRows(ByVal lo As ListObject, ByVal desiredRows As Long) As Boolean
    Const PROC As String = "ResizeTableDataRows"
    On Error GoTo EH

    Dim ws As Worksheet
    Dim firstCell As Range
    Dim newRange As Range
    Dim colCount As Long

    ResizeTableDataRows = False
    If lo Is Nothing Then Exit Function

    If desiredRows < 1 Then desiredRows = 1

    Set ws = lo.Parent
    Set firstCell = lo.Range.Cells(1, 1)
    colCount = lo.ListColumns.Count

    Set newRange = ws.Range(firstCell, firstCell.Offset(desiredRows, colCount - 1))
    lo.Resize newRange

    ResizeTableDataRows = True
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
End Function

Private Sub EnsureTableRows(ByVal lo As ListObject, ByVal desiredRows As Long)
    Const PROC As String = "EnsureTableRows"
    On Error GoTo EH

    If lo Is Nothing Then Exit Sub
    If desiredRows < 1 Then desiredRows = 1

    If Not ResizeTableDataRows(lo, desiredRows) Then
        Err.Raise 9801, PROC, "ResizeTableDataRows failed for table " & lo.Name
    End If

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Sub ResetTableToSingleBlankRow(ByVal lo As ListObject)
    Const PROC As String = "ResetTableToSingleBlankRow"
    On Error GoTo EH

    If lo Is Nothing Then Exit Sub

    EnsureTableRows lo, 1
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.ClearContents

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Sub LoadArrayIntoTable(ByVal lo As ListObject, ByRef arr() As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    Const PROC As String = "LoadArrayIntoTable"
    On Error GoTo EH

    If lo Is Nothing Then Exit Sub

    If rowCount <= 0 Then
        ResetTableToSingleBlankRow lo
        Exit Sub
    End If

    EnsureTableRows lo, rowCount
    lo.DataBodyRange.Cells(1, 1).Resize(rowCount, colCount).Value = arr

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Function GetSortedKeyCollection(ByVal dict As Object) As Collection
    Const PROC As String = "GetSortedKeyCollection"
    On Error GoTo EH

    Dim result As Collection
    Dim rawKeys As Variant
    Dim keyText As String
    Dim inserted As Boolean
    Dim i As Long
    Dim k As Long

    Set result = New Collection

    If dict Is Nothing Then
        Set GetSortedKeyCollection = result
        Exit Function
    End If

    If dict.Count = 0 Then
        Set GetSortedKeyCollection = result
        Exit Function
    End If

    rawKeys = dict.Keys

    For k = LBound(rawKeys) To UBound(rawKeys)
        keyText = CStr(rawKeys(k))
        inserted = False

        If result.Count = 0 Then
            result.Add keyText
            inserted = True
        Else
            For i = 1 To result.Count
                If StrComp(keyText, CStr(result(i)), vbTextCompare) < 0 Then
                    result.Add keyText, Before:=i
                    inserted = True
                    Exit For
                End If
            Next i

            If Not inserted Then
                result.Add keyText
            End If
        End If
    Next k

    Set GetSortedKeyCollection = result
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
    Set GetSortedKeyCollection = New Collection
End Function

Private Function GetSortedKeysFromDict(ByVal dict As Object, ByRef keyCount As Long) As String()
    Const PROC As String = "GetSortedKeysFromDict"
    On Error GoTo EH

    Dim keys() As String
    Dim keyCol As Collection
    Dim i As Long

    Set keyCol = GetSortedKeyCollection(dict)
    keyCount = keyCol.Count

    If keyCount = 0 Then
        ReDim keys(1 To 1)
        GetSortedKeysFromDict = keys
        Exit Function
    End If

    ReDim keys(1 To keyCount)
    For i = 1 To keyCount
        keys(i) = CStr(keyCol(i))
    Next i

    GetSortedKeysFromDict = keys
    Exit Function

EH:
    LogError PROC, Err.Number, Err.Description
    keyCount = 0
    ReDim keys(1 To 1)
    GetSortedKeysFromDict = keys
End Function

Private Sub ClearPlainOutputSheet(ByVal ws As Worksheet)
    Const PROC As String = "ClearPlainOutputSheet"
    On Error GoTo EH

    If ws Is Nothing Then Exit Sub
    ws.Cells.ClearContents
    ws.Cells.ClearFormats

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Sub WriteBasicHeaders(ByVal ws As Worksheet)
    Dim hdr As Variant
    hdr = Array("StockKey", "BuyLen_in", "PieceID", "CutLen_in", "Bar#", "RemainingBefore_in", "RemainingAfter_in", "IsBarEnd", "LeftoverAtEnd_in", "RemnantStatus")
    ws.Range("A1").Resize(1, 10).Value = hdr
End Sub

Private Sub WriteArrayToRange(ByVal target As Range, ByRef arr() As Variant, ByVal rowCount As Long, ByVal colCount As Long)
    If rowCount <= 0 Then Exit Sub
    target.Resize(rowCount, colCount).Value = arr
End Sub

Private Function FinalNestingHasRows() As Boolean
    Dim lo As ListObject

    Set lo = GetTable(WS_FINAL, TBL_FINAL)
    If lo Is Nothing Then Exit Function
    If IsTableEffectivelyBlank(lo) Then Exit Function

    FinalNestingHasRows = True
End Function

'=========================================================
' INPUT PART CALCULATED COLUMNS
'=========================================================
Private Sub ApplyTblPartsCalculatedColumns(ByVal lo As ListObject)
    Const PROC As String = "ApplyTblPartsCalculatedColumns"
    On Error GoTo EH

    Dim idxReq As Long
    Dim idxTotal As Long
    Dim idxStock As Long

    If lo Is Nothing Then Exit Sub
    If lo.ListRows.Count = 0 Then EnsureTableRows lo, 1

    idxReq = GetListColumnIndex(lo, "RequiredCutLen_in", PROC)
    idxTotal = GetListColumnIndex(lo, "TotalLenReq_in", PROC)
    idxStock = GetListColumnIndex(lo, "StockKey", PROC)

    lo.DataBodyRange.Columns(idxReq).Formula = _
        "=IFERROR(N([@[FinishLen_in]])+N([@[ExtraStock_in]])+IF(OR([@[Kerf_in]]="""",ISBLANK([@[Kerf_in]])),DefaultKerf_in,N([@[Kerf_in]]))+IF(OR([@[FaceAllow_in]]="""",ISBLANK([@[FaceAllow_in]])),DefaultFaceAllow_in,N([@[FaceAllow_in]])),0)"

    lo.DataBodyRange.Columns(idxTotal).Formula = _
        "=IFERROR(N([@[Qty]])*N([@[RequiredCutLen_in]]),0)"

    lo.DataBodyRange.Columns(idxStock).Formula = _
        "=TRIM([@[Material]])&""|""&TEXT([@[Diameter_in]],""0.000"")"

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

'=========================================================
' VALIDATION HELPERS
'=========================================================
Private Sub ApplyDecimalValidation(ByVal rng As Range, ByVal minVal As Double, ByVal maxVal As Double)
    Const PROC As String = "ApplyDecimalValidation"
    On Error GoTo EH

    With rng.Validation
        On Error Resume Next
        .Delete
        On Error GoTo EH
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=CStr(minVal), Formula2:=CStr(maxVal)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Diameter_in"
        .ErrorTitle = "Invalid Diameter"
        .InputMessage = "Enter a numeric diameter > 0."
        .ErrorMessage = "Diameter_in must be numeric and > 0."
        .ShowInput = True
        .ShowError = True
    End With
    Exit Sub

EH:
    LogInfo PROC, "Validation warning: " & Err.Description
End Sub

Private Sub ApplyWholeValidation(ByVal rng As Range, ByVal minVal As Long, ByVal maxVal As Long)
    Const PROC As String = "ApplyWholeValidation"
    On Error GoTo EH

    With rng.Validation
        On Error Resume Next
        .Delete
        On Error GoTo EH
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=CStr(minVal), Formula2:=CStr(maxVal)
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Qty"
        .ErrorTitle = "Invalid Qty"
        .InputMessage = "Enter a whole number >= 1."
        .ErrorMessage = "Qty must be a whole number >= 1."
        .ShowInput = True
        .ShowError = True
    End With
    Exit Sub

EH:
    LogInfo PROC, "Validation warning: " & Err.Description
End Sub

Private Sub ApplyListValidationAllowFreeText(ByVal rng As Range, ByVal Formula1 As String)
    Const PROC As String = "ApplyListValidationAllowFreeText"
    On Error GoTo EH

    With rng.Validation
        On Error Resume Next
        .Delete
        On Error GoTo EH
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=Formula1
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Material"
        .ErrorTitle = "Material"
        .InputMessage = "Choose from list or type a new material."
        .ErrorMessage = "Choose from list or type a new material."
        .ShowInput = True
        .ShowError = False
    End With
    Exit Sub

EH:
    LogInfo PROC, "Validation warning: " & Err.Description
End Sub

'=========================================================
' PAGE BREAKS / PRINT
'=========================================================
Private Sub ResetSheetPageBreaks(ByVal ws As Worksheet)
    Const PROC As String = "ResetSheetPageBreaks"
    On Error GoTo EH

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.ResetAllPageBreaks
    On Error GoTo EH

    Do While ws.HPageBreaks.Count > 0
        ws.HPageBreaks(1).Delete
    Loop

    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

'=========================================================
' NAMED RANGE HELPERS
'=========================================================
Private Sub SetNamedRangeValue(ByVal rangeName As String, ByVal newValue As Variant)
    Const PROC As String = "SetNamedRangeValue"
    On Error GoTo EH

    Dim rng As Range
    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then Exit Sub
    rng.Value = newValue
    Exit Sub

EH:
    LogError PROC, Err.Number, Err.Description
    Err.Raise Err.Number, PROC, Err.Description
End Sub

Private Function GetNamedRange(ByVal rangeName As String) As Range
    On Error Resume Next
    Set GetNamedRange = ThisWorkbook.Names(rangeName).RefersToRange
    On Error GoTo 0
End Function

Private Function GetNamedRangeNumber(ByVal rangeName As String, ByVal defaultValue As Double) As Double
    Dim rng As Range
    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then
        GetNamedRangeNumber = defaultValue
    ElseIf Trim$(CStr(rng.Value)) = vbNullString Then
        GetNamedRangeNumber = defaultValue
    Else
        GetNamedRangeNumber = NzDouble(rng.Value, defaultValue)
    End If
End Function

Private Function GetNamedRangeDateValue(ByVal rangeName As String) As Date
    Dim rng As Range
    Set rng = GetNamedRange(rangeName)
    If rng Is Nothing Then Exit Function
    If Trim$(CStr(rng.Value)) = vbNullString Then Exit Function
    If IsDate(rng.Value) Then
        GetNamedRangeDateValue = CDate(rng.Value)
    Else
        GetNamedRangeDateValue = 0
    End If
End Function

'=========================================================
' ARRAY / TYPE CONVERSION HELPERS
'=========================================================
Private Function GetColumnFrom2DArray(ByRef src() As Variant, ByVal colIdx As Long) As Variant
    Dim i As Long
    Dim rowCount As Long
    Dim outArr() As Variant

    rowCount = UBound(src, 1)
    ReDim outArr(1 To rowCount, 1 To 1)
    For i = 1 To rowCount
        outArr(i, 1) = src(i, colIdx)
    Next i
    GetColumnFrom2DArray = outArr
End Function

Private Function NzDouble(ByVal v As Variant, Optional ByVal defaultValue As Double = 0#) As Double
    If IsError(v) Then
        NzDouble = defaultValue
    ElseIf IsNumeric(v) Then
        NzDouble = CDbl(v)
    ElseIf Trim$(CStr(v)) = vbNullString Then
        NzDouble = defaultValue
    Else
        NzDouble = defaultValue
    End If
End Function

Private Function NzBool(ByVal v As Variant, Optional ByVal defaultValue As Boolean = False) As Boolean
    If IsError(v) Then
        NzBool = defaultValue
    ElseIf VarType(v) = vbBoolean Then
        NzBool = CBool(v)
    ElseIf Trim$(UCase$(CStr(v))) = "TRUE" Then
        NzBool = True
    ElseIf Trim$(UCase$(CStr(v))) = "FALSE" Then
        NzBool = False
    ElseIf IsNumeric(v) Then
        NzBool = (CDbl(v) <> 0)
    Else
        NzBool = defaultValue
    End If
End Function

Private Function RoundToPlaces(ByVal v As Double, ByVal places As Long) As Double
    RoundToPlaces = WorksheetFunction.Round(v, places)
End Function

Private Function FormatNumberSafe(ByVal v As Double) As String
    FormatNumberSafe = Format$(v, "0.000")
End Function