' Copyright © 2024 Mohamed Safaa Ismail
' Licensed under the MIT License. See the LICENSE file for details.
Sub FillCellsWithRandomNumbers()
    Dim rng As Range
    Dim cell As Range
    Dim total As Long
    Dim randNum As Long
    Dim targetSum As Long
    Dim numCells As Long
    Dim i As Long
    Dim adjustmentRequired As Boolean
    Dim currentCell As Range
    Dim possibleValues() As Variant
    Dim value As Variant
    Dim maxValue As Long
    Dim maxPossibleSum As Long
    Dim userInput As String
    Dim validRange As Boolean
    Dim valuesString As String
    Dim valuesArray() As 

    ' طلب المجموع المطلوب من المستخدم
    targetSum = Val(InputBox("أدخل المجموع المطلوب:", "مجموع مطلوب"))
    If targetSum <= 0 Then
        MsgBox "المجموع المطلوب غير صحيح. العملية ستتوقف."
        Exit Sub
    End If

    ' طلب القيم العشوائية من المستخدم
    valuesString = InputBox("أدخل القيم العشوائية التي تريد استخدامها (فصل القيم بفواصل):", "القيم العشوائية", "5,10,15,25")
    If valuesString = "" Then
        MsgBox "لم يتم إدخال القيم العشوائية. العملية ستتوقف."
        Exit Sub
    End If

    ' تحويل القيم العشوائية إلى مصفوفة
    valuesArray = Split(valuesString, ",")
    ReDim possibleValues(LBound(valuesArray) To UBound(valuesArray))
    For i = LBound(valuesArray) To UBound(valuesArray)
        possibleValues(i) = Val(valuesArray(i))
    Next i

    validRange = False
    Do While Not validRange
        ' اختيار النطاق من ورقة العمل
        On Error Resume Next
        Set rng = Application.InputBox("حدد نطاق الخلايا (مثل A1:E7):", Type:=8)
        On Error GoTo 0

        ' التحقق من صحة النطاق
        If rng Is Nothing Then
            MsgBox "لم يتم تحديد نطاق الخلايا. العملية ستتوقف."
            Exit Sub
        End If

        numCells = rng.Cells.Count
        rng.ClearContents

        ' حساب الحد الأقصى الممكن للمجموع بالنطاق المحدد
        maxPossibleSum = numCells * Application.WorksheetFunction.Max(possibleValues)

        ' التحقق إذا كان النطاق يكفي لتحقيق المجموع المطلوب
        If maxPossibleSum < targetSum Then
            userInput = MsgBox("النطاق المحدد صغير جدًا لتحقيق المجموع المطلوب (" & targetSum & "). هل تريد تحديد نطاق أكبر؟", vbYesNo + vbExclamation, "النطاق غير كافٍ")
            If userInput = vbNo Then
                MsgBox "لم يتم تحديد نطاق مناسب. العملية ستتوقف."
                Exit Sub
            End If
            ' إعادة محاولة اختيار نطاق جديد
        Else
            validRange = True
        End If
    Loop

    ' ملء الخلايا بأرقام عشوائية
    total = 0
    For i = 1 To numCells
        ' اختيار رقم عشوائي من القيم الممكنة
        randNum = possibleValues(Application.WorksheetFunction.RandBetween(0, UBound(possibleValues)))

        ' التحقق إذا كان إضافة الرقم العشوائي سيبقي المجموع ضمن المجموع المطلوب
        If total + randNum <= targetSum Then
            rng.Cells(i).Value = randNum
            total = total + randNum
        Else
            Exit For
        End If
    Next i

    ' إذا لم يصل المجموع إلى المجموع المطلوب، ضبط القيم
    If total < targetSum Then
        adjustmentRequired = True
        Dim diff As Long
        diff = targetSum - total

        ' ضبط القيم في الخلايا لضبط المجموع النهائي
        For Each currentCell In rng
            If diff <= 0 Then Exit For
            If currentCell.Value < Application.WorksheetFunction.Max(possibleValues) Then
                maxValue = Application.WorksheetFunction.Min(Application.WorksheetFunction.Max(possibleValues), diff + currentCell.Value)
                If maxValue > currentCell.Value Then
                    total = total - currentCell.Value + maxValue
                    currentCell.Value = maxValue
                    diff = targetSum - total
                End If
            End If
        Next currentCell

        ' إذا كانت القيم أكبر من المجموع المطلوب، تقليل القيم لتناسب المجموع
        If total > targetSum Then
            For Each currentCell In rng
                If total <= targetSum Then Exit For
                If currentCell.Value > Application.WorksheetFunction.Min(possibleValues) Then
                    maxValue = Application.WorksheetFunction.Max(Application.WorksheetFunction.Min(possibleValues), currentCell.Value - (total - targetSum))
                    If maxValue < currentCell.Value Then
                        total = total - currentCell.Value + maxValue
                        currentCell.Value = maxValue
                    End If
                End If
            Next currentCell
        End If
    End If

    MsgBox "تم الانتهاء من التعبئة! المجموع النهائي هو: " & total
End Sub
