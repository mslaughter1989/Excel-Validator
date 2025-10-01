Attribute VB_Name = "Specifics_HealthKarmaTBL"
Sub Specifics_HealthKarmaTBL()
    Dim ws As Worksheet
    Dim headerDict As Object
    Dim desiredOrder As Variant
    Dim i As Long, colIndex As Variant
    Dim currentHeader As String
    Dim tempSheet As Worksheet

    Set ws = ActiveSheet
    Set headerDict = CreateObject("Scripting.Dictionary")

    ' Build dictionary of current headers, normalize Internal Code to MetaTag1
    For i = 1 To ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
        currentHeader = Trim(ws.Cells(1, i).Value)
        If LCase(currentHeader) = "internal code" Then
            currentHeader = "MetaTag1"
        End If
        If Not headerDict.Exists(LCase(currentHeader)) Then
            headerDict(LCase(currentHeader)) = i
        End If
    Next i

    ' Define desired header order
    desiredOrder = Array( _
        "LastName", "FirstName", "Gender", "DateOfBirth", "AddressLine1", "AddressLine2", _
        "City", "State", "ZipCode", "CountryCode", "MobilePhone", "EmailAddress", _
        "EffectiveStart", "EffectiveEnd", "MemberType", "ClientMemberID", _
        "SecondaryClientMemberID", "ClientPrimaryMemberID", "ServiceOffering", _
        "GroupID", "GroupName", "MetaTag1", "MetaTag2", "MetaTag3", "MetaTag4", "MetaTag5" _
    )

    ' Create temporary worksheet to rearrange columns
    Set tempSheet = Worksheets.Add
    tempSheet.Name = "TempSheet12345"

    ' Rearrange by copying columns in desired order
    For i = 0 To UBound(desiredOrder)
        currentHeader = LCase(desiredOrder(i))
        If headerDict.Exists(currentHeader) Then
            ws.Columns(headerDict(currentHeader)).Copy Destination:=tempSheet.Columns(i + 1)
        Else
            tempSheet.Cells(1, i + 1).Value = desiredOrder(i) ' Blank column with correct header
        End If
    Next i

    ' Overwrite original worksheet
    ws.Cells.Clear
    tempSheet.UsedRange.Copy Destination:=ws.Range("A1")
    Application.DisplayAlerts = False
    tempSheet.Delete
    Application.DisplayAlerts = True

End Sub


