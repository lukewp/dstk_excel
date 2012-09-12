Attribute VB_Name = "geoCode"
' Excel VBA functions retrieve a TIGER-based long/lat geocode for an address
' from the datasciencetoolkit.org address2coordinates API
' Written 5/25/2011 by Luke Peterson luke.peterson@gmail.com
'
' Usage: Feed a USA street address to getGeocode() to get an API response,
' which you can pass through getLatitude() and getLongitude() in order to put
' the Long/Lat into separate columns. Best-practice is to pull the API response
' into a single field with getGeocode(), then parse that field with
' getLatitude() and getLongitude(), rather than using
' getLatitude(getGeocode([Address])) and getLongitude(getGeocode([Address])),
' since the latter will double the number of requests to the dstk server.
'
' Dependencies: Enable "Microsoft XML, v6.0" in Excel's Tools...References dialog.
'
' URLEncode function borrowed from here: http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba


Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function


Public Function getGeocode(address As String) As String
    
    Dim geocodeService As New XMLHTTP60
    
    geocodeService.Open "GET", "http://www.datasciencetoolkit.org/street2coordinates/" + URLEncode(address, True)
    geocodeService.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    geocodeService.send
    
    If geocodeService.Status = "200" Then
               
        getGeocode = geocodeService.responseText
    Else
        getGeocode = "Error"
           
    End If

End Function

Public Function getLatitude(geocode As String) As String
    If InStr(geocode, "latitude") = 0 Then
        getLatitude = "None"
    Else
        getLatitude = Mid(geocode, InStr(geocode, "latitude") + 10, InStr(Mid(geocode, InStr(geocode, "latitude") + 10), ",") - 1)
    End If
End Function

Public Function getLongitude(geocode As String) As String
    If InStr(geocode, "longitude") = 0 Then
        getLongitude = "None"
    Else
        getLongitude = Mid(geocode, InStr(geocode, "longitude") + 11, InStr(Mid(geocode, InStr(geocode, "longitude") + 11), ",") - 1)
    End If
End Function
