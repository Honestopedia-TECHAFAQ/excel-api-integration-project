Function GetDigiKeyPartInfo(apiKey As String, partNumber As String) As String
    Dim http As Object
    Dim url As String
    Dim response As String
    
    ' Create the URL for the API request
    url = "https://api.digikey.com/services/product-information/v4/part-details/" & partNumber
    
    ' Create the HTTP object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    ' Open the request
    http.Open "GET", url, False
    http.setRequestHeader "X-DIGIKEY-Client-Id", apiKey
    http.setRequestHeader "X-DIGIKEY-Locale-Site", "US"
    http.setRequestHeader "Content-Type", "application/json"
    
    ' Send the request
    http.Send
    
    ' Get the response
    response = http.responseText
    
    ' Return the response as a string
    GetDigiKeyPartInfo = response
End Function

Sub ButtonClickHandler()
    Dim apiKey As String
    Dim partNumber As String
    Dim response As String
    
    ' Get the API key and part number from the worksheet
    apiKey = Sheets("Sheet1").Range("B1").Value
    partNumber = Sheets("Sheet1").Range("A1").Value
    
    ' Call the function to get the product info
    response = GetDigiKeyPartInfo(apiKey, partNumber)
    
    ' Output the response in cell A2
    Sheets("Sheet1").Range("A2").Value = response
End Sub
