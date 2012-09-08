Option Explicit
'Script to download BHAVCOPY data from NSE and import in AmiBroker
'Usage cscript SimpleLoops.vbs <proxyserver> <username> <password>
'	example : cscript SimpleLoops.vbs mandar_limaye lamepassword

'Constants

Const cstrFilename  = "cm{DD}{MMM}{YYYY}bhav.csv"

Const cstrBHAVCOPYURL  = "http://www.nseindia.com/content/historical/EQUITIES/{YYYY}/{MMM}/"

Const cstrUNZIPURL  = "http://dipoletech.com/temp/unzip.exe"


Const ForReading = 1, ForWriting = 2, ForAppending = 8 'Constants for FSO


'Dim strProxyServer, strUserName, strPassword

Dim gdtmDateFrom, gdtmDateTo, gstrProxyServer, gstrProxyUsername, gstrProxyPassword

'Call CheckUnzip()
'WScript.Quit()

On Error Resume Next
Dim AmiBroker
Set AmiBroker = CreateObject( "Broker.Application" )
If Err.Number <> 0 Then
  WScript.Echo "ERROR!!!: Could not detect Amibroker installation. Exiting now!"
  WScript.Quit(1)
End If

On Error Goto 0

Call Main()




'****************************************************
'Main Sub
'****************************************************
Sub Main()
	Dim strResult, strBHAVCOPYurl, dtmDate, intNumDays
	
  Call InitialiseGlobals()

	'dtmDate = DateAdd("d" , -4, Date)
  
  intNumDays = DateDiff("d", gdtmDateFrom, gdtmDateTo)
  
  If intNumDays < 0 Then
    Call Usage("The ""To"" date cannot be before ""From"" date")
  End If
  
  Call CreateTempFolder()
  WriteFormatFile(GetTempFolderPath() & "\nsedaily.format")
  Call CheckUnzip()
  
  Dim i 
  
  'Do Indices
  Call GetIndices()
  
  'Do stocks
  For i = 0 to intNumDays
    dtmDate = DateAdd("d", i, gdtmDateFrom)
    strBHAVCOPYurl = GetBHAVCOPYurl(dtmDate)
    If DateValue(dtmDate) > DateValue(#01/12/2009#) Then 
        strBHAVCOPYurl = strBHAVCOPYurl & ".zip"
        If SaveWebBinary(strBHAVCOPYurl, GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate) & ".zip") Then
            InflateFile(GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate) & ".zip")
            Call RemoveDuplicatesFromFile(GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate), dtmDate)
        End If
    ElseIf SaveWebBinary(strBHAVCOPYurl, GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate)) Then
        Call RemoveDuplicatesFromFile(GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate), dtmDate)
    End If
  Next
	
	Call AmiBroker.RefreshAll()
	Call AmiBroker.SaveDatabase()
	Call AmiBroker.Quit()

End Sub


'****************************************************
'Sub prints the correct usage for the script and exits
'****************************************************
Sub InflateFile(strFile)
    Dim objShell, strCommand, fso
    Set objShell = WScript.CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    
    strCommand = GetTempFolderPath() & "\" & "unzip -o """ & strFile & """ -d """ & GetTempFolderPath() & """"
    objShell.Run strCommand, 0, true
    'fso.DeleteFile (strFile)
    'WScript.Echo strCommand

End Sub


'****************************************************
'Sub prints the correct usage for the script and exits
'****************************************************
Sub Usage(strMessage)
  If strMessage <> "" Then
    WScript.Echo "ERROR !!! : " & strMessage
  End If
  const L_Echo1_Text = "NSE Data Fetcher for AmiBroker"
	const L_Echo2_Text = "    - this tools fetches BHAVCOPY from NSE website in CSV format and uploads it to Amibroker"
	const L_Echo3_Text = ""
	const L_Echo4_Text = "Usage:	cscript.exe NSEDataFetch.vbs [-dF DateFrom] [-dT DateTo] [-dDelta num] [/pS ProxyServerName] [-pU ProxyUserName] [-pP ProxyPassword]"
	const L_Echo5_Text = ""
	const L_Echo6_Text = ""
	const L_Echo7_Text = "  -dF	- (Optional) Date from which the tool will fetch data, default is today's date"
	const L_Echo8_Text = "  -dF	- (Optional) Date to which the tool will fetch data, default is today's date"
	const L_Echo9_Text = "  -pS	- (Optional) Specify a proxy server for connecting to the internet"
	const L_Echo10_Text = "  -pU	- (Optional) Specify a proxy Username for connecting to the internet"
	const L_Echo11_Text = "  -pP	- (Optional) Specify a proxy Password for connecting to the internet"
  const L_Echo12_Text = "  -Refresh	- (Optional) Refresh the equity tickers, names and industries"
  const L_Echo13_Text = "  -dDelta	- (Optional) Fetch all data for the specified number of days prior to today"

	WScript.Echo L_Echo1_Text
	WScript.Echo L_Echo2_Text
	WScript.Echo L_Echo3_Text
	WScript.Echo L_Echo4_Text
	WScript.Echo L_Echo5_Text
	WScript.Echo L_Echo6_Text
	WScript.Echo L_Echo7_Text
	WScript.Echo L_Echo8_Text
	WScript.Echo L_Echo9_Text
	WScript.Echo L_Echo10_Text
	WScript.Echo L_Echo11_Text
	WScript.Echo L_Echo12_Text
	WScript.Echo L_Echo13_Text
	WScript.Echo 
  WScript.Quit(1)
End Sub


Const L_Question1_Text = "-?"
Const L_Question2_Text = "-h"

Const L_DateFrom = "-dF"
Const L_DateTo = "-dT"
Const L_DateDelta = "-dDelta"

Const L_ProxyServer = "-pS"
Const L_ProxyUsername = "-pU"
Const L_ProxyPassword = "-pP"

Const L_Refresh = "-Refresh"

'****************************************************
'Sub initialises gloabs from arguments
'****************************************************
Sub InitialiseGlobals()
	Dim boolRefresh
	boolRefresh = False
  'Exit if no args passed
  If WScript.Arguments.Count = 0 Then
    Call Usage("")
  End If
    
  'Show help if only one arg is passed
  If  ( WScript.Arguments.Count = 1 And ( WScript.Arguments(0) = L_Question1_Text or WScript.Arguments(0) = L_Question2_Text ) ) then
		Call Usage("")
  End If
   
  'Initialise default values
  gdtmDateFrom = Date: gdtmDateTo = Date
  gstrProxyServer = "":gstrProxyUsername = "":gstrProxyPassword = ""
  Dim i 
  i = 0
  Do While i < WScript.Arguments.Count
		If WScript.Arguments(i) = L_Question1_Text Or WScript.Arguments(i) = L_Question2_Text Then
			Call Usage("")
		Elseif WScript.Arguments(i) = L_DateFrom Then
			i = i + 1
			If i >= WScript.Arguments.Count Then
				Call Usage("")
			End If
			gdtmDateFrom = DateValue(WScript.Arguments(i))
    Elseif WScript.Arguments(i) = L_DateTo Then
			i = i + 1
			If i >= WScript.Arguments.Count Then
				Call Usage("")
			End If
			gdtmDateTo = DateValue(WScript.Arguments(i))
    Elseif WScript.Arguments(i) = L_ProxyServer Then
			i = i + 1
			If i >= WScript.Arguments.Count Then
				Call Usage("")
			End If
			gstrProxyServer = WScript.Arguments(i)
    Elseif WScript.Arguments(i) = L_ProxyUsername Then
			i = i + 1
			If i >= WScript.Arguments.Count Then
				Call Usage("")
			End If
			gstrProxyUsername = WScript.Arguments(i)
    Elseif WScript.Arguments(i) = L_ProxyPassword Then
			i = i + 1
			If i >= WScript.Arguments.Count Then
				Call Usage("")
			End If
			gstrProxyPassword = WScript.Arguments(i)
    Elseif WScript.Arguments(i) = L_Refresh Then
			boolRefresh = True
		ElseIf WScript.Arguments(i) = L_DateDelta Then
			i = i + 1
			If i >= WScript.Arguments.Count Then
				Call Usage("")
			End If
			gdtmDateFrom = DateAdd("d" , CInt("-" & WScript.Arguments(i)), Date)
			gdtmDateTo = Date
    End If
    i = i + 1
  Loop

  Call PrintOptions()
  If boolRefresh Then
  	Call RefreshTickers()
  End If
End Sub


'****************************************************
'Sub fetch Indices data
'****************************************************
Sub PrintOptions()
'gdtmDateFrom, gdtmDateTo, gstrProxyServer, gstrProxyUsername, gstrProxyPassword
  WScript.Echo "***************************************************************"
  WScript.Echo "NSEDataFetcher starting with settings:"
  WScript.Echo "***************************************************************"
  WScript.Echo "Date From            :" & gdtmDateFrom
  WScript.Echo "Date To              :" & gdtmDateTo
  WScript.Echo "Amibroker Database   :" & Amibroker.DatabasePath
  WScript.Echo "***************************************************************"
  WScript.Echo ""
End Sub


'****************************************************
'Sub fetch Indices data
'****************************************************
Sub GetIndices()
  Dim strIndicesHTML, strIndicesURL, regEx, Matches, i, objStock, strIdxTicker, strIdxName
  
  strIndicesURL = "http://www.nseindia.com/content/indices/ind_histvalues.htm"
  
  strIndicesHTML = ReadFileFromWeb(strIndicesURL)
  
  Set regEx = New RegExp
	regEx.Pattern = "<option value=\""(.*)\""(.*)>(.*)<"
  regEx.Global = True
  Set Matches = regEx.Execute( strIndicesHTML )


  For i = 0 to Matches.Count - 1
    strIdxTicker = Matches(i).SubMatches(0)
    strIdxName = Matches(i).SubMatches(2)
   
    'strIdxTicker = "CNX INFRASTRUCTURE"
    'strIdxName = "CNX Infrastructure"
    Set objStock = Amibroker.Stocks.Add(strIdxTicker)
    
    objStock.FullName = strIdxName
    
    objStock.GroupID = 254
    'WScript.Echo "GroupId = " & objStock.GroupID
    Call AmiBroker.RefreshAll()
  	On Error Resume Next
    Call DoIndex(objStock)
    If Err.Number <> 0 Then
    	WScript.Echo vbTab & "Function:DoIndices; ERROR: " & Err.Number & " " & Err.Description
    End If
    Err.Clear
    On Error GoTo 0
    'WScript.Quit(0)
  Next
  

  Call Amibroker.RefreshAll()

  
  'WScript.Echo strIndicesHTML

End Sub


'****************************************************
'Sub get data for given index and update
'****************************************************
Sub DoIndex(objStock)
  Dim strCsvURL, strCsvData, astrRows, astrFields, dtmDate, i, j, objQuote, strParams
  
  WScript.Echo "Processing Index: " & objStock.Ticker
  strCsvURL = GetCsvURL(objStock.Ticker)
  
  'WScript.Echo "Debug: DoIndex; Index=" & objStock.Ticker & " Url1=" & strCsvURL & vbCrLf
  
  If Trim(strCsvURL) = "" Then
  	Exit Sub
  End If
  
  strCsvData = ReadFileFromWeb(strCsvURL)
  
  'WScript.Echo "Debug: DoIndex: strCsvData=" & strCsvData
  
  astrRows = Split(strCsvData, vbLf)
  
  For i = 1 to UBound(astrRows)
  'WScript.Echo ""
  'WScript.Echo ""
    'WScript.Echo "Debug: DoIndex: Processing Row=" & astrRows(i)
    If Trim(astrRows(i)) <> "" Then
      astrFields = Split(astrRows(i),",")
    
    	'WScript.Echo "Function: DoIndex; astrFields(0)=" & astrFields(1)
      Set objQuote = objStock.Quotations.Add(CDate(Replace(astrFields(0),"""","")))
      
      If InStr(astrFields(2),"-") < 1 Then objQuote.High = CSng(Trim(Replace(astrFields(2),"""","")))
      If InStr(astrFields(3),"-") < 1 Then objQuote.Low = CSng(Trim(Replace(astrFields(3),"""","")))
      If InStr(astrFields(1),"-") < 1 Then objQuote.Open = CSng(Trim(Replace(astrFields(1),"""","")))
      If InStr(astrFields(4),"-") < 1 Then objQuote.Close = CSng(Trim(Replace(astrFields(4),"""","")))
    
      
      If UBound(astrFields) > 4 Then
        If InStr(astrFields(5),"-") < 1 Then objQuote.Volume = CSng(Trim(Replace(astrFields(5),"""","")))
      End If
      'WScript.Echo "Ticker=" & objStock.Ticker & " Close=" & objQuote.Close
    End If
    
  Next
  
Call Amibroker.RefreshAll()
End Sub


'****************************************************
'Function returns URL of the form to gen index data
'****************************************************
Function GetCsvURL(strIdxTicker)
  Dim strFormURL, strCsvURL, strYear, strMonth, strDay, strFormHTML
  strFormURL = "http://www.nseindia.com/marketinfo/indices/histdata/historicalindices.jsp?fromDate={dd}-{mm}-{yyyy}&indexType={TICKER}&toDate={ddT}-{mmT}-{yyyyT}"
  
  strFormURL = Replace(strFormURL, "{TICKER}", Replace(strIdxTicker, "&", "%26"))
  
  strYear = DatePart("yyyy", gdtmDateFrom)
  strFormURL = Replace(strFormURL, "{yyyy}", strYear)
	strMonth = Pad(DatePart("m", gdtmDateFrom), 2, "0")
  strFormURL = Replace(strFormURL, "{mm}", strMonth)
	strDay = Pad(DatePart("d", gdtmDateFrom), 2, "0")
  strFormURL = Replace(strFormURL, "{dd}", strDay)
  
  strYear = DatePart("yyyy", gdtmDateTo)
  strFormURL = Replace(strFormURL, "{yyyyT}", strYear)
	strMonth = Pad(DatePart("m", gdtmDateTo), 2, "0")
  strFormURL = Replace(strFormURL, "{mmT}", strMonth)
	strDay = Pad(DatePart("d", gdtmDateTo), 2, "0")
  strFormURL = Replace(strFormURL, "{ddT}", strDay)
  
  'WScript.Echo "Debug: GetCsvParams: Params=" & strFormURL
  strFormHTML = ReadFileFromWeb(strFormUrl)
  
  Dim regEx, Matches
  
  Set regEx = New RegExp
  regEx.Pattern = "href=\""(.*)\.csv\"" "
  regEx.Global = True
  
  Set Matches = regEx.Execute(strFormHTML)
  
  If Matches.Count > 0 Then
    If Matches(0).SubMatches.Count > 0 Then
      strCsvURL = "http://www.nseindia.com" & Matches(0).SubMatches(0) & ".csv"
    End If
  End If
  
  'WScript.Echo strFormURL
  'WScript.Echo strCsvURL
  'WScript.Echo ""
  
  GetCsvUrl = strCsvURL
End Function


'****************************************************
'Sub refreshes tickers
'****************************************************
Sub RefreshTickers()
  Dim strURL, strList, astrEquities, i, astrFields, strTicker, strStockName
  
  
  strURL = "http://nseindia.com/content/equities/EQUITY_L.csv"
  
  strList = ReadFileFromWeb(strURL)
  
  astrEquities = Split(strList, vbLf)
  
  Dim objStock
  
  For i = 1 to UBound(astrEquities)
    If Len(Trim(astrEquities(i))) > 0 Then
      astrFields = Split(astrEquities(i), ",")
      strTicker = astrFields(0)
      strStockName = astrFields(1)
      WScript.Echo "Ticker=" & strTicker & "  StockName=" & strStockName
      Set objStock = AmiBroker.Stocks.Add(strTicker)
      objStock.FullName = strStockName
    End If
  Next
  
  Call AmiBroker.RefreshAll()
  
  Call AmiBroker.SaveDatabase()
  
  'Call AmiBroker.Quit()
  
  WScript.Quit(0)
  
End Sub

'****************************************************
'Sub removes duplicates from file and saves
'****************************************************
Sub RemoveDuplicatesFromFile(strFile, dtmDate)
	Dim dicEquity, strFileContents, astrFileRows, strHeaderRow, astrDicRows, strOutputFile
	Dim i, strKey
	
	'Read file into dictionary
	Set dicEquity = CreateObject("Scripting.Dictionary")
	
	strFileContents = ReadFromFile(strFile)
	
	astrFileRows = Split(strFileContents, vbLf)
	
	strHeaderRow = astrFileRows(0)
	
	For i=1 to UBound(astrFileRows)
		If Len(Trim(astrFileRows(i))) > 0 Then
			strKey = GetKeyForRow(astrFileRows(i))
			'If strKey = "ABAN" Then WScript.Echo astrFileRows(i)
			If dicEquity.Exists(strKey) Then
				If (InStr(astrFileRows(i),"EQ") > 0) Then
					dicEquity.Item(strKey) = astrFileRows(i)
					'WScript.Echo "Has EQ hence importing:" & dicEquity.Item(strKey)
				End If 
			Else
				dicEquity.Add strKey, astrFileRows(i)
			End If
		End If
	Next

	astrDicRows = dicEquity.Items
	
	strOutputFile = strHeaderRow

	For i=0 to UBound(astrDicRows)
		'Call AddRowToAmibroker(astrDicRows(i), dtmDate)
		strOutputFile = strOutputFile & vbLf & astrDicRows(i) 
	Next
	
	strOutputFile = strOutputFile & vbLf
	
	Call WriteToFile(GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate), strOutputFile)
	
	Call AddRowToAmibroker("", dtmDate)
	
End Sub


'****************************************************
'Sub checks if unzip.exe exists.. if not downlaod it
'****************************************************
Sub CheckUnzip()
  Dim oFso, strPath
  Set oFso= createobject("Scripting.FileSystemObject")
  strPath = GetTempFolderPath() & "\unzip.exe"
  If Not oFso.FileExists(strPath) Then
    Call SaveWebBinary(cstrUNZIPURL, strPath)
  End If
End Sub


'****************************************************
'Sub returns path for temp folder
'****************************************************
Function GetTempFolderPath()
  Dim oFso, strPath
  Set oFso= createobject("Scripting.FileSystemObject")
  'strPath = oFso.getAbsolutePathName("")
  'GetTempFolderPath = CStr(strPath & "\temp")
  GetTempFolderPath = oFso.GetSpecialFolder(2)
End Function

'****************************************************
'Sub creates temp folder
'****************************************************
Sub CreateTempFolder()
  Dim oFso, strPath, newfolder
  Set oFso= createobject("Scripting.FileSystemObject")
  strPath = oFso.getAbsolutePathName("")
  If  Not oFso.FolderExists(strPath & "\temp") Then
   newfolder = oFso.CreateFolder (strPath & "\temp")
  End If
  WriteFormatFile(strPath & "\temp\nsedaily.format")
End Sub

'****************************************************
'Sub writes out the format file to temp folder
'****************************************************
Sub WriteFormatFile(strFile)
  Dim strFormat
  strFormat = "$FORMAT Ticker, Skip, Open, High, Low, Close, Skip, Skip, Volume, Skip, Date_DMY, Skip" & vbCrLf
  strFormat = strFormat & "$SKIPLINES 1" & vbCrLf
  strFormat = strFormat & "$SEPARATOR ," & vbCrLf
  strFormat = strFormat & "$CONT 1" & vbCrLf
  strFormat = strFormat & "$GROUP 255" & vbCrLf
  strFormat = strFormat & "$AUTOADD 1" & vbCrLf
  strFormat = strFormat & "$DEBUG 1" & vbCrLf
  strFormat = strFormat & "" & vbCrLf
  Call WriteToFile(strFile, strFormat)
End Sub


'****************************************************
'Sub adds a row to AmiBroker
'****************************************************
Sub AddRowToAmibroker(strRow, dtmDate)
	Dim astrFields, strStock, objStock, objQuote, lngResult

	'Set AmiBroker = CreateObject( "Broker.Application" )
	
	astrFields = Split(strRow, ",")
	
	strStock = GetKeyForRow(strRow)
	
	'Set objStock = AmiBroker.Stocks.Add(strStock)
	
	'Set objQuote = objStock.Quotations.Add(dtmDate)
	
	'SYMBOL,SERIES,OPEN,HIGH,LOW,CLOSE,LAST,PREVCLOSE,TOTTRDQTY,TOTTRDVAL,TIMESTAMP,
  '  0      1      2    3   4    5     6      7         8         9         10
	
	'objQuote.High = astrFields(3)
	'objQuote.Low = astrFields(4)
	'objQuote.Open = astrFields(2)
	'objQuote.Close = astrFields(5)
	
	'objQuote.Volume = astrFields(5)
	
	WScript.Echo "Importing file=" & GetFilenameForDate(dtmDate)
	
	lngResult = AmiBroker.Import(0,GetTempFolderPath() & "\" & GetFilenameForDate(dtmDate),GetTempFolderPath() & "\nsedaily.format")
	
	If lngResult <> 0 Then
    WScript.Echo "ERROR !!!: There was an error importing file:" & GetFilenameForDate(dtmDate)
  End If
	
	Call AmiBroker.RefreshAll()
	
	'WScript.Echo "Stock=" & strStock & vbTab & "Volume=" & objQuote.Volume
	

End Sub


'****************************************************
'Function returns key for a given row
'****************************************************
Function GetKeyForRow(strRow)
	Dim i
	If Len(Trim(strRow)) > 2 Then
		i = InStr(1, strRow, ",", 1) - 1
		
		If i > 0 Then
			GetKeyForRow = Left(strRow, i)
		Else
			GetKeyForRow = ""
		End If
	Else
		GetKeyForRow = ""
	End If

End Function



'****************************************************
'Function filename for a given date object
'****************************************************
Function GetFilenameForDate(dtmDate)
	Dim strFilename, monthNames, strYear, strMonth, strDay
	strFilename = ""
	
	monthNames = Array("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
	
	strYear = DatePart("yyyy", dtmDate)
	strMonth = monthNames(DatePart("m", dtmDate))
	strDay = DatePart("d", dtmDate)
	
	strFilename = Replace(cstrFilename, "{YYYY}", strYear)
	strFilename = Replace(strFilename, "{MMM}", strMonth)
  strFilename = Replace(strFilename, "{DD}", Pad(strDay, 2, "0"))

	'WScript.Echo strYear & "/" & strMonth & "/" & strDay
	'WScript.Echo strBHAVCOPYurl

	GetFilenameForDate = strFilename
End Function

'****************************************************
'Function for padding month and day
'****************************************************
Function Pad(Value, Width, Char) 
  Pad = Right(String(Width, CStr(Char)) & CStr(Value), Width) 
End Function

'****************************************************
'Function returns modified URL for BHAVCOPY
'****************************************************
Function GetBHAVCOPYurl(dtmDate)
	Dim strBHAVCOPYurl, monthNames, strYear, strMonth, strDay
	strBHAVCOPYurl = ""
	
	monthNames = Array("", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
	
	'dtmDate = DateAdd("d" , -1, Date)
	'dtmDate = Date
	strYear = DatePart("yyyy", dtmDate)
	strMonth = monthNames(DatePart("m", dtmDate))
	strDay = DatePart("d", dtmDate)
	
	strBHAVCOPYurl = Replace(cstrBHAVCOPYurl, "{YYYY}", strYear)
	strBHAVCOPYurl = Replace(strBHAVCOPYurl, "{MMM}", strMonth)
	strBHAVCOPYurl = Replace(strBHAVCOPYurl, "{DD}", strDay)

	'WScript.Echo strYear & "/" & strMonth & "/" & strDay
	'WScript.Echo strBHAVCOPYurl

	GetBHAVCOPYurl = strBHAVCOPYurl & GetFilenameForDate(dtmDate)
End Function



'****************************************************
'Function reads from a given file and returns string
'****************************************************
Function ReadFromFile(strFile)
	Dim objFSO, objFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFile, ForReading, True)
	
	ReadFromFile = objFile.ReadAll()
	
End Function

'****************************************************
'Function writes to a given file
'****************************************************
Function WriteToFile(strFile, strText)
	Dim objFSO, objFile
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(strFile, ForWriting, True)
	
	objFile.Write(strText)
	
End Function

Dim gstrNSETest
gstrNSETest = ""

'****************************************************
'Function to fetch a file from web
'****************************************************
Function ReadFileFromWeb(strURL)
	On Error Resume Next
	Dim objXML
	
	'WScript.Echo "Debug: ReadFileFromWeb: strUrl=" & strUrl & vbCrLf
	
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.6.0")
	
	objXML.Open "GET", strURL, False
  objXML.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
  If gstrNSETest <> "" Then
    objXML.setrequestheader "Cookie", "NSE-TEST=" & gstrNSETest & ";"
  End If
  
  If gstrProxyServer <> "" Then
    objXML.setProxy 2, gstrProxyServer, "<local>"
    objXML.setProxyCredentials gstrProxyUsername, gstrProxyPassword
  End If


	objXML.Send				   
  
  If Err.Number <> 0 Then
    WScript.Echo "ERROR!!!: ReadFileFromWeb: Could not download URL=" & strURL
    WScript.Echo vbTab & "Description: " & Err.Description
  End If
  
  
  Dim temp, regEx, Matches
  temp = objXML.getResponseHeader("Set-Cookie")
  Set regEx = New RegExp
  regEx.Pattern = "NSE-TEST=(.*?);"
  Set Matches = regEx.Execute(temp)
  If Matches.Count > 0 Then
    If Matches(0).SubMatches.Count > 0 Then
      gstrNSETest = Matches(0).SubMatches(0)
    End If
  End If
  
	ReadFileFromWeb = objXML.ResponseText

End Function



'****************************************************
'Function to fetch a file from web, pass params as Dictionary
'****************************************************
Function ReadFileFromPOST(strURL, strParams)
	On Error Resume Next
	Dim objXML
	
	'WScript.Echo "Debug: ReadFileFromWebPOST: strUrl=" & strUrl & vbCrLf
	
	Set objXML = CreateObject("MSXML2.ServerXMLHTTP.6.0")
	
	objXML.Open "POST", strURL, False
  objXML.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
  objXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
  If gstrNSETest <> "" Then
    objXML.setrequestheader "Cookie", "NSE-TEST=" & gstrNSETest & ";"
  End If
  
  If gstrProxyServer <> "" Then
    objXML.setProxy 2, gstrProxyServer, "<local>"
    objXML.setProxyCredentials gstrProxyUsername, gstrProxyPassword
  End If


	objXML.Send	strParams			   
  
  If Err.Number <> 0 Then
    WScript.Echo "ERROR!!!: ReadFileFromWebPOST: Could not download URL=" & strURL
    WScript.Echo vbTab & "Description: " & Err.Description
  End If
  
  
  Dim temp, regEx, Matches
  temp = objXML.getResponseHeader("Set-Cookie")
  Set regEx = New RegExp
  regEx.Pattern = "NSE-TEST=(.*?);"
  Set Matches = regEx.Execute(temp)
  If Matches.Count > 0 Then
    If Matches(0).SubMatches.Count > 0 Then
      gstrNSETest = Matches(0).SubMatches(0)
    End If
  End If
  
	ReadFileFromPOST = objXML.ResponseText

End Function


'****************************************************
'Function to extract image names from the shtml
'****************************************************
Function GetImageArray(strHTML)
	'On Error Resume Next
	Dim objRegExp, objMatch, objMatches
	
	set objRegExp = new regexp
	
	objRegExp.Pattern = "/gms(.*).jpg"
	
	objRegExp.Global = True
	
	set objMatches = objRegExp.Execute(strHTML)
	
	set GetImageArray = objMatches
	
End Function


'****************************************************
'Function to save a binary file from the web
'****************************************************
Function SaveWebBinary(strUrl, strFile) 'As Boolean
  On Error Resume Next
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
	Const ForWriting = 2
	
	Dim web, varByteArray, strData, strBuffer, lngCounter, ado
    Dim objXML
	
	Set web = CreateObject("MSXML2.ServerXMLHTTP.6.0")
	web.Open "GET", strURL, False
	web.setrequestheader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)"
  If gstrProxyServer <> "" Then
    web.setProxy 2, gstrProxyServer, "<local>"
    web.setProxyCredentials gstrProxyUsername, gstrProxyPassword
  End If
	web.Send		
	
    If Err.Number <> 0 Then
        WScript.Echo "ERROR!!!: Could not download URL=" & strURL
        WScript.Echo vbTab & "Description: " & Err.Description
        SaveWebBinary = False
        Set web = Nothing
        Exit Function
    End If
    'WScript.echo web.Status
    If web.Status <> "200" Then
        SaveWebBinary = False
        Set web = Nothing
        Exit Function
    End If
    varByteArray = web.ResponseBody
    
    Set web = Nothing
    'Now save the file with any available method
    'On Error Resume Next
    Set ado = Nothing
    Set ado = CreateObject("ADODB.Stream")
    If ado Is Nothing Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set ts = fs.OpenTextFile(strFile, ForWriting, True)
        strData = ""
        strBuffer = ""
        For lngCounter = 0 to UBound(varByteArray)
            ts.Write Chr(255 And Ascb(Midb(varByteArray,lngCounter + 1, 1)))
        Next
        ts.Close
    Else
        ado.Type = adTypeBinary
        ado.Open
        ado.Write varByteArray
        ado.SaveToFile strFile, adSaveCreateOverWrite
        ado.Close
    End If
    SaveWebBinary = True
    Exit Function
End Function



'************************************************************************************************************************************
'Sectors&Industries code follows
'************************************************************************************************************************************




'****************************************************
'Function sectors file
'****************************************************
Function GetSectorsFile()
  Dim strFilePath, strFolderPath
  
  strFolderPath = GetAmibrokerFolder()
  
  If strFolderPath <> "" Then
    strFilePath = strFolderPath & "\broker.sectors"
  End If
  GetSectorsFile = strFilePath
End Function


'****************************************************
'Function industries file
'****************************************************
Function GetIndustriesFile()
  Dim strFilePath, strFolderPath
  
  strFolderPath = GetAmibrokerFolder()
  
  If strFolderPath <> "" Then
    strFilePath = strFolderPath & "\broker.industries"
  End If
  GetIndustriesFile = strFilePath
End Function

'****************************************************
'Function gets amibroker install folder
'****************************************************
Function GetAmibrokerFolder()
  Dim strFolderPath, Sh, strCLSID
  
  Set Sh = CreateObject("WScript.Shell")

  strCLSID = Sh.RegRead("HKCR\Broker.Application\CLSID\")
  
  If strCLSID <> "" Then
    strFolderPath = Sh.RegRead("HKCR\CLSID\{2DCDD57C-9CC9-11D3-BF72-00C0DFE30718}\LocalServer32\")
    If strFolderPath <> "" Then
      Dim objFSO
      Set objFSO = CreateObject("Scripting.FileSystemObject")
      strFolderPath = objFSO.GetParentFolderName(strFolderPath)
    End If
  End If
  GetAmibrokerFolder = strFolderPath
End Function


'****************************************************
'Function get sector names in a dictionary
'****************************************************
Function GetSectors()
  Dim dicSectors, strBaseUrl, strSectorsFile, astrSectors
  Set dicSectors = CreateObject("Scripting.Dictionary")
  
  strBaseUrl = "http://www.nseindia.com/content/indices/"
  
  dicSectors.Add "CNX IT Index", strBaseUrl & "ind_cnxitlist.csv"
  dicSectors.Add "CNX Bank Index", strBaseUrl & "ind_cnxbanklist.csv"
  dicSectors.Add "CNX FMCG Index", strBaseUrl &  "ind_cnxfmcglist.csv"
  dicSectors.Add "CNX PSE Index", strBaseUrl & "ind_cnxpselist.csv"
  dicSectors.Add "CNX MNC Index", strBaseUrl & "ind_cnxmnclist.csv"
  dicSectors.Add "CNX Service Sector Index", strBaseUrl & "ind_cnxservicelist.csv"
  dicSectors.Add "CNX Energy Index", strBaseUrl & "ind_cnxenergylist.csv"
  dicSectors.Add "CNX Pharma Index", strBaseUrl & "ind_cnxpharmalist.csv"
  dicSectors.Add "CNX Infrastructure Index", strBaseUrl & "ind_cnxinfralist.csv"
  dicSectors.Add "CNX PSU BANK Index", strBaseUrl & "ind_cnxpsubanklist.csv"
  
  strSectorsFile = GetSectorsFile()
  
  astrSectors = dicSectors.Keys
  
  If strSectorsFile <> "" Then
    Dim i, strSectorFileOutputText
    For i = 0 to UBound(astrSectors)
      strSectorFileOutputText = strSectorFileOutputText & astrSectors(i) & vbCrLf
    Next
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FileExists(strSectorsFile) And Not(objFSO.FileExists(strSectorsFile & ".bkp"))Then
      objFSO.CopyFile strSectorsFile, strSectorsFile & ".bkp", False
    End If
    'WScript.Echo strSectorFileOutputText
    Call WriteToFile(strSectorsFile, strSectorFileOutputText)
    Call Amibroker.RefreshAll()
  Else
    WScript.Echo "ERROR: Cannot update sectors since sectors file was not found."
  End If
  Set GetSectors = dicSectors
End Function

'****************************************************
'Function get industry names in a dictionary
'****************************************************
Function GetIndustries(dicSectors)
  Dim dicIndustries, astrSectors
  Set dicIndustries = CreateObject("Scripting.Dictionary")

  astrSectors = dicSectors.Items
  
  Dim i 
  For i = 0 to UBound(astrSectors)
    
  Next
  
  Set GetIndustries = dicIndustries
End Function
