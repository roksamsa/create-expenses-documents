Dim objExcel1, objExcel2, objExcel3, objWorkbook1, objWorkbook2, objWorkbook3

Set objExcel1 = CreateObject("Excel.Application")
Set objExcel2 = CreateObject("Excel.Application")
Set objExcel3 = CreateObject("Excel.Application")

objExcel1.Visible = false
objExcel1.DisplayAlerts = False
objExcel2.Visible = false
objExcel2.DisplayAlerts = False
objExcel3.Visible = false
objExcel3.DisplayAlerts = False

currentLocation = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

costsYear = WScript.Arguments.Named("costsYear")
costsMonth = WScript.Arguments.Named("costsMonth")
tripsThisMonth = WScript.Arguments.Named("tripsThisMonth")

'Only 1 trip this month
If tripsThisMonth = "1" Then

  firstTravel = WScript.Arguments.Named("firstTravel")
  secondTravel = WScript.Arguments.Named("secondTravel")
  travelStartTime = WScript.Arguments.Named("travelStartTime")
  travelEndTime = WScript.Arguments.Named("travelEndTime")

  fileLocation1 = currentLocation + "\" + costsYear + "\" + costsMonth + "\1\Daily-allowance_Spesenblatt_SamsaR_" + costsMonth + costsYear + "_01.xlsx"

  Set objWorkbook1 = objExcel1.Workbooks.Open(fileLocation1)

  objExcel1.Range("B1").value = costsMonth + costsYear + "-" + "01"
  objExcel1.Range("H1").value = Date
  objExcel1.Range("C6").value = firstTravel + "." + costsMonth + "." + costsYear
  objExcel1.Range("C10").value = secondTravel + "." + costsMonth + "." + costsYear

'2 trips this month
ElseIf tripsThisMonth = "2" Then
  firstTravel = WScript.Arguments.Named("firstTravel")
  secondTravel = WScript.Arguments.Named("secondTravel")
  thirdTravel = WScript.Arguments.Named("thirdTravel")
  forthTravel = WScript.Arguments.Named("forthTravel")
  travelStartTime1 = WScript.Arguments.Named("travelStartTime1")
  travelStartTime2 = WScript.Arguments.Named("travelStartTime2")
  travelEndTime1 = WScript.Arguments.Named("travelEndTime1")
  travelEndTime2 = WScript.Arguments.Named("travelEndTime2")

  fileLocation1 = currentLocation + "\" + costsYear + "\" + costsMonth + "\1\Daily-allowance_Spesenblatt_SamsaR_" + costsMonth + costsYear + "_01.xlsx"
  fileLocation2 = currentLocation + "\" + costsYear + "\" + costsMonth + "\2\Daily-allowance_Spesenblatt_SamsaR_" + costsMonth + costsYear + "_02.xlsx"

  Set objWorkbook1 = objExcel1.Workbooks.Open(fileLocation1)
  Set objWorkbook2 = objExcel2.Workbooks.Open(fileLocation2)

  objExcel1.Range("B1").value = costsMonth + costsYear + "-" + "01"
  objExcel1.Range("H1").value = Date
  objExcel1.Range("C6").value = firstTravel + "." + costsMonth + "." + costsYear
  objExcel1.Range("C10").value = secondTravel + "." + costsMonth + "." + costsYear
  objExcel1.Range("G6").value = travelStartTime1
  objExcel1.Range("D28").value = travelEndTime1

  objExcel2.Range("B1").value = costsMonth + costsYear + "-" + "02"
  objExcel2.Range("H1").value = Date
  objExcel2.Range("C6").value = thirdTravel + "." + costsMonth + "." + costsYear
  objExcel2.Range("C10").value = forthTravel + "." + costsMonth + "." + costsYear
  objExcel2.Range("G6").value = travelStartTime2
  objExcel2.Range("D28").value = travelEndTime2

  objWorkbook1.Save
  objWorkbook1.Close
  objExcel1.Quit
  Set objExcel1 = Nothing
  Set objWorkbook1 = Nothing
  objWorkbook2.Save
  objWorkbook2.Close
  objExcel2.Quit
  Set objExcel2 = Nothing
  Set objWorkbook2 = Nothing

'3 trips this month
ElseIf tripsThisMonth = "3" Then
  fileLocation1 = currentLocation + "\" + costsYear + "\" + costsMonth + "\1\Daily-allowance_Spesenblatt_SamsaR_" + costsMonth + costsYear + "_01.xlsx"
  fileLocation2 = currentLocation + "\" + costsYear + "\" + costsMonth + "\2\Daily-allowance_Spesenblatt_SamsaR_" + costsMonth + costsYear + "_02.xlsx"
  fileLocation3 = currentLocation + "\" + costsYear + "\" + costsMonth + "\3\Daily-allowance_Spesenblatt_SamsaR_" + costsMonth + costsYear + "_03.xlsx"

  Set objWorkbook1 = objExcel1.Workbooks.Open(fileLocation1)
  Set objWorkbook2 = objExcel2.Workbooks.Open(fileLocation2)
  Set objWorkbook3 = objExcel3.Workbooks.Open(fileLocation3)

  objExcel1.Range("B1").value = costsMonth + costsYear + "-" + "01"
  objExcel2.Range("B1").value = costsMonth + costsYear + "-" + "02"
  objExcel3.Range("B1").value = costsMonth + costsYear + "-" + "03"

Else
   WScript.Echo "You were travelling more than 3 times this month!"
End If

'tempstr1 = replace(firstTravel, "-", ".")
