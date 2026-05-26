# VBA-Outlook
```
Function AddOrdinalSuffix(ByVal d As Integer) As String

    Dim suffix As String
    
    Select Case d Mod 100
        Case 11, 12, 13
            suffix = "th"
        Case Else
            Select Case d Mod 10
                Case 1: suffix = "st"
                Case 2: suffix = "nd"
                Case 3: suffix = "rd"
                Case Else: suffix = "th"
            End Select
    End Select
    
    AddOrdinalSuffix = d & suffix

End Function
```
```
Sub CarparkReportTemplate()

    Dim mail As MailItem
    Dim today As Date
    Dim start As Date
    Dim rangeStr As String
    
    today = Date
    start = DateAdd("d", -3, today)
    
    rangeStr = Format(start, "d/m") & " - " & Format(today, "d/m")
    
    ' === Create new mail ===
    Set mail = Application.CreateItem(olMailItem)
    
    ' === Predefined fields ===
    With mail
        .To = "kinyung.looi@shcsb.com.my; WeiSheng.Lee@shcsb.com.my;L1Support@shcsb.com.my "
        .CC = "penghon.yong@shcsb.com.my; khairil.hishamuddin@shcsb.com.my; tanushaa.t@shcsb.com.my; mahallaxmie.d@shcsb.com.my"
        .subject = rangeStr & " Carpark Registration, Carpark Termination, Invoice Payment & Carpark Change Report"
        ' Remove signature completely
        .htmlBody = ""
    End With
    
    
    ' === Today's date ===
    Dim todayStr As String
    todayStr = AddOrdinalSuffix(Day(Date)) & " " & Format(Date, "mmmm yyyy")
    
    ' === Start date (3 days before today) ===
    Dim startDateStr As String
    startDateStr = AddOrdinalSuffix(Day(start)) & " " & Format(start, "mmmm yyyy")
    
    ' === Your custom body template ===
    Dim bodyContent As String
    bodyContent = "<p>[**ATTACH CSV**]</p>" & _
                  "<p>Hi all,</p>" & _
                  "<p>Attached above are the carpark registration, carpark termination, invoice payment & carpark change of bay status report as of " & startDateStr & " - " & todayStr & ".</p>" & _
                  "<p>For your reference, there are no failed transactions.</p>" & _
                  "<p>[**SCREENSHOT**]</p>" & _
                  "<p>[**YOUR OWN SIGNATURE**]</p>" & _
                  "<br>"
    
    ' === Combine body + signature ===
    mail.htmlBody = bodyContent
    
    ' === Display email ===
    mail.Display

End Sub
```
