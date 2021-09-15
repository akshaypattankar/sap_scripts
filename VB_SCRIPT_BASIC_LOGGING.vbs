StartTime = timer
StartDateTime = Now()

'Add the logic of the script after current line
'
'
'
'

log_report_name = "SAMPLE REPORT" 'Add report name for logs
log_csv_location = "C:\VB_SCRIPT_LOGS.CSV" 'Add file location for log

EndTime = timer
EndDateTime = Now()

DeltaTime = FormatNumber( EndTime - StartTime , 2 )

LogText = log_report_name & "," & StartDateTime & "," & EndDateTime & "," & DeltaTime

Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(log_csv_location,8,true)
objFileToWrite.WriteLine(LogText)
objFileToWrite.Close
Set objFileToWrite = Nothing

MsgBox ( log_report_name & " script completed in " & DeltaTime  & " seconds")
