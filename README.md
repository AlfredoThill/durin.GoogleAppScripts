# durin.GoogleAppScripts

#Sheets API & Drive API

In order to perform multiple tasks, a spreadsheet was made with a linked project and two sheets, one where the user chooses the task to be performed and another that shows the results. The operations involve a few hundred spreadsheets and consist of: permission management, report generation, exports such as csv to the same user drive and a comparison between the spreadsheets and a query on the database of the platform "moodle" The project makes use of "google.script.run" to chain calls to the google server by dividing the work in batches of approximately 2.5 minutes by passing a key between calls to advance the iteration and a simple sidebar to operate.