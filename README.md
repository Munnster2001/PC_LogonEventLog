Advanced Logon History Viewer
This PowerShell script provides a graphical user interface (GUI) to query Windows Security Event Logs for logon and logoff events. It can be run on a local machine or a remote computer (with appropriate permissions). The script filters for common logon-related event IDs, presents the data in a user-friendly table, and allows for exporting the results to a CSV file.

Features
Intuitive GUI: Easy-to-use interface for non-technical users.

Local & Remote Queries: Query the event logs of the local computer or any remote machine on the network.

Time-Based Filtering: Specify the number of days to search back.

Data Export: Export the filtered event data to a .csv file for further analysis.

Detailed Event View: Select a row in the table to view the full details of a specific event.

Requirements
PowerShell 5.1 or newer: The script relies on cmdlets and syntax introduced in this version.

Administrator Privileges: The script must be run as an Administrator to access the Security event log.

Remote Access: To query a remote computer, the WinRM service must be running and the appropriate firewall rules must be enabled on the target machine.

How to Use
Run the script: Open PowerShell as an Administrator and execute the script file.

Enter Computer Name: Type localhost to query the local machine, or enter a remote computer's name (e.g., SERVER01).

Specify Days: Enter the number of days you wish to search back in the event logs.

Click "Get Advanced History": The script will query the event log, and the results will populate the table.

View Details: Click on any row in the table to see a detailed breakdown of that specific event.

Export Data: Click the Export to CSV button to save the current data table to a file.

Common Event IDs Used
4624: Logon Successful

4625: Logon Failed

4634: Logoff

4647: User-initiated Logoff

4800: Workstation Locked

4801: Workstation Unlocked

4672: Special Logon (admin privileges assigned)

This README provides a great overview and will help anyone quickly get started with your excellent script!
