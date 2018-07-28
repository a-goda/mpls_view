# mpls_view
A script which outputs an overview of a Cisco MPLS environment by showing VRF related information, leaks, L3 interfaces, and summary of each VRF static routes
The output is an Excel workbook, each workbook sheet contains a specific PE related information - by default, but could aggregate multiple PE in one sheet

The analysis is done on log files of the commands output "show run" and "sh vlan" for each PE. Each PE on a separate log file.

###Inputs
1. A directory that contains PEs' log file
2. Extension to filter files based on
	
###Outputs - on the same directory of log files
1. MPLS overview Excel file
2. L3 interfaces Excel file
