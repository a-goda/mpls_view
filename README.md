# mpls_view
A script which outputs an overview of a Cisco MPLS environment by showing VRF related information, leaks, L3 interfaces, and summary of each VRF static routes
The output is an Excel workbook, each workbook sheet contains a specific PE related information - by default, but could aggregate multiple PE in one sheet

The analysis is done on log files of the commands output "show run" and "sh vlan" for each PE. Each PE on a separate log file.
### Prerequisites:
    1. install openpyxl package
        $pip install openpyxl
        
### Inputs: 
	1. A directory that contains PEs' log file
	2. Extension to filter files based on
	
### Outputs - on the same directory of log files: 
	1. MPLS overview Excel file
	2. L3 interfaces Excel file


## Log files
    1. Save each PE log file by the extention .log
        Path and the expected extention could be changed at the beginning of the main function
    2. PE log file name options
        Option 1: Preceed the name with number (ex. "1 PE.log")
            The number is used to order Excel workbook sheets
        
        Option 2: Log file name without preceding number (ex. "PE2.log") 
            Log file information would be collected with other PEs - without precedding number - log files 
            on a workbook sheet called Branches

        
          