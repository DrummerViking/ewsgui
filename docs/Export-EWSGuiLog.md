# Export-EWSGuiLog

All tasks performed by the EWSGui module will generate logs. These logs are available in a local folder, but you can also use the `Export-EWSGuiLog` which can easily extract the logs.

## Parameters

### FilePath  
Defines the file path to export the CSV file.  
Default value is the user's Desktop with a file name like "yyyy-MM-dd HH_mm_ss" - EWSGui logs.csv"

## OutputType  
Defines the output types available. Can be a single output or combined.  
Current available options are CSV, GridView. Default value is 'GridView'.

## DaysOld
Defines how old we will go to fetch the logs. Valid range is between 1 through 7 days old. Default Value is 1.  

# Example

For example you can run:
``` powershell
PS C:\> Export-EWSGuiLog -OutputType CSV,GridView -DaysOld 5
```  

In this example, the script will fetch all logs within the last 5 days, export to CSV to default location at the user's Desktop and also displays in powershell's GridView.