Class EWSTraceListener : Microsoft.Exchange.WebServices.Data.ITraceListener {
    
    [void] Trace ([string] $traceType, [string] $traceMessage) {
        CreateXMLTextFile($traceType,$traceMessage.ToString())
    }

    [void] CreateXMLTextFile ([string] $FileName, [string] $TraceContent) {
        # Create a new XML file for the trace information.
        try {
            # If the trace data is valid XML, create an XmlDocument object and save.
            $xmlDoc = New-Object System.Xml.XmlDocument
            $xmlDoc.Load($TraceContent)
            $xmlDoc.Save("$env:temp\EWSGui Logging\tracelog.xml")
        }
        catch {
            # If the trace data is not valid XML, save it as a text document.
            [system.IO.File]::WriteAllText("$env:temp\EWSGui Logging\tracelog.txt", $TraceContent)
        }
    }
}