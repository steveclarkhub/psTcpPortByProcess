$ProcColl = @{}
$listenPorts = Get-NetTCPConnection | Where-Object {$_.OwningProcess -ne 0} | Select-Object LocalAddress, LocalPort, RemoteAddress, OwningProcess

foreach ($port in $listenPorts) {
    $ProcNm = Get-CimInstance Win32_Process -Filter "ProcessId = '$($port.OwningProcess)'"
    if ($ProcNm) {
        $ProcessKey = "{0}-{1}-{2}" -f $ProcNm.ProcessId, $ProcNm.ProcessName, (Invoke-CimMethod -InputObject $ProcNm -MethodName GetOwner).user
        if (-not $ProcColl.ContainsKey($ProcessKey)) {
            $output = [PSCustomObject]@{
                ProcessName = $ProcNm.ProcessName
                ProcessID = $ProcNm.ProcessId
                ProcessOwner = (Invoke-CimMethod -InputObject $ProcNm -MethodName GetOwner).user
                TcpPorts = [System.Collections.Generic.List[int]]::new() # Use List[int]
                WorkingSetMB = [Math]::Round(($ProcNm.WorkingSetSize / 1MB), 2)
                ThreadCount = $ProcNm.ThreadCount
                HandleCount = $ProcNm.HandleCount
            }
            $ProcColl[$ProcessKey] = $output
        }
        if ($ProcColl[$ProcessKey].TcpPorts -notcontains $port.LocalPort) { # Check for existing ports
            $ProcColl[$ProcessKey].TcpPorts.Add($port.LocalPort) # Use Add() for List
        }
    }
}

# Output the results
$ProcColl.Values | ForEach-Object {
    [PSCustomObject]@{
        ProcessID = $_.ProcessID
        ProcessName = $_.ProcessName
        ProcessOwner = $_.ProcessOwner
        TcpPorts = ($_.TcpPorts | Sort-Object) -join ',' # Directly sort and join the List
        WorkingSetMB = if ($_.WorkingSetMB -lt 1) { "<1" } else { $_.WorkingSetMB }
        ThreadCount = $_.ThreadCount
        HandleCount = $_.HandleCount
    }
} | Sort-Object ProcessID | Format-Table -AutoSize