$vnemployee = Import-Excel 'D:\Work\04. Projects\201811_Cadena_AD_Sync\VM_EMP.xlsx' -WorkSheetname DB
#$vnemployee | ogv 
$vnhavemail = $vnemployee |
    where {$_.employeeid -ne $null -and $_.employeeid -ne "" -and $_.email -like "*@adenservices.com" -and $_.EmployeeStatus -eq "Active"} | 
    select employeeid, email
$vnhavemail | ogv
$vnhavemail.Count

for ($i = 0; $i -lt $vnhavemail.Count; $i++)
{ 
    $employeeid = $vnhavemail[$i].EmployeeID
    $sam = ($vnhavemail[$i].email -split "@")[0]
    #$sam
    try
    {
        Set-ADUser $sam -EmployeeID $employeeid -ErrorAction Stop
        #$employeeid + " " + $sam + " OK"
        #get-aduser $sam -Properties employeeid
    }
    catch [System.Exception]
    {
        Write-Host "Error: $_"
    }
    
}