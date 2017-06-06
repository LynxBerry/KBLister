<#

0 = Not started
1 = In progress
2 = Succeeded
3 = Succeeded With Errors
4 = Failed
5 = Aborted


#>


$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()
$HistoryCount = $Searcher.GetTotalHistoryCount()
$Updates = $Searcher.QueryHistory(0,$HistoryCount)
foreach ($Upd in $Updates) {
$OutPut = New-Object -Type PSObject -Prop @{
        'ComputerName'=$env:computername
        'Product' = $($Upd.Categories).Name
        'UpdateDate'=$Upd.date
        'KB'=[regex]::match($Upd.Title,'KB(\d+)')
        'UpdateTitle'=$Upd.title
        'UpdateDescription'=$Upd.Description
        'SupportUrl'=$Upd.SupportUrl
        'UpdateId'=$Upd.UpdateIdentity.UpdateId
        'RevisionNumber'=$Upd.UpdateIdentity.RevisionNumber
        'Operation' = $Upd.operation
        'ResultCode' = $Upd.resultcode
    }
    
    write-output $OutPut    
}