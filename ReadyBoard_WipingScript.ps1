#Connect-PnPOnline -Url "http://share-internal.deere.com/teams/IEsuptJDWW/Components"
# Put Core URL Above
# Put List Name Below (In both Lines of the script)
#Will it works at ~1 Item per Second the Rough Estimate is: Total Items / 60 
#$items = Get-PnPListItem -List "IE Testing Board"

#foreach ($Item in $items){
#    try{
#        Remove-PnPListItem -List "IE Testing Board" -Identity $Item.Id -Force
#        }
#    catch{
#        Write-Host "Error"
#        }
#}