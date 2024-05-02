Function Generate-Complex-Domain-Password ([Parameter(Mandatory=$true)][int]$PassLenght)
{

Add-Type -AssemblyName System.Web
$requirementsPassed = $false
do {
$newPassword=[System.Web.Security.Membership]::GeneratePassword($PassLenght,1)
If ( ($newPassword -cmatch "[A-Z\p{Lu}\s]") `
-and ($newPassword -cmatch "[a-z\p{Ll}\s]") `
-and ($newPassword -match "[\d]") `
-and ($newPassword -match "[^\w]")
)
{
$requirementsPassed=$True
}
} While ($requirementsPassed -eq $false)
return $newPassword
}
Generate-Complex-Domain-Password (8)