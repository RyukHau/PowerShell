#Check MSOnline Module
if(!(get-module msonline))
{
    install-module msonline -Confirm:$false
}

#Upgrade MSOnline Module
import-module msonline

#Display Successd
write-host "Enter Azure credentials to connect Azure AD" -ForegroundColor yellow

#Connect to Azure AD
Connect-MsolService

#Import CSV Data, csv data is Username and MFA Statu
$users = Import-Csv 'Path\BulkUpdateMFASampleFile.csv'

#Setting MFA
foreach ($user in $users)
{
    #Get CSV Data
    $NewState= $user.'MFA Status'
    
    #Check List
    if($NewState -eq "Enabled")
    {
        #Enabled MFA
        $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
        $st.RelyingParty = "*"
        $st.State = $NewState
        $sta = @($st)

        Set-MsolUser -UserPrincipalName $user.Username -StrongAuthenticationRequirements $sta

        #Check User MFA
		$CurrentState =  (Get-MsolUser -UserPrincipalName $user.Username | Select -ExpandProperty StrongAuthenticationRequirements).state

		#Display Result
		if($CurrentState -eq "Enabled" -or $CurrentState -eq "Enforced")
        {
            write-host "Current state of MFA for user $user is - Done" -ForegroundColor Green
        }
        else
        {
            write-host "Current state of MFA for user $user is - Fail" -ForegroundColor Red
        }
    }

    else
    {
        #Disabled MFA
		$cfa = @()

        Set-MsolUser -UserPrincipalName $user.Username -StrongAuthenticationRequirements $cfa

        #Check User MFA
		$CurrentState =  (Get-MsolUser -UserPrincipalName $user.Username | Select -ExpandProperty StrongAuthenticationRequirements).state

        #Display Result
		if($CurrentState -eq "Enabled" -or $CurrentState -eq "Enforced")
        {
            write-host "Current state of MFA for user $user is Fail" -ForegroundColor Red
        }
        else
        {
            write-host "Current state of MFA for user $user is Done" -ForegroundColor Green
        }
    }
}

#Export MFA result
$MFAStatu=Get-MsolUser -All | Select UserPrincipalName, @{N="Statu"; E={$_.StrongAuthenticationRequirements.State}}

$MFAStatu | Export-Csv -Path 'Path\MFAStatu.csv' -NoTypeInformation -Encoding UTF8

#Logout
Disconnect-AzureAD
