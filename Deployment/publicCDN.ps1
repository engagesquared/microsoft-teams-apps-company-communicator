function IsValidParam {
    [OutputType([bool])]
    param
    (
        [Parameter(Mandatory = $true)]
        $param
    )

    return -not([string]::IsNullOrEmpty($param.Value)) -and ($param.Value -ne '<<value>>')
}

# Validate input parameters.
function ValidateParameters {
    $isValid = $true
    if (-not(IsValidParam($parameters.adminLogin))) {
        WriteE -message "Invalid adminLogin."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.password))) {
        WriteE -message "Invalid password."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.tenantAdminUrl))) {
        WriteE -message "Invalid tenantAdminUrl."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.publicCDNSiteUrl))) {
        WriteE -message "Invalid publicCDNSiteUrl."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.publicCDNSiteTitle))) {
        WriteE -message "Invalid publicCDNSiteTitle."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.orgAssetsLibraryTitle))) {
        WriteE -message "Invalid orgAssetsLibraryTitle."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.resourceGroupName))) {
        WriteE -message "Invalid resourceGroupName."
        $isValid = $false;
    }
    if (-not(IsValidParam($parameters.appServiceName))) {
        WriteE -message "Invalid appServiceName."
        $isValid = $false;
    }
    
    return $isValid
}


# Load Parameters from JSON meta-data file
$parametersListContent = Get-Content '.\publicCDNParameters.json' -ErrorAction Stop

# Validate all the parameters.
$parameters = $parametersListContent | ConvertFrom-Json
if (-not(ValidateParameters)) {
    EXIT
}


$credential = New-Object System.Management.Automation.PSCredential($parameters.adminLogin.Value, (ConvertTo-SecureString $parameters.password.Value -AsPlainText -Force))
Connect-PnPOnline -Url $parameters.tenantAdminUrl.Value -Credentials $credential
Set-PnPTenantCdnEnabled -CdnType Public -Enable $true


$site = Get-PnPTenantSite -Identity $parameters.publicCDNSiteUrl.Value -ErrorAction SilentlyContinue
if (!$site)
{
    New-PnPSite -Type CommunicationSite -Title $parameters.publicCDNSiteTitle.Value -Url $parameters.publicCDNSiteUrl.Value -Owner $parameters.adminLogin.Value
}

Connect-PnPOnline -Url $parameters.publicCDNSiteUrl.Value -Credentials $credential
$publicCdnSite = Get-PnPSite -Includes Id
$web = Get-PnPWeb
New-PnPList -Title $parameters.orgAssetsLibraryTitle.Value -Template DocumentLibrary
$list = Get-PnPList -Identity $parameters.orgAssetsLibraryTitle.Value
Set-PnPWebPermission -User "Everyone except external users" -AddRole "Contribute"
Add-PnPOrgAssetsLibrary -LibraryUrl "$($parameters.publicCDNSiteUrl.Value)/$($parameters.orgAssetsLibraryTitle.Value)" -CdnType Public


Connect-AzureRmAccount -Credential $credential
$appService = Get-AzureRmWebApp -ResourceGroupName $parameters.resourceGroupName.Value -Name $parameters.appServiceName.Value
$appSettings = $appService.SiteConfig.AppSettings
$newAppSettings = @{}
ForEach ($item in $appSettings) {
    $newAppSettings[$item.Name] = $item.Value
}

$newAppSettings.PublicCDNSiteId = "$($publicCdnSite.Id)"
$newAppSettings.PublicCDNWebId = "$($web.Id)"
$newAppSettings.PublicCDNListId = "$($list.Id)"
$newAppSettings.SharepointHostName = $parameters.tenantAdminUrl.Value.Substring(8, $parameters.tenantAdminUrl.Value.IndexOf("-admin") - 8)

Set-AzureRmWebApp -AppSettings $newAppSettings -Name $parameters.appServiceName.Value -ResourceGroupName $parameters.resourceGroupName.Value