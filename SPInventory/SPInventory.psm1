﻿<#
    .NOTES
    --------------------------------------------------------------------------------
     Code generated by:  SAPIEN Technologies, Inc., PowerShell Studio 2018 v5.5.152
     Generated on:       6/11/2018 2:50 AM
     Generated by:       Administrator
    --------------------------------------------------------------------------------
    .DESCRIPTION
        Script generated by PowerShell Studio 2018
#>


<#	
	 Created by:   	Evgeni Petrov
	 Filename:     	SPInventory.psm1
	 Module Name: SPInventory
#>



function Export-SPIObject
{
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[string]$Output
	)
	
	$spiobject = Get-SPIObject
	$spiobject | Export-Clixml -LiteralPath $Output
}

Export-ModuleMember -Function Export-SPIObject


function Get-SPIObject
{
	[CmdletBinding()]
	param ()
	
	Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
	
	$spFarm = Get-SPFarm
	$spDatabase = Get-SPDatabase
	$spServer = Get-SPServer
	$spServiceInstance = Get-SPServiceInstance
	$spWebApplication = Get-SPWebApplication -IncludeCentralAdministration
	$spSite = $spWebApplication | Get-SPSite -Limit All
	$spAlternateUrl = Get-SPAlternateURL
	$quotaTemplates = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.QuotaTemplates
	$spEnterpriseSearchServiceApplication = Get-SPEnterpriseSearchServiceApplication
	$spServiceApplication = Get-SPServiceApplication
	$spFeature = Get-SPFeature
	$spSolution = Get-SPSolution
	$spManagedAccount = Get-SPManagedAccount
	
	
	$object = New-Object -TypeName System.Management.Automation.PSObject
	
	$member = Get-SPIFarmOverview
	$object | Add-Member -NotePropertyName 'FarmOverview' -NotePropertyValue $member
	
	$member = Get-SPIServersInFarm
	$object | Add-Member -NotePropertyName 'ServersInFarm' -NotePropertyValue $member
	
	$member = Get-SPIWebApplicationsAndSiteCollections
	$object | Add-Member -NotePropertyName 'WebApplicationsAndSiteCollections' -NotePropertyValue $member
	
	$member = Get-SPIContentDatabases
	$object | Add-Member -NotePropertyName 'ContentDatabases' -NotePropertyValue $member
	
	$member = Get-SPIServicesOnServer
	$object | Add-Member -NotePropertyName 'ServicesOnServer' -NotePropertyValue $member
	
	$member = Get-SPIWebApplicationList
	$object | Add-Member -NotePropertyName 'WebApplicationList' -NotePropertyValue $member
	
	$member = Get-SPIAlternateAccessMappings
	$object | Add-Member -NotePropertyName 'AlternateAccessMappings' -NotePropertyValue $member
	
	$member = Get-SPIIISSettings
	$object | Add-Member -NotePropertyName 'IISSettings' -NotePropertyValue $member
	
	$member = Get-SPISiteCollectionList
	$object | Add-Member -NotePropertyName 'SiteCollectionList' -NotePropertyValue $member
	
	$member = Get-SPISiteCollectionUsageAndProperties
	$object | Add-Member -NotePropertyName 'SiteCollectionUsageAndProperties' -NotePropertyValue $member
	
	$member = Get-SPIQuotaTemplates
	$object | Add-Member -NotePropertyName 'QuotaTemplates' -NotePropertyValue $member
	
	$member = Get-SPISiteCollectionQuotas
	$object | Add-Member -NotePropertyName 'SiteCollectionQuotas' -NotePropertyValue $member
	
	$member = Get-SPIContentSources
	$object | Add-Member -NotePropertyName 'ContentSources' -NotePropertyValue $member
	
	$member = Get-SPIScopes
	$object | Add-Member -NotePropertyName 'Scopes' -NotePropertyValue $member
	
	$member = Get-SPIServiceApplicationList
	$object | Add-Member -NotePropertyName 'ServiceApplicationList' -NotePropertyValue $member
	
	$member = Get-SPIUPSAAdministrators
	$object | Add-Member -NotePropertyName 'UPSAAdministrators' -NotePropertyValue $member
	
	$member = Get-SPIServiceApplicationPermissions
	$object | Add-Member -NotePropertyName 'ServiceApplicationPermissions' -NotePropertyValue $member
	
	$member = Get-SPIWebApplicationAssociations
	$object | Add-Member -NotePropertyName 'WebApplicationAssociations' -NotePropertyValue $member
	
	$member = Get-SPIContentDatabaseDetails
	$object | Add-Member -NotePropertyName 'ContentDatabaseDetails' -NotePropertyValue $member
	
	$member = Get-SPIFarmDatabaseDetails
	$object | Add-Member -NotePropertyName 'FarmDatabaseDetails' -NotePropertyValue $member
	
	$member = Get-SPIOutgoingEmailSettings
	$object | Add-Member -NotePropertyName 'OutgoingEmailSettings' -NotePropertyValue $member
	
	$member = Get-SPIFarmFeatures
	$object | Add-Member -NotePropertyName 'FarmFeatures' -NotePropertyValue $member
	
	$member = Get-SPISolutions
	$object | Add-Member -NotePropertyName 'Solutions' -NotePropertyValue $member
	
	$member = Get-SPIFarmAdministrators
	$object | Add-Member -NotePropertyName 'FarmAdministrators' -NotePropertyValue $member
	
	$member = Get-SPIManagedAccounts
	$object | Add-Member -NotePropertyName 'ManagedAccounts' -NotePropertyValue $member
	
	$member = Get-SPIServiceAccounts
	$object | Add-Member -NotePropertyName 'ServiceAccounts' -NotePropertyValue $member
	
	$member = Get-SPISharePointDesignerSettings
	$object | Add-Member -NotePropertyName 'SharePointDesignerSettings' -NotePropertyValue $member

    Remove-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
	
	Write-Output $object
}


function Get-SPIFarmOverview
{
	[CmdletBinding()]
	param ()
	
	$installedSharePointVersion = $spFarm.BuildVersion.ToString()
	$license = Get-SPIFarmLicense
	$configurationDatabase = $spDatabase | Where-Object { $_.Type -eq 'Configuration Database' } | Select-Object -ExpandProperty Name
	
	$property = [ordered]@{
		'Installed SharePoint version'	   = $installedSharePointVersion
		'License'						   = $license
		'Configuration database'		   = $configurationDatabase
	}
	$payload = New-Object -TypeName System.Management.Automation.PSObject -Property $property
	
	$title = "Farm Overview"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload -TypeName 'PSCustomObject'
	Write-Output $output
}


function Get-SPIFarmLicense {
    $license = Get-SPUserLicense
    
    $licenses = @()
    foreach($l in $license){
        $licenses += $l.License
    }
    if ($licenses -contains 'Enterprise'){
        Write-Output 'Enterprise'
    }
    elseif($licenses -contains 'Standard'){
        Write-Output 'Standard'
    }
    elseif($licenses -contains '???'){
        Write-Output 'SharePoint Foundation'
    }
}


function Get-SPIServersInFarm
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($server in $spServer)
	{
		$role = 'Other'
		if ($server.Role -ne 'Invalid')
		{
			$role = $server.Role
		}
		$operatingSystem = Invoke-Command -ComputerName $server.Name { Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -ExpandProperty Caption }
		$memory = Invoke-Command -ComputerName $server.Name { (Get-CimInstance -ClassName 'Cim_PhysicalMemory' | Measure-Object -Property Capacity -Sum).Sum }
		$property = [ordered]@{
			'ServerName'    = $server | Select-Object -ExpandProperty Name
			'Role'		    = $role
			'OperatingSystem' = $operatingSystem
			'Memory'	    = $memory/1GB
		}
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Servers in farm"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIWebApplicationsAndSiteCollections
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		$spSites = $spSite | Where-Object { $_.WebApplication.Name -eq ($webApplication | Select-Object -ExpandProperty Name) }
		foreach ($site in $spSites)
		{
			$siteAdmins = $site.Owner.DisplayName
			if ($site.SecondaryContact -ne $null)
			{
				$siteAdmins = $siteAdmins + "; " + $site.SecondaryContact.DisplayName
			}
			$property = [ordered]@{
				'WebApplication'    = $webApplication | Select-Object -ExpandProperty DisplayName
				'SiteCollection'    = $site.Url
				'SiteAdmins'	    = $siteAdmins
			}
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
		
	}
	
	$title = "Web applications and site collections"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIContentDatabases
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		foreach ($contentDatabase in $webApplication.ContentDatabases)
		{
			$property = [ordered]@{
				'Name'	   = $contentDatabase.Name
			}
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
		
	}
	
	$title = "Content Databases"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIServicesOnServer
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	$servers = $spServer | Where-Object { $_.Role -eq 'Application' }
	$typeNames = $spServiceInstance | Select-Object -ExpandProperty TypeName -Unique
	foreach ($typeName in $typeNames)
	{
		$property = @{
			'ServiceName'	  = $typeName
		}
		
		foreach ($server in $servers)
		{
			$serverName = $server.Address
			$serviceStatus = ($spServiceInstance | Where-Object { $_.TypeName -eq $typeName -and $_.Server.Name -eq $serverName }).Status
			
			$property.Add($serverName, $serviceStatus)
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Services on Server"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIWebApplicationList
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		$property = @{
			'WebApplication'    = $webApplication.DisplayName
			'Url'			    = $webApplication.Url
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Content Databases"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIAlternateAccessMappings
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		
		$alternateUrls = Get-SPAlternateURL -WebApplication $webApplication
		
		foreach ($alternateUrl in $alternateUrls)
		{
			$property = @{
				'DisplayName'	  = $webApplication.DisplayName
				'InternalUrl'	  = $alternateUrl.IncomingUrl
				'Zone'		      = $alternateUrl.Zone
				'Url'			  = $alternateUrl.PublicUrl
			}
		}
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Alternate Access Mappings"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIIISSettings
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		
		$alternateUrls = Get-SPAlternateURL -WebApplication $webApplication
		
		foreach ($alternateUrl in $alternateUrls)
		{
			$authenticationProvider = Get-SPAuthenticationProvider -WebApplication $webApplication -Zone $alternateUrl.Zone
			$ssl = $webApplication.Url -like 'https*'
			if ($authenticationProvider.DisableKerberos -eq $true -and $authenticationProvider.UseWindowsIntegratedAuthentication -eq $true)
			{
				$authentication = 'NTLM'
			}
			elseif ($authenticationProvider.DisableKerberos -eq $false -and $authenticationProvider.UseWindowsIntegratedAuthentication -eq $true)
			{
				$authentication = 'Kerberos'
			}
			
			$property = @{
				'DisplayName'    = $webApplication.DisplayName
				'Url'		     = $alternateUrl.IncomingUrl
				'Zone'		     = $alternateUrl.Zone
				'Authentication' = $authentication
				'ApplicationPoolName' = $webApplication.ApplicationPool.Name
				'ApplicationPoolIdentity' = $webApplication.ApplicationPool.Username
				'SSl'		     = $ssl
				'ClaimsAuth'	 = $webApplication.UseClaimsAuthentication
				'CEIP'		     = $webApplication.BrowserCEIPEnabled
			}
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
	}
	
	$title = "IIS Settings"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPISiteCollectionList
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		$spSites = $spSite | Where-Object { $_.WebApplication.Name -eq ($webApplication | Select-Object -ExpandProperty Name) }
		foreach ($site in $spSites)
		{
			$owners = $site.Owner.DisplayName
			if ($site.SecondaryContact -ne $null)
			{
				$siteAdmins = $siteAdmins + "; " + $site.SecondaryContact.DisplayName
			}
			$properties = @{
				'WebApplication'    = $webApplication.DisplayName
				'Title'			    = $site.RootWeb.Title
				'Url'			    = $site.Url
				'ContentDatabase'   = $site.ContentDatabase.Name
				'Owners'		    = $owners
			}
			
			$object = New-Object -TypeName PSObject -Property $properties
			
			$payload += $object
		}
	}
	
	$title = "Site Collection List"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPISiteCollectionUsageAndProperties
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		$spSites = $spSite | Where-Object { $_.WebApplication.Name -eq ($webApplication | Select-Object -ExpandProperty Name) }
		foreach ($site in $spSites)
		{
			$properties = @{
				'WebApplication'    = $webApplication.DisplayName
				'Title'			    = $site.RootWeb.Title
				'NumberOfWebs'	    = $site.AllWebs.Count
				'Storage'		    = $site.Usage.Storage/1GB
			}
			
			$object = New-Object -TypeName PSObject -Property $properties
			
			$payload += $object
		}
	}
	
	$title = "Site collection usage and properties"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIQuotaTemplates
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($quotaTemplate in $quotaTemplates)
	{
		$property = @{
			'TemplateName'	   = $quotaTemplate.Name
			'StorageMaximumLevel' = $quotaTemplate.StorageMaximumLevel
			'StorageWarningLevel' = $quotaTemplate.StorageWarningLevel
			'InvitedUserMaximumLevel' = $quotaTemplate.InvitedUserMaximumLevel
			'UserCodeMaximumLevel' = $quotaTemplate.UserCodeMaximumLevel
			'UserCodeWarningLevel' = $quotaTemplate.UserCodeWarning.Level
			'WarningLevelEmail' = '?' # TODO
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		
		$payload += $object
	}
	
	$title = "Quota templates"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPISiteCollectionQuotas
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($site in $spSite)
	{
		$siteQuota = $siteQuotas | Where-Object { $_.QuotaID -eq $site.Quota.QuotaID }
		$properties = [ordered]@{
			'SiteCollection'   = $site.RootWeb.Title
			'Url'			   = $site.Url
			'QuotaName'	       = $siteQuota.Name
			'LockStatus'	   = $site.ReadOnly
			'StorageMaximumLevel' = $site.Quota.StorageMaximumLevel
			'StorageWarningLevel' = $site.Quota.StorageWarningLevel
			'Usage'		       = $site.Usage.Storage
		}
		
		$object = New-Object -TypeName PSObject -Property $properties
		$payload += $object
	}
	
	$title = "Content Databases"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIContentSources
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($searchApplication in $spEnterpriseSearchServiceApplication)
	{
		$spContentSource = Get-SPEnterpriseSearchCrawlContentSource -SearchApplication $spEnterpriseSearchServiceApplication
		foreach ($contentSource in $spContentSource)
		{
			$property = [ordered]@{
				'ServiceApplication'   = $spEnterpriseSearchServiceApplication.Name
				'Name'				   = $contentSource.Name
				'Type'				   = $contentSource.Type
			}
			
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
		
	}
	
	$title = "Content sources"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIScopes
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($searchApplication in $spEnterpriseSearchServiceApplication)
	{
		$spQueryScope = Get-SPEnterpriseSearchQueryScope -SearchApplication $spEnterpriseSearchServiceApplication
		foreach ($queryScope in $spQueryScope)
		{
			$property = [ordered]@{
				'ServiceApplication'   = $spEnterpriseSearchServiceApplication.Name
				'DisplayName'		   = $queryScope.Name
				'LastModifiedBy'	   = $queryScope.LastModifiedBy
				'AlternateResultsPage' = $queryScope.AlternateResultsPage
				'DefaultSearchResults' = '' # TODO
				'DifferentPageForSearching' = '' # TODO
				'RuleCount'		       = $queryScope.Count
			}
			
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
		
	}
	
	$title = "Scopes"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIServiceApplicationList
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($serviceApplication in $spServiceApplication)
	{
		$status = 'Offline'
		foreach ($serviceInstance in $serviceApplication.ServiceInstances)
		{
			if ($serviceInstance.Status -eq 'Online')
			{
				$status = 'Online'
			}
		}
		
		$property = [ordered]@{
			'Name'	   = $serviceApplication.DisplayName
			'Type'	   = $serviceApplication.TypeName
			'Status'   = $status
			'ApplicationPool' = $serviceApplication.ApplicationPool.Name
			'Identity' = $serviceApplication.ApplicationPool.ProcessAccountName
			'Databases' = '' # TODO
			'Server'   = '' # TODO
			'FailoverServer' = '' # TODO
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
		
	}
	
	$title = "Service Application List"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIUPSAAdministrators
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	$upsa = $spServiceApplication | Where-Object { $_.TypeName -eq 'User Profile Service Application' }
	$administrationAccessControl = $upsa.GetAdministrationAccessControl()
	foreach ($accountName in $administrationAccessControl.AccessRules)
	{
		foreach ($permission in $administrationAccessControl.NamedAccessRights)
		{
			$property = [ordered]@{
				'ServiceApplication'   = $upsa.DisplayName
				'AccountName'		   = $accountName.Name
				'Permission'		   = $permission.Name
			}
			
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
		
	}
	
	$title = "UPSA Administrators"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIServiceApplicationPermissions
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($serviceApplication in $spServiceApplication)
	{
		$administrationAccessControl = $serviceApplication.GetAdministrationAccessControl()
		
		foreach ($permission in $administrationAccessControl.NamedAccessRights)
		{
			if (-not $administrationAccessControl.AccessRules)
			{
				$property = [ordered]@{
					'ServiceApplication'   = $serviceApplication.DisplayName
					'AccountName'		   = 'Local farm'
					'Permission'		   = $permission.Name
				}
				
				$object = New-Object -TypeName PSObject -Property $property
				$payload += $object
			}
			else
			{
				foreach ($accountName in $administrationAccessControl.AccessRules)
				{
					$property = [ordered]@{
						'ServiceApplication'   = $serviceApplication.DisplayName
						'AccountName'		   = $accountName.Name
						'Permission'		   = $permission.Name
					}
					
					$object = New-Object -TypeName PSObject -Property $property
					$payload += $object
				}
			}
		}
	}
	
	$title = "Service Application Permissions"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIWebApplicationAssociations
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		
		$property = [ordered]@{
			'WebApplication' = $webApplication.DisplayName
			'ProxyGroup'	 = $webApplication.ServiceApplicationProxyGroup.FriendlyName
			'Applications'   = ($webApplication.ServiceApplicationProxyGroup.Proxies | Select-Object -ExpandProperty DisplayName) # TODO Expand Applications
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Web application associations"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIContentDatabaseDetails
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	$spContentDatabase = $spDatabase | Where-Object { $_.Type -eq 'Content Database' }
	foreach ($contentDatabase in $spContentDatabase)
	{
		
		$property = [ordered]@{
			'DatabaseName' = $contentDatabase.Name
			'DatabaseServer' = $contentDatabase.Server
			'WebApplication' = $contentDatabase.WebApplication.DisplayName
			'Status'	   = $contentDatabase.Status
			'CurrentSiteCount' = $contentDatabase.CurrentSiteCount
			'WarningSiteCount' = $contentDatabase.WarningSiteCount
			'MaximumSiteCount' = $contentDatabase.MaximumSiteCount
			'SizeGB'	   = $contentDatabase.DiskSizeRequired/1GB
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Content database details"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIFarmDatabaseDetails
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	$databases = $spDatabase | Where-Object { $_.Type -ne 'Content Database' }
	foreach ($database in $databases)
	{
		
		$property = [ordered]@{
			'DatabaseName' = $database.Name
			'Type'		   = $database.Type
			'DatabaseServer' = $database.Server.Name
			'Status'	   = $database.Status
			'RecoveryModel' = '' # TODO
			'SizeGB'	   = $database.DiskSizeRequired/1GB
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Farm Database details"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIOutgoingEmailSettings
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		
		$property = [ordered]@{
			'WebApplication' = $webApplication.DisplayName
			'OutboundEmailServer' = $webApplication.OutboundMailServiceInstance.Server.Name
			'FromAddress'    = $webApplication.OutboundMailSenderAddress
			'ReplyToAddress' = $webApplication.OutboundMailReplyToAddress
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Outgoing email settings"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIFarmFeatures
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	$spFarmFeature = $spFeature | Where-Object { $_.Scope -eq 'Farm' }
	foreach ($feature in $spFarmFeature)
	{
		$custom = $false
		$solutionName = $null
		foreach ($solution in $spSolution)
		{
			if ($feature.SolutionId -eq $solution.SolutionId.Guid)
			{
				$custom = $true
				$solutionName = $solution.Name
			}
		}
		$property = [ordered]@{
			'Title' = $feature.DisplayName
			'FeatureId' = $feature.Id
			'Hidden' = $feature.Hidden
			'Custom' = $custom
			'Solution' = $solutionName
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Farm features"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPISolutions
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($solution in $spSolution)
	{
		$property = [ordered]@{
			'Name' = $solution.DisplayName
			'DeploymentState' = $solution.DeploymentState
			'ContainsCasPolicy' = $solution.ContainsCasPolicy
			'DeployedTo' = $solution.DeployedWebApplications
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Solutions"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIFarmAdministrators
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	$caWebApplication = $spWebApplication | Where-Object { $_.IsAdministrationWebApplication -eq $true }
	$farmAdmins = (Get-SPSite -Identity $caWebApplication.Url).RootWeb.SiteGroups["Farm Administrators"].Users
	foreach ($admin in $farmAdmins)
	{
		$property = [ordered]@{
			'LoginName' = $admin.UserLogin
			'FullName'  = $admin.DisplayName
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Farm administrators"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIManagedAccounts
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($managedAccount in $spManagedAccount)
	{
		$property = [ordered]@{
			'Username' = $managedAccount.Username
			'AutomaticChange' = $managedAccount.AutomaticChange
			'LastChange' = $managedAccount.PasswordLastChanged
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "Managed accounts"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPIServiceAccounts
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		$property = [ordered]@{
			'Component' = 'Web application pool - {0}' -f $webApplication.DisplayName
			'Account'   = $webApplication.ApplicationPool.Username
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	foreach ($serviceApplication in $spServiceApplication)
	{
		if ($serviceApplication.ApplicationPool.ProcessAccountName)
		{
			$property = [ordered]@{
				'Component' = 'Service application pool - {0}' -f $serviceApplication.Name
				'Account'   = $serviceApplication.ApplicationPool.ProcessAccountName
			}
			
			$object = New-Object -TypeName PSObject -Property $property
			$payload += $object
		}
		
	}
	
	$title = "Service accounts"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}


function Get-SPISharePointDesignerSettings
{
	[CmdletBinding()]
	param ()
	
	$payload = @()
	foreach ($webApplication in $spWebApplication)
	{
		$property = [ordered]@{
			'WebApplication' = $webApplication.DisplayName
			'Enabled'	     = $webApplication.AllowDesigner
			'EnableDetachingPages' = $webapplication.MasterPageReferenceEnabled
			'EnableCustomizingPages' = $webapplication.AllowMasterPageEditing
			'EnableManaging' = '' # TODO
		}
		
		$object = New-Object -TypeName PSObject -Property $property
		$payload += $object
	}
	
	$title = "SharePoint designer settings"
	
	$output = New-Object -TypeName System.Management.Automation.PSObject
	$output | Add-Member -MemberType NoteProperty -Name 'Title' -Value $title
	$output | Add-Member -MemberType NoteProperty -Name 'Payload' -Value $payload
	Write-Output $output
}
