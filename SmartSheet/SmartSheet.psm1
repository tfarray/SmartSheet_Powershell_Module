# SmartSheet Version 2.0
# Developped by Thomas Farray .:|:. Cisco .:|:.
# Updated to work with API V2

if (!(get-module ModulesUpdater)) { import-module ModulesUpdater }

Function Get-SmartSheetAPIToken {
 <#
	.Synopsis
		Tool used to retrieve a smartsheet API token and store it into a PSCredential object. The file is encrypted and may only be read by the user account which created it. This script will consider that you are using the same Smartsheet username than your actual session username. if not, please specify with the -username switch
	.Description
		The Get-SmartSheetAPIToken go fetch credentials that are stored into a file in the user's profile directory. It returns a System.Management.Automation.PSCredential object. The file is encrypted using the Windows Data Protection API (DPAPI) standard string representation
	.Example
		$SSToken = get-SmartSheetAPIToken
		If no credentials were previously set, this command will prompt for a user/passowrd. In either case, it will return a System.Management.Automation.PSCredential object
	.Example
		$SSToken = get-SmartSheetAPIToken -user mySmartSheetUserName
		If no credentials are set, Prompts for a passowrd and returns a System.Management.Automation.PSCredential object
		Optional : The name of the credential's account
	.Parameter pwdfile
		Optional : file FQDN to store the hash
		Default : The file is stored in you profile in "\appdata\\Roaming\[login].pwd"
	.Parameter pwdfile
		Optional : file FQDN to store the hash
		Default : The file is stored in you profile in "\appdata\\Roaming\[login].pwd"
   .Inputs
		null
   .OutPuts
		[PSCredential]
 #>
   [cmdletbinding()]
   param(
		[string]$pwdfile = 	$env:APPDATA + "\" + ($env:username).split(".")[0] + "_SSToken.pwd",
		[string]$user = ($env:username).split(".")[0]
    )
	process {
		# Getting Credentials either from file or from prompt
		if (test-path $pwdfile) {
			$SecureString = Get-Content $pwdfile  | convertto-securestring
		} else {
			$SecureString = Read-Host Please enter your SmartSheet API Token for the account ($user) -AsSecureString 
		}
		$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $user, $SecureString
		write-verbose "SmartSheet token read for $user"
		return $Credential
	}		
}
Export-ModuleMember -Function Get-SmartSheetAPIToken

Function Set-SmartSheetAPIToken {
 <#
	.Synopsis
		Tool used to store SmartSheet API token in the users's profile directory. Created credentials may only be read by the user who created them
	.Description
		The Set-SmartSheetAPIToken is meant to store a SmartSheet API Token to an encrypted file.
	.Example
		Set-SmartSheetAPIToken
		Will prompt for a passowrd for the username of the actual session. It then stores the credentials into a hashed file called in your profile
	.Example
		Set-SmartSheetAPIToken 
		Prompts for a password, and stores the credentials into a hashed file.
	.Example
		Set-SmartSheetAPIToken -user MySmartSheetUserAccount
		Prompts for a password for MySmartSheetUserAccount, and stores the credentials into a hashed file.
	.Parameter user
		Optional : The name of the smartsheet account account
	.Parameter pwdfile
		Optional : file FQDN to store the hash
		Default : The file is stored in you profile in "\appdata\Roaming\[login].pwd"
   .Inputs
		[int]
   .OutPuts
		[bool]
 #>
   param(
		[string]$pwdfile = 	$env:APPDATA + "\" + ($env:username).split(".")[0] + "_SSToken.pwd",
		[string]$user = ($env:username).split(".")[0]
    )
	
	process {
		# Getting Credentials either from file or from prompt
		write-verbose ("login : $username, file : $pwdfile, validation = $NoValidation")
		$SecureString = Read-Host Please enter your password for the account ($user) -AsSecureString  
		$Credential = new-object -typename System.Management.Automation.PSCredential -argumentlist $user, $SecureString 
		$SecureString | ConvertFrom-SecureString | out-file $pwdfile
		write-host Crendentials stored to $pwdfile -foregroundcolor darkgreen
	}
}
Export-ModuleMember -Function Set-SmartSheetAPIToken

function Invoke-Smartsheet {
	[cmdletbinding()]
	Param(
		[ValidateSet("contacts","favorites","folders","groups","home","reports","search","serverinfo","sheets","sheets","templates","token","users","users","workspaces")]
			[string]$SSMainFuntion = "sheets",
		[ValidateSet("rows","columns","attachments","discussions","updaterequests")] 
			[string]$SSFuntion,
		[ValidateSet("Get","Head","Post","Put","Delete","Trace","Options","Merge","Patch")]
			[string]$Method="Get",
		[string]$SSData,
		$Row,
		[Hashtable]$Query,
		$SS
	)
	
	#Some variables
	if (!$global:SmartSheetToken) { $global:SmartSheetToken = Get-SmartSheetAPIToken }
	$SS_API_Version = "2.0"
	$URL = "https://api.smartsheet.com/$SS_API_Version/$SSMainFuntion/" 
	$headers = @{Authorization = "Bearer $(($global:SmartSheetToken).GetNetworkCredential().password)"}
	# write-verbose "Smartsheet ID : $($ss.id)"
	# Building the URL that will be queried
	if ($SSMainFuntion -like "sheets" -and ($SS -or $Row)) {
		if (!$SS) { $SS = $Row.__SmartSheet }
		$URL += [string]($SS.ID) + "/" 
		if ($SSFuntion) {  	$URL += "$SSFuntion"  }
		if ($Row) { $URL += "/" + [string]($Row.__ID) }
	}
	if ($Query) { $URL += "?" + ( $Query.keys | % { "$_=" + ( $Query.$_ -join "," ) }  ) }
	
	write-verbose ("URL : $URL `nMethod : $Method `n Data : $($SSData -join " ") `nHeaders : " + (($headers.Keys | % { "$_ = $($headers.$_)" } ) -join "|" ))
	
	try {
		switch -regex ($method) {
			"get|delete" { 
				$Invoked = Invoke-WebRequest $URL -Headers $headers  -method $Method}
			"put|post" { 
				$Invoked = Invoke-WebRequest $URL -Headers $headers -method $Method -Body $SSdata -ContentType "application/json" }
		}
	} catch {
		write-host "Invoke-Smartsheet Commmand failed with error : $(($_.ErrorDetails.Message | ConvertFrom-Json).message)" -foregroundcolor red
		return
	}
	return ConvertFrom-Json $Invoked.content
}
Export-ModuleMember -Function Invoke-Smartsheet

function ConvertFrom-SmartSheetRowObject {
	param(
		$ROW,
		$ReferenceRow,
		$ExtraProperties
	)
	process {
		$RowProperties = @{}
		if ($ROW) {
			if (!$row.__SmartSheet) { $row | Add-Member -MemberType noteproperty -Name __SmartSheet -Value $SS }
			if ($row.__id) { $RowProperties.Add("id",$row.__id) }
			$CellProperties = @($ROW.psobject.properties | select -expandproperty name | ? { $_ -notmatch "__" } | % {
				new-object psobject -property @{ 
					columnId = $row.__SmartSheet.Colname2ID.$_
					value    = $Row.$_
				}
			})
			$RowProperties.add("cells",$CellProperties)
		}
		$ExtraProperties.keys | % { if ($_) { $RowProperties.Add($_,$ExtraProperties.$_) }}
		$JasonData = new-object psobject -property $RowProperties | ConvertTo-Json
			
		return "[ $JasonData ]"
	}
}
Export-ModuleMember -Function  ConvertFrom-SmartSheetRowObject

function ConvertTo-SmartSheetRowObject {
	param($JasonRow,$SS)
	process {
		# write-verbose $JasonRow.count
		foreach ($JSR in @($JasonRow)) {
			# First we put all data and we create the object
			# $Properties = $ss.Colname2ID.Clone()
			$Properties = @{}
			$Properties.add("__id",$JSR.id)
			$Properties.add("__smartsheet",$SS)
			$ValuesToShow = @()
			$JSR.cells | select columnId,value | ? { $_.columnId -in $SS.ID2Colname.Keys } | % {
				$ColName = $SS.ID2Colname.($_.columnId)
				$Properties.add($ColName,$_.value)
				$ValuesToShow+=$ColName
			}
			$NewRow = New-Object psobject -Property $Properties
			#write-host $ValuesToShow
			#Fromatting output
			$NewRow.PSObject.TypeNames.Insert(0,'SmartSheet.Table')
			$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet",[string[]]$ValuesToShow)
			$NewRow | Add-Member MemberSet PSStandardMembers ([System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet))
		
		
			# write-verbose "qsdfdsf" # $JSR
			#The we check parentship
			if ($JSR.parentId) {
				# write-verbose "Addin Data To parent & child"
				$Parent = $ss.Table_ID2PSo.($JSR.parentId)
				add-member -inputobject $NewRow noteproperty -name "__ParentNode" -value @($Parent)
				if ($Parent.__Childnode) { 
					$Parent.__Childnode += $NewRow
				} else { 
					add-member -inputobject $Parent noteproperty -name "__Childnode" -value @($NewRow)
				}
			} else {
				$SS.table += @($NewRow)
			}

			# We now add to the Smartsheet the new row and we add All methods
			$SS.Table_ID2PSo.add($JSR.id,$NewRow)
			Add-Member -InputObject $NewRow -MemberType ScriptMethod -Name 'Update' -Value $SSUpdateRowBlock #-PassThru
			Add-Member -InputObject $NewRow -MemberType ScriptMethod -Name 'AddChild' -Value $SSAddNewChild #-PassThru
			Add-Member -InputObject $NewRow -MemberType ScriptMethod -Name 'AddRow' -Value $SSAddNewRow #-PassThru
			Add-Member -InputObject $NewRow -MemberType ScriptMethod -Name 'Delete' -Value $SSRemoveRow #-PassThru
		}
	}
}
Export-ModuleMember -Function  ConvertTo-SmartSheetRowObject

function Fix-SmartSheet_Row_Methods_and_Links {
	param(
		[Parameter(Mandatory=$true)][PsObject]$Row,
		[PsObject]$SS=$null,
		[PsObject]$LinkedRow=$null,
		$Placement=$null,
		$ID=$null
	)
	Process {
		#Fixing Associated Methods
		if (!$Row.psobject.methods.Item("AddRow")) {
			Add-Member -InputObject $Row -MemberType ScriptMethod -Name 'Update' -Value $SSUpdateRowBlock
			Add-Member -InputObject $Row -MemberType ScriptMethod -Name 'AddChild' -Value $SSAddNewChild
			Add-Member -InputObject $Row -MemberType ScriptMethod -Name 'AddRow' -Value $SSAddNewRow
			Add-Member -InputObject $Row -MemberType ScriptMethod -Name 'Delete' -Value $SSRemoveRow
		}
		# Fixing Row "__" records
		foreach ($prop in "__id","__SmartSheet") {
			if (!($row.psobject.Properties | ? { $_.name -like $prop })) {
				Add-member -InputObject $Row -MemberType noteproperty -Name $prop -Value ""
			}
			if (!$row.$prop) {
				switch ($prop) {
					"__id" { $row.$prop = $ID }
					"__SmartSheet" { $row.$prop = $SS }
				}
			}
			if (!$row.$prop) { write-error "unable to find a way to fill in $prop for this row" ; return }
		}
		
		#Fixing Row reference in the smartsheet object
		if ($row.__SmartSheet.Table_ID2PSo.keys -notcontains $row.__id) {
			$row.__SmartSheet.Table_ID2PSo.add($row.__id,$row) 
		}	
	}
}

function Fix-SmarSheet_New_Object_Placement {
	# Row is the Reference Row
	# Placement, is where it is places
	param($Row,$Placement)
	process {
		$NewObjProperties = @{}
		if ($Row) {
			if ($Placement -contains "ParentId") { 
				$NewObjProperties.add("parentId",$Row.__id) 
				if ($Placement -contains "toBottom") { $NewObjProperties.add("toBottom","true") }
			} else {
				$NewObjProperties.add("siblingId",$Row.__id) 
				if ($Placement -contains "above") { $NewObjProperties.add("above","true") }
			}
		} else {
			if ($Placement -contains "toTop") { 
				$NewObjProperties.add("toTop","true")
			} else { 
				$NewObjProperties.add("toBottom","true") 
			}
		}
		return $NewObjProperties
	}
}
# Export-ModuleMember -Function  Fix-SmarSheet_New_Object_Placement

function Clone-SmartSheetRow {
	param(
		[Parameter(Mandatory=$true)]$ROW
	)
	process {
		$NewROW = $Row.psobject.copy()
		$NewRow.psobject.properties | ? { $_.name -match "__" } | % { $NewRow.psobject.properties.remove($_.name) }
		return $NewRow
	}
}

function Update-SmartsheetRow {
	[cmdletbinding()] 	
	param(
		[Parameter(Position=0,Mandatory=$true)]$row,
		[string[]]$Column
	)
	process {
		if ($Column) {
			$JasonData = ConvertFrom-SmartSheetRowObject ($row | Select-Object($Column + "__*"))
		} else {
			$JasonData = ConvertFrom-SmartSheetRowObject $row
		}
		$Result = Invoke-Smartsheet -SSfuntion "rows"  -method "PUT" -ss $row.__smartsheet -SSData $JasonData
	}
}
Export-ModuleMember -Function Update-SmartsheetRow

function Get-SmartsheetRow {
	param([Parameter(Mandatory=$true)]$Row,[switch]$Format,[switch]$discussions,[switch]$attachments)
	$SS = $row.__smartsheet
	$Selection = "columnId","value"
	if ($Format -or $discussions -or $attachments ) {
		$Query = @{ include = @() }
		if ($Format) { $Query.include += "format" ; $Selection += "format" }
		if ($discussions) { $Query.include += "discussions" ; $Selection += "discussions" }
		if ($attachments) { $Query.include += "attachments" ; $Selection += "attachments" }
		$Invoked = Invoke-Smartsheet -SSfuntion "rows" -Row $Row -Query $Query
	} else {
		$Invoked = Invoke-Smartsheet -SSfuntion "rows" -Row $Row
	}
	if ($invoked) {
		$SSRawRow = $invoked.cells | select $Selection
		$Properties = @{}
		$SSRawRow | % { 
			if ($ss.ID2Colname.($_.columnId)) { $Properties.add($ss.ID2Colname.($_.columnId),$_.value)} 
		}
		return new-object PsObject -property $Properties
	}
}
Export-ModuleMember -Function Get-SmartsheetRow

function Add-SmartsheetRow {
	<#
	.Synopsis
		Add a new row to your SmartSheet
	.Description
		This will allow you to add a new row to you smartsheet
		 - The smartsheet object ($smartsheet)
		 - opt. A reference row
		 - Where to store the row
		 - The new values to store
	.Parameter SS
		Mandtory : A smartsheet object
	.Parameter Row
		Optional : A row of a table of the smartsheetobject
	.Parameter Values
		Mandatory : A list of columns to update, type Dictionnary
	.Example
		$MySmartSheet = Get-Smartsheet "VIF B*"
		###### TO UPDATE Set-SmartsheetRow -SS $MySmartSheet -Row $MySmartSheet.Table[0] -Column Note -NewValue NewVAL
   .Inputs
		PsObject,PsObject,string,string
   .OutPuts
		Nothing, but it will update the variable with the new value
 #>
	[cmdletbinding()]
	param(
		[PsObject]$SS,
		[PsObject]$Row,
		[PsObject]$ReferenceRow,
		[ValidateSet("toTop","toBotom","above","siblingId","parentId")][String[]]$Placement="toTop"
	)
	process {
		if (!$SS.ID) { $SS = $ReferenceRow.__smartsheet }
		if (!$SS.ID) { write-error "Please specify a reference Row or SmartSheet." ; return }
		
		# First we set correctly positioning
		$NewObjProperties = Fix-SmarSheet_New_Object_Placement -Row $ReferenceRow -Placement $Placement
		
		# Then we convert to Jason to push to Smartsheet
		$JasonData = ConvertFrom-SmartSheetRowObject -ExtraProperties $NewObjProperties
							
		# Running the new row addition against smartsheet
		write-verbose " Invoke Jason Data : $JasonData"
		$Result = Invoke-Smartsheet -SSfuntion rows -method POST -SSData $JasonData -ss $SS
			
		# return $Result
		if ($Result.resultCode -eq 0) {
			if (!$result.result) { write-warning "The command ran, but no changes were made" ; return }
			ConvertTo-SmartSheetRowObject -JasonRow $Result.result -ss $SS
		} else {
			write-host (" Unable to add a new") -foregroundcolor red
		}
	}
}
Export-ModuleMember -Function Add-SmartsheetRow

function Remove-SmartsheetRow {	
	<#
	.Synopsis
		Removes a row to your SmartSheet
	.Description
		This will simply remove the sent row
		 - The smartsheet object ($smartsheet)
		 - A reference row
	.Parameter SS
		Mandatory : A smartsheet object
	.Parameter Row
		Mandatory : Row of the smartsheetobject to remove
	.Parameter Recurse
		Optional : This will remove all the childs of the row
	.Example
		$MySmartSheet = Get-Smartsheet "VIF B*"
		Remove-SmartsheetRow -SS $MySmartSheet -Row $MySmartSheet.Table[0]
   .Inputs
		PsObject,PsObject
   .OutPuts
		Nothing, but it will update the variable with the new value
 #>
	[cmdletbinding()]
	param(
		[Parameter(Mandatory=$true)]$Row
	)
	process {
		$Result = Invoke-Smartsheet -SSfuntion rows -method DELETE -row $ROW
		
		if ($Result.resultCode -eq 0) { 	
			# write-verbose " Row(s) $($Result.result -join ',') have beed deleted"
			if ($row.__ParentNode) {
				$row.__ParentNode.__childnode = @($row.__ParentNode.__childnode | ? { $_.__id -notlike $row.__id } )
			} else {
				$row.__smartsheet.table = $row.__smartsheet.table  | ? { $_.__id -notlike $row.__id }
			}
			$Result.result
			$Result.result | % { 
				# $row.__smartsheet.Table_ID2PSo.$_ = $null
				$row.__smartsheet.Table_ID2PSo.remove($row.__id)
			}
		}
		
		# return $result
	}
}
Export-ModuleMember -Function Remove-SmartsheetRow

$SSGenerateLeafTableBlock = { 
	write-verbose "LeafTable : Building it !"
	$LTable = @() ; $LTable_ID2PSo  = @{}
	foreach ($node in ($this.Table_ID2PSo.values | ? {!$_.__Childnode} )) {
		# We need to duplicate the object, else we will modify the Table list of objects
		$NewObj = $node.psobject.copy()
		$NewObj.psobject.properties | ? { $_.name -match "__Childnode|__ParentNode" } | % { $NewObj.psobject.properties.remove($_.name) }
		Add-Member -InputObject $NewObj -MemberType NoteProperty -Name '__OriginalRow' -Value $node
		$LTable_ID2PSo.add($NewObj.__id,$NewObj)
		
		# Then, we need to update each line with flattended  values
		$ActualNode = $node
		[System.Collections.ArrayList]$MissingProperties = $ActualNode.psobject.properties | ? { $_.name -notmatch "__Childnode|__ParentNode" -and !$_.value } | select -expandproperty name
		while ($ActualNode.__ParentNode -and $MissingProperties) {
			$ActualNode = $ActualNode.__ParentNode
			$ToRemove = @($MissingProperties | % { if ($ActualNode.$_) { $NewObj.$_ = $ActualNode.$_ ; ($_) } })
			$ToRemove | % { $MissingProperties.remove($_)} # can not remove an item while enumerating
		}
		$LTable += $NewObj
	}
	Add-Member -InputObject $this -MemberType NoteProperty -Name 'LeafTable' -Value $LTable
	Add-Member -InputObject $this -MemberType NoteProperty -Name 'LTable_ID2PSo' -Value $LTable_ID2PSo
}

$SSUpdateRowBlock = {
	param([string[]]$Parameters2update)
	if ($Parameters2update) {
		Update-SmartsheetRow -row $this -Column ($Parameters2update)
	} else { Update-SmartsheetRow -row $this  }
}

$SSAddNewChild = {
	param([PsObject]$NewNode,[bool]$ToBottom=$false)
	if ($ToBottom) {
		Add-SmartsheetRow -parentRow $this -NewRow $NewNode -ToBottom -Placement parentId
	} else {
		Add-SmartsheetRow -parentRow $this -NewRow $NewNode -Placement parentId
	}
}
	
$SSAddNewRow = {
	param($NewNode)
	$Result = Add-SmartsheetRow -parentRow $this -NewRow $NewNode -Placement siblingId 
}

$SSGenerateServerInfo = {
	# Add-Member -InputObject $this -MemberType Noteproperty ServerInfo (invoke-smartsheet -SSMainFuntion "serverinfo" -SS $this)
	write-host $This.name
}

$SSRemoveRow = { Remove-SmartsheetRow -row $this }

function Get-Smartsheet {
<#
	.Synopsis
		This function is meant to fetch a SmartSheet and return a PSobject that will represent your SmartSheet
	.Description
		1 - Pre-requise
			Before beeing able to use this tool, you will need to create a SmartSheet API token.
			To get this token, please go to http://Smartsheet.cisco.com
			Then > Account (top left) > Personal settings > API Access > Generate a token`n
			You will be requested the token when launching the tool. You can also store it using the Set-SmartSheetAPIToken SmartSheet
		2 - If multiple replies
			You will get a list of all the seets. In case of mutiple similar names, please specify with -ID switch
		3 - When a Single SmartSheet is selected, you will get a new object
			Example : $MySmartSheet = get-smartsheet -ID 012345678910
			This object represent the smartsheet, here are the base sub-objects
				ID : the ID of the SpreadSheet 
				SStoken : A Secured variable used to store your SmartSheet API token
				ID2Colname : A dictionnary of ID / Columns name
				Colname2ID : A dictionnary of Columns name / ID
				Table : The represented SmartSheet. This is what you will use almost all the time
				Table_ID2PSo : A way to directly go from a Row ID to the Row Object (no need to filter)
			To this smartsheet Object, you may run additional commands (ex :  $MySmartSheet.GetServerInfo())
				GenerateLeafTable : Will generate a new table that only contains leaf object, but that will copy (if not not empty) all data of the parent. It will generate 2 variables : $MySmartSheet.LeafTable & $MySmartSheet.LTable_ID2PSo - equivalent to the previous one for Leaf Table
				GetServerInfo : Will retrieve the server info (which holds, for example, color shemes - please check API manual) and store it to a new variable $MySmartSheet.ServerInfo
	.Parameter name
		Optional : Name of the smartsheet. Wildcards accepted
	.Parameter ID
		Optional : ID of the smartsheet. Strict ID requested
	.Example
		You must specify the name or the [ID] of the smartsheet (-name or -ID), below are sheets corresponding to your filter to :
			> API [12312312312312]
			> Another Test API [12312312312312]
			...
	.Example
		$APITest = Get-Smartsheet "API"
		$APITest.Table
			(Get-Smartsheet API).table[0]
				__id         : 2610625433102212
				Nom          : AZER
				Prenom       : Thomas
				__SmartSheet : @{ID=12312312312312; SStoken=System.Management.Autom...}
   .Inputs
		String,Long
   .OutPuts
		pscustomobject
 #>
	[cmdletbinding()]
	param(
		[Parameter(Position=0)][String]$name="*",$ID
	)
	process {	
		if (!$ID) {
			$global:SmartSheetSheets = (Invoke-Smartsheet).data 
			$found =  $global:SmartSheetSheets | ? { $_.name -like $name }
			if (!$found) { write-host "No sheet was found using the search string : $name" -foregroundcolor red ; return }
			if ($found -is [array]) { 
				write-host " Multiple answers found. Please either specify another name, or use the -ID switch" -foregroundcolor darkred
				$found | % { write-host "  > $($_.name) [$($_.id)]" -foregroundcolor darkgreen }
				return
			}
			$ID = $found.ID
		}
		# write-verbose " Fetchin ID : $ID"
		$Invoked = Invoke-Smartsheet -SS (new-object psobject -property @{ID=$ID})

		# Preparing the SmartSheet Global result 
		$SmartSheetProperties = @{
			Colname2ID = @{}
			ID2Colname = @{}
			ID = $Invoked.id
			name = $Invoked.name
			Table =  @() 
			Table_ID2PSo =  @{}
		}
			
		$Invoked.columns | ? {!$_.systemColumnType} | % { 
			$SmartSheetProperties.ID2Colname.add($_.id,$_.title)
			$SmartSheetProperties.Colname2ID.add($_.title,$_.id)
		}
		
		$SmartSheet = new-object psobject -property $SmartSheetProperties
		# Tranforming to Table
		write-verbose "Table : Building the main table"
		$null = ConvertTo-SmartSheetRowObject -JasonRow $Invoked.rows -SS $SmartSheet		
		$SmartSheet = Add-Member -InputObject $SmartSheet -MemberType ScriptMethod -Name 'GenerateLeafTable' -Value $SSGenerateLeafTableBlock -PassThru
		$SmartSheet = Add-Member -InputObject $SmartSheet -MemberType ScriptMethod -Name 'GetServerInfo' -Value $SSGenerateServerInfo -PassThru
		
		# hiding some output : This portion only hide columns when the PsObject is shown
		$SmartSheet.PSObject.TypeNames.Insert(0,"SmartSheet")
		$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet("DefaultDisplayPropertySet",[string[]]('Name','Id','Table'))
		$SmartSheet | Add-Member MemberSet PSStandardMembers ([System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet))
		
		#Outputting result
		return $SmartSheet 
	}
}
Export-ModuleMember -Function Get-Smartsheet

function update-SmartSheetModule {
	copy \\tsclient\c\Users\tfarray.cisco\Documents\WindowsPowerShell\Modules\SmartSheet\SmartSheet.psm1 (Get-Module SmartSheet | select -ExpandProperty path)
}
Export-ModuleMember -Function update-SmartSheetModule