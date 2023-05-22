#----------------------------------------------------------------------------------------------
#  Copyright (c) CompuTec S.A.. All rights reserved.
#  Licensed under the MIT License. See LICENSE.txt in the project root for license information.
#----------------------------------------------------------------------------------------------

using module .\lib\CTLogger.psm1;
using module .\lib\CTProgress.psm1;
using module .\lib\CTTimer.psm1;
Clear-Host
[System.Reflection.Assembly]::LoadWithPartialName("CompuTec.ProcessForce.API")
$FORM_STATE_MAXIMIZED = 'From state: maximized';
$FORM_STATE_REGULAR = 'From state: regular';

$ItemsDictionary = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[psobject]]';
$ResourcesList = New-Object 'System.Collections.Generic.List[string]';
$OperationsList = New-Object 'System.Collections.Generic.List[string]';
$RoutingsList = New-Object 'System.Collections.Generic.List[string]';
$ItemsDictionary.Add('MAKE', (New-Object 'System.Collections.Generic.List[psobject]'));
$ItemsDictionary.Add('BUY', (New-Object 'System.Collections.Generic.List[psobject]'));
$CreatedMors = New-Object 'System.Collections.Generic.List[psobject]';

$connectionXMLFilePath = $PSScriptRoot + "\conf\Connection.xml";

if (!(Test-Path $connectionXMLFilePath -PathType Leaf)) {
	Write-Host -BackgroundColor Red ([string]::Format("File: {0} does not exist.", $connectionXMLFilePath));
	return;
}
[xml] $connectionConfigXml = Get-Content -Encoding UTF8 $connectionXMLFilePath
$xmlConnection = $connectionConfigXml.SelectSingleNode("/CT_CONFIG/Connection");
$xmlTestInfo = $connectionConfigXml.SelectSingleNode("/CT_CONFIG/Test");
$testType = $xmlTestInfo.Type;

if ($testType -eq '3') {
	$testConfigXMLFilePath = $PSScriptRoot + "\conf\TestConfigLong.xml";
}
elseif ($testType -eq '2') {
	$testConfigXMLFilePath = $PSScriptRoot + "\conf\TestConfigMedium.xml";
}
else {
	$testConfigXMLFilePath = $PSScriptRoot + "\conf\TestConfigShort.xml";
}

if (!(Test-Path $testConfigXMLFilePath -PathType Leaf)) {
	Write-Host -BackgroundColor Red ([string]::Format("File: {0} does not exist.", $testConfigXMLFilePath));
	return;
}

[xml] $TestConfigXml = Get-Content -Encoding UTF8 $testConfigXMLFilePath
$MDConfigXml = $TestConfigXml.SelectSingleNode("/CT_CONFIG/MasterData");
$UIConfigXML = $TestConfigXml.SelectSingleNode("/CT_CONFIG/UI");

$date = Get-Date
$RESULT_FOLDER = [string]::Format("{0}\RESULTS_{1}{2}{3}_{4}{5}", $PSScriptRoot, ([string] $date.Year).PadLeft(4, '0'), ([string] $date.Month).PadLeft(2, '0'), ([string] $date.Day).PadLeft(2, '0'), ([string] $date.Hour).PadLeft(2, '0'), ([string] $date.Minute).PadLeft(2, '0') );

if ((Test-Path -Path $RESULT_FOLDER) -eq $false) {
	New-Item -Path $RESULT_FOLDER -ItemType Directory
}

$RESULT_FILE = $RESULT_FOLDER + "\Results_Details.csv";
$RESULT_FILE_CONF = $RESULT_FOLDER + "\Result_Enviroment.csv";

$pfcCompany = $null;

function connectDI() {
	[CTLogger] $logJobs = New-Object CTLogger ('DI', 'Connection', $RESULT_FILE)
	$logJobs.startSubtask('Connection');
	$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::CreateCompany();
	$pfcCompany.Databasename = $xmlConnection.CompanyDB;
	$pfcCompany.DbServerType = [SAPbobsCOM.BoDataServerTypes]::[string]$xmlConnection.DbServerType;
	$pfcCompany.SQLServer = $xmlConnection.DBServer;
	$pfcCompany.SLDAddress = $xmlConnection.SLDServer;
	$pfcCompany.UserName = $xmlConnection.Username;
	$pfcCompany.Password = $xmlConnection.Password;

	write-host -backgroundcolor yellow -foregroundcolor black  "Trying to connect..."
	$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
	write-host -backgroundcolor green -foregroundcolor black "PF API Library:" $version';' 'Host:'(Get-CimInstance Win32_OperatingSystem).CSName';' 'OS Architecture:' (Get-CimInstance Win32_OperatingSystem).OSArchitecture

	try {
		[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', 'dummy')]
		$dummy = $pfcCompany.Connect()

		write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "SAP Business One company version: " $pfcCompany.SapCompany.Version
	}
	catch {
		#Show error messages & stop the script
		write-host "Connection failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
		#$logJobs.endSubtask('Connection', 'F', $_.Exception.Message);
		write-host "CompanyDB" $pfcCompany.Databasename
		write-host "DbServerType:" $pfcCompany.DbServerType
		write-host "DBServer:" $pfcCompany.SQLServer
		write-host "SLDServer:" $pfcCompany.SLDAddress
		write-host "Username:" $pfcCompany.UserName
		exit;
	}

	#If company is not connected - stops the script
	if (-not $pfcCompany.IsConnected) {
		write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected"
		$logJobs.endSubtask('Connection', 'F', 'Company is not connected');
		return;
	}
	$logJobs.endSubtask('Connection', 'S', '');

	return $pfcCompany;
	#endregion
}


function testChoosedDatabaseConfiguration($pfcCompany) {
	$validationPassed = $true;
	function validateWarehouse($pfcCompany, $warehouseCode) {
		# Check warehouses
		try {
			$qm = New-Object CompuTec.Core.DI.Database.QueryManager
			$qm.SimpleTableName = "OWHS";
			$qm.SetSimpleResultFields("WhsCode");
			$qm.SetSimpleWhereFields("WhsCode", "Inactive");
			$rs = $qm.ExecuteSimpleParameters($pfcCompany.Token, $warehouseCode, 'N');
			if ($rs.RecordCount -lt 1) {
				Throw [System.Exception] ([string]::Format("Warehouse with code: {0} does not exists.", $warehouseCode));
			}
		}
		catch {
			$err = [string]::Format("Warehouses {0} validation failed: {1}", $warehouseCode, [string]$_.Exception.Message);
			Throw [System.Exception] ($err);
		}
	}
	function checkIfItemsExists($pfcCompany, $itemPrefix) {
		try {
			$qm = New-Object CompuTec.Core.DI.Database.QueryManager
			$ItemCodePrefix = $itemPrefix + "%";
			$qm.CommandText = "SELECT ""ItemCode"" FROM OITM WHERE ""ItemCode"" LIKE @ItemCodePrefix";
			$qm.AddParameter("ItemCodePrefix", $ItemCodePrefix)
			$rs = $qm.Execute($pfcCompany.Token);
			if ($rs.RecordCount -gt 0) {
				Throw [System.Exception] ("It seems that the script was already run on this company database.");
			}
		}
		catch {
			$err = [string]$_.Exception.Message;
			Throw [System.Exception] ($err);
		}
	}
	function isNumberingSharedSeriesOptionIsOn($pfcCompany) {
		$qm = New-Object CompuTec.Core.DI.Database.QueryManager
			$qm.SimpleTableName = "CINF";
			$qm.SetSimpleResultFields("DocNmMtd");
			$rs = $qm.ExecuteSimpleParameters($pfcCompany.Token);
			if ($rs.RecordCount -gt 0) {
				if($rs.Fields.Item("DocNmMtd").Value -eq 'Y') {
					return $true;
				}
			} else {
				Throw [System.Exception] ("Could not read configuration");
			}
		return $false;
	}
	function validateSeries($pfcCompany, $objectCode, $objectName) {
		try {

			$sharedNumbering = isNumberingSharedSeriesOptionIsOn $pfcCompany;

			$whereValues = New-Object 'System.Collections.Generic.List[string]';
			$whereFields = New-Object 'System.Collections.Generic.List[object]';
			$whereFields.Add("IsManual");
			$whereValues.Add('Y');
			$whereFields.Add("Locked");
			$whereValues.Add('N');
			if($sharedNumbering) {
				$whereFields.Add("SeriesType");
				$whereValues.Add('I');
			} else {
				$whereFields.Add("ObjectCode");
				$whereValues.Add($objectCode);
			}
			$qm = New-Object CompuTec.Core.DI.Database.QueryManager
			$qm.SimpleTableName = "NNM1";
			$qm.SetSimpleResultFields("Series");
			$qm.SetSimpleWhereFields($whereFields);
			$rs = $qm.ExecuteSimpleParameters($pfcCompany.Token, $whereValues.ToArray());
			
			if ($rs.RecordCount -lt 1) {
				Throw [System.Exception] ("Manual numbering series need to be unlocked.");
			}
		}
		catch {
			$err = [string]::Format("Object {0} validation failed: {1}", $objectName, [string]$_.Exception.Message);
			Throw [System.Exception] ($err);
		}
	}
	function checkAutoITW($pfcCompany) {
		try {
			$qm = New-Object CompuTec.Core.DI.Database.QueryManager
			$qm.SimpleTableName = "OADM";
			$qm.SetSimpleResultFields("AutoITW");
			$qm.SetSimpleWhereFields("AutoITW");
			$rs = $qm.ExecuteSimpleParameters($pfcCompany.Token, 'Y');
			if ($rs.RecordCount -gt 0) {
				Throw [System.Exception] ("Auto. Add All Warehouses to New and Existing Items setting needs to be disabled. You can find it in General Settings -> Stock -> Items");
			}
		}
		catch {
			$err = [string]$_.Exception.Message;
			Throw [System.Exception] ($err);
		}
	}


	#region warehouses validation
	try {
		$w1 = $MDConfigXml.SelectNodes("//@WarehouseCode");
		$w2 = $MDConfigXml.SelectNodes("//@ItemsWarehouseCode");
		$warehouses = New-Object 'System.Collections.Generic.List[string]'
		foreach ($whsCode in $w1) {
			if ($warehouses.Contains($whsCode.Value) -eq $false) {
				$warehouses.Add($whsCode.Value);
			}
		}
		foreach ($whsCode in $w2) {
			if ($warehouses.Contains($whsCode.Value) -eq $false) {
				$warehouses.Add($whsCode.Value);
			}
		}

		foreach ($whsCode in $warehouses) {
			validateWarehouse -pfcCompany $pfcCompany -warehouseCode $whsCode
		}
	}
	catch {
		$validationPassed = $false;
		$err = [string]$_.Exception.Message;
		Write-Host -BackgroundColor DarkRed -ForegroundColor White $err;
	}
	#endregion

	#region check if manual numbering possible in case of all objects
	try {
		validateSeries -pfcCompany $pfcCompany -objectCode ([int][SAPbobsCOM.BoObjectTypes]::oItems) -objectName "Item Master Data";

	}
	catch {
		$validationPassed = $false;
		$err = [string]$_.Exception.Message;
		Write-Host -BackgroundColor DarkRed -ForegroundColor White $err
	}


	#endregion


	#region check Auto. Add All Warehouses to New and Existing Items
	try {
		checkAutoITW -pfcCompany $pfcCompany
	}
 catch {
		$validationPassed = $false;
		$err = [string]$_.Exception.Message;
		Write-Host -BackgroundColor DarkRed -ForegroundColor White $err
	}
	#end regions

	#region check if items already exist - this means that the script already run
	$xmlItems = $MDConfigXml.SelectSingleNode([string]::Format("ItemMasterData"));
	$itemPrefix = [string] $xmlItems.Prefix
	try {
		checkIfItemsExists -pfcCompany $pfcCompany -itemPrefix $itemPrefix
	}
 catch {
		$err = [string]$_.Exception.Message;
		Write-Host -BackgroundColor DarkRed -ForegroundColor White $err
		if ($validationPassed -eq $true) {
			Write-Host "Are you sure that you want to continue? (Not Recommended) [y/n]: " -backgroundcolor Yellow -foregroundcolor DarkBlue -NoNewline
			$confirmation = Read-Host
			if (($confirmation -ne 'y') -and ($confirmation -ne 'Y')) {
				$validationPassed = $false;
			}
		}
	}
	#endregion


	return $validationPassed
}
function Imports($pfcCompany) {
	[CTLogger] $logJobs = New-Object CTLogger ('DI', 'Import', $RESULT_FILE)

	#region connection
	$logJobs.startSubtask('Import');
	$sapCompany = $pfcCompany.SapCompany;


	function importIMD($sapCompany) {
		[CTLogger] $logIMD = New-Object CTLogger ('DI', 'Import Item Master Data', $RESULT_FILE)
		#region import of Item Master Data
		write-host ''
		write-host 'Import of Item Master Data: ' -NoNewline;
		$xmlItems = $MDConfigXml.SelectSingleNode([string]::Format("ItemMasterData"));

		$numberOfItems = [int] $xmlItems.NumberOfItems;
		$numberOfMakeItems = [int] $xmlItems.NumberOfMakeItems;
		$itemCodeLength = ([string]$numberOfItems).Length;
		$itemPrefix = [string] $xmlItems.Prefix
		$warehouseCode = [string] $xmlItems.WarehouseCode
		[CTProgress] $progress = New-Object CTProgress ($numberOfItems);
		for ($i = 0; $i -lt $numberOfItems; $i++) {
			try {
				$progress.next();
				$logIMD.startSubtask('Get Item Master Data');
				$sapIMD = $sapCompany.GetBusinessObject([SAPbobsCOM.BoObjectTypes]::oItems);

				$ItemCode = $itemPrefix + ([string]$i).PadLeft($itemCodeLength, '0');

				if ($i -lt $numberOfMakeItems) {
					$ItemsDictionary['MAKE'].Add([psobject]@{
							ItemCode  = $ItemCode
							Revisions = New-Object 'System.Collections.Generic.List[string]';
						});
				}
				else {
					$ItemsDictionary['BUY'].Add([psobject]@{
							ItemCode  = $ItemCode
							Revisions = New-Object 'System.Collections.Generic.List[string]';
						});
				}

				$retValue = $sapIMD.GetByKey($ItemCode)

				if ($retValue -eq $true) {
					$logIMD.endSubtask('Get Item Master Data', 'S', 'Item Already Exists');
					continue;
				}
				$logIMD.endSubtask('Get Item Master Data', 'S', '');
				$logIMD.startSubtask('Add Item Master Data');

				$sapIMD.ItemCode = $ItemCode;
				$sapIMD.ItemName = $ItemCode;

				$sapIMD.WhsInfo.WarehouseCode = $warehouseCode;
				$sapIMD.DefaultWarehouse = $warehouseCode;

				$message = $sapIMD.Add();

				if ($message -lt 0) {
					$err = $sapCompany.GetLastErrorDescription();
					Throw [System.Exception] ($err);
				}
				$logIMD.endSubtask('Add Item Master Data', 'S', '');
			}
			Catch {
				$err = $_.Exception.Message;
				$logIMD.endSubtask('Add Item Master Data', 'F', $err);
				continue;
			}
		}
		#endregion
	}
	function importItemDetails($pfcCompany) {
		[CTLogger] $logPFIMD = New-Object CTLogger ('DI', 'Import Item Details', $RESULT_FILE)
		write-host ''
		write-host 'Import of Item Details: ' -NoNewline;
		$xmlItemDetails = $MDConfigXml.SelectSingleNode([string]::Format("ItemDetails"));
		$numberOfRevisions = [int] $xmlItemDetails.NumberOfRevisions;
		$revisionCodeLength = ([string]$numberOfRevisions).Length;


		[CTProgress] $progress = New-Object CTProgress (($ItemsDictionary['MAKE'].Count + $ItemsDictionary['BUY'].Count));
		foreach ($itemType in $ItemsDictionary.Keys) {
			foreach ($item in $ItemsDictionary[$itemType]) {
				$progress.next()
				$itemCode = $item.ItemCode;
				try {
					$logPFIMD.startSubtask('Get Item Details');
					$itemDetails = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ItemDetails);
					$itemDetailsExists = $itemDetails.GetByItemCode($itemCode);
					$logPFIMD.endSubtask('Get Item Details', 'S', '');
				}
				catch {
					$err = $_.Exception.Message;
					$logPFIMD.endSubtask('Get Item Details', 'F', $err);
					continue;
				}
				try {
					if ($itemDetailsExists) {
						$logPFIMD.startSubtask('Update Item Details');
						# $count = $itemDetails.Revisions.Count
						# for ($i = 0; $i -lt $count; $i++) {
						#     $dummy = $itemDetails.Revisions.DelRowAtPos(0);
						# }
					}
					else {
						$logPFIMD.startSubtask('Add Item Details');
						$itemDetails.U_ItemCode = $itemCode;
					}
					$itemDetails.Revisions.SetCurrentLine($itemDetails.Revisions.Count - 1);

					for ($i = 0; $i -lt $numberOfRevisions; $i++) {
						$revisionCode = 'code' + ([string]$i).PadLeft($revisionCodeLength, '0');
						$item.Revisions.Add($revisionCode);
						if ($itemDetailsExists -eq $false) {
							$itemDetails.Revisions.U_Code = $revisionCode;
							$itemDetails.Revisions.U_Description = $revisionCode;
							$itemDetails.Revisions.U_Status = 1;

							if ($i -eq 0) {
								$itemDetails.Revisions.U_Default = 1;
								$itemDetails.Revisions.U_IsMRPDefault = 1;
								$itemDetails.Revisions.U_IsCostingDefault = 1;
							}
							else {
								$itemDetails.Revisions.U_Default = 2;
								$itemDetails.Revisions.U_IsMRPDefault = 2;
								$itemDetails.Revisions.U_IsCostingDefault = 2;
							}

							$dummy = $itemDetails.Revisions.Add()
						}
					}

					$message = 0

					if ($itemDetailsExists ) {
						$message = $itemDetails.Update()
						RefreshHeader($itemDetails)
					}
					else {
						$message = $itemDetails.Add()
						RefreshHeader($itemDetails)
					}

					if ($message -lt 0) {
						$err = $pfcCompany.GetLastErrorDescription()
						Throw [System.Exception] ($err);
					}
					if ($itemDetailsExists) {
						$logPFIMD.endSubtask('Update Item Details', 'S', '');
					}
					else {
						$logPFIMD.endSubtask('Add Item Details', 'S', '');
					}

				}
				catch {
					$err = $_.Exception.Message;
					if ($itemDetailsExists) {
						$logPFIMD.endSubtask('Update Item Details', 'F', $err);
					}
					else {
						$logPFIMD.endSubtask('Add Item Details', 'F', $err);
					}
					continue;
				}

			}
		}
	}
	function RefreshHeader($udo) {
		try {
			$dummy = $udo.RefreshHeaderData();
		}
		catch {
			$err = $_.Exception.Message;
			$logPFIMD.endSubtask('RefreshHeader', 'F', $err);
		}
	}
	function ImportBOMStructure($pfcCompany) {
		[CTLogger] $log = New-Object CTLogger ('DI', 'Import BOM Structure', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Import of BOM:' -NoNewline;
		$xmlBOM = $MDConfigXml.SelectSingleNode([string]::Format("BOM"));
		$numberOfItems = [int] $xmlBOM.NumberOfItems;
		$numberOfBoms = [int] $xmlBOM.NumberOfBoms;
		$warehouseCode = [string] $xmlBOM.WarehouseCode;
		$itemsWarehouseCode = [string] $xmlBOM.ItemsWarehouseCode;
		[CTProgress] $progress = New-Object CTProgress ($numberOfBoms);
		for ($iBOM = 0; $iBOM -lt $numberOfBoms; $iBOM++) {
			try {
				$progress.next();
				$bomItemCode = $ItemsDictionary['MAKE'][$iBOM].ItemCode;
				$bomRevisionCode = $ItemsDictionary['MAKE'][$iBOM].Revisions[0];
				$log.startSubtask('Get BOM');
				$bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BillOfMaterial);
				$exists = $bom.GetByItemCodeAndRevision($bomItemCode, $bomRevisionCode);
				if ($exists -eq -1) {
					$bomExists = $false;
				}
				else {
					$bomExists = $true;
				}
				$log.endSubtask('Get BOM', 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Get BOM', 'F', $err);
				continue;
			}
			try {
				if ($bomExists) {
					$log.startSubtask('Update BOM');
					$count = $bom.Items.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $bom.Items.DelRowAtPos(0);
					}
				}
				else {
					$log.startSubtask('Add BOM');
					$bom.U_ItemCode = $bomItemCode;
					$bom.U_Revision = $bomRevisionCode;
					$bom.U_WhsCode = $warehouseCode;
				}

				for ($iItems = 0; $iItems -lt $numberOfItems; $iItems++) {
					$itemCode = $ItemsDictionary['BUY'][$iItems].ItemCode;
					$revisionCode = $ItemsDictionary['BUY'][$iItems].Revisions[0];
					#$bom.Items.U_Sequence = ($iItems * 10);
					$bom.Items.U_ItemCode = $itemCode;
					$bom.Items.U_Revision = $revisionCode;
					$bom.Items.U_WhsCode = $itemsWarehouseCode;
					$bom.Items.U_Factor = 1
					$bom.Items.U_Quantity = 1
					$bom.Items.U_ScrapPercentage = 0
					$bom.Items.U_IssueType = 'M'
					$bom.Items.U_SubRecepitItem = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No;
					$dummy = $bom.Items.Add()
				}

				$message = 0

				if ($bomExists ) {
					$message = $bom.Update()
					RefreshHeader($bom)

				}
				else {
					$message = $bom.Add()
					RefreshHeader($bom)

				}

				if ($message -lt 0) {
					$err = $pfcCompany.GetLastErrorDescription()
					Throw [System.Exception] ($err);
				}
				if ($bomExists) {
					$log.endSubtask('Update BOM', 'S', '');
				}
				else {
					$log.endSubtask('Add BOM', 'S', '');
				}

			}
			catch {
				$err = $_.Exception.Message;
				if ($bomExists) {
					$log.endSubtask('Update BOM', 'F', $err);
				}
				else {
					$log.endSubtask('Add BOM', 'F', $err);
				}
				continue;
			}

		}
	}
	function ImportResources($pfcCompany) {
		[CTLogger] $log = New-Object CTLogger ('DI', 'Import Resources', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Import of Resources:' -NoNewline
		$xmlResource = $MDConfigXml.SelectSingleNode([string]::Format("Resource"));
		$numberOfResources = [int] $xmlResource.NumberOfResources;
		$resourceCodeLength = ([string]$numberOfResources).Length;
		$resourcePrefix = [string] $xmlResource.Prefix
		[CTProgress] $progress = New-Object CTProgress ($numberOfResources);
		for ($i = 0; $i -lt $numberOfResources; $i++) {
			try {
				$progress.next();
				$ResourceCode = $resourcePrefix + ([string]$i).PadLeft($resourceCodeLength, '0');
				$ResourcesList.Add($ResourceCode);

				$log.startSubtask('Get Resource');
				$resource = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Resource);
				$exists = $resource.GetByRscCode($ResourceCode);
				if ($exists -eq -1) {
					$resourceExists = $false;
				}
				else {
					$resourceExists = $true;
				}
				$log.endSubtask('Get Resource', 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Get Resource', 'F', $err);
				continue;
			}
			try {
				if ($resourceExists) {
					$log.startSubtask('Update Resource');
				}
				else {
					$log.startSubtask('Add Resource');
					$resource.U_RscType = 1;
					$resource.U_RscCode = $ResourceCode
					$resource.U_RscName = $ResourceCode
				}

				$message = 0

				if ($resourceExists ) {
					$message = $resource.Update()
					RefreshHeader($resource)

				}
				else {
					$message = $resource.Add()
					RefreshHeader($resource)

				}

				if ($message -lt 0) {
					$err = $pfcCompany.GetLastErrorDescription()
					Throw [System.Exception] ($err);
				}
				if ($resourceExists) {
					$log.endSubtask('Update Resource', 'S', '');
				}
				else {
					$log.endSubtask('Add Resource', 'S', '');
				}

			}
			catch {
				$err = $_.Exception.Message;
				if ($resourceExists) {
					$log.endSubtask('Update Resource', 'F', $err);
				}
				else {
					$log.endSubtask('Add Resource', 'F', $err);
				}
				continue;
			}

		}
	}
	function ImportOperations($pfcCompany) {
		[CTLogger] $log = New-Object CTLogger ('DI', 'Import Operations', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Import of Operations:' -NoNewline
		$xmlOperation = $MDConfigXml.SelectSingleNode([string]::Format("Operation"));
		$numberOfOperations = [int] $xmlOperation.NumberOfOperations;
		$numberOfResources = [int] $xmlOperation.NumberOfResources;
		$operationCodeLength = ([string]$numberOfOperations).Length;
		$operationPrefix = [string] $xmlOperation.Prefix
		[CTProgress] $progress = New-Object CTProgress ($numberOfOperations);
		for ($iOperation = 0; $iOperation -lt $numberOfOperations; $iOperation++) {
			try {
				$progress.next();
				$operationCode = $operationPrefix + ([string]$iOperation).PadLeft($operationCodeLength, '0');
				$OperationsList.Add($operationCode);

				$log.startSubtask('Get Operation');
				$operation = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Operation);
				$exists = $operation.GetByOprCode($operationCode);
				if ($exists -eq -1) {
					$operationExists = $false;
				}
				else {
					$operationExists = $true;
				}
				$log.endSubtask('Get Operation', 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Get Operation', 'F', $err);
				continue;
			}
			try {
				if ($operationExists) {
					$log.startSubtask('Update Operation');
					$count = $operation.OperationResources.Count - 1
					for ($i = $count - 1; $i -ge 0; $i--) {
						$dummy = $operation.OperationResources.DelRowAtPos(0);
					}
				}
				else {
					$log.startSubtask('Add Operation');
					$operation.U_OprCode = $operationCode;
					$operation.U_OprName = $operationCode;
				}


				for ($iResources = 0; $iResources -lt $numberOfResources; $iResources++) {
					$ResourceCode = $ResourcesList[$iResources];

					$operation.OperationResources.U_RscCode = $ResourceCode
					if ($iResources -eq 0) {
						$operation.OperationResources.U_IsDefault = 'Y';
					}
					else {
						$operation.OperationResources.U_IsDefault = 'N';
					}

					$operation.OperationResources.U_HasCycles = [CompuTec.ProcessForce.API.Enumerators.YesNoType]::No

					$dummy = $operation.OperationResources.Add()
				}


				$message = 0

				if ($operationExists ) {
					$message = $operation.Update()
					RefreshHeader($operation)

				}
				else {
					$message = $operation.Add()
					RefreshHeader($operation)

				}

				if ($message -lt 0) {
					$err = $pfcCompany.GetLastErrorDescription()
					Throw [System.Exception] ($err);
				}
				if ($operationExists) {
					$log.endSubtask('Update Operation', 'S', '');
				}
				else {
					$log.endSubtask('Add Operation', 'S', '');
				}

			}
			catch {
				$err = $_.Exception.Message;
				if ($operationExists) {
					$log.endSubtask('Update Operation', 'F', $err);
				}
				else {
					$log.endSubtask('Add Operation', 'F', $err);
				}
				continue;
			}

		}
	}
	function ImportRoutings($pfcCompany) {
		[CTLogger] $log = New-Object CTLogger ('DI', 'Import Routings', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Import of Routings:' -NoNewline
		$xmlRouting = $MDConfigXml.SelectSingleNode([string]::Format("Routing"));
		$numberOfOperations = [int] $xmlRouting.NumberOfOperations;
		$numberOfRoutings = [int] $xmlRouting.NumberOfRoutings;
		$routingCodeLength = ([string]$numberOfRoutings).Length;
		$routingPrefix = [string] $xmlRouting.Prefix
		[CTProgress] $progress = New-Object CTProgress ($numberOfRoutings);
		for ($i = 0; $i -lt $numberOfRoutings; $i++) {
			try {
				$progress.next();
				$routingCode = $routingPrefix + ([string]$i).PadLeft($routingCodeLength, '0');
				$RoutingsList.Add($routingCode);

				$log.startSubtask('Get Routing');
				$routing = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::Routing);
				$exists = $routing.GetByRtgCode($routingCode);
				if ($exists -eq -1) {
					$routingExists = $false;
				}
				else {
					$routingExists = $true;
				}
				$log.endSubtask('Get Routing', 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Get Routing', 'F', $err);
				continue;
			}
			try {
				if ($routingExists) {
					$log.startSubtask('Update Routing');
					$count = $routing.Operations.Count
					for ($j = 0; $j -lt $count; $j++) {
						$dummy = $routing.Operations.DelRowAtPos(0);
					}
					$count = $routing.OperationResources.Count
					for ($k = 0; $k -lt $count; $k++) {
						$dummy = $routing.OperationResources.DelRowAtPos(0);
					}
				}
				else {
					$log.startSubtask('Add Routing');
					$routing.U_RtgCode = $routingCode
					$routing.U_RtgName = $routingCode
					$routing.U_Active = 1
				}


				for ($iOperations = 0; $iOperations -lt $numberOfOperations; $iOperations++) {
					$OperationCode = $OperationsList[$iOperations];

					$routing.Operations.U_OprCode = $OperationCode
					$dummy = $routing.Operations.Add()
				}

				$message = 0

				if ($routingExists ) {
					$message = $routing.Update()
					RefreshHeader($routing)
				}
				else {
					$message = $routing.Add()
					RefreshHeader($routing)
				}

				if ($message -lt 0) {
					$err = $pfcCompany.GetLastErrorDescription()
					Throw [System.Exception] ($err);
				}
				if ($routingExists) {
					$log.endSubtask('Update Routing', 'S', '');
				}
				else {
					$log.endSubtask('Add Routing', 'S', '');
				}

			}
			catch {
				$err = $_.Exception.Message;
				if ($routingExists) {
					$log.endSubtask('Update Routing', 'F', $err);
				}
				else {
					$log.endSubtask('Add Routing', 'F', $err);
				}
				continue;
			}

		}
	}
	function ImportProductionProcesses($pfcCompany) {
		[CTLogger] $log = New-Object CTLogger ('DI', 'Import Production Processes', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Import of Production Processes:' -NoNewline;
		$xmlProductionProcess = $MDConfigXml.SelectSingleNode([string]::Format("ProductionProcess"));
		$numberOfBoms = [int] $xmlProductionProcess.NumberOfBoms;
		$numberOfRoutings = [int] $xmlProductionProcess.NumberOfRoutings;
		[CTProgress] $progress = New-Object CTProgress ($numberOfBoms);
		for ($iBOM = 0; $iBOM -lt $numberOfBoms; $iBOM++) {
			try {
				$progress.next();
				$bomItemCode = $ItemsDictionary['MAKE'][$iBOM].ItemCode;
				$bomRevisionCode = $ItemsDictionary['MAKE'][$iBOM].Revisions[0];
				$log.startSubtask('Get Production Process');
				$bom = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::BillOfMaterial);
				$exists = $bom.GetByItemCodeAndRevision($bomItemCode, $bomRevisionCode);
				if ($exists -eq -1) {
					$bomExists = $false;
				}
				else {
					$bomExists = $true;
				}
				$log.endSubtask('Get Production Process', 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Get Production Process', 'F', $err);
				continue;
			}
			try {
				if ($bomExists) {
					$log.startSubtask('Update Production Process');
					$count = $bom.Routings.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $bom.Routings.DelRowAtPos(0);
					}

					$count = $bom.RoutingOperations.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $bom.RoutingOperations.DelRowAtPos(0);
					}

					$count = $bom.RoutingOperationResources.Count
					for ($i = 0; $i -lt $count; $i++) {
						$dummy = $bom.RoutingOperationResources.DelRowAtPos(0);
					}
				}
				else {
					$log.startSubtask('Add Production Process');
					$bom.U_ItemCode = $bomItemCode;
					$bom.U_Revision = $bomRevisionCode;
				}

				for ($iRoutings = 0; $iRoutings -lt $numberOfRoutings; $iRoutings++) {
					$routingCode = $RoutingsList[$iRoutings];
					$bom.Routings.U_RtgCode = $routingCode;
					if ($iRoutings -eq 0) {
						$bom.Routings.U_IsDefault = 'Y';
						$bom.Routings.U_IsRollUpDefault = 'Y'
					}
					else {
						$bom.Routings.U_IsDefault = 'N'
						$bom.Routings.U_IsRollUpDefault = 'N'
					}
					#  $bom.RoutingOperationResources
					$dummy = $bom.Routings.Add()
				}

				$message = 0

				if ($bomExists ) {
					$message = $bom.Update()
					RefreshHeader($bom)

				}
				else {
					$message = $bom.Add()
					RefreshHeader($bom)

				}

				if ($message -lt 0) {
					$err = $pfcCompany.GetLastErrorDescription()
					Throw [System.Exception] ($err);
				}
				if ($bomExists) {
					$log.endSubtask('Update Production Process', 'S', '');
				}
				else {
					$log.endSubtask('Add Production Process', 'S', '');
				}

			}
			catch {
				$err = $_.Exception.Message;
				if ($bomExists) {
					$log.endSubtask('Update Production Process', 'F', $err);
				}
				else {
					$log.endSubtask('Add Production Process', 'F', $err);
				}
				continue;
			}

		}
	}
	function CreateManufacturingOrders($pfcCompany) {
		[CTLogger] $log = New-Object CTLogger ('DI', 'Add Manufacturing Orders', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Adding Manufacturing Orders:' -NoNewline;
		$xmlProductionProcess = $MDConfigXml.SelectSingleNode([string]::Format("MOR"));
		$numberOfMors = [int] $xmlProductionProcess.NumberOfMors;
		[CTProgress] $progress = New-Object CTProgress ($numberOfMors);
		for ($iMOR = 0; $iMOR -lt $numberOfMors; $iMOR++) {
			try {
				$progress.next();
				$bomItemCode = $ItemsDictionary['MAKE'][$iMOR].ItemCode;
				$bomRevisionCode = $ItemsDictionary['MAKE'][$iMOR].Revisions[0];
				$log.startSubtask('Get MOR');
				$mor = $pfcCompany.CreatePFObject([CompuTec.ProcessForce.API.Core.ObjectTypes]::ManufacturingOrder);
				$log.endSubtask('Get MOR', 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Get MOR', 'F', $err);
				continue;
			}
			try {
				
				$log.startSubtask('Add MOR');
				$mor.U_ItemCode = $bomItemCode;
				$mor.U_Revision = $bomRevisionCode;
				
				$message = 0
				$message = $mor.Add()
			
				if ($message -lt 0) {  
					$err = $pfcCompany.GetLastErrorDescription()
					Throw [System.Exception] ($err);
				}
				RefreshHeader $mor;
				$log.endSubtask('Add MOR', 'S', '');
				$CreatedMors.Add([psobject]@{
						Series = $mor.Series
						DocNum = $mor.DocNum
					});
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Add MOR', 'F', $err);
				continue;
			}

		}
	}

	$logJobs.startSubtask('Import IMD');
	importIMD $sapCompany
	$logJobs.endSubtask('Import IMD', 'S', '');

	$logJobs.startSubtask('Import Item Details');
	importItemDetails $pfcCompany
	$logJobs.endSubtask('Import Item Details', 'S', '');

	#restore Item Costing

	$logJobs.startSubtask('Import BOM Structure');
	ImportBOMStructure $pfcCompany
	$logJobs.endSubtask('Import BOM Structure', 'S', '');

	$logJobs.startSubtask('Import Resources');
	ImportResources $pfcCompany
	$logJobs.endSubtask('Import Resources', 'S', '');

	$logJobs.startSubtask('Import Operations');
	ImportOperations $pfcCompany
	$logJobs.endSubtask('Import Operations', 'S', '');

	$logJobs.startSubtask('Import Routings');
	ImportRoutings $pfcCompany
	$logJobs.endSubtask('Import Routings', 'S', '');

	$logJobs.startSubtask('Import Production Processes');
	ImportProductionProcesses $pfcCompany
	$logJobs.endSubtask('Import Production Processes', 'S', '');

	$logJobs.startSubtask('Import MOR');
	CreateManufacturingOrders $pfcCompany
	$logJobs.endSubtask('Import MOR', 'S', '');

	$logJobs.endSubtask('Import', 'S', '');

	$pfcCompany.Disconnect();
}

function UITests() {
	[CTLogger] $logJobs = New-Object CTLogger ('UI', 'GET', $RESULT_FILE)

	#region connection
	$logJobs.startSubtask('Get');
	$logJobs.startSubtask('Connection');
	Write-Host -BackgroundColor Blue 'Connecting...'

	$app = $null;

	#region connection
	write-host -backgroundcolor yellow -foregroundcolor black  "Trying to connect..."
	$version = [CompuTec.Core.CoreConfiguration+DatabaseSetup]::AddonVersion
	write-host -backgroundcolor green -foregroundcolor black "ProcessForce API library:" $version';' 'Host:'(Get-CimInstance Win32_OperatingSystem).CSName';' 'OS Architecture:' (Get-CimInstance Win32_OperatingSystem).OSArchitecture

	try {
		$pfcCompany = [CompuTec.ProcessForce.API.ProcessForceCompanyInitializator]::ConnectUI([ref] $app, $true)
		write-host -backgroundcolor green -foregroundcolor black "Connected to:" $pfcCompany.SapCompany.CompanyName "/ " $pfcCompany.SapCompany.CompanyDB"" "SAP Business One Company version: " $pfcCompany.SapCompany.Version
	}
	catch {
		#Show error messages & stop the script
		write-host "Connection Failure: " -backgroundcolor red -foregroundcolor white $_.Exception.Message
		#$logJobs.endSubtask('Connection', 'F', $_.Exception.Message);
		write-host "CompanyDB" $pfcCompany.Databasename
		write-host "DbServerType:" $pfcCompany.DbServerType
		write-host "DBServer:" $pfcCompany.SQLServer
		write-host "SLDServer:" $pfcCompany.SLDAddress
		write-host "Username:" $pfcCompany.UserName
	}

	#If company is not connected - stops the script
	if ($pfcCompany.SapCompany.Connected -ne 1) {
		write-host -backgroundcolor yellow -foregroundcolor black "Company is not connected."
		$logJobs.endSubtask('Connection', 'F', 'Company is not connected.');
		return
	}
	#If company is connected to a wrong company database - stops the script
	if ($pfcCompany.SapCompany.CompanyDB -ne $xmlConnection.CompanyDB) {
		write-host -backgroundcolor yellow -foregroundcolor black "Company is connected to a wrong company database.";
		$logJobs.endSubtask('Connection', 'F', 'Company is connected to wrong company database.');
		return;
	}
	#endregion
	$logJobs.endSubtask('Connection', 'S', '');

	function setStateOfForm($app, $menuItemName, $maximized = $false) {
		$formOpenMenu = $app.Menus.Item($menuItemName);
		$formOpenMenu.Activate();
		$form = $app.Forms.ActiveForm();
		if ($maximized) {
			$form.State = [SAPbouiCOM.BoFormStateEnum]::fs_Maximized;
		}
		else {
			$form.State = [SAPbouiCOM.BoFormStateEnum]::fs_Restore;
		}
		$form.Close();
	}


	function openItemDetailsForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open Item Details', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open Item Details:' -NoNewline;
		$xmlOpenItemDetails = $UIConfigXml.SelectSingleNode([string]::Format("ItemDetails"));
		$repeatOpenForm = [int] $xmlOpenItemDetails.repeatOpenForm;

		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open Item Details Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_1');
				$formOpenMenu.Activate();
				$log.endSubtask('Open Item Details Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open Item Details Form', 'F', $err);
				continue;
			}
		}
	}
	<#
		$firstEntityKeyValues  = New-Object 'System.Collections.Generic.Dictionary[string,string]'; $firstEntityKeyValues.Add("keyName","valueName");
	#>
	function _loadForm($app, $stateName, $menuItemName, $repeatOpenForm, $firstEntityKeyValues, $log, $subtaskNameOpen, $subtaskNameLoad) {
		$next = $app.Menus.Item('1288');
		$find = $app.Menus.Item('1281');
		[CTProgress] $progress = New-Object CTProgress ($recordsToGoThrough);
		$subtaskNameLoad = "Item Details Load Data";
		try {
			$log.startSubtask($subtaskNameOpen, $stateName);
			$formOpenMenu = $app.Menus.Item($menuItemName);
			$formOpenMenu.Activate();
			if ($find.Enabled -eq $true) {
				$find.Activate();
			}
			$form = $app.Forms.ActiveForm
			$log.endSubtask($subtaskNameOpen, 'S', '');
		}
		catch {
			$err = $_.Exception.Message;
			$log.endSubtask($subtaskNameOpen, 'F', $err);
			continue;
		}
		for ($iRecord = 0; $iRecord -lt $recordsToGoThrough; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask($subtaskNameLoad, $stateName);
				if ($iRecord -eq 0) {
					foreach ($entityKey in $firstEntityKeyValues.Keys) {
						$InputField = $form.Items($entityKey);

						if ($InputField.Type -eq 16) {
							$InputField.Specific.String = [string] $firstEntityKeyValues[$entityKey];
						}
						if ($InputField.Type -eq 113) {
							$InputField.Specific.Select([string] $firstEntityKeyValues[$entityKey]);
						}
						$app.SendKeys('{TAB}');

					}
					$app.SendKeys('{ENTER}');
				}
				else {
					$next.Activate();
				}
				$log.endSubtask($subtaskNameLoad, 'S', '');
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask($subtaskNameLoad, 'F', $err);
				continue;
			}
		}
		$form.Close();
	}
	function loadItemDetails ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Open Item Details', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load Item Details:';
			$xmlOpenItemDetails = $UIConfigXml.SelectSingleNode([string]::Format("ItemDetails"));
			$recordsToGoThrough = [int] $xmlOpenItemDetails.recordsToGoThrough;
			$firstItemCode = $ItemsDictionary['MAKE'][0].ItemCode;
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			$firstEntityKeyValues.Add("idtItCTbx", $firstItemCode);
			$menuItemName = 'CT_PF_1';
			$subtaskNameOpen = "Open Item Details Form";
			$subtaskNameLoad = "Item Details Load Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true;
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;

			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false;
			#_loadItemDetails -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -firstItemCode $firstItemCode -log $log
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load Item Details:{0}', [string]$err);
			Write-Host $msg;
		}
	}

	function openBOMForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open BOM', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open BOM:' -NoNewline;
		$xmlOpenBOM = $UIConfigXml.SelectSingleNode([string]::Format("BOM"));
		$repeatOpenForm = [int] $xmlOpenBOM.repeatOpenForm;

		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open BOM Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_2');
				$formOpenMenu.Activate();
				$log.endSubtask('Open BOM Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open BOM Form', 'F', $err);
				continue;
			}
		}
	}

	function loadBOM ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Load BOMs', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load BOM:'
			$xmlOpenBOM = $UIConfigXml.SelectSingleNode([string]::Format("BOM"));
			$recordsToGoThrough = [int] $xmlOpenBOM.recordsToGoThrough;
			$firstItemCode = $ItemsDictionary['MAKE'][0].ItemCode;
			$firstRevision = $ItemsDictionary['MAKE'][0].Revisions[0]
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			$firstEntityKeyValues.Add("7", $firstItemCode);
			$firstEntityKeyValues.Add("13", $firstRevision);
			$menuItemName = 'CT_PF_2';
			$subtaskNameOpen = "Open BOM Form";
			$subtaskNameLoad = "Load BOM Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true;
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;

			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false;
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load BOM:{0}', [string]$err);
			Write-Host $msg;
		}
	}

	function openProductionProcessForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open Production Process', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open Production Process:' -NoNewline;
		$xmlOpenProductionProcess = $UIConfigXml.SelectSingleNode([string]::Format("ProductionProcess"));
		$repeatOpenForm = [int] $xmlOpenProductionProcess.repeatOpenForm;

		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open Production Process Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_81');
				$formOpenMenu.Activate();
				$log.endSubtask('Open Production Process Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open Production Process Form', 'F', $err);
				continue;
			}
		}
	}

	function loadProductionProcess ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Load Production Processes', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load Production Processes:';
			$xmlOpenProductionProcess = $UIConfigXml.SelectSingleNode([string]::Format("ProductionProcess"));
			$recordsToGoThrough = [int] $xmlOpenProductionProcess.recordsToGoThrough;
			$firstItemCode = $ItemsDictionary['MAKE'][0].ItemCode;
			$firstRevision = $ItemsDictionary['MAKE'][0].Revisions[0]
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			$firstEntityKeyValues.Add("7", $firstItemCode);
			$firstEntityKeyValues.Add("RevNameTbx", $firstRevision);
			$menuItemName = 'CT_PF_81';
			$subtaskNameOpen = "Open Production Process Form";
			$subtaskNameLoad = "Load Production Process Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true;
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;

			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false;
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load Production Process:{0}', [string]$err);
			Write-Host $msg;
		}
	}

	function openResourceForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open Resources', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open Resources:' -NoNewline;
		$xmlOpenResource = $UIConfigXml.SelectSingleNode([string]::Format("Resource"));
		$repeatOpenForm = [int] $xmlOpenResource.repeatOpenForm;

		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open Resource Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_12');
				$formOpenMenu.Activate();
				$log.endSubtask('Open Resource Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open Resource Form', 'F', $err);
				continue;
			}
		}
	}

	function loadResources ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Load Resources', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load Resources:';
			$xmlOpenResource = $UIConfigXml.SelectSingleNode([string]::Format("Resource"));
			$recordsToGoThrough = [int] $xmlOpenResource.recordsToGoThrough;
			$firstResourceCode = $ResourcesList[0]
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			$firstEntityKeyValues.Add("rscCodBox", $firstResourceCode);
			$menuItemName = 'CT_PF_12';
			$subtaskNameOpen = "Open Resource Form";
			$subtaskNameLoad = "Load Resource Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true;
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;

			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false;
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load Resources:{0}', [string]$err);
			Write-Host $msg;
		}
	}

	function openOperationForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open Operations', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open Operations:' -NoNewline;
		$xmlOpenOperation = $UIConfigXml.SelectSingleNode([string]::Format("Operation"));
		$repeatOpenForm = [int] $xmlOpenOperation.repeatOpenForm;

		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open Operation Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_14');
				$formOpenMenu.Activate();
				$log.endSubtask('Open Operation Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open Operation Form', 'F', $err);
				continue;
			}
		}
	}

	function loadOperations ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Load Operations', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load Operations:';
			$xmlOpenOperation = $UIConfigXml.SelectSingleNode([string]::Format("Operation"));
			$recordsToGoThrough = [int] $xmlOpenOperation.recordsToGoThrough;
			$firstOprationCode = $OperationsList[0]
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			$firstEntityKeyValues.Add("oprCodBox", $firstOprationCode);
			$menuItemName = 'CT_PF_14';
			$subtaskNameOpen = "Open Operation Form";
			$subtaskNameLoad = "Load Operation Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true;
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;

			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false;
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load Production Process:{0}', [string]$err);
			Write-Host $msg;
		}

	}

	function openRoutingForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open Routings', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open Routings:' -NoNewline;
		$xmlOpenRouting = $UIConfigXml.SelectSingleNode([string]::Format("Routing"));
		$repeatOpenForm = [int] $xmlOpenRouting.repeatOpenForm;

		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open Routing Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_13');
				$formOpenMenu.Activate();
				$log.endSubtask('Open Routing Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open Routing Form', 'F', $err);
				continue;
			}
		}
	}

	function loadRoutings ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Load Routings', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load Routings:';
			$xmlOpenRouting = $UIConfigXml.SelectSingleNode([string]::Format("Routing"));
			$recordsToGoThrough = [int] $xmlOpenRouting.recordsToGoThrough;
			$firstRoutingCode = $RoutingsList[0]
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]';
			$firstEntityKeyValues.Add("rtgCodBox", $firstRoutingCode);
			$menuItemName = 'CT_PF_13';
			$subtaskNameOpen = "Open Routing Form";
			$subtaskNameLoad = "Load Routing Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true;
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;

			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false;
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load Production Process:{0}', [string]$err);
			Write-Host $msg;
		}
	}
	function openMORForm ($app) {
		[CTLogger] $log = New-Object CTLogger ('UI', 'Open MOR', $RESULT_FILE)
		Write-Host '';
		Write-Host 'Open MOR:' -NoNewline;
		$xmlOpenRouting = $UIConfigXml.SelectSingleNode([string]::Format("MOR"));
		$repeatOpenForm = [int] $xmlOpenRouting.repeatOpenForm;
			
		[CTProgress] $progress = New-Object CTProgress ($repeatOpenForm);
		for ($iRecord = 0; $iRecord -lt $repeatOpenForm; $iRecord++) {
			try {
				$progress.next();
				$log.startSubtask('Open MOR Form');
				$formOpenMenu = $app.Menus.Item('CT_PF_6'); 
				$formOpenMenu.Activate();
				$log.endSubtask('Open MOR Form', 'S', '');
				$form = $app.Forms.ActiveForm()
				$form.Close();
			}
			catch {
				$err = $_.Exception.Message;
				$log.endSubtask('Open MOR Form', 'F', $err);
				continue;
			}
		}
	}

	function loadMORs ($app) {
		try {
			[CTLogger] $log = New-Object CTLogger ('UI', 'Load MOR', $RESULT_FILE)
			Write-Host '';
			Write-Host 'Load Routings:';
			$xmlOpenRouting = $UIConfigXml.SelectSingleNode([string]::Format("MOR"));
			$recordsToGoThrough = [int] $xmlOpenRouting.recordsToGoThrough;
			$firstMOR = $CreatedMors[0];
			$firstMORSeries = $firstMOR.Series;
			$firstMORDocNum = $firstMOR.DocNum;
			$firstEntityKeyValues = New-Object 'System.Collections.Generic.Dictionary[string,string]'; 
			$firstEntityKeyValues.Add("Series", $firstMORSeries);
			$firstEntityKeyValues.Add("5", $firstMORDocNum);
			$menuItemName = 'CT_PF_6';
			$subtaskNameOpen = "Open MOR Form";
			$subtaskNameLoad = "Load MOR Data";
			Write-Host '* Maximized:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $true; 
			_loadForm -app $app -stateName $FORM_STATE_MAXIMIZED -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
	
			Write-Host '';
			Write-Host '* Regular:' -NoNewline;
			setStateOfForm -app $app -menuItemName $menuItemName -maximized $false; 
			_loadForm -app $app -stateName $FORM_STATE_REGULAR -repeatOpenForm $recordsToGoThrough -menuItemName $menuItemName -log $log -subtaskNameOpen $subtaskNameOpen -subtaskNameLoad $subtaskNameLoad -firstEntityKeyValues $firstEntityKeyValues;
		}
		catch {
			$err = $_.Exception.Message;
			$msg = [string]::Format('Unexpected exception while running Load MOR:{0}', [string]$err);
			Write-Host $msg;
		}
	}

	
	openItemDetailsForm $app;

	loadItemDetails $app;

	openBOMForm $app;

	loadBOM $app;

	openProductionProcessForm $app;

	loadProductionProcess $app;

	openResourceForm $app;

	loadResources $app;

	openOperationForm $app;

	loadOperations $app;

	openRoutingForm $app;

	loadRoutings $app;

	openMORForm $app;

	loadMORs $app;

	$logJobs.endSubtask('Get', 'S', '');
}

function saveTestConfiguration() {

	[CTProgress] $progress = New-Object CTProgress (10);
	Write-Host 'Checking enviroment:' -NoNewline

	Add-Content -path $RESULT_FILE_CONF ([string]::Format("Test started at: {0}", (Get-Date)));

	Add-Content -Path $RESULT_FILE_CONF '';
	$os = Get-Ciminstance Win32_OperatingSystem;
	Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Total Memory: {0} GB", [int]($os.TotalVisibleMemorySize / 1mb)) );
	Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Free Memory: {0} GB", [math]::Round($os.FreePhysicalMemory / 1mb, 2)) );
	$progress.next();

	Add-Content -Path $RESULT_FILE_CONF '';
	$processor = Get-CimInstance win32_processor
	Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Processor: {0}", $processor.Name) );
	Add-Content -Path $RESULT_FILE_CONF ( [string]::Format("Processor average usage: {0} %", ($processor | Measure-Object -property LoadPercentage -Average | Select-Object Average).Average) );

	Add-Content -Path $RESULT_FILE_CONF '';
	$progress.next();

	# TestConfig.xml
	Add-Content -Path $RESULT_FILE_CONF 'TestConfig.xml:'
	Add-Content -path $RESULT_FILE_CONF $TestConfigXml.InnerXml;
	Add-Content -Path $RESULT_FILE_CONF ''
	$progress.next();

	# Connection.xml
	Add-Content -Path $RESULT_FILE_CONF 'Connection.xml:'
	Add-Content -path $RESULT_FILE_CONF $connectionConfigXml.InnerXml;
	Add-Content -Path $RESULT_FILE_CONF ''
	$dbServer = ($xmlConnection.DBServer).Split(':')[0]
	# Extracting SAP HANA server name in case of SAP HANA 2.0 multitenat prefix e.g., NDB
	if ($dbServer.Contains("@")) {
		$dbServer = $dbServer.Split('@')[1]
	}
	$sldServer = ($xmlConnection.SLDServer).Split(':')[0]
	Add-Content -Path $RESULT_FILE_CONF '';
	$progress.next();

	$pingToDbServer = Test-Connection $dbServer -Count 20
	$progress.next();
	$pingToDbServer += Test-Connection $dbServer -Count 20
	$progress.next();
	$pingToDbServer += Test-Connection $dbServer -Count 20

	Add-Content -Path $RESULT_FILE_CONF 'Ping database server:'
	foreach ($pingResponse in  $pingToDbServer) {
		Add-Content -Path $RESULT_FILE_CONF ([string]::Format("Source:{0}, Destination:{1}, IPV4Address:{2}, IPV6Address{3}, ResponseTime: {4}",
				$pingResponse.PSComputerName, $pingResponse.Address , $pingResponse.IPV4Address, $pingResponse.IPV6Address, $pingResponse.ResponseTime ))
	}
	Add-Content -Path $RESULT_FILE_CONF '';
	$progress.next();

	$pingToSLDServer = Test-Connection $sldServer -Count 20
	$progress.next();
	$pingToSLDServer += Test-Connection $sldServer -Count 20
	$progress.next();
	$pingToSLDServer += Test-Connection $sldServer -Count 20
	Add-Content -Path $RESULT_FILE_CONF 'Ping SLD server:'
	foreach ($pingResponse in  $pingToSLDServer) {
		Add-Content -Path $RESULT_FILE_CONF ([string]::Format("Source:{0}, Destination:{1}, IPV4Address:{2}, IPV6Address{3}, ResponseTime: {4}",
				$pingResponse.PSComputerName, $pingResponse.Address , $pingResponse.IPV4Address, $pingResponse.IPV6Address, $pingResponse.ResponseTime ))
	}
	Add-Content -Path $RESULT_FILE_CONF '';
	$progress.next();
}

function logToConsole($task, $timer) {
	if($null -eq $timer) {
		$timer = New-Object CTTimer;
		Write-Host $task;
		return $timer;
	} else {
		$msg = [string]::Format('{0} took: {1} s', $task, $timer.totalSeconds());
		Write-Host $msg;
	}
}
$timer = logToConsole 'Connecting to database';
$pfcCompany = connectDI;
logToConsole 'Connecting to database' $timer;

$timer = logToConsole 'Testing database configuration';
$configurationTest = testChoosedDatabaseConfiguration -pfcCompany $pfcCompany;
logToConsole 'Testing database configuration' $timer;

if ($configurationTest -eq $true) {
	$timer = logToConsole 'Testing enviroment';
	saveTestConfiguration ;
	logToConsole 'Testing enviroment' $timer;
	$timer = logToConsole 'Imporitng data';
	Imports -pfcCompany $pfcCompany ;
	logToConsole 'Imporitng data' $timer;
	$timer = logToConsole 'UI Tests';
	UITests ;
	logToConsole 'UI Tests' $timer;
}
else {
	Write-Host -BackgroundColor DarkRed -ForegroundColor White "Configuration test failed. To perform Performance Test please fix your configuration."
}