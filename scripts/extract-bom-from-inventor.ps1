# PowerShell script to extract BOM data from Inventor .idw file
# This script uses COM interop to access Inventor's API
# Returns JSON data with BOM information

param(
    [Parameter(Mandatory=$true)]
    [string]$IdwFilePath
)

# Convert path to absolute if needed
$IdwFilePath = [System.IO.Path]::GetFullPath($IdwFilePath)

if (-not (Test-Path $IdwFilePath)) {
    Write-Error "File not found: $IdwFilePath"
    exit 1
}

try {
    # Create Inventor Application object
    $inventorApp = New-Object -ComObject "Inventor.Application"
    $inventorApp.Visible = $false
    
    # Open the drawing document
    $drawingDoc = $inventorApp.Documents.Open($IdwFilePath, $false)
    
    # Extract job number from document properties
    $jobNumber = "UNKNOWN"
    try {
        $fullPartNumber = $drawingDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value
        if ($fullPartNumber -and $fullPartNumber.Contains(" ")) {
            $jobNumber = $fullPartNumber.Split(" ")[0]
        } else {
            $jobNumber = $fullPartNumber
        }
    } catch {
        # Use default if property not found
    }
    
    # Initialize result structure
    $result = @{
        jobNumber = $jobNumber
        printPackage = @()
        sheets = @()
    }
    
    # Helper function to get column index
    function Get-ColumnIndex {
        param($partsList, $columnTitle)
        for ($i = 1; $i -le $partsList.PartsListColumns.Count; $i++) {
            if ($partsList.PartsListColumns.Item($i).Title -eq $columnTitle) {
                return $i
            }
        }
        return -1
    }
    
    # Helper function to clean sheet name
    function Clean-Name {
        param($name)
        $bad = ":/\?*[]()"
        foreach ($ch in $bad.ToCharArray()) {
            $name = $name.Replace($ch, "_")
        }
        return $name.Trim()
    }
    
    # Extract part number prefix from sheet name (e.g., "35221-01.1_HEATER WELDMENT" -> "35221-01.1")
    function Get-PartNumberPrefix {
        param($sheetName)
        # Try to match pattern: {P/N}_{rest} or {P/N}.{suffix}_{rest}
        if ($sheetName -match '^([\d]+-[\d]+(?:\.[\d]+)?)') {
            return $matches[1]
        }
        return $null
    }
    
    # Process Print Package (all BOM items with UM="BOM")
    $currentRow = 0
    foreach ($sheet in $drawingDoc.Sheets) {
        if ($sheet.PartsLists.Count -gt 0) {
            for ($i = 1; $i -le $sheet.PartsLists.Count; $i++) {
                $pl = $sheet.PartsLists.Item($i)
                
                foreach ($row in $pl.PartsListRows) {
                    if (-not $row.Visible) { continue }
                    
                    # Check UM column
                    $umCol = Get-ColumnIndex $pl "UM"
                    $umValue = ""
                    if ($umCol -gt 0) {
                        try {
                            $umValue = $row.Item($umCol).Value.ToString().ToUpper().Trim()
                        } catch {}
                    }
                    
                    # Only process assemblies (UM = "BOM")
                    if ($umValue -eq "BOM") {
                        # Get part number
                        $partNumCol = Get-ColumnIndex $pl "KEMCO PART NUMBER"
                        $bomPartNumber = ""
                        if ($partNumCol -gt 0) {
                            try {
                                $bomPartNumber = $row.Item($partNumCol).Value.ToString().Trim()
                            } catch {}
                        }
                        
                        # Only add if valid part number
                        if ($bomPartNumber -and $bomPartNumber -ne "N/A" -and $bomPartNumber -ne "Error") {
                            $bomItem = @{
                                sheetName = $sheet.Name
                                kemcoPartNumber = $bomPartNumber
                                kemcoDescription = ""
                                qty = ""
                                um = $umValue
                                amount = ""
                            }
                            
                            # Get description
                            $descCol = Get-ColumnIndex $pl "KEMCO DESCRIPTION"
                            if ($descCol -gt 0) {
                                try {
                                    $bomItem.kemcoDescription = $row.Item($descCol).Value.ToString()
                                } catch {}
                            }
                            
                            # Get QTY
                            $qtyCol = Get-ColumnIndex $pl "QTY"
                            if ($qtyCol -gt 0) {
                                try {
                                    $bomItem.qty = $row.Item($qtyCol).Value.ToString()
                                } catch {}
                            }
                            
                            # Get AMOUNT
                            $amtCol = Get-ColumnIndex $pl "AMOUNT"
                            if ($amtCol -gt 0) {
                                try {
                                    $bomItem.amount = $row.Item($amtCol).Value.ToString()
                                } catch {}
                            }
                            
                            $result.printPackage += $bomItem
                        }
                    }
                }
            }
        }
    }
    
    # Process individual sheets
    foreach ($sheet in $drawingDoc.Sheets) {
        if ($sheet.PartsLists.Count -eq 0) { continue }
        
        $sheetName = $sheet.Name
        $cleanedName = Clean-Name $sheetName
        
        # Determine final tab name (matching VBA logic)
        $finalTabName = ""
        $upperName = $cleanedName.ToUpper()
        
        if ($upperName.Contains("HEATER FINAL")) {
            $finalTabName = "$jobNumber`_HEATER FINAL"
        } elseif ($upperName.Contains("HEATER WELDMENT")) {
            $finalTabName = "$jobNumber`.1_HEATER WELDMENT"
        } elseif ($upperName.Contains("HEATER SHELL")) {
            $finalTabName = "$jobNumber`.2_HEATER SHELL"
        } elseif ($upperName.Contains("HEATER STACK")) {
            $finalTabName = "$jobNumber`.3_HEATER STACK"
        } elseif ($upperName.Contains("GAS TRAIN")) {
            $finalTabName = "$jobNumber`.4_GAS TRAIN"
        } elseif ($upperName.Contains("MODULAR PIPING")) {
            $finalTabName = "$jobNumber`.5_MODULAR PIPING"
        } elseif ($upperName.Contains("TANK") -and $upperName.Contains("FFA")) {
            $finalTabName = "$jobNumber`_TANK FFA"
        } elseif ($upperName.Contains("SHELL") -and $upperName.Contains("FFA")) {
            $finalTabName = "$jobNumber`.1_SHELL FFA"
        } elseif ($upperName.Contains("SUCTION FITTING")) {
            $finalTabName = "SUCTION FITTING"
        } else {
            # Use cleaned name (truncated to 28 chars)
            $finalTabName = if ($cleanedName.Length -gt 28) { $cleanedName.Substring(0, 28) } else { $cleanedName }
        }
        
        # Extract part number prefix from final tab name
        $partNumberPrefix = Get-PartNumberPrefix $finalTabName
        
        # Get column headers from first PartsList
        $pl = $sheet.PartsLists.Item(1)
        $columns = @()
        for ($c = 1; $c -le $pl.PartsListColumns.Count; $c++) {
            $columns += $pl.PartsListColumns.Item($c).Title
        }
        
        # Collect all BOM rows from all PartsLists on this sheet
        $bomRows = @()
        for ($i = 1; $i -le $sheet.PartsLists.Count; $i++) {
            $pl = $sheet.PartsLists.Item($i)
            
            foreach ($row in $pl.PartsListRows) {
                if (-not $row.Visible) { continue }
                
                $rowData = @{}
                for ($c = 1; $c -le $columns.Count; $c++) {
                    try {
                        $rowData[$columns[$c - 1]] = $row.Item($c).Value.ToString()
                    } catch {
                        $rowData[$columns[$c - 1]] = ""
                    }
                }
                
                # Apply filters (matching VBA logic)
                $qtyVal = if ($rowData.ContainsKey("QTY")) { $rowData["QTY"].ToString().ToUpper().Trim() } else { "" }
                $partVal = if ($rowData.ContainsKey("KEMCO PART NUMBER")) { $rowData["KEMCO PART NUMBER"].ToString().Trim() } else { "" }
                
                # Skip if QTY=0, QTY=REF, or PART starts with 900-
                if ($qtyVal -eq "0" -or $qtyVal -eq "REF" -or ($partVal.Length -ge 4 -and $partVal.Substring(0, 4) -eq "900-")) {
                    continue
                }
                
                # Process QTY/AMOUNT override (matching VBA logic)
                if ($rowData.ContainsKey("AMOUNT") -and $rowData.ContainsKey("QTY")) {
                    $amtVal = $rowData["AMOUNT"].ToString().Trim()
                    if ($amtVal) {
                        $rowData["QTY"] = $amtVal
                    }
                }
                
                # Add WAREHOUSE column (before QTY, matching VBA logic)
                $rowData["WAREHOUSE"] = "000"
                
                # Process UM updates (matching VBA logic)
                if ($rowData.ContainsKey("UM")) {
                    $currentUM = $rowData["UM"].ToString().ToUpper().Trim()
                    if ($currentUM -eq "BOM") {
                        $isHeaterFinal = $finalTabName.ToUpper().Contains("HEATER FINAL")
                        $description = if ($rowData.ContainsKey("KEMCO DESCRIPTION")) { $rowData["KEMCO DESCRIPTION"].ToString().ToUpper() } else { "" }
                        
                        if ($isHeaterFinal -and ($description.Contains("GAS TRAIN") -or $description.Contains("HEATER, WELD"))) {
                            $rowData["UM"] = "PEGGED SUPPLY"
                        } else {
                            $rowData["UM"] = "PHANTOM"
                        }
                    }
                }
                
                $bomRows += $rowData
            }
        }
        
        # Create sheet data object
        $sheetData = @{
            sheetName = $sheetName
            finalTabName = $finalTabName
            partNumberPrefix = $partNumberPrefix
            columns = $columns + "WAREHOUSE"  # Add WAREHOUSE to columns list
            bomRows = $bomRows
        }
        
        $result.sheets += $sheetData
    }
    
    # Close document without saving
    $drawingDoc.Close($false)
    
    # Return JSON result
    $result | ConvertTo-Json -Depth 10
    
} catch {
    Write-Error "Error processing Inventor file: $_"
    exit 1
} finally {
    # Clean up Inventor application
    if ($inventorApp) {
        try {
            $inventorApp.Quit()
        } catch {}
    }
}

