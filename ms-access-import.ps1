# Powershell Functions for MS Access :)
# Here are several functions that cover tasks I tend to do most often in Access (hopefully more to come soon). These are mostly  
# just wrappers around Access VBA methods, but this helps me avoid constant googling every time I need to do something in Access.
# example usage is provided at the bottom of the script


# Import CSV file into MS Access database
# office vba documentation: https://docs.microsoft.com/en-us/office/vba/api/access.docmd.transferspreadsheet

function Import-MsAccessCsv
{
    param ( 
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $Path,
        [Parameter(Mandatory = $True)]  [string] $TableName,
        [Parameter(Mandatory = $False)] [switch] $HasFieldNames,
        [Parameter(Mandatory = $False)] [string] $SpecificationName=$null
    )

    $transferType = 0 
    $DoCmd = $Access.DoCmd
    $DoCmd.TransferText( $transferType, $SpecificationName, $TableName, $Path, [bool]$HasFieldNames )
}



# Import Excel file into MS Access database
# office vba documentation: https://docs.microsoft.com/en-us/office/vba/api/access.docmd.transferspreadsheet

function Import-MsAccessExcel
{
    param (
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $Path,
        [Parameter(Mandatory = $True)]  [string] $TableName,
        [Parameter(Mandatory = $True)]  [switch] $HasFieldNames,
        [Parameter(Mandatory = $False)] [string] $Range=$null,
        [Parameter(Mandatory = $False)] 
        [ValidateSet('Current','2010','2000','1997','1995','5.0','4.0','3.0')]
        [string[]]$Version='Current'
    )

    # see https://docs.microsoft.com/en-us/office/vba/api/access.acspreadsheettype for acSpreadsheetType
    $acSpreadsheetTypes = @{
        '3.0'  = 0;
        '4.0'  = 6;
        '5.0'  = 5;
        '1995' = 5;
        '1997' = 8;
        '2000' = 8;
        '2010' = 9;
        'Current' = 10;
    }

    $transferType = 0 
    $spreadsheetType = $acSpreadsheetTypes.item( [string]$Version )
    $DoCmd = $Access.DoCmd
    $DoCmd.TransferSpreadsheet( $transferType, $spreadsheetType, $TableName, $Path, [bool]$HasFieldNames, $Range, $null )
}



# Tests whether a QueryDef exists

function Test-MsAccessQueryDef
{
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $QueryName
    )

    $db = $Access.CurrentDb()

    try { 
        $queryDef = $db.QueryDefs.item($QueryName)
    } catch { 
        return $false
    }
    return $true
}



# Create a new MS Access query
# office vba documentation: https://docs.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/database-createquerydef-method-dao

function New-MsAccessQuery
{
    [OutputType([__ComObject])]
    param (
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $QueryName,
        [Parameter(Mandatory = $False)] [string] $SQL=$null
    )

    # check whether querydef with the same name already exists
    if( Test-MsAccessQueryDef $Access $QueryName ) {
        throw "Error: A query by the name of '$QueryName' already exists."
    } elseif( $SQL -ne $null ) { 
        $db = $Access.CurrentDb()
        return $db.CreateQueryDef( $QueryName, $SQL )
    } else {
        return $db.CreateQueryDef( $QueryName )
    }
}



# Close MS Access query
# office vba documentation: https://docs.microsoft.com/en-us/office/vba/api/access.docmd.close

function Close-MsAccessQuery
{
    param (
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $QueryName,
        [Parameter(Mandatory = $False)] [switch] $Save
    )

    $DoCmd = $Access.DoCmd

    if( $Save ) { 
        $DoCmd.close(1, $QueryName, 1)
    } else {
        $DoCmd.close(1, $QueryName, 2)
    }
}



# Open a query in MS Access
# note that if this runs an existing query and includes a sql argument, the existing query will be saved-over
# office vba documentation: https://docs.microsoft.com/en-us/office/vba/api/access.docmd.openquery

function Open-MsAccessQuery
{
    param (
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $QueryName,
        [Parameter(Mandatory = $False)] [string] $SQL=$null
    )   

    $DoCmd = $access.DoCmd

    # create new query if not exists
    if( !( Test-MsAccessQueryDef $Access $QueryName )) {
        New-MsAccessQuery $Access $QueryName $SQL
    }

    # update query sql if SQL argument is provided 
    if( $SQL ) {
        $db = $Access.CurrentDb()
        $queryDef = $db.QueryDefs.Item($QueryName)
        $queryDef.sql = $SQL

        # save & close query after updating
        Close-MsAccessQuery -Access $Access -QueryName $QueryName -Save
    }

    # open query in current ms access window
    $Access.Visible = $true
    $DoCmd.OpenQuery($QueryName)
}



# Convert MS Access RecordSet (table or query) to Powershell object array having the same field names
# Note that this is very inefficient/slow because each individual row and column must be iterated through, so it is a good idea 
# to only use this with a limited amount of data

function Get-MsAccessData
{
    [OutputType([Object[]])]
    param (
        [Parameter(Mandatory = $True)] [__ComObject] $Access,
        [Parameter(Mandatory = $True)] [ValidateSet('Query','Table')] [string[]]$ObjectType,
        [Parameter(Mandatory = $True)] [string] $Name,
        [Parameter(Mandatory = $False)] [ValidateRange(0,[int]::MaxValue)] [int] $Limit
    )

    $data = @() 
    $db = $access.CurrentDb()
    
    # open recordset from query/table, as applicable
    switch($ObjectType) {
        'Query' {
            $queryDef = $db.QueryDefs.item($Name)
            $recordSet = $queryDef.OpenRecordset()
        }
        'Table' {
            $tableDef = $db.TableDefs.Item($Name)
            $recordSet = $tableDef.OpenRecordset()
        }
    }

    $fieldNames = $recordSet.Fields | Select-Object -ExpandProperty name

    # loop through each row of recordset & convert row to custom object, then add to data array
    $i = 0
    while( !$recordSet.EOF ) { 

        $record = [psCustomObject]@{}

        foreach( $name in $fieldNames ) { 

            $record | Add-Member -MemberType NoteProperty `
                                 -name $name `
                                 -value $recordSet.Fields.item($name).value
        }

        $data += $record
        $recordSet.MoveNext()
        
        # if limit is given exit loop when limit is reached
        if( $Limit -and ++$i -ge $Limit ) { break }
    }

    $recordSet.Close()
    return $data 
}

function New-MsAccessDB
{
    param ( 
        [Parameter(Mandatory = $True)]  [__ComObject] $Access,
        [Parameter(Mandatory = $True)]  [string] $Path
    )

    if (Test-Path $Path) {
        Write-Host 'There is an existing Access DB. Please use another name.'
        Exit
    }
    $Access.NewCurrentDataBase($Path)
    $Access.CloseCurrentDataBase()
    # $Access.Quit()
}

# ----------------------------------------------------------------------------------------
# Examples:

$dbPath = $PSScriptRoot + '\database.accdb'
$xlFile = $PSScriptRoot + '\ICA_G3_Cert_Ownership_internal_2019-12.xlsx'

$access = New-Object -ComObject Access.Application

# create a new Access DB
if (Test-Path $dbPath) {
    Remove-Item $dbPath
}
New-MsAccessDB -Access $access -Path $dbPath

# open an existing ms access database:
# $access = New-Object -ComObject Access.Application
$access.OpenCurrentDatabase($dbPath)
$access.Visible = $true

# import csv file into a new table
# Import-MsAccessCsv -Access $access -Path 'C:\path\to\csv_file.csv' -TableName 'csv_data' -HasFieldNames



# import excel file into a new table 
$tableName = 'webserver'
$range = "01_WebServer!" 
Import-MsAccessExcel -Access $access -Path $xlFile -TableName $tableName -HasFieldNames $range

# run a query (this creates & saves a new query but also work using an existing query)
# Open-MsAccessQuery -Access $access -QueryName 'new_query' -SQL 'SELECT top 20 * FROM excel_data'

# pull  records of table/query data into powershell object arrays (careful, very slow!)
# Get-MsAccessData -Access $access -ObjectType Table -Name 'csv_data' -Limit 10 | Out-GridView
# Get-MsAccessData -Access $access -ObjectType Query -Name 'new_query' -Limit 20 | Out-GridView

# close access
$access.Quit()
