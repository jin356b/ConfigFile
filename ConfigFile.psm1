#=================================================================================================
# NAME       : Configuration File Module
# AUTHOR     : Jonathon Bauer
# DATE       : 11/1/2017
# VERSION    : 3.1.1
# DESCRIPTION: This module provides functions to read and write data from a configuration file.
# DEPENDENCY : PowerShell v3.0 or higher.
# FUNCTIONS  : $cfg = New-Cfg [-Path] <string> [-Password <string>|<securestring>] [-CachedMode] [-Force]  [<CommonParameters>]
#              $cfg.Get([String]$varName, [ [String]|[SecureString]$Password ]) 
#              $cfg.Set([String]$varName, [Object]$value, [ "DPAPI"|"AES256" ], [ [String]|[SecureString]$Password ])
#              $cfg.Read([String[]]$varName, [ [String]|[SecureString]$Password ]) 
#              $cfg.Write([String[]]$varName, [ "DPAPI"|"AES256" ], [ [String]|[SecureString]$Password ])
#              $cfg.Remove([String[]]$varName)
#              $cfg.Flush()
#=================================================================================================

#-------------------------------
#Class Definitions
#-------------------------------

$CfgClass = {
    [CmdletBinding()]
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [String]$_Path,

        [Parameter(Position=1,Mandatory=$true)]
        [AllowNull()]
        $_Password,

        [Parameter(Position=2,Mandatory=$true)]
        [bool]$_CachedMode
    )

#PRIVATE VARIABLES

    [HashTable[]]$_ConfigData = @()
    #[String]$_Path
    #$_Password              #Should either be $null or a SecureString
    #[bool]$_CachedMode
    
#PUBLIC VARIABLES

    #<none> All public variables are added as ScriptProperty members in the Set-ConfigFile Function

#PRIVATE METHODS

    #Writes the config file, using the private $_ConfigData array
    Function Write-ConfigData
    {
        [CmdletBinding()]param()

        #Ensure Cached Mode isn't enabled
        if($_CachedMode){ Return }

        #Ensure there is data to write
        if($_ConfigData.count -eq 0){ Write-Verbose "No variables to write"; Return }

        Write-Debug "Writing $($_ConfigData.count) variable(s) to $_Path"

        #Format the config data as string data
        [String]$Text = Join-VarRecordsToText $_ConfigData
        
        #Write the string data out
        if(Test-Path -LiteralPath $_Path) #Try Literal Path
        {
            Try  { [System.IO.File]::WriteAllText($_Path, $Text) }
            Catch{ Write-Error "Unable to write config file: $_Path"; Return }
        }
        else
        { Write-Error "Unable to locate config file at: $_Path"; Return }
    }

    #Reads the config file, and populates the private $_ConfigData array
    Function Read-ConfigData
    {
        [CmdletBinding()]param()
        Write-Debug "Reading Config Data from $_Path"

        #Check if the config file & path exist
        if(!(Test-Path -LiteralPath $_Path))
        { Write-Error "Unable to locate config file at: $_Path"; Return }
        
        #Read the String data in
        Try  { [string]$Text = [System.IO.File]::ReadAllText($_Path) }
        Catch{ Write-Error "Unable to read config file: $_Path"; Return }

        #Populate $configData with variable names, types, and Value as a string
        $Script:_ConfigData = @()
    
        #Ensure the Config file contains some data
        if($Text.Trim()){ $Script:_ConfigData = Split-TextToVarRecords $Text }
        else            { Write-Verbose "Config file is empty: $_Path" }
    }
    
    #Takes an array of Variable Records (Hashtables), and converts it to a text string
    Function Join-VarRecordsToText
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
 		    [HashTable[]]$varRecords,                  # A variable record HashTable containing 'Name', 'DataType', and 'Value' keys

            [Parameter(Position=1, Mandatory=$false)]
            [String]$Parent                            # Name of the parent object, to prefex onto the variable record
        )

        #Init Vars
        [string[]]$strData = @()

        foreach($varRecord in $varRecords)
        {
            #Identify Elements of the variable record
            if($Parent){ $Name = "$Parent.$($varRecord.Name)" }
            else{ $Name = $varRecord.Name }
            $Type = $varRecord.DataType
            $Value = $varRecord.Value
            Write-Debug "J-VRTD Name: $Name"
            Write-Debug "J-VRTD Type: $Type"
            
            #Handle Comments and Whitespace
            if($Type -eq "Cfg.File.Format")
            { 
                $strData += $Value
                Continue
            }
            
            #Apply Single and Multi-Line formatting
            if($Value.GetType().BaseType.Name -eq "Array") # Has child values
            {
                if($Value){ $Value = "`r`n$(Join-VarRecordsToText -varRecords $Value -Parent $Name)" }
                else      { $Value = "" } #Handle parent objects with no children
            }
            elseif($Value.Contains("`r`n")) # Multi line var data (i.e. [String:3]) 
            {
                $lines = ($Value -split "`r`n").count
                $Type += ":$lines"
                $Value = "`r`n$Value"
            }
            else # Single line string
            { $Value = "=$Value" }
            Write-Debug "J-VRTD Value: $Value"

            #Format Variable Type
            if($Type -eq "String"){ $Type = "" }
            else{ $Type = "[$Type]" }
    
            $strData += "$Name$Type$Value"
        }

        #Join the string array, and return the text
        [String]$Text = $strData -join "`r`n"
        Return $Text
    }

    #Takes a text string, and converts it to an array of Variable Records (Hashtables)
    Function Split-TextToVarRecords
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
            [AllowEmptyString()]
 		    [String]$Text
        )

        #Initialize variables
        $varRecords = @()
        [string[]]$strData = $Text -split "`r`n"

        #Loop through Text data, and assemble an array of variables, but don't handle data type conversion yet
        for($i=0; $i -lt $strData.Length; $i++) # We use a for-loop so we can manually advance $i when reading multi-line text
        {
            $line = $strData[$i]
        
            # Handle White space and comments
            if(($line.Trim() -eq "") -or ($line[0] -eq "#")) 
            {
                if($line[0] -eq "#"){ $vName = "Comment" }
                else{ $vName = "WhiteSpace" }

                $varRecords += @{"Name"=$vName;"DataType"="Cfg.File.Format";"Value"=$line}
                Write-Debug "Format Data: $line"
                Continue #Proceed to next line
            }

            # Process Variable
            $vData = "" # Init first-level HashTables and Arrays to an empty string
            $isChild = $false

            # Split Variable's name/datatype prefix from it's value
            if($line.Contains("=")) # This is a single line variable
            {
                #Split variable across the first equals sign
                $x = $line.IndexOf("=")
                $prefix = $line.Substring(0,$x)
                $vData = $line.Substring($x+1)
            }
            else # Multi-line variable
            { $prefix = $line }
            Write-Debug "Prefix: $prefix"
            
            #Determine if this is a child variable
            $dot = $prefix.IndexOf(".")
            $brk = $prefix.IndexOf("[")
            if(    $dot -eq -1 -and $brk -eq -1){ $isChild = $false }
            elseif($dot -eq -1 -and $brk -gt  0){ $isChild = $false }
            elseif($dot -gt  0 -and $brk -eq -1){ $isChild = $true }
            elseif($dot -lt $brk){ $isChild = $true }
            else{ $isChild = $false }
            Write-Debug "Is Child: $isChild"
            
            #Identify Variable Name and Type
            if($brk -gt 0) # Process Data type
            {
                $vName = $prefix.Substring(0,$brk)
                $vType = $prefix.Substring($brk+1,$prefix.Length-$brk-2)
            }
            else # No Data Type, so assume string
            {
                $vName = $prefix
                $vType = "String" 
            }
            Write-Debug "Variable Name: $vName"

            #Handle Multi-Line String data    
            $MLcount = 0
            if($vType.Contains(":"))
            {
                #Get count of Multi-Line Objects
                $x = $vType.IndexOf(":")
                Try{ [int]$MLcount = $vType.Substring($x+1) }Catch{}
                if($MLcount -le 0){ Write-Warning "Ignoring malformed multi-line text data: $line"; Continue }
                $vType = $vType.Substring(0,$x) #Remove line count from type name
                Write-Debug "Multi-Line Count: $MLcount"
                    
                #Grab the number of lines specified
                $a = $i+1
                $b = $i+$MLcount
                $vData = $strData[$a..$b] -join "`r`n"
                $i = $b #### WARNING!!! Moving position forward in the for loop! $i will be skewed below this point
            }
            Write-Debug "Variable Type: $vType"
                
            if($isChild)# Dot Notation child. We don't care about Type data (except for multi-line strings)
            {
                #Get the name of the parent variable to populate
                $parentName = $prefix.Substring(0,$dot)
                Write-Debug "Parent Name: $parentName"

                #Get the child variable data
                $childData = $line.Substring($dot+1)#Strip out the Parent's name
                if($MLcount){ $childData = "$childData`r`n$vData" } # Add in Multi-Line data

                #Add child variable data to the parent's array. Uses most recent parent with that name
                $varRecords | Where { $_.Name -eq $parentName -AND $_.DataType -ne "Cfg.File.Format" } | Select-Object -Last 1 | % { if($_.Value){ $_.Value += "`r`n$childData" }else{ $_.Value += $childData } }
                Write-Debug "Child Data: $childData"
                Continue #Proceed to next line
            }

            #Create new variable record
            $varRecords += @{"Name"=$vName;"DataType"=$vType;"Value"=$vData}
            Write-Debug "Variable Data: $vData"
        }

        #Convert Datatypes with child objects
        foreach($varRecord in $varRecords)
        {
            if($varRecord.DataType -eq "HashTable" -OR $varRecord.DataType.EndsWith("[]"))
            {
                if($varRecord.Value.Trim()) #Ensure the array or Hashtable isn't null
                { $varRecord.Value = @(Split-TextToVarRecords $varRecord.Value) }
                else
                { $varRecord.Value = @() } #Handle objects with no children
            }
        }

        Return $varRecords
    }

    #Gets the string equivalent of the variable object
    Function Get-VarString
    {
        [CmdletBinding()] 
        param(
            [Parameter(Position=0, Mandatory=$true)]
            [String]$Name,                             # Name of the variable to convert. Needed for dot-notation variables

            [Parameter(Position=1, Mandatory=$true)]
            [AllowEmptyString()]
            $Value,                                    # Object to convert to a string

            [Parameter(Position=2, Mandatory=$false)]
            [String]$Type,                             # (Optional) Used to specify the type of encyption (DPAPI or AES256)

            [Parameter(Position=3, Mandatory=$false)]
            $Password                                  # (Optional) Password to encrypt variables with AES 256
        )

        #Init Variables
        if(!$Type){ $Type = $Value.GetType().Name }

        #Convert Objects value data to strings
        if($Type -eq "DateTime") # [DateTime]
        { $newValue = Get-date $Value -format "yyyy-MM-dd HH:mm:ss.fffffff" }
        elseif($Type -eq "XmlDocument") # [XmlDocument]
        { $newValue = $Value.OuterXml }
        elseif($Type -eq "HashTable") # [HashTable]
        {
            $newValue = @()            
            foreach($item in $Value.GetEnumerator())
            {
                $keyData = Get-VarString -Name "Key"   -Value $item.Name
                $valData = Get-VarString -Name "Value" -Value $item.Value
                if(!$keyData -or !$valData){ Write-Error "Unable to convert Hashtable element for '$Name' to string data"; Continue}

                $newValue += $keyData
                $newValue += $valData
            }
        }
        elseif($Type.EndsWith("[]")) # [Array]
        {
            $newValue = @()
            for($i=0; $i -lt $Value.Length; $i++)
            { 
                $newValue += Get-VarString -Name "$i" -Value $Value[$i]
                if(!$newValue[-1]){ Write-Error "Unable to convert Array element for '$Name' to string data" } #Will add a $null element if ErrorAction is Continue, to preserve array index positions
            }
        }
        elseif($Type -eq "SecureString")
        {
            #Convert SecureStrings to the Base64 format 
            Try  { $newValue = SecureStringToBase64 $Value -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert '$Name' to a Base64 String."; Return }
        }
        elseif($Type -eq "PSCredential")
        {
            #Convert Username to SecureString
            Try  { $userValue = ConvertTo-SecureString -String $Value.UserName -AsPlainText -Force -ErrorAction Stop }
            Catch{ Write-Error "Unable to encrypt the username value of '$Name' using $Type"; Return }

            #Convert User & Password SecureString to the Base64 format 
            Try  { $passValue = SecureStringToBase64 $Value.Password -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert the password value of '$Name' to a Base64 String."; Return }
            Try  { $userValue = SecureStringToBase64 $userValue -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert the username value of '$Name' to a Base64 String."; Return }

            $newValue = "$userValue,$passValue"
        }
        elseif($Type -eq "DPAPI")
        {
            #Convert the variable's data type and value to a string
            $valData = Get-VarString -Name "d" -Value $Value
            if(!$valData){ Write-Error "Unable to convert '$Name' to string data"; Return }
            $varString = Join-VarRecordsToText $valData

            #Encrypt string using DPAPI
            Try  { $newValue = ConvertTo-SecureString -String $varString -AsPlainText -Force -ErrorAction Stop}
            Catch{ Write-Error "Unable to encrypt the variable '$Name' using $Type"; Return }

            #Convert SecureStrings to the Base64 format 
            Try  { $newValue = SecureStringToBase64 $newValue -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert '$Name' to a Base64 String."; Return }
        }
        elseif($Type -eq "AES256")
        {
            #Validate the Password data
            $Password = Get-Password $Password -ErrorAction Continue
            if(!$Password){ Write-Error "Unable to encrypt the variable '$Name' using $Type"; Return }

            #Convert the variable's data type and value to a string
            $valData = Get-VarString -Name "a" -Value $Value
            if(!$valData){ Write-Error "Unable to convert '$Name' to string data"; Return }
            $varString = Join-VarRecordsToText $valData

            #Encrypt string using AES-256, and format as base64 string
            Try
            {
                $AEScsp = New-Object System.Security.Cryptography.AesCryptoServiceProvider -ErrorAction Stop
                $AEScsp.Key = ConvertToKey $Password -ErrorAction Stop
                $AEScsp.IV = $AEScsp.Key[8..23]
            
                $inBlock = [System.Text.UnicodeEncoding]::Unicode.getbytes($varString) 
                $xfrm = $AEScsp.CreateEncryptor() 
                $outBlock = $xfrm.TransformFinalBlock($inBlock, 0, $inBlock.Length) 
                $newValue = [System.Convert]::ToBase64String($outBlock)
            }
            Catch{ Write-Error "Unable to encrypt the variable '$Name' using $Type"; Return }
        }
        else # All other data types
        { $newValue = "$Value" }

        #The $newValue can NOT be $Null
        if($newValue -eq $null){ $newValue = ""; Write-Warning "Variable '$Name' had a `$null value set to an empty string" }

        #Return a new $varRecord
        Return @{"Name"=$Name;"DataType"=$Type;"Value"=$newValue}
    }

    #Gets the Object equivalent of the Variable Record's string value
    Function Get-VarObject
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
 		    [HashTable]$varRecord,                    # A variable record HashTable containing 'Name', 'DataType', and 'Value' keys

            [Parameter(Mandatory=$false)]
            $Password                                 # (Optional) Password to decrypt AES 256 encoded variables.
        )

        #Don't do any conversions on Strings or formatters
        if($varRecord.DataType -eq "String" -OR $varRecord.DataType -eq "Cfg.File.Format") #Formatters shouldn't be resolved, but just in case
        { $retObj = $varRecord.Value }
        elseif($varRecord.DataType -eq "HashTable") #Convert string data to a Hashtable
        {
            $retObj = @{}
            for($i=0; $i -lt $varRecord.Value.count; $i+=2)
            {
                #Error Checking
                if($varRecord.Value[$i].Name -ne "Key" -OR $varRecord.Value[$i+1].Name -ne "Value")
                {                    
                    Write-Debug "KeyName: $($varRecord.Value[$i].Name) KeyType: $($varRecord.Value[$i].DataType) KeyValue: $($varRecord.Value[$i].Value)"
                    Write-Debug "ValueName: $($varRecord.Value[$i+1].Name) ValueType: $($varRecord.Value[$i+1].DataType) ValueValue: $($varRecord.Value[$i+1].Value)"
                    Write-Warning "Malformed HashTable element ignored." 
                    Continue
                }

                #Get the Key & Value objects
                $keyObj = Get-VarObject $varRecord.Value[$i]
                $valObj = Get-VarObject $varRecord.Value[$i+1]
                if(!$keyObj -or !$valObj){ Write-Warning "Malformed HashTable element ignored."; Continue }
                
                #Add the Key/Value pair
                $retObj.Add($keyObj.Object,$valObj.Object) 
            }
        }
        elseif($varRecord.DataType.EndsWith("[]")) #Convert string data to an Array
        {
            $retObj = New-Object System.Collections.Generic.List[System.Object]
            $records = $varRecord.Value | sort { $_.Name -as [int] } | % { Get-VarObject $_ } # Sort and retrieve the array elements
            foreach($item in $records){ $retObj.Add($item.Object) } #Populate the list
            
            $retObj = $retObj.ToArray() -as (ConvertTo-DataType $varRecord.DataType) # Convert list to an array, and fix data type
        }
        elseif($varRecord.DataType -eq "SecureString") #Convert SecureStrings from their Base64 format to a SecureString
        {
            Try  { $retObj = Base64ToSecureString $varRecord.Value -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert '$($varRecord.Name)' to a Secure String."; Return }
        }
        elseif($varRecord.DataType -eq "PSCredential") #Convert PSCredential from their Base64 format to a SecureString
        {
            #Split the string data into username/password
            $items = $varRecord.Value -split ','

            #Convert User & Password from Base64 to SecureString 
            Try  { $userValue = Base64ToSecureString $items[0] -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert the username portion of '$($varRecord.Name)' to a Secure String."; Return }
            Try  { $passValue = Base64ToSecureString $items[1] -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert the password portion of '$($varRecord.Name)' to a Secure String."; Return }

            #Decrypt the username
            Try  { $userValue = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($userValue)) }
            Catch{ Write-Error "Unable to decrypt the username portion of '$($varRecord.Name)' using $($varRecord.DataType)"; Return }

            #Assemble the PSCredential Object
            Try  { $retObj = New-Object System.Management.Automation.PSCredential ($userValue, $passValue) -ErrorAction Stop }
            Catch{ Write-Error "Unable to create the PSCredential object for '$($varRecord.Name)'"; Return }
        }
        elseif($varRecord.DataType -eq "DPAPI") #Decrypt DPAPI variables
        {
            #Convert DPAPI from their Base64 format to a SecureString
            Try  { $SecureStr = Base64ToSecureString $varRecord.Value -ErrorAction Stop }
            Catch{ Write-Error "Unable to convert '$($varRecord.Name)' to a Secure String."; Return }

            # Perform DPAPI decryption
            Try  { $varString = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureStr)) }
            Catch{ Write-Error "Unable to decrypt the variable '$($varRecord.Name)' using $($varRecord.DataType)"; Return }

            # Unpack the Single Variable Object
            Return Get-VarObject (Split-TextToVarRecords $varString)
        }
        elseif($varRecord.DataType -eq "AES256") #Decrypt AES256 variables
        {
            # Validate the Password data
            $Password = Get-Password $Password -ErrorAction Continue
            if(!$Password){ Write-Error "Unable to decrypt the variable '$($varRecord.Name)' using $($varRecord.DataType)"; Return }

            # Perform AES-256 decryption
            Try
            {
                $AEScsp = New-Object System.Security.Cryptography.AesCryptoServiceProvider -ErrorAction Stop
                $AEScsp.Key = ConvertToKey $Password -ErrorAction Stop
                $AEScsp.IV = $AEScsp.Key[8..23]

                $inBlock = [System.Convert]::FromBase64String($varRecord.Value)
                $xfrm = $AEScsp.CreateDecryptor()
                $outBlock = $xfrm.TransformFinalBlock($inBlock, 0, $inBlock.Length) 
                $varString = [System.Text.UnicodeEncoding]::Unicode.GetString($outBlock)
            }
            Catch{ Write-Error "Unable to decrypt the variable '$($varRecord.Name)' using $($varRecord.DataType)"; Return }

            # Unpack the Single Variable Object
            Return Get-VarObject (Split-TextToVarRecords $varString)
        }
        else #Handle all other data types, to include: Int, Guid, DateTime, Byte, XmlDocument   
        {
            Try  { $retObj = $varRecord.Value -as (ConvertTo-DataType $varRecord.DataType -ErrorAction Stop) }
            Catch{ Write-Error "Unable to convert variable '$($varRecord.Name)' to an object of type '$($varRecord.DataType)'"; Return }
        }

        #Return variable in a hashtable to prevent array type being lost
        Return @{"Object"=$retObj}
    }
    
    #Returns a DataType object, based on the supplied string
    Function ConvertTo-DataType
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
		    [String]$Name
        )

        #Handle Special cases
        if($Name -eq "XmlDocument"){ $Name = "Xml.XmlDocument" }

        #Convert to Datatype
        Try  { $dt = [Type]$Name }
        Catch{ Write-Warning "Unable to resolve data type '$Name'. Falling back to [String] data type."; $dt = [String] }
    
        Return $dt
    }
    
    #Converts a SecureString of any length to a 32-bit encryption key
    Function ConvertToKey
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
		    [SecureString]$Value
        )

        $SHAcsp = new-Object System.Security.Cryptography.SHA256CryptoServiceProvider
        Return $SHAcsp.ComputeHash([System.Text.Encoding]::UTF8.GetBytes([Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($Value))))
    }

    #Converts a Secure String to a Base 64 string
    Function SecureStringToBase64
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
		    [Alias("value")][SecureString]$SecureStr
        )
    
        $hexString = ConvertFrom-SecureString -SecureString $SecureStr #Get text version of encrypted string
        [byte[]]$hashbytes = $hexString -split '(..)' | ? { $_ } | % { [convert]::ToByte($_,16) } #Convert hex string to byte array
        $base64Str = [System.Convert]::ToBase64String($hashbytes) #Convert byte array to Base64

        Return $base64Str
    }

    #Converts a Base 64 string to a Secure String
    Function Base64ToSecureString
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
		    [Alias("value")][String]$base64Str
        )
    
        [byte[]]$hashbytes = [System.Convert]::FromBase64String($base64Str) #Convert Base64 to byte array
        $value = ([System.BitConverter]::ToString($hashbytes)).replace('-','') #Convert byte array to hex string
        $SecureStr = ConvertTo-SecureString -String $value.Trim()

        Return $SecureStr
    }

    #Formats the supplied password as a Secure String, or uses the private $_Password
    Function Get-Password
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$false)]
		    $Password
        )

        if($Password)
        {
            if($Password.getType() -eq [String])
            { Return ConvertTo-SecureString -String $Password -AsPlainText -Force }
            elseif($Password.getType() -eq [SecureString])
            { Return $Password }
            else
            { Write-Error "The 'Password' parameter must be of type [String] or [SecureString]."; Return }
        }
        elseif($_Password)
        { Return $_Password }

        Write-Error "A Password must be specified to encrypt/decrypt variable data."
        Return
    }

    #Formats and validates the variable name syntax
    Function Get-VariableName
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]      
		    [Alias("VarName")][String]$VariableName      # Name of variable to validate
        )

        #Trim the first dollar sign
        if($VariableName[0] -eq '$'){ $VariableName = $VariableName.Substring(1) } 

        #Validate syntax
        if($VariableName -notmatch '^[a-zA-Z0-9_]')
        { Write-Error "The variable name '$VariableName' contains invalid characters."; Return }

        Return $VariableName
    }

    #Debug function used to dump the contents of $_ConfigData
    Function ShowConfigData
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Mandatory=$false)]      
		    [HashTable[]]$varRecords = $Script:_ConfigData,

            [Parameter(Mandatory=$false)]
            [String]$prefix = ""
        )

        foreach($varRecord in $varRecords)
        {
            Write-Host "$prefix ======================="
            Write-Host "$prefix Name: $($varRecord.Name)"
            Write-Host "$prefix Type: $($varRecord.DataType)"
            if($varRecord.Value.GetType().BaseType.Name -eq "Array")
            {
                ShowConfigData -varRecords $varRecord.Value -prefix "$prefix    "
            }
            else
            { Write-Host "$prefix Value: $($varRecord.Value)" }
        }
    }

#PUBLIC METHODS

    # .Get("name",[pass])
    # Returns the value of a single config file variable
    Function Get
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]      # Name of variable to read, and return it's data. Use caution when getting array data,
		    [Alias("VarName")][String]$VariableName,      # as PowerShell doesn't always return it properly. Use $obj.Read($VariableName) instead.

            [Parameter(Mandatory=$false)]
            $Password                                     # (Optional) Password to decrypt AES 256 encoded variables.
        )

        Write-Debug ".Get() called for variable: $VariableName"

        #Format and validate the variable name syntax
        $VariableName = Get-VariableName $VariableName
        if(!$VariableName){ Return }

        #Get the Variable data
        $varRecord = $Script:_ConfigData | Where { $_.Name -eq $VariableName -AND $_.DataType -ne "Cfg.File.Format" } | Select -First 1
        if(!$varRecord){ Write-Error "Variable '$VariableName' not found in the config file: $_Path"; Return }

        #Get the Variable Object
        $varObj = Get-VarObject $varRecord -Password $Password
        if(!$varObj){ Write-Error "Variable '$VariableName' could not be read from the config file: $_Path"; Return }

        Return $varObj.Object
    }

    # .Set("name",$value,[Crypto],[Pass])
    # Sets the value of a single config file variable
    Function Set
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]      
		    [Alias("VarName")][String]$VariableName,      # Name of variable to assign a value to

            [Parameter(Position=1, Mandatory=$true)]
            [AllowEmptyString()]
            $Value,                                       # Value to assign the variable

            [Parameter(Position=2, Mandatory=$false)]
            [String]$Crypto,                              # (Optional) Encryption method

            [Parameter(Position=3, Mandatory=$false)]
            $Password,                                    # (Optional) Password to encrypt AES 256 encoded variables

            ## FOR SYSTEM USE ONLY ##

            [Parameter(Position=4, Mandatory=$false)]
            [bool]$ReturnAction,                          # When enabled, Returns $true if variable was set, otherwise $false

            [Parameter(Position=5, Mandatory=$false)]
            [bool]$noWrite                                # When enabled, the config file will not be updated
        )
    
        Write-Debug ".Set() called for variable: $VariableName"
        
        #Format and validate the variable name syntax
        $VariableName = Get-VariableName $VariableName
        if(!$VariableName){ if($ReturnAction){ Return $false }else{ Return } }
        
        #Determine the Data Type
        if($Crypto){ $DataType = $Crypto }
        else{ $DataType = $Value.GetType().Name }

        #Convert the $Value object to a string
        $newVar = Get-VarString -Name $VariableName -Value $Value -Type $DataType -Password $Password
        if(!$newVar)
        { 
            Write-Error "Unable to convert '$VariableName' to string data"
            if($ReturnAction){ Return $false }else{ Return }
        }

        #Attempt to get the Variable data, if it already exists
        $varRecord = $Script:_ConfigData | Where { $_.Name -eq $VariableName -AND $_.DataType -ne "Cfg.File.Format" } | Select -First 1
        if($varRecord)
        {
            if($varRecord.DataType -eq $newVar.DataType -AND $varRecord.Value -eq $newVar.Value)
            {
                Write-Verbose "The variable '$VariableName' has not changed. Skipping variable update"
                if($ReturnAction){ Return $false }else{ Return }
            }

            Write-Verbose "Overwriting '$VariableName'"
            if($varRecord.DataType -ne $newVar.DataType)
            {
                Write-Warning "The data type for '$VariableName' is changing from $($varRecord.DataType) to $($newVar.DataType)" 
                $varRecord.DataType = $newVar.DataType
            }

            $varRecord.Value = $newVar.Value
        }
        else
        {
            Write-Verbose "Adding new variable '$VariableName'"
            $Script:_ConfigData += $newVar
        }

        if(!$noWrite){ Write-ConfigData } #Update the config file

        if($ReturnAction){ Return $True }
    }

    # .Read([@"name"],[Pass]) 
    # Reads the value of the specified variables into the current scope
    # Reads all variables when no names are provided
    Function Read
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$false)]      
		    [Alias("VarName")][String[]]$VariableName,    # Name(s) of variable(s) to read from the config file

            [Parameter(Position=1, Mandatory=$false)]
            $Password                                     # (Optional) Password to decrypt AES 256 encoded variables
        )
        
        #Read all variable if none were specified
        if(!$VariableName)
        { [String[]]$VariableName = ($_ConfigData | Where { $_.DataType -ne "Cfg.File.Format" }).Name }

        #Loop through each variable to process
        foreach($var in $VariableName)
        {
            #Format and validate the variable name syntax
            $var = Get-VariableName $var
            if(!$var){ Continue }

            #Get the Variable data
            $varRecord = $Script:_ConfigData | Where { $_.Name -eq $var -AND $_.DataType -ne "Cfg.File.Format" } | Select -First 1
            if($varRecord)
            { 
                $varObj = Get-VarObject $varRecord -Password $Password
                if(!$varObj){ Continue }
            }
            else
            { Write-Error "Variable '$var' not found in the config file: $_Path"; Continue }

            #Check if the variable already exists
            Try  { $currentVar = (Get-Variable -scope 2 -name $var -ErrorAction Stop).Value }
            Catch{ Write-Verbose "The variable name '$var' does not exist in the current scope, and will be created"; $currentVar = $null }

            #If the variable exists, check it's data type
            if($currentVar)
            {
                if($currentVar.gettype() -ne $varObj.Object.gettype())
                { Write-Warning "The data type for variable '$var' is changing from $($currentVar.gettype().Name) to $($varObj.Object.gettype().Name)" } 
            }

            #Set the variable value
            Try
            {
                Set-Variable -scope 2 -Name $var -Value $varObj.Object -ErrorAction Stop
                Write-Debug "Updating variable '$var' with data from config file."
            }
            Catch
            { Write-Error "Unable to update variable '$var' in the parent scope."; Continue }
        }
    }

    # .Write(@"name",[Crypto],[Pass])
    # Writes the values of the specified variables in the current scope, to the config file
    Function Write
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]      
		    [Alias("VarName")][String[]]$VariableName,    # Name(s) of variable(s) to write to the config file

            [Parameter(Position=1, Mandatory=$false)]
            [String]$Crypto,                              # (Optional) Encryption method

            [Parameter(Position=2, Mandatory=$false)]
            $Password                                     # (Optional) Password to encrypt AES 256 encoded variables
        )
        
        #Init
        $wroteVars = $false
       
        foreach($var in $VariableName)
        {
            #Format and validate the variable name syntax
            $var = Get-VariableName $var
            if(!$var){ Continue }

            Try  { $value = (Get-Variable -scope 2 -name $var -ErrorAction Stop).Value }
            Catch{ Write-Error "The variable name '$VariableName' does not exist in the current scope."; Continue }

            if($This.Set($var, $value, $Crypto, $Password, $true, $true))
            { $wroteVars = $true }
        }

        #Write new variable data to the Config file
        if($wroteVars){ Write-ConfigData }
    }

    # .Remove(@"name")
    # Removes the specified variable name(s) from the config file
    Function Remove
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]      # Name(s) of variable(s) to remove from the config file
		    [Alias("VarName")][String[]]$VariableName
        )

        #Init
        [String[]]$varList = @()
        $preCount = $_ConfigData.count

        #Format and validate the variable name syntax
        foreach($var in $VariableName)
        {
            $var = Get-VariableName $var
            if($var){ $varList += $var }
        }

        #Rebuild the $_ConfigData, ignoring any of the supplied variable names
        $Script:_ConfigData = $_ConfigData | Where { $varList -notcontains $_.Name }
        
        #Handle results
        if($preCount -eq $_ConfigData.count)
        { Write-Verbose "No variables were removed"; Return }
        else
        {
            Write-Verbose "$($preCount - $_ConfigData.count) variables were removed"
            Write-ConfigData #Update the config file
        }
    }

    #Overwrite the ToString() Method
    Function ToString
    {
        [CmdletBinding()]param() 
        Return $Script:_Path
    }

    #Returns private properties using the supplied name
    Function Private
    {
        [CmdletBinding()] 
        param(
		    [Parameter(Position=0, Mandatory=$true)]
            [String]$Name,

            [Parameter(Position=1, Mandatory=$false)]
            $Value
        )
        
        if($Name -eq "Path"){ Return $_Path }
        if($Name -eq "Count"){ Return ($_ConfigData | Where { $_.DataType -ne "Cfg.File.Format" }).Count }
        if($Name -eq "Names"){ Return ($_ConfigData | Where { $_.DataType -ne "Cfg.File.Format" }).Name }
        if($Name -eq "Dump"){ Return ShowConfigData }
        if($Name -eq "Password")
        {
            if($Value)
            { 
                $tmpPass = Get-Password $Value
                if($tmpPass){ $Script:_Password = $tmpPass }
            }
            else
            { 
                if($_Password){ Return $true}
                else{ Return $false} 
            }
        } 
    }

    # Writes all variable data to the config file
    # Should only be used when operating in Cached Mode
    Function Flush
    {
        [CmdletBinding()]param() 
        
        #Preserve original settings
        $origMode = $_CachedMode
        
        if(!$_CachedMode){ Write-Warning "The Flush() Method is only needed when using Cached Mode." }

        #Write variable data to the config file
        $Script:_CachedMode = $false
        Write-ConfigData
        $Script:_CachedMode = $origMode
    }

#CLASS INITIALIZATION
    
    #Validate the Password data
    if($_Password)
    { $_Password = Get-Password $_Password }

    #Check if the config file exists
    if(Test-Path -LiteralPath $_Path)
    { $_Path = Get-Item -LiteralPath $_Path }
    elseif(Test-Path -Path $_Path)
    { $_Path = Get-Item -Path $_Path }
    else
    { Write-Error "Unable to locate config file at: $_Path"; Return }

    #Read the Config File into the $_ConfigData Array of Hashtables
    Read-ConfigData
    
    #Declare the class definition
    $PubMethod= 'Get',
                'Set',
                'Read',
                'Write',
                'Remove',
                'Flush',
                'ToString',
                'Private'

    Export-ModuleMember -Function $PubMethod 
}


#-------------------------------
#Module Functions
#-------------------------------

Function New-ConfigFile
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true)]
		[String]$Path,

        [Parameter(Mandatory=$false)]
		$Password,

        [Switch]$CachedMode,

        [Switch]$Force
	)

    $ShowLastException = { Write-Warning ($Error[0].Exception.Message+"`r`n"+($Error[0].ScriptStackTrace -split "`r`n")[0]+"`r`n"+$Error[0].InvocationInfo.PositionMessage) }

    #Check if the config file exists
    if(!(Test-Path -LiteralPath $Path) -AND !(Test-Path -Path $Path))
    {
        if($Force)
        {
            Try  { New-Item -ItemType File -Path $Path -Force -ErrorAction Stop > $null }
            Catch{ Write-Error "Unable to create the config file at $Path"; Return }
        }
        else
        { Write-Error "Config file not found at: $Path`nUse the -Force switch to automatically create it."; Return }
    }

    #Construct the ConfigFile object 
    $cfObj = New-Module -AsCustomObject -ScriptBlock $CfgClass -ArgumentList $Path, $Password, $CachedMode
    $cfObj.PSTypeNames.Clear()
    $cfObj.PSTypeNames.Add('Config.File')

    #Create the object's Script Property members
    $cfObj | Add-Member -MemberType ScriptProperty -Name "Path" -Value { $This.Private("Path") }
    $cfObj | Add-Member -MemberType ScriptProperty -Name "Count" -Value { $This.Private("Count") }
    $cfObj | Add-Member -MemberType ScriptProperty -Name "Names" -Value { $This.Private("Names") }
    $cfObj | Add-Member -MemberType ScriptProperty -Name "Password" -Value { $This.Private("Password") }{ param($pass); $This.Private("Password",$pass) }

    #Return the object
    Return $cfObj
}


#-------------------------------
#Initialize Module
#-------------------------------

New-Alias -Name 'New-Cfg' -Value 'New-ConfigFile'
$aliases = 'New-Cfg'
$functions = 'New-ConfigFile'
Export-ModuleMember -Function $functions -Alias $aliases
