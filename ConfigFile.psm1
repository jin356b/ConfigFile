#=================================================================================================
# NAME       : Configuration File Module
# AUTHOR     : Jonathon Bauer
# DATE       : 11/8/2016
# VERSION    : 2.1.0
# DESCRIPTION: This module provides functions to read and write data from a configuration file.
# DEPENDENCY : PowerShell v3.0 or higher.
# FUNCTIONS  : Set-ConfigFile [-Path] <string> [-Password <string>] [-NoSet]  [<CommonParameters>]
#              Read-Var   [[-All]] [-Path <string>] [-Password <string>] [-Return] [-Force]  [<CommonParameters>]
#              Read-Var   [-Variable] <string> [-Path <string>] [-Password <string>] [-Return] [-Force]  [<CommonParameters>]
#              Write-Var  [-Variable] <string> [[-Value] <Object>] [[-Path] <string>] [-NoClobber] [-EncryptDPAPI]  [<CommonParameters>]
#              Write-Var  [-Variable] <string> [[-Value] <Object>] [[-Path] <string>] [-NoClobber] [-EncryptAES] [-Password <string>]  [<CommonParameters>]
#              Remove-Var [-Variable] <string> [-Path <string>]  [<CommonParameters>]
#=================================================================================================

#-------------------------------
#Module Variables
#-------------------------------

[string]$script:pConfigFile = $null
[string]$script:pPassword = $null

#-------------------------------
#Module Functions
#-------------------------------

#Reads the the value of a variable (not case sensitive) from within the config file, and assigns it to the parent scope variable, unless -Return is used.
#Blank lines and lines starting with '#' will be ignored from the config file.
#Unless -Force is used, does not load the variable if an equivalent variable hasn't been declared in the script.
Function Read-ConfigVariable
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true, ParameterSetName="specify")]
		[Alias("Var")][String]$Variable,          #Name of variable to read. Read multiple variables by piping an array of strings to the function.

        [Parameter(Mandatory=$false)]
 		[Alias("File")][String]$Path,             #(Optional) The config file's path. Any previously set Config File will be temporarily ignored.

        [Parameter(Mandatory=$false)]
 		[Alias("pass")][String]$Password,         #(Optional) Password to decrypt AES 256 encoded variables.

        [Parameter(Position=0, ParameterSetName="all")]
 		[Switch]$All,                             #Read all variables from the file

        [Switch]$Return,                          #Return the variable data, instead of assigning it directly to the parent variable

        [Switch]$Force                            #Loads variables as strings if they're not declared in the parent script
	)

    Begin
    {
        #Initialize Variables
        $VarProcessed = 0

        Write-Debug "[String] Path: $Path"
        Write-Debug "[String] Password: $Password"
        Write-Debug "[Switch] All: $All"
        Write-Debug "[Switch] Return: $Return"

        #Validate $Path Parameters
        if($Path)
        {
            $Path = Set-ConfigFile $Path -NoSet
        }
        elseif($script:pConfigFile)
        {
            $Path = $script:pConfigFile
        }
        else
        {
            Throw "Unable to read variable(s). Must provide the config file's path."
        }
        Write-Verbose "Using config file: $Path"

        #Validate & Format Input Parameters
        if($All)
        {
            Write-Verbose "Reading all variables."
        }
        elseif($PSCmdlet.ParameterSetName -eq "specify")
        {
            Write-Verbose "Reading only specific variables."
        }
        else
        {
            Throw "Unable to read variable(s). Must supply the variable name(s) or specify -All."
        }

        #Load Password Data
        if(!$Password -and $script:pPassword)
        {
            $Password = $script:pPassword
            Write-Debug "[String] Password: $Password"
        }

        #Read Config File
        $err = $false
        Try
        {
            [string[]]$ConfLines = Get-Content $Path -ErrorAction Stop
        }
        Catch
        {
            Write-Debug $_.Exception.Message
            $err = $true
        }
        if($err)
        {
            Throw "Unable to read config file: $Path"
        }
    }

    Process
    {
        if(!$All)
        {
            Write-Debug "[String] Variable: $Variable"

            #Validate variable name
            if($Variable -notmatch '^[a-zA-Z0-9\-_.]*$')
            {
                Write-Warning "The variable name `$$Variable contains invalid characters."
                Return
            }
        }

        #Parse the config file line by line
        $value = $null
        $found = $false
        foreach($line in $ConfLines)
        {
            $line = $line.Trim() #Remove whitespaces

            if(($line[0] -ne '#') -and ($line[0] -ne '=') -and ($line -like '*=*')) #ensure line is not a comment, and contains an equals sign
            {
                #Extract the variable name from the config file
                $ConfVarName = $line.Substring(0,$line.IndexOf('='))

                #Handle specially formatted data
                $type = $null
                if($ConfVarName.EndsWith("[array]") -or $ConfVarName.EndsWith("[]")){ $type = "array" } #Include [] for backward compatibility
                elseif($ConfVarName.EndsWith("[dpapi]")){ $type = "dpapi" }
                elseif($ConfVarName.EndsWith("[aes256]")){ $type = "aes256" }

                #Remove type data for the name
                if($ConfVarName.Contains("["))
                {
                    $ConfVarName = $ConfVarName.Substring(0,$ConfVarName.IndexOf('['))
                }

                #Compare the variable name from the config file with the one we're searching for
                if(($Variable -eq $ConfVarName) -or $All)
                {
                    Write-Debug "Processing variable `$$ConfVarName"
                    $found = $true
                    $value = $line.Substring($line.IndexOf('=')+1) 

                    $err = $false
                    Try
                    {
                        $ParentVar = Get-Variable -scope 2 -name $ConfVarName -ErrorAction Stop
                        if($ParentVar.Value -eq $null)
                        {
                            $err = $true
                        }
                    }
                    Catch
                    {
                        Write-Debug $_.Exception.Message
                        $err = $true
                    }

                    if($err -and !$Force)
                    {
                        Write-Warning "The variable `$$ConfVarName has not been declared yet. It will not be imported."
                    }
                    elseif($err -and $Force)
                    {
                        Write-Verbose "Reading variable data for `$$ConfVarName as a string."
                        $ParentVarType = [string]
                        $err = $false
                    }
                    else
                    {
                        Write-Verbose "Pulling data type for `$$ConfVarName from the parent script."      
                        $ParentVarType = $ParentVar.Value.GetType() #Preserve data type
                    }

                    if(!$err) #Used to skip an iteration when an error occurs and -All is set
                    {
                        # Format the variable data
                        if($type -eq "array") # ARRAY DATA
                        {
                            $value = Invoke-Expression $value #Convert string into an array of variables
                        }
                        elseif($ParentVarType.Name -eq "Boolean") # BOOLEAN DATA
                        {
                            if($value -eq "False")
                            {
                                $value = $false
                            }
                            else
                            {
                                $value = $true
                            }
                        }
                        elseif($ParentVarType.Name -eq "String") # STRING DATA
                        {
                            if($type -eq "dpapi") #Read DPAPI encrypted string
                            {
                                $err=$false
                                Try
                                {
                                    [byte[]]$hashbytes = [System.Convert]::FromBase64String($value) #Convert Base64 to byte array
                                    $value = ([System.BitConverter]::ToString($hashbytes)).replace('-','') #Convert byte array to hex string
                                    $SecureStr = ConvertTo-SecureString -String $value.Trim() -ErrorAction Stop
                                    $value = (New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList "user", $SecureStr -ErrorAction Stop).GetNetworkCredential().Password 
                                }
                                Catch
                                {
                                    Write-Debug $_.Exception.Message
                                    $err=$true
                                    Write-Warning "Unable to decrypt `$$ConfVarName using Windows DPAPI."
                                }
                            }
                            elseif($type -eq "aes256") #Read AES256 encrypted string
                            {
                                if(!$Password) #Verify the password was specified
                                {
                                    Write-Warning "The -Password parameter was not specified. Unable to decrypt `$$ConfVarName using 256-bit AES."
                                    $err=$true
                                }
                                else
                                {
                                    $err=$false
                                    Try
                                    {
                                        $AEScsp = New-Object System.Security.Cryptography.AesCryptoServiceProvider
                                        $AEScsp.Key = ConvertToKey $Password
                                        $AEScsp.IV = $AEScsp.Key[8..23]

                                        $inBlock = [System.Convert]::FromBase64String($Value)
                                        $xfrm = $AEScsp.CreateDecryptor()
                                        $outBlock = $xfrm.TransformFinalBlock($inBlock, 0, $inBlock.Length) 
                                        $Value = [System.Text.UnicodeEncoding]::Unicode.GetString($outBlock) 
                                    }
                                    Catch
                                    {
                                        Write-Debug $_.Exception.Message
                                        $err=$true
                                        Write-Warning "Unable to decrypt `$$ConfVarName using 256-bit AES."
                                    }
                                }
                            }

                            $value = $value.replace('`n',"`n") #Handle new line chars
                        }
                        else # OTHER DATA
                        {
                            $value = $value -as $ParentVarType #Preserve data type
                        }

                        #If there were no errors formatting the value data
                        if(!$err)
                        {
                            # Output the variable
                            if($Return)
                            {
                                Write-Verbose "Leaving data for variable `$$ConfVarName on the pipeline."
                                $VarProcessed++
                                $value
                            }
                            else
                            {
                                Try
                                {
                                    Set-Variable -scope 2 -Name $ConfVarName -Value $value -ErrorAction Stop #Set parent variable
                                    Write-Verbose "Updating `$$ConfVarName with data from config file."
                                    $VarProcessed++
                                }
                                Catch
                                {
                                    Write-Debug $_.Exception.Message
                                    Write-Warning "Unable to update `$$ConfVarName in the parent scope."
                                }
                            }
                        }
                    }

                    if(!$All){ Return } #If we're not reading everything, we can move onto the next variable in the pipe
                }
            }
        }

        if(!$found -and $Variable)
        {
            Write-Warning "Unable to find variable `$$Variable within the config file."
        }
    }

    End
    {
        if($VarProcessed -gt 0)
        {
            Write-Verbose "Read $VarProcessed variable(s) from the config file: $Path"
        }
        else
        {
            Throw "No variables were read from the config file."
        }
    }
}

#Writes the specified variable to the config file. If the variable already exists, it's value is overwritten, unless -NoClobber is used.
#Variable names must be alphanumeric (not case sensitive), with the only allowed special characters being '.', '-', and '_'. Spaces are not allowed.
#If a string variable contains new line characters, they will be replaced with `n in the config file.
Function Write-ConfigVariable
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[Alias("Var")][String]$Variable,          #Name of variable to write. Write multiple variables by piping an array of strings to the function.

        [Parameter(Position=1, Mandatory=$false)]
        [AllowEmptyString()]
 		[Alias("Val")]$Value,                     #(Optional) The data to assign to the new variable. If multiple variables are passed by the pipeline, this value will be assigned to all of them.

        [Parameter(Position=2, Mandatory=$false)]
 		[Alias("File")][String]$Path,             #(Optional) The config file's path. Any previously set Config File will be temporarily ignored.

 		[Alias("nc")][Switch]$NoClobber,          #Will prevent the overwriting of variables when set. Variables are overwritten by default.

        [Parameter(ParameterSetName="DPAPI")]
        [Alias("dpapi")][Switch]$EncryptDPAPI,    #Will encrypt string variables with the Windows Data Protection API (DPAPI).

        [Parameter(ParameterSetName="AES")]
        [Alias("aes")][Switch]$EncryptAES,        #Will encrypt string variables with the Windows Data Protection API (DPAPI).

        [Parameter(Mandatory=$false, ParameterSetName="AES")]
 		[Alias("pass")][String]$Password          #(Optional) Password to encrypt string variables using FIPS compliant AES 256.
	)

    Begin
    {
        #Initialize Variables
        $VarProcessed = 0
        $UseParent = $false

        Write-Debug "[String] Path: $Path"
        Write-Debug "[String] Password: $Password"
        Write-Debug "[Switch] Encrypt: $Encrypt"
        Write-Debug "[Switch] NoClobbber: $NoClobber"

        #Validate $Path Parameters
        if($Path)
        {
            $Path = Set-ConfigFile $Path -NoSet
        }
        elseif($script:pConfigFile)
        {
            $Path = $script:pConfigFile
        }
        else
        {
            Throw "Unable to write variable(s). Must provide the config file's path."
        }
        Write-Verbose "Using config file: $Path"

        #Determine value source
        if($Value -eq $null) #value wasn't specified
        {
            $UseParent = $true
        }

        #Load Password Data if -Encrypt was specified
        if($EncryptAES -and !$Password)
        {
            if($script:pPassword)
            {
                $Password = $script:pPassword
                Write-Debug "[String] Password: $Password"
            }
            else
            {
                Throw "The -Password parameter must be specified to use AES 256 encryption."
            }
        }

        #Read Config File
        $err = $false
        Try
        {
            [string[]]$ConfLines = Get-Content $Path
        }
        Catch
        {
            Write-Debug $_.Exception.Message -ErrorAction Stop
            $err = $true
        }
        if($err)
        {
            Throw "Unable to read config file: $Path"
        }
    }
    
    Process
    {
        Write-Debug "[String] Variable: $Variable"
        Write-Debug "[Object] Value: $Value"

        #Validate variable name
        if($Variable -notmatch '^[a-zA-Z0-9\-_.]*$')
        {
            Write-Warning "The variable name `"$Variable`" contains invalid characters."
            Return
        }
        
        #Populate variable value
        if($UseParent)
        {
            Write-Verbose "Pulling `"$Variable`"'s value from the parent script."
            #Value wasn't specified, so we need to pull it from the parent scope
            Try
            {
                $ParentVar = Get-Variable -scope 2 -name $Variable -ErrorAction Stop
            }
            Catch
            {
                Write-Warning "No value specified, and the variable `"$Variable`" has not been declared. Skipping."
                return
            }

            $Value = $ParentVar.Value

            Write-Debug "[Object] Value: $Value"
        }

        #Determine if variable is already present in config file
        $found = $false
        for($i=0;$i -lt $ConfLines.count;$i++)
        {
            $line = $ConfLines[$i]
            $line = $line.Trim() #Remove whitespaces

            if(($line[0] -ne '#') -and ($line[0] -ne '=') -and ($line -like '*=*')) #ensure line is not a comment, and contains an equals sign
            {
                $ConfVarName = $line.Substring(0,$line.IndexOf('='))
            
                if($ConfVarName.Contains("[")) #Handle type names
                {
                    $ConfVarName = $line.Substring(0,$line.IndexOf('['))
                }
                
                if($ConfVarName -eq $Variable)
                {
                    $found = $true
                    break #Preserves $i at the correct index
                }
            }
        }

        #Skip variables that exist when -NoClobber is set
        if($found -and $NoClobber)
        {
            Write-Verbose "NoClobber is enabled. Skipping $Variable because it already exists within $Path."
            $VarProcessed++
            Return
        }

        #Construct DataType specific Value string
        if($Value.GetType().IsArray) #ARRAY
        {
            $valstring = "$Variable[array]="
            foreach($val in $Value) #Assemble string representation of array data
            {
                if($val.GetType().Name -like "Int*") #Handling number data
                {
                    $valstring += "$val, "
                }
                else #Handle all other data as string
                {
                    $valstring += "`"$val`", "
                }
            }
            $valstring = $valstring.Substring(0,$valstring.LastIndexOf(',')) #remove final comma and space
        }
        elseif($Value.GetType() -eq [string]) #STRING
        {
            #Handle line breaks
            $Value = $Value.Replace("`n",'`n')

            if($EncryptAES) #Use AES256 encryption
            {
                $AEScsp = New-Object System.Security.Cryptography.AesCryptoServiceProvider
                $AEScsp.Key = ConvertToKey $Password
                $AEScsp.IV = $AEScsp.Key[8..23]
            
                $inBlock = [System.Text.UnicodeEncoding]::Unicode.getbytes($Value) 
                $xfrm = $AEScsp.CreateEncryptor() 
                $outBlock = $xfrm.TransformFinalBlock($inBlock, 0, $inBlock.Length) 
                $Value = [System.Convert]::ToBase64String($outBlock)
                                 
                $valstring = "$Variable[aes256]=$Value"
            }
            elseif($EncryptDPAPI) #Use DPAPI encryption
            {
                $SecureStr = ConvertTo-SecureString -String $Value -AsPlainText -Force #Encrypt string using DPAPI
                $Value = ConvertFrom-SecureString -SecureString $SecureStr #Get text version of encrypted string

                [byte[]]$hashbytes = $Value -split '(..)' | ? { $_ } | % { [convert]::ToByte($_,16) } #Convert hex string to byte array
                $Value = [System.Convert]::ToBase64String($hashbytes) #Convert byte array to Base64

                $valstring = "$Variable[dpapi]=$Value"
            }
            else #Normal Strings
            {
                $valstring = "$Variable=$Value"
            }
        }
        elseif($Value.GetType() -eq [datetime]) #DATE TIME
        {
            $valstring = "$Variable=$(Get-date $Value -format "yyyy-MM-dd HH:mm:ss.fffffff")"
        }
        else #All other data types
        {
            $valstring = "$Variable=$Value" 
        }

        #Add new value data to $lines
        if($found)
        {
            $ConfLines[$i] = $valstring #Substitute the new value
            Write-Verbose "Overwriting $Variable within $Path."
        }
        else
        {
            $ConfLines += $valstring #Add the new value to the end
            Write-Verbose "Adding $Variable to $Path."
        }

        $VarProcessed++ #Increment the write count
    }

    End
    {
        if($VarProcessed -gt 0)
        {
            #Output the new data
            $err = $false
            Try
            {
                Set-Content -Path $Path -Value $ConfLines -ErrorAction Stop
            }
            Catch
            {
                Write-Debug $_.Exception.Message
                $err = $true
            }
            if($err)
            {
                Throw "Unable to write data back to config file: $Path"
            }
        
            Write-Verbose "Wrote $VarProcessed variable(s) to the config file: $Path"
        }
        else
        {
            Throw "No variables were written back to the config file."
        }
    }
}

#Removes the specified variables from the config file
Function Remove-ConfigVariable
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[Alias("Var")][String]$Variable,          #Name of variable to remove. Remove multiple variables by piping an array of strings to the function.

        [Parameter(Mandatory=$false)]
 		[Alias("File")][String]$Path              #(Optional) The config file's path. Any previously set Config File will be temporarily ignored.
	)

    Begin
    {
        #Initialize Variables
        $VarProcessed = 0

        Write-Debug "[String] Path: $Path"

        #Validate $Path Parameters
        if($Path)
        {
            $Path = Set-ConfigFile $Path -NoSet
        }
        elseif($script:pConfigFile)
        {
            $Path = $script:pConfigFile
        }
        else
        {
            Throw "Unable to remove variable(s). Must provide the config file's path."
        }
        Write-Verbose "Using config file: $Path"

        #Read Config File
        $err = $false
        Try
        {
            [string[]]$ConfLines = Get-Content $Path -ErrorAction Stop
        }
        Catch
        {
            Write-Debug $_.Exception.Message
            $err = $true
        }
        if($err)
        {
            Throw "Unable to read config file: $Path"
        }
    }
    
    Process
    {
        Write-Debug "[String] Variable: $Variable"

        #Validate variable name
        if($Variable -notmatch '^[a-zA-Z0-9\-_.]*$')
        {
            Write-Warning "The variable name `"$Variable`" contains invalid characters."
            Return
        }
        
        #Copy all variables, except for the one we're removing
        [string[]]$NewLines = @()
        for($i=0;$i -lt $ConfLines.count;$i++)
        {
            $line = $ConfLines[$i]
            $line = $line.Trim() #Remove whitespaces

            if(($line[0] -ne '#') -and ($line[0] -ne '=') -and ($line -like '*=*')) #ensure line is not a comment, and contains an equals sign
            {
                $ConfVarName = $line.Substring(0,$line.IndexOf('='))
            
                if($ConfVarName.Contains("[")) #Handle special types
                {
                    $ConfVarName = $line.Substring(0,$line.IndexOf('['))
                }
                
                if($ConfVarName -eq $Variable)
                {
                    Write-Verbose "Removing $Variable from $Path."
                    $VarProcessed++ #Increment the delete count
                }
                else
                {
                    $NewLines += $ConfLines[$i]
                }
            }
            else
            {
                $NewLines += $ConfLines[$i]
            }
        }

        #Update main data array
        $ConfLines = $NewLines
    }

    End
    {
        if($VarProcessed -gt 0)
        {
            #Output the new data
            $err = $false
            Try
            {
                Set-Content -Path $Path -Value $ConfLines -ErrorAction Stop
            }
            Catch
            {
                Write-Debug $_.Exception.Message
                $err = $true
            }
            if($err)
            {
                Throw "Unable to write data back to config file: $Path"
            }
        
            Write-Verbose "Removed $VarProcessed variable(s) from the config file: $Path"
        }
        else
        {
            Write-Verbose "No variables were removed from the config file."
        }
    }
}


#<void> Set-ConfigFile ".\config.ini" ( -NoSet )
Function Set-ConfigFile
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true)]
		[Alias("file")][String]$Path,             #The full or relative path to the config file.
        
        [Parameter(Mandatory=$false)]
 		[Alias("pass")][String]$Password,         #(Optional) Password to decrypt AES 256 encoded variables.

        [Switch]$NoSet                            #The config file path is validated and initialized, but not set. Instead the full path is returned.
	)

    Write-Debug "[String] Path is: $Path"
    Write-Debug "[String] Password is: $Password"
    Write-Debug "[Switch] NoSet is: $NoSet"
    Write-Debug "Current config file: $script:pConfigFile"

    #Store password
    if($Password)
    {
        $script:pPassword = $Password
    }

    #Ensure config file exists, and resolve full path
    if(Test-Path $Path)
    {
        Write-Debug "Config file $Path already exists"
        #Resolve Full Path
        $Path = (Resolve-Path $Path).ProviderPath 
    }
    else
    {
        Write-Debug "Config file $Path does not exist"
        #Create File
        Try
        {
            New-Item -ItemType File -Path $Path -Force -ErrorAction Stop > $null # Init the config file
            $Path = (Resolve-Path $Path -ErrorAction Stop).ProviderPath
        }
        Catch
        {
            Throw "The config file $Path could not be created. $($_.Exception.Message)"
        }

        Write-Verbose "Config file created at $Path"
    }

    #Set or return config file path
    if($NoSet)
    {
        Write-Verbose "Validated the config file $Path"
        Return $Path
    }
    else
    {
        $script:pConfigFile = $Path
        Write-Verbose "Config file set to $Path"
    }
}


#Converts a string of any length to a 32-bit encryption key
Function ConvertToKey($str)
{
    $SHAcsp = new-Object System.Security.Cryptography.SHA256CryptoServiceProvider
    
    Return $SHAcsp.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($str))
}


#-------------------------------
#Initialize Module
#-------------------------------

New-Alias -Name 'Read-Var' -Value 'Read-ConfigVariable'
New-Alias -Name 'Write-Var' -Value 'Write-ConfigVariable'
New-Alias -Name 'Remove-Var' -Value 'Remove-ConfigVariable'

$aliases = 'Read-Var', 'Write-Var', 'Remove-Var'

$functions = 'Read-ConfigVariable', 'Write-ConfigVariable', 'Remove-ConfigVariable', 'Set-ConfigFile'

Export-ModuleMember -Function $functions -Alias $aliases