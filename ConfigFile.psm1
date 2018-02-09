#=================================================================================================
# NAME       : Configuration File Module
# AUTHOR     : Jonathon Bauer
# DATE       : 10/11/2016
# VERSION    : 1.0.1
# DESCRIPTION: This module provides functions to read and write data from a configuration file.
# DEPENDENCY : PowerShell v3.0 or higher.
# FUNCTIONS  : <void> Read-Var "VariableName" ( -Path ".\config.ini" ) ( -All ) ( -Return )
#              <void> Write-Var "VariableName" ( -Path ".\config.ini" )
#              <void> Set-VarFile ".\config.ini"
#=================================================================================================

#-------------------------------
#Module Variables
#-------------------------------

[string]$script:priv_ConfigFile = $null

#-------------------------------
#Module Functions
#-------------------------------

#Reads the the value of a variable (not case sensitive) from within the config file, and assigns it to the parent scope variable as a string.
#Blank lines and lines starting with '#' will be ignored from the config file.
#Does not load the variable if an equivalent variable hasn't been declared in the script.
Function Read-Var
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$false)]
		[Alias("var")][String]$VarName="",        #Name of variable to read. Must be a String, and optionally contain a comma seperated list of variables.

        [Parameter(Position=1, Mandatory=$false)]
 		[Alias("file")][String]$Path,             #Optionally, specify the config file's path

        [Switch]$All,                             #Read all variables from the file

        [Switch]$Return                           #Return the variable data, instead of assigning it directly to the parent variable
	)

    #Validate & Format Input Parameters
    if(!$VarName -and !$All)
    {
        Throw "Unable to read variable(s). Must supply the variable name or specify -All."
    }
    if(!$Path -and !$script:priv_ConfigFile)
    {
        Throw "Unable to read variable(s). Must provide the config file's path."
    }
    if(!$Path)
    {
        $Path = $script:priv_ConfigFile
    }

    if($VarName)
    {
        [string[]]$VarList = $VarName.split(",")
        [string[]]$TmpList = $null
        foreach($Var in $VarList)
        {
            if($Var -match '^[a-zA-Z0-9\-_.]*$')
            {
                $TmpList += $Var
            }
            else
            {
                Write-Verbose "The variable name `"$Var`" contains invalid characters. Ignoring."
            }
        }
        $VarList = $TmpList
        Write-Verbose "Reading $($VarList.Count) variable(s)."    
    }
    else
    {
        [string[]]$VarList = $null
        Write-Verbose "Reading all variables."
    }
    
    #Ensure the config file exists
    if(!(Test-Path $Path)) 
    {
        Throw "Unable to read variable(s). The config file $Path could not be found."
    }

    #Read in the config file
    Try
    {
        [string[]]$lines = Get-Content -Path $Path 
    }
    Catch
    {
        Throw "Unable to read variable(s). The config file $Path could not be read."
    }

    #Parse the config file line by line
    $rc = 0
    $value = $null
    foreach($line in $lines)
    {
        $line = $line.Trim() #Remove whitespaces

        if(($line[0] -ne '#') -and ($line[0] -ne '=') -and ($line -like '*=*')) #ensure line is not a comment, and contains an equals sign
        {
            #Extract the variable name from the config file
            $IsArray = $false
            $ConfVarName = $line.Substring(0,$line.IndexOf('='))
            if($ConfVarName.EndsWith("[]")) #Handle arrays
            {
                $ConfVarName = $ConfVarName.Substring(0,$ConfVarName.IndexOf('['))
                $IsArray = $true
            }

            #Compare the variable name from the config file with the one we're searching for
            if(($VarList -contains $ConfVarName) -or $All)
            {
                $value = $line.Substring($line.IndexOf('=')+1) 
                $ParentVar = Get-Variable -scope 2 -name $ConfVarName -erroraction SilentlyContinue

                if($ParentVar -eq $null) 
                {
                    # The Parent variable doesn't exists
                    Write-Verbose "The variable `$$ConfVarName has not been declared yet. It will not be imported."
                }
                else 
                {
                    # Format the variable data
                    if($IsArray) # ARRAY DATA
                    {
                        $value = Invoke-Expression $value #Convert string into an array of variables
                    }
                    elseif($ParentVar.Value.GetType().Name -eq "Boolean") # BOOLEAN DATA
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
                    else # OTHER DATA
                    {
                        $value = $value -as $ParentVar.Value.GetType() #Preserve data type
                    }

                    # Output the variable
                    if($Return)
                    {
                        $value 
                        Write-Verbose "Put data for variable $ConfVarName on the pipeline."
                    }
                    else
                    {
                        Set-Variable -scope 2 -Name $ConfVarName -Value $value #Set parent variable
                        Write-Verbose "Updating $ConfVarName with data from config file."
                    }

                    $rc++ #Increment read counter    
                }
            }
        }
    }

    Write-Verbose "Successfully read $rc variable(s) from $Path."
}

#Writes the specified variable to the config file. If the variable already exists, it's value is overwritten.
#Variable names must be alphanumeric (not case sensitive), with the only allowed special characters being '.', '-', and '_'. Spaces are not allowed.
#A string variable's value must not contain new line characters.
Function Write-Var
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true)]
		[Alias("var")][String]$VarName,           #Name of variable to write. Must be a String, and optionally contain a comma seperated list of variables.

        [Parameter(Position=1, Mandatory=$false)]
 		[Alias("file")][String]$Path              #Optionally, specify the config file's path
	)

    #Validate & Format Input Parameters
    if(!$Path -and !$script:priv_ConfigFile)
    {
        Throw "Unable to write variable(s). Must provide the config file's path."
    }
    if(!$Path)
    {
        $Path = $script:priv_ConfigFile
    }

    [string[]]$VarList = $VarName.split(",")
    [string[]]$TmpList = $null
    foreach($Var in $VarList)
    {
        if($Var -match '^[a-zA-Z0-9\-_.]*$')
        {
            $TmpList += $Var
        }
        else
        {
            Write-Verbose "The variable name `"$Var`" contains invalid characters. Ignoring."
        }
    }
    $VarList = $TmpList
    Write-Verbose "Writing $($VarList.Count) variable(s)."    

    #Ensure the config file exists
    if(!(Test-Path $Path)) 
    {
        Write-Verbose "Config file $Path not found. Generating a new one."
    
        Try
        {
            New-Item -ItemType File -Path $Path -Force  > $null # Init the config file
        }
        Catch
        {
            Throw "Unable to write variable(s). The config file $Path could not be created."
        }
    }

    #Read in the config file
    Try
    {
        [string[]]$lines = Get-Content -Path $Path 
    }
    Catch
    {
        Throw "Unable to write variable(s). The config file $Path could not be read."
    }

    #Loop through each variable we want to write
    $wc = 0
    foreach ($Var in $VarList)
    {
        #Get Variable's data
        $ParentVar = Get-Variable -scope 2 -name $Var -erroraction SilentlyContinue
        if($ParentVar -eq $null)
        {
            Write-Verbose "The variable `"$Var`" has not been declared. Skipping."
            break
        }
        else
        {
            $value = $ParentVar.Value
        }

        #Handle line breaks
        if($value.GetType() -eq [string])
        {
            $value = $value.Replace("`n"," ")
        }

        #Determine if variable is already present in config file
        $found = $false
        for($i=0;$i -lt $lines.count;$i++)
        {
            $line = $lines[$i]
            $line = $line.Trim() #Remove whitespaces

            if(($line[0] -ne '#') -and ($line[0] -ne '=') -and ($line -like '*=*')) #ensure line is not a comment, and contains an equals sign
            {
                $ConfVarName = $line.Substring(0,$line.IndexOf('='))
            
                if($ConfVarName.EndsWith("[]")) #Handle Array names
                {
                    $ConfVarName = $line.Substring(0,$line.IndexOf('['))
                }
                
                if($ConfVarName -eq $Var)
                {
                    $found = $true
                    break #Preserves $i at the correct index
                }
            }
        }

        #Construct DataType specific Value string
        if($value.GetType().IsArray) #ARRAY
        {
            $valstring = "$Var[]="
            foreach($val in $value) #Assemble string representation of array data
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
        else #All other data types
        {
            $valstring = "$Var=$value" 
        }

        #Add new value data to $lines
        if($found)
        {
            $lines[$i] = $valstring #Substitute the new value
            Write-Verbose "Overwriting $Var within $Path."
        }
        else
        {
            $lines += $valstring #Add the new value to the end
            Write-Verbose "Adding $Var to $Path."
        }

        $wc++ #Increment the write count
    }

    #Output the new data
    Try
    {
        Set-Content -Path $Path -Value $lines
    }
    Catch
    {
        Throw "Unable to write variable(s). Could not data back to the config file $Path."
    }

    Write-Verbose "Wrote $wc variable(s) to $Path."
}

#<void> Set-VarFile ".\config.ini"
Function Set-VarFile
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true)]
		[Alias("file")][String]$Path
	)

    $script:priv_ConfigFile = $Path
}


#-------------------------------
#Initialize Module
#-------------------------------

$functions = 'Read-Var', 'Write-Var', 'Set-VarFile'

Export-ModuleMember -Function $functions
