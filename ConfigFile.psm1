#=================================================================================================
# NAME       : Configuration File Module
# AUTHOR     : Jonathon Bauer
# DATE       : 11/1/2016
# VERSION    : 2.0.0
# DESCRIPTION: This module provides functions to read and write data from a configuration file.
# DEPENDENCY : PowerShell v3.0 or higher.
# FUNCTIONS  : <void> Read-ConfigVariable "VariableName" ( -Path ".\config.xml" ) ( -All ) ( -Return )
#              <void> Write-ConfigVariable "VariableName" ( -Path ".\config.xml" )
#              <void> Set-ConfigFile ".\config.xml" ( -NoSet )
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
Function Read-ConfigVariable
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

Function Open-ConfigFile
{
    [CmdletBinding()] 
    param(
        [Parameter(Position=0, Mandatory=$true)]
 		[Alias("File")][String]$Path
	)

    #Read Config File
    $err = $false
    Try
    {
        $data = Get-Content $Path
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

    #Load XML Data
    $XmlData = New-Object System.Xml.XmlDocument
    $err = $false
    Try
    {
        $XmlData.LoadXml($data)
    }
    Catch
    {
        Write-Debug $_.Exception.Message
        $err = $true
    }
    if($err)
    {
        Throw "Unable to load XML from config file: $Path"
    }

    Return $XmlData
}

Function Resolve-DataType
{
    [CmdletBinding()] 
    param(
        [Parameter(Position=0, Mandatory=$true)]
 		[Alias("Type")][String]$typeName
	)

###TODO: Handle arrays like System.Byte[], System.String[], etc...

    $TypeSets = @(@("string","System.String"),
                  @("int","System.Int32"),
                  #@("array","System.Object[]"),
                  @("bool","System.Boolean"),
                  @("hashtable","System.Collections.Hashtable"),
                  @("dictionary","System.Collections.Specialized.OrderedDictionary"),
                  @("char","System.Char"),
                  @("byte","System.Byte"),
                  @("long","System.Int64"),
                  @("decimal","System.Decimal"),
                  @("single","System.Single"),
                  @("double","System.Double"),
                  @("datetime","System.DateTime"),
                  @("xml","System.Xml.XmlDocument"),
                  @("xlmnt","System.Xml.XmlElement"),
                  @("guid","System.Guid"),
                  @("null","null"))

    if($typeName.EndsWith("[]"))
    {
        Return "array"
    }

    foreach($set in $TypeSets)
    {
        if($set[1] -eq $typeName)
        {
            Return $set[0]
        }
    }

    Return $typeName
}

#Writes the specified variable to the config file. If the variable already exists, it's value is overwritten.
#Variable names must be alphanumeric (not case sensitive), with the only allowed special characters being '.', '-', and '_'. Spaces are not allowed.
Function Write-ConfigVariable
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
		[Alias("Var")][String]$Variable,          #Name of variable to write. Write multiple variables by piping an array of strings to the function.

        [Parameter(Position=1, Mandatory=$false)]
        [AllowEmptyString()]
 		[Alias("Val")]$Value,                     #(Optional) The data to assign to the new variable. If multiple variables are passed by the pipeline, this value will be assigned to all of them.

        [Parameter(Mandatory=$false)]
 		[Alias("File")][String]$Path,             #(Optional) The config file's path. Any previously set Config File will be temporarily ignored.

 		[Alias("NC")][Switch]$NoClobber,          #Will prevent the overwriting of variables when set. Variables are overwritten by default.

        [Switch]$Force                            #Will update a variable's data type. Default behavior is to not update a variable if it already exists and the data type is different.
	)

    Begin
    {
        #Initialize Variables
        $VarProcessed = 0
        $simpleTypes = @("string","int","byte","long","decimal","single","double","guid","bool","datetime","char","xml","xlmnt")
        $complexTypes = @("array","hashtable","dictionary")

        Write-Debug "[String] Path: $Path"
        Write-Debug "[Switch] NoClobbber: $NoClobber"

        #Validate $Path Parameters
        if($Path)
        {
            $Path = Set-ConfigFile $Path -NoSet
        }
        elseif($script:priv_ConfigFile)
        {
            $Path = $script:priv_ConfigFile
        }
        else
        {
            Throw "Unable to write variable(s). Must provide the config file's path."
        }
        Write-Verbose "Using config file: $Path"

        #Load XML Data
        [xml]$XmlData = Open-ConfigFile $Path
        $VarList = $XmlData.SelectNodes("/ConfigData/*").Name
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
        if($Value -eq $null) ##########fix to be in begin
        {
            Write-Verbose "Pulling `"$Variable`"'s value from the parent script."
            #Value wasn't specified, so we need to pull it from the parent scope
            Try
            {
                $ParentVar = Get-Variable -scope 2 -name $Variable
            }
            Catch
            {
                Write-Warning "No value specified, and the variable `"$Variable`" has not been declared. Skipping."
                return
            }

            $Value = $ParentVar.Value

            Write-Debug "[Object] Value: $Value"
        }

        #Buildout Variable Template
        $ArrayType = $null
        Try
        {
            $VarDataType = $Value.GetType().ToString()
        }
        Catch
        {
            $VarDataType = "null"
        }
        $VarDataType = Resolve-DataType $VarDataType # convert to common names for known data types
        $newNode = $XmlData.CreateElement($VarDataType)
        $newNode.SetAttribute("name","$Variable")

        #Determine if variable is already present
        if($VarList -contains $Variable)
        {
            #Variable is present in config
            if($NoClobber)
            {
                Write-Verbose "Skipping variable `"$Variable`" because it already exists, and NoClobber is set."
                $VarProcessed++
                Return
            }

            #Get the exact variable case, to prevent XML matching problems
            $Variable = $VarList[$VarList.ToLower().IndexOf($Variable.ToLower())]

            #Determine data type of variable in config file
            $node = $XmlData.SelectSingleNode("/ConfigData/*[@name=`"$Variable`"]")
            $ConfVarDataType = $node.ToString()
            
            #Check for data type mis-match
            if(($VarDataType -ne $ConfVarDataType) -and (!$Force)) #Data types are different and -Force isn't set
            {
                Write-Warning "Skipping variable `"$Variable`" because it's datatype ($VarDataType) doesn't match the config file's ($ConfVarDataType). Override using the -Force switch."
                Return
            }

            #Reset the config variable's data
            $XmlData.SelectSingleNode('/ConfigData').ReplaceChild($newNode, $node) > $null
        }
        else #Variable is not present
        {
            #Create the variable node
            $XmlData.SelectSingleNode("/ConfigData").AppendChild($newNode) > $null 
        }

        #Write Variable Data based on data type
        if($simpleTypes -contains $VarDataType)
        {
            if($VarDataType -eq "datetime")
            {
                $varString = Get-date $Value -format "yyyy-MM-dd HH:mm:ss.fffffff"
            }
            elseif($VarDataType -eq "char")
            {
                $varString = ([byte]$Value).ToString()
            }
            elseif(($VarDataType -eq "xml") -or ($VarDataType -eq "xlmnt"))
            {
                $varString = $xmlData.OuterXml
            }
            else
            {
                $varString = $Value.ToString()
            }

            $newNode.AppendChild($XmlData.CreateTextNode($varString)) > $null
            Write-Verbose "Set `"$Variable`" to: $varString"
        }
        elseif($complexTypes -contains $VarDataType)
        {
            if($VarDataType -eq "array")
            {

            }
            elseif($VarDataType -eq "hashtable")
            {

            }
            elseif($VarDataType -eq "dictionary")
            {

            }
        }

        $VarProcessed++
    }

    End
    {
        if($VarProcessed -gt 0)
        {
            #Output the new data
            $err = $false
            Try
            {
                $XmlData.Save($Path)
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

Function Remove-ConfigVariable
{

}


#<void> Set-ConfigFile ".\config.xml" ( -NoSet )
Function Set-ConfigFile
{
    [CmdletBinding()] 
    param(
		[Parameter(Position=0, Mandatory=$true)]
		[Alias("file")][String]$Path,             #The full or relative path to the config file.
        
        [Switch]$NoSet                            #The config file path is validated and initialized, but not set. Instead the full path is returned.
	)

    Write-Debug "[String] Path is: $Path"
    Write-Debug "[Switch] NoSet is: $NoSet"
    Write-Debug "Current config file: $script:priv_ConfigFile"

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
            New-Item -ItemType File -Path $Path -Force  > $null # Init the config file
            $Path = (Resolve-Path $Path).ProviderPath
        }
        Catch
        {
            Throw "The config file $Path could not be created. $($_.Exception.Message)"
        }

        Write-Verbose "Config file created at $Path"
    }

    #Validate config file format
    $err = $false
    Try
    {
        $XmlData = Open-ConfigFile $Path
    }
    Catch
    {
        $XmlData = New-Object System.Xml.XmlDocument
        $err = $true
    }

    if($err -or !$XmlData.ConfigData)
    {
        #Create ConfigData element
        Write-Verbose "Initializing root element of config file."
        $XmlData.AppendChild($XmlData.CreateElement("ConfigData")) > $null
        $XmlData.Save($Path)
    }

    #Set or return config file path
    if($NoSet)
    {
        Write-Verbose "Validated the config file $Path"
        Return $Path
    }
    else
    {
        $script:priv_ConfigFile = $Path
        Write-Verbose "Config file set to $Path"
    }
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
