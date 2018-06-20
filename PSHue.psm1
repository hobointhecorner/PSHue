$defaultHueConfigPath = "$env:APPDATA\PSHue\config.json"

#######
#region OBJECTS
#######

class HuePref
{
    [string]$ComputerName
    [string]$Username
}

class HueLight
{
    [string]$Name
    [bool]$On
    [int]$Id
    [string]$Type
    [string]$ProductName
    [string]$ModelId
    [string]$ManufacturerName
    [bool]$Certified
    [object]$State
    [string]$UniqueId
    [object]$HueObject

    [string]ToString()
    {
        return $this.Name
    }

    [int]ToInt()
    {
        return $this.Id
    }
}

class HueRoom
{
    [string]$Name
    [int]$Id
    [string]$Type
    [HueLight[]]$Lights
    [object]$HueObject

    [string] ToString()
    {
        return $this.Name
    }

    [int] ToInt()
    {
        return $this.Id
    }
}

class HueScene
{
    [string]$Name
    [string]$Id
    [HueLight[]]$Lights
    [HueRoom]$Room
    [object]$HueObject

    [string] ToString()
    {
        return $this.Name
    }
}

#######
#endregion
#######

#######
#region COLORS
#######

$colorMap = @{
    Blue =         @{ hue = 46920 ; sat = 254 }
    BlueGreen =    @{ hue = 35828 ; sat = 254 }
    Green =        @{ hue = 25500 ; sat = 254 }
    Red =          @{ hue = 1     ; sat = 254 }
    Orange =       @{ hue = 2046  ; sat = 254 }
    Peach =        @{ hue = 1     ; sat = 150 }
    Purple =       @{ hue = 48771 ; sat = 254 }
    White =        @{ hue = 1     ; sat = 1   }
    Yellow =       @{ hue = 10821 ; sat = 254 }
    YellowOrange = @{ hue = 9138  ; sat = 254 }
}

<#
    .SYNOPSIS
    Converts RGB color codes to X,Y format to be used with Hue devices

    .DESCRIPTION
    This command converts a user-defined RGB color value to the equivalent X,Y value.  This is useful for quick color conversion when setting Hue light states.
    It was adapted from the Objective-C example here: https://developers.meethue.com/documentation/color-conversions-rgb-xy
#>
function Convert-RgbToXy
{
    param(
        [parameter(Mandatory=$true)]
        [ValidateRange(0,255)]
        [int]$R,
        
        [parameter(Mandatory=$true)]
        [ValidateRange(0,255)]
        [int]$G,
        
        [parameter(Mandatory=$true)]
        [ValidateRange(0,255)]
        [int]$B
    )

    #Convert RGB values to a percentage of 255
    $rP = $R / 255
    $gP = $G / 255
    $bP = $B / 255

    #Apply gamma correction
    if ($rP > .04045) { $red = [math]::Pow((($rP + .055) / (1 + .055)), 2.4) }
    else              { $red = $rP / 12.92 }

    if ($gP > .04045) { $green = [math]::Pow((($gP + .055) / (1 + .055)), 2.4) }
    else              { $green = $gP / 12.92 }
        
    if ($bP > .04045) { $blue = [math]::Pow((($bP + .055) / (1 + .055)), 2.4) }
    else              { $blue = $bp / 12.92 }

    #Convert to XYZ color space
    $x = $red * 0.664511 + $green * 0.154324 + $blue * 0.162028
    $y = $red * 0.283881 + $green * 0.668433 + $blue * 0.047685
    $z = $red * 0.000088 + $green * 0.072310 + $blue * 0.986039
        
    #Calculate XY values
    $valSum = $x + $y + $z
    $x = [math]::Round(($x / $valSum), 4)
    $y = [math]::Round(($y / $valSum), 4)

    #Build and return object
    New-Object psobject -Property @{ X = $x ; Y = $y }
}

#######
#endregion
#######

#######
#region PREFRENCES
#######

function Get-HuePref
{
    [cmdletbinding()]
    param(
        [string]$Path = $defaultHueConfigPath
    )

    process
    {
        if (Test-Path $Path)
        {
            #Pref file found, build and return Pref object
            Get-Content $Path |
                ConvertFrom-Json |
                foreach {
                    [HuePref]@{
                        ComputerName = $_.ComputerName
                        Username = $_.Username
                    }
                }
        }
        else
        {
            #No pref file found, attempt new connection to Hue Bridge
            Connect-HueBridge
        }
    }
}

function Set-HuePref
{
    [cmdletbinding()]
    param(
        [parameter(ValueFromPipeline = $true)]
        [HuePref[]]$HuePref,

        [string]$ComputerName,
        [string]$Username,

        [string]$Path = $defaultHueConfigPath,

        [switch]$PassThru
    )

    begin
    {
        #Get hue pref if object not passed in, will attempt to connect to Hue bridge if not already connected.
        if (!$HuePref) { $HuePref = Get-HuePref }
    }

    process
    {
        if ($HuePref)
        {
            foreach ($pref in $HuePref)
            {
                if (!(Test-Path $Path)) { New-Item $Path -Force -ItemType File | Out-Null }

                if ($ComputerName) { $pref.ComputerName = $ComputerName }
                if ($Username) { $pref.Username = $Username }

                ConvertTo-Json $pref | Out-File $Path -Force -ErrorAction Stop
                $defaultHuePref = $pref

                if ($PassThru) { Write-Output $pref }
            }
        }
        else
        {
            Write-Error "No hue preferences configured.  Use Connect-HueBridge to configure Hue preferences."
        }
    }
}

#######
#endregion
#######

#######
#region REQUESTS
#######

<#
    .SYNOPSIS
    Converts request parameter hashtable to JSON
#>
function Format-HueRequestBody
{
    param(
        [hashtable]$Body
    )

    $Body | ConvertTo-Json -Compress
}

<#
    .SYNOPSIS
    Converts HuePref objects to hashtables, useful for parameter splatting
#>
function Format-HueRequestParam
{
    param(
        [HuePref]$HuePref
    )

    @{ ComputerName = $HuePref.ComputerName ; Username = $HuePref.Username }
}

<#
    .SYNOPSIS
    Submits a REST request to the defined Hue Hub

    .PARAMETER ComputerName
    The IP or DNS name of your Hue Hub

    .PARAMETER Username
    The username initially fetched from your Hue when Connect-HueBridge was run

    .PARAMETER Resource
    The trailing resource to be used in the URI

    .PARAMETER Method
    The REST method to be used

    .PARAMETER Body
    The JSON string body to be sent with the request
#>
function Invoke-HueRequest
{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$ComputerName,
        [string]$Username,
        [string]$Resource,
        [ValidateSet('GET','POST','PUT')]
        [string]$Method = 'GET',
        [string]$Body
    )
    
    begin
    {
        #URL setup
        $uri = "http://$computername/api"
        if ($Username) { $uri += "/$Username" }
        if ($Resource) { $uri += "/$Resource" }

        #Invoke command parameter setup
        $param_Invoke = @{
            Method = $Method
            Uri = $uri
        }

        if ($Body)
        {
            $param_Invoke.Add('Body',$Body)
            Write-Verbose "Sending content: $Body"
        }
    }

    process
    {
        #Invoke REST method using parameters in $param_Invoke and parse the return (if any)
        if ($response = Invoke-RestMethod @param_Invoke)
        {
            if ($response | Get-Member | where { $_.Name -ieq 'error' })
            {
                Write-Verbose "Error type: $($response.Error.Type)`nAddress: $($response.Error.Address)`nDescription: $($response.Error.Description)"
                Write-Warning $($response.Error.Description)
            }
            else { Write-Output $response }
        }
    }
}

<#
    .SYNOPSIS
    Connnects to your Hue Bridge and stores login information for future connections

    .DESCRIPTION
    This connects to your Hue Bridge to request a username to use in future communication.
    After starting the command, the user has to press the Link button on the top of the Hue Hub within the MaxTimeSec time (Default of 60).

    .PARAMETER ComputerName
    The DNS name or IP address of the Hue Bridge to connect to

    .PARAMETER DeviceType
    The device name to register with when communicating with the Hue

    .PARAMETER MaxTimeSec
    The maximum time to wait for the user to press the Hue Hub Link button
#>
function Connect-HueBridge
{
    [cmdletbinding()]
    param(
        [string]$ComputerName = (Read-Host -Prompt "Enter the DNS name or IP address of your Hue Bridge"),
        [string]$DeviceType = "$env:USERNAME`.$env:COMPUTERNAME`:PSHue",
        [int]$MaxTimeSec = 60
    )

    begin
    {
        #Variable setup
        $oldInfoPref = $InformationPreference
        $InformationPreference = 'Continue'

        $username = $null
        $checkTime = (Get-Date).AddSeconds($MaxTimeSec)
        $requestBody = Format-HueRequestBody @{ devicetype = $DeviceType }        
    }

    process
    {
        Write-Information "Attempting to connect to hue bridge at $ComputerName.  Please press the link button on the bridge to continue."
        while ((Get-Date) -lt $checkTime)
        {
            #Query if the requested DeviceType is authorized until successful, or the timeout period expires
            if ($connectRequest = Invoke-HueRequest -ComputerName $ComputerName -Method POST -Body $requestBody -ErrorAction Stop)
            {
                if ($connectRequest.Success)
                {
                    Write-Verbose "Got username: $($connectRequest.Success.Username)"
                    $username = $connectRequest.Success.Username
                    break
                }
            }

            sleep 1
        }

        if ($username)
        {
            #Module authorized successfully, save connection information
            Set-HuePref -HuePref ([HuePref]@{ ComputerName = $ComputerName ; Username = $username }) -passthru
        }
        else
        {
            throw "Failed to connect to Hue Bridge."
        }
    }

    end
    {
        $InformationPreference = $oldInfoPref
    }
}

#######
#endregion
#######

#######
#region LIGHTS AND ROOMS
#######

<#
    .SYNOPSIS
    Gets a list of Hue lights associated with your Hue Bridge

    .PARAMETER Name
    Filters lights by name using a simple wildcard search

    .PARAMETER Id
    Fetches light(s) from a list of one or more IDs

    .PARAMETER HueGroup
    Fetches lights associated with defined HueRoom or HueScene object(s).  Essentially just a shortcut for 'Get-Hue[Light | Scene] | select -ExpandProperty Lights'

    .PARAMETER HuePref
    The preferences used to connect to the Hue Bridge.  The default value is set when Connect-HueBridge is successfully run.
#>
function Get-HueLight
{
    [cmdletbinding()]
    param(
        [string]$Name = '*',

        [int[]]$Id,
        [Parameter(ValueFromPipeline=$true)]
        [object[]]$HueGroup,
        
        [ValidateNotNullOrEmpty()]
        [HuePref]$HuePref = $defaultHuePref
    )

    begin
    {
        #Parameter setup
        $param_HueRequest = Format-HueRequestParam $HuePref
    }  

    process
    {
        if ($Id)
        {
            foreach ($i in $Id)
            {
                #Fetch and build HueLight object(s)
                Invoke-HueRequest @param_HueRequest -Resource "lights/$i" |
                    where { $_.Name -like $Name } |
                    sort Name | 
                    foreach {
                        [HueLight]@{
                            Name = $_.Name
                            On = $_.State.On -ieq 'true'
                            Id = $i
                            Type = $_.Type
                            ProductName = $_.ProductName
                            ModelId = $_.ModelId
                            ManufacturerName = $_.ManufacturerName
                            Certified = $_.Capabilities.Certified -ieq 'true'
                            State = $_.State
                            UniqueId = $_.UniqueId
                            HueObject = $_
                        }
                    }
            }
        }
        elseif ($HueGroup)
        {
            $HueGroup | select -ExpandProperty Lights
        }
        else
        {
            #Build a list of light ID's to query the Hue API individually for full info
            $id = Invoke-HueRequest @param_HueRequest -Resource "lights" |
                    Get-Member |
                        where { $_.Name -match "\d{1,4}" } |
                        select -ExpandProperty Name
                    
            #Run another query to get full info
            Get-HueLight -Id $id -Name $Name -HuePref $HuePref     
        }
    }
}

<#
    .SYNOPSIS
    Gets a list of rooms associated with your Hue Bridge

    .PARAMETER Name
    Finds room(s) by name using a simple wildcard search

    .PARAMETER Id
    Fetches rooms(s) from a list of one or more IDs

    .PARAMETER LightList
    List of lights to search through in order to list lights available in the room.  You shouldn't have to worry about using this.

    .PARAMETER HuePref
    The preferences used to connect to the Hue Bridge.  The default value is set when Connect-HueBridge is successfully run.
#>
function Get-HueRoom
{
    [cmdletbinding()]
    param(
        [string]$Name = '*',
        [int[]]$Id,

        [HueLight[]]$LightList = (Get-HueLight),
        
        [ValidateNotNullOrEmpty()]
        [HuePref]$HuePref = $defaultHuePref
    )

    begin
    {        
        #Parameter setup
        $param_HueRequest = Format-HueRequestParam $HuePref
    }

    process
    {
        if ($Id)
        {
            foreach ($i in $Id)
            {
                #Fetch groups of type 'room' and build HueRoom object(s)
                Invoke-HueRequest @param_HueRequest -Resource "groups/$i" |
                    where { $_.Type -ieq 'room' } |
                    where { $_.Name -like $Name } | 
                    sort Name |
                    foreach {
                        $room = $_
                        [HueRoom]@{
                            Name = $_.Name
                            Type = $_.Class
                            Id = $i
                            Lights =  $LightList | where { $_.Id -in $room.Lights }
                            HueObject = $_
                        }
                    }
            }
        }
        else
        {
            #Build a list of room IDs to query the Hue API individually for full info
            $id = Invoke-HueRequest @param_HueRequest -Resource "groups" |
                    Get-Member |
                        where { $_.Name -match "\d{1,3}" } |
                        select -ExpandProperty Name
            
            #Run another query to get the full info for each light
            Get-HueRoom -Id $id -Name $Name -LightList $LightList -HuePref $HuePref
        }
    }
}

<#
    .SYNOPSIS
    Changes the state of the defined lights or rooms

    .DESCRIPTION
    This cmdlet is used to modify the state of defined lights, or lights in defined rooms.  You can use this to turn the lights on/off, set brightness, color, and other settings.

    .PARAMETER HueObject
    The Hue Light(s) or Room(s) whose state you wish to change

    .PARAMETER Light
    Defines that lights are being set.  This is assumed by default

    .PARAMETER Room
    Defines that rooms are being set.

    .PARAMETER PowerState
    The power state of the light (on/off)

    .PARAMETER BrightnessPercent
    The 0-100 value representing the brightness of the light(s)

    .PARAMETER Effect
    Turns on special light effects
    
      NOTES:
        -Be sure to set to 'none' when done.
        -Overrides all other non-power parameters


    .PARAMETER Color
    Sets light(s) to a predefined color

      NOTES:
        -Parameter will be overridden by Effect, X/Y, and Temperature parameters


    .PARAMETER Saturation
    Sets the saturation value of the light(s) color (0-254)

      NOTES:
        -Parameter will be overridden by all other non-power parameters

    .PARAMETER Hue
    Sets the hue value of the light(s) color (0-65535).  This is a wraparound value, with 0 and 65535 being red, 46920 blue, and 25500 green

      NOTES:
        -Parameter will be overridden by all other non-power parameters

    .PARAMETER X
    Sets the X value in an X,Y color value pair (0-1)

      NOTES:
        -Must be used in conjunction with the X parameter
        -Will be overridden by Effect parameter
    

    .PARAMETER Y
    Sets the Y value in an X,Y color value pair (0-1).
    
      NOTES:
        -Must be used in conjunction with the X parameter
        -Will be overridden by Effect parameter
    

    .PARAMETER Temperature
    Sets the white color temperature of lights that support this feature

      NOTES:
        -Will be overridden by Effect and X/Y parameters

    .PARAMETER HuePref
    The preferences used to connect to the Hue Bridge.  The default value is set when Connect-HueBridge is successfully run.
#>
function Set-HueState
{
    [cmdletbinding()]
    param(
        [parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [object[]]$HueObject, 

        [ValidateSet('On','Off')]
        [string]$PowerState,
            
        [ValidateRange(0,100)]
        [int]$BrightnessPercent,        
        
        [ValidateSet('ColorLoop','None')]
        [string]$Effect,

        [ValidateSet('Blue','BlueGreen','Green','Red','Orange','Peach','Purple','White','Yellow','YellowOrange')]
        [string]$Color,        

        [ValidateRange(0,254)]
        [int]$Saturation,        
        [ValidateRange(0,65535)]
        [int]$Hue,        
        
        [ValidateRange(0,1)]
        [float]$X,       
        [ValidateRange(0,1)]
        [float]$Y,        
        
        [ValidateRange(0,500)]
        [int]$Temperature,
        
        [ValidateNotNullOrEmpty()]
        [HuePref]$HuePref = $defaultHuePref
    )

    begin
    {    
        #Parameter setup
        if (($x -and !$y) -or ($y -and !$x)) { throw "You must define both X and Y when using X,Y color notation." }            
        $param_HueRequest = Format-HueRequestParam $HuePref        
        
        #Set up request body
        $body = @{}
                
        #Set power state
        if ($PowerState) { $body.Add('on',($PowerState -ieq 'on')) }

        #Set brightness.  Max value in the Hue API is 254
        if ($BrightnessPercent) { $body.Add('bri',([math]::Round(($BrightnessPercent * .01) * 254))) }

        #Set light color
        if ($Effect)          { $body.Add('effect',$Effect.ToLower()) }
        elseif ($X -or $Y)    { $body.Add('xy',@($X,$Y)) }
        elseif ($Temperature) { $body.Add('ct',$Temperature) }
        elseif ($Color)       { 'sat','hue' | foreach { $body.Add($_,$colorMap[$Color][$_]) } }
        else
        {   
            if ($Saturation) { $body.Add('sat',$Saturation) }            
            if ($Hue) { $body.Add('hue',$Hue) }
        }
        
        if ($body -eq @{}) { throw "You must define a setting to change." }
        $body = Format-HueRequestBody $body
    }

    process
    {
        foreach ($o in $HueObject)
        {
            #Get uri resource to use based on type of object passed in
            switch ($o.GetType().FullName)
            {
                "HueLight" { $resource = "lights/$($o.Id)/state" ; break }
                "HueRoom" { $resource = "groups/$($o.Id)/action" ; break }
                default { throw "Could not detect object type $($o.gettype().fullname))" }
            }

            Invoke-HueRequest @param_HueRequest -Resource $resource -Method PUT -Body $body
        }
    }
}

#######
#endregion
#######

#######
#region SCENES
#######

<#
    .SYNOPSIS
    Gets a list of Scenes associated with your Hue Bridge

    .PARAMETER Name
    Finds scenes(s) by name using a simple wildcard search

    .PARAMETER RoomName
    Finds scenes(s) by room name using a simple wildcard search

    .PARAMETER Id
    Fetches rooms(s) from a list of one or more IDs
    
    .PARAMETER RoomList
    List of rooms possibly associated with the scene.  You shouldn't have to worry about using this.
    
    .PARAMETER LightList
    List of lights possibly associated with the scene.  You shouldn't have to worry about using this.

    .PARAMETER HuePref
    The preferences used to connect to the Hue Bridge.  The default value is set when Connect-HueBridge is successfully run.
#>
function Get-HueScene
{
    [cmdletbinding()]
    param(
        [string]$Name = '*',

        [parameter(ValueFromPipeline=$true)]
        [object[]]$Room,        
        [string[]]$Id,

        [HueRoom[]]$RoomList = (Get-HueRoom),
        [HueLight[]]$LightList = (Get-HueLight),

        [ValidateNotNullOrEmpty()]
        [HuePref]$HuePref = $defaultHuePref
    )

    begin
    {    
        $param_HueRequest = Format-HueRequestParam $HuePref
    }

    process
    {
        if ($Id)
        {
            foreach ($i in $Id)
            {
                #Fetch and build HueScene object(s)
                Invoke-HueRequest @param_HueRequest -Resource "scenes/$i" |
                    where { $_.Name -like $Name } |
                    sort Name |
                    foreach {
                        $scene = $_
                        [HueScene]@{
                            Name = $_.Name
                            Id = $i
                            Room = $RoomList | where { ($scene.lights | select -f 1) -in $_.lights.id }
                            Lights = $LightList | where { $_.Id -in $scene.Lights }
                            HueObject = $_
                        }
                    }
            }
        }
        elseif ($Room)
        {
            Get-HueScene -Name $Name -RoomList $RoomList -LightList $LightList -HuePref $HuePref | where { $_.Room.Name -in $Room.Name }
        }
        else
        {
            #Build a list of scene IDs to query the Hue API individually for full info
            $id = Invoke-HueRequest @param_HueRequest -Resource "scenes" |
                    Get-Member |
                        where { $_.MemberType -ieq 'noteproperty' } |
                        select -ExpandProperty Name 
            
            #Run another query to get full scene info
            Get-HueScene -Id $id -Name $Name -RoomList $RoomList -LightList $LightList -HuePref $HuePref
        }
    }
}

<#
    .SYNOPSIS
    Starts a Hue Scene

    .PARAMETER Scene
    The HueScene object to start

    .PARAMETER HuePref
    The preferences used to connect to the Hue Bridge.  The default value is set when Connect-HueBridge is successfully run.
#>
function Start-HueScene
{
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [object[]]$Scene,
        
        [ValidateNotNullOrEmpty()]
        [HuePref]$HuePref = $defaultHuePref
    )

    begin
    {        
        #Parameter setup
        $param_HueRequest = Format-HueRequestParam $HuePref
    }

    process
    {
        foreach ($s in $Scene)
        {
            $body = Format-HueRequestBody @{ scene = $s.Id }
            Invoke-HueRequest @param_HueRequest -Resource "groups/$($s.Room.Id)/action" -Method PUT -Body $body
        }
    }
}

#######
#endregion
#######

#Sets the default HuePref to be used in Hub communications
$defaultHuePref = Get-HuePref -ErrorAction Stop