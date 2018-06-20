# PSHue
PowerShell toolkit for controlling Philips Hue lights

## Description
This toolkit can be used to get rooms, scenes, and lights associated with a Hue Bridge and control them.  This can be useful for binding Scenes or power states to macro keys on your keyboard or for orchestrating complex automation sequences.

## Setup
The module can be imported with Import-Module using the path to the .psd file in the same directory as the rest of the module files:
```PowerShell
Import-Module <<path to PSHue module files>>\PSHue.psd1
```
It can be also be imported by name by creating a directory at either $env:programfiles\WindowsPowerShell\Modules\PSHue or $env:userprofile\Documents\WindowsPowerShell\Modules\PSHue and copying the module files into the directory you created
```PowerShell
#Copy all files from the repo into their own folder, then enter the path to that folder in the below variable:
$moduleSource = <<Directory containing ONLY PSHue module files>>

#Enter the desired module destination in the below variable, generally either $env:programfiles\WindowsPowerShell\Modules\PSHue or $env:userprofile\Documents\WindowsPowerShell\Modules\PSHue
$moduleDestination = <<Module directory>>

if (!(Test-Path $moduleDestination)) { New-Item $moduleDestination -ItemType Directory -Force }
Copy-Item (Join-Path $moduleSource '*') $moduleDestination -Recurse -Force
Import-Module PSHue
```

The first time the module is imported, it should request the IP or DNS name of your Hue Bridge.  Once entered, it will ask that you press the Link button on the top of your Hue Bridge to authorize the module's access to the Bridge.

## Help
To get a list of available commands, use the following (Module must be imported):
```PowerShell
Get-Command -Module PSHue
```
To get more information about a specific command, use the following (Module must be imported):
```PowerShell
Get-Help <<Command>>
```

## Examples
### Get all hue lights
```PowerShell
Get-HueLight
```

### Get a list of lights for a room
```PowerShell
Get-HueRoom Bedroom | Get-HueLight
```

### Turn a light on
```PowerShell
Get-HueLight 'Ceiling Light' | Set-HueState -PowerState On
```

### Start a scene
```PowerShell
Get-HueScene Bright -RoomName 'Living Room' | Start-HueScene
```

### Set a light to a predefined color
```PowerShell
Get-HueLight 'Splash Lamp' | Set-HueSate -Color Blue
```

### Start color loop on a light
```PowerShell
Get-HueLight 'Splash Lammp' | Set-HueState -Effect ColorLoop
```

### Convert an RGB value to XY notation and set all lights in a room to that color
```PowerShell
$xy = Convert-RgbToXy -R 0 -G 0 -B 255
Get-HueRoom 'Living Room' | Get-HueLight | Set-HueState -X $xy.x -Y $xy.y
```

### Dim the brightness of a group of lights in a room to 10% over the next 30 minutes
```PowerShell
Get-HueRoom Bedroom | Get-HueLight Ceiling* | Set-HueState -BrightnessPercent 10 -TransitionTime "00:30:00.0000"
```
