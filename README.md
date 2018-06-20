# PSHue
PowerShell toolkit for controlling Philips Hue lights

## Description
This toolkit can be used to get rooms, scenes, and lights associated with a Hue Bridge and control them.  This can be useful for binding Scenes or power states to macro keys on your keyboard or for orchestrating complex automation sequences.

To get a list of available commands, use the following:
```PowerShell
Get-Command -Module PSHue
```
To get more information about a specific command, use the following:
```PowerShell
Get-Help <<Command>>
```

## Setup
The module can be imported with Import-Module using the path to the .psd file in the same directory as the rest of the module files:
```PowerShell
Import-Module c:\PSHue\PSHue.psd
```
It can be also be imported by name by creating a directory at either %programfiles%\WindowsPowerShell\Modules\PSHue or %UserProfile%\Documents\WindowsPowerShell\Modules\PSHue and copying the module files into the directory you created
```PowerShell
Import-Module PSHue
```

The first time the module is imported, it should request the IP or DNS name of your Hue Bridge.  Once entered, it will ask that you press the Link button on the top of your Hue Bridge to authorize the module's access to the Bridge.

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

### Set the brightness of a group of lights in a room to 25%
```PowerShell
Get-HueRoom Bedroom | Get-HueLight Ceiling* | Set-HueState -BrightnessPercent 25
```

### Start a scene
```PowerShell
Get-HueScene Bright -RoomName 'Living Room' | Start-HueScene
```

### Start color loop on a light
```PowerShell
Get-HueLight 'Splash Lammp' | Set-HueState -Effect ColorLoop
```

### Set a light to a predefined color
```PowerShell
Get-HueLight 'Splash Lamp' | Set-HueSate -Color Blue
```

### Convert an RGB value to XY notation and set all lights in a room to that color
```PowerShell
$xy = Convert-RgbToXy -R 0 -G 0 -B 255
Get-HueRoom 'Living Room' | Get-HueLight | Set-HueState -X $xy.x -Y $xy.y
```
