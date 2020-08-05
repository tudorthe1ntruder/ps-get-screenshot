function Get-ScreenShot
{
    [CmdletBinding(DefaultParameterSetName='OfWholeScreen')]
    param(    
    # If set, takes a screen capture of the current window
    [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        ParameterSetName='OfWindow')]
    [Switch]$OfWindow,
    
    # If set, takes a screenshot of a location on the screen.
    # If two numbers are passed, the screenshot will be from 0,0 to first (left), second (top)
    # If four numbers are passed, the screenshot will be from first (Left), second(top), third (width), fourth (height)
    [Parameter(Mandatory=$true,
        ValueFromPipelineByPropertyName=$true,
        ParameterSetName='OfLocation')]    
    [Double[]]$OfLocation,
    
    # The path for the screenshot.
    # If this isn't set, the screenshot will be automatically saved to a file in the current directory named ScreenCapture
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [string]
    $Path,
    
    # The image format used to store the screen capture
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [ValidateSet('PNG', 'JPEG', 'TIFF', 'GIF', 'BMP')]
    [string]
    $ImageFormat = 'JPEG',
    
    # The time before and after each screenshot
    [Parameter(ValueFromPipelineByPropertyName=$true)]
    [Timespan]$ScreenshotTimer = "0:0:0.125"
    )


    begin {
        Add-Type -AssemblyName System.Drawing, System.Windows.Forms
        $saveScreenshotFromClipboard = {
            if ([Runspace]::DefaultRunspace.ApartmentState -ne 'STA') {
                # The clipboard isn't accessible in MTA, so save the image in background runspace
                $cmd = [PowerShell]::Create().AddScript({
                    $bitmap = [Windows.Forms.Clipboard]::GetImage()    
                    $bitmap.Save($args[0], $args[1], $args[2])                    
                    $bitmap.Dispose()
                }).AddParameters(@("${screenCapturePathBase}${c}.$ImageFormat",$Codec, $ep))
                $runspace = [RunspaceFactory]::CreateRunspace()
                $runspace.ApartmentState = 'STA'
                $runspace.ThreadOptions = 'ReuseThread'
                $runspace.Open()
                $cmd.Runspace = $runspace
                $cmd.Invoke()
                $runspace.Close()
                $runspace.Dispose()
                $cmd.Dispose()
            } else {            
                $bitmap = [Windows.Forms.Clipboard]::GetImage()    
                $bitmap.Save("${screenCapturePathBase}${c}.$ImageFormat", $Codec, $ep)                    
                $bitmap.Dispose()
            }
        }
    }
    process {
        #region Codec Info
        $Codec = [Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | 
            Where-Object { $_.FormatDescription -eq $ImageFormat }

        $ep = New-Object Drawing.Imaging.EncoderParameters  
        if ($ImageFormat -eq 'JPEG') {
            $ep.Param[0] = New-Object Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality, [long]100)  
        }
        #endregion Codec Info
        

        #region PreScreenshot timer
        if ($ScreenshotTimer -and $ScreenshotTimer.TotalMilliseconds) {
            Start-Sleep -Milliseconds $ScreenshotTimer.TotalMilliseconds
        }
        #endregion Prescreenshot Timer
        
        #region File name
        if (-not $Path) {
            $screenCapturePathBase = "$pwd\ScreenCapture"
        } else {
            $screenCapturePathBase = $Path
        }
        $c = 0
        while (Test-Path "${screenCapturePathBase}${c}.$ImageFormat") {
            $c++
        }
        #endregion
        

        
        if ($psCmdlet.ParameterSetName -eq 'OfWindow') {
            [Windows.Forms.Sendkeys]::SendWait("%{PrtSc}")        
            #region PostScreenshot timer
            if ($ScreenshotTimer -and $ScreenshotTimer.TotalMilliseconds) {
                Start-Sleep -Milliseconds $ScreenshotTimer.TotalMilliseconds
            }
            #endregion Postscreenshot Timer
            . $saveScreenshotFromClipboard 
            Get-Item -ErrorAction SilentlyContinue -Path "${screenCapturePathBase}${c}.$ImageFormat"
        } elseif ($psCmdlet.ParameterSetName -eq 'OfLocation') {
            if ($OfLocation.Count -ne 2 -and $OfLocation.Count -ne 4) {
                Write-Error "Must provide either a width and a height, or a top, left, width, and height"                
                return
            }
            if ($OfLocation.Count -eq 2) {
                $bounds  = New-Object Drawing.Rectangle -Property @{
                    Width = $OfLocation[0]
                    Height = $OfLocation[1]
                }                
            } else {
                $bounds  = New-Object Drawing.Rectangle -Property @{
                    X = $OfLocation[0]
                    Y = $OfLocation[1]
                    Width = $OfLocation[2]
                    Height = $OfLocation[3]
                }
            }
            
            $bitmap = New-Object Drawing.Bitmap $bounds.width, $bounds.height
            $graphics = [Drawing.Graphics]::FromImage($bitmap)
            $graphics.CopyFromScreen($bounds.Location, [Drawing.Point]::Empty, $bounds.size)
            #region PostScreenshot timer
            if ($ScreenshotTimer -and $ScreenshotTimer.TotalMilliseconds) {
                Start-Sleep -Milliseconds $ScreenshotTimer.TotalMilliseconds
            }
            #endregion Postscreenshot Timer

            $bitmap.Save("${screenCapturePathBase}${c}.$ImageFormat", $Codec, $ep)                    
            $graphics.Dispose()
            $bitmap.Dispose()
            Get-Item -ErrorAction SilentlyContinue -Path "${screenCapturePathBase}${c}.$ImageFormat"
        } elseif ($psCmdlet.ParameterSetName -eq 'OfWholeScreen') {
            [Windows.Forms.Sendkeys]::SendWait("{PrtSc}")        
            #region PostScreenshot timer
            if ($ScreenshotTimer -and $ScreenshotTimer.TotalMilliseconds) {
                Start-Sleep -Milliseconds $ScreenshotTimer.TotalMilliseconds
            }
            #endregion Postscreenshot Timer
            . $saveScreenshotFromClipboard             
            Get-Item -ErrorAction SilentlyContinue -Path "${screenCapturePathBase}${c}.$ImageFormat"
        }
        
                
                
    }
}

Get-ScreenShot -Path $env:temp\temp
