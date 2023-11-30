# Main start time
$mainStart = Get-Date

 

 

# Start PowerPoint and time it, total time is listed
$startPowerPointTime = Get-Date
Start-Process "powerpnt.exe"
Start-Sleep -Seconds 5
$endPowerPointTime = Get-Date

 

 

# Load necessary assembly for SendKeys and PowerPoint COM object
Add-Type -AssemblyName System.Windows.Forms

 

 

# Send keystrokes to create new slides and time it
$startSlideCreationTime = Get-Date
1..5 | ForEach-Object {
    [System.Windows.Forms.SendKeys]::SendWait("^m")
    Start-Sleep -Seconds 2
}
$endSlideCreationTime = Get-Date

 

 

# Populate data and time it
$startDataPopulationTime = Get-Date

 

 

# Reference to PowerPoint application and other necessary preparations
$powerpoint = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
$presentation = $powerpoint.ActivePresentation
$players = @("Xander Bogaerts", "Rafael Devers", "Chris Sale", "J.D. Martinez", "Nathan Eovaldi")
$numbers = 1..$players.Length | ForEach-Object { Get-Random -Minimum 1 -Maximum 500 }

 

 

# Populate the tables on the slides
1..4 | ForEach-Object {
    $slide = $presentation.Slides[$_]
    $table = $slide.Shapes.AddTable($players.Length + 1, 2, 100, 100, 500, 300).Table
    $table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "Player"
    $table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "Number"

 

 

    for ($i = 0; $i -lt $players.Length; $i++) {
        $table.Cell($i + 2, 1).Shape.TextFrame.TextRange.Text = $players[$i]
        $table.Cell($i + 2, 2).Shape.TextFrame.TextRange.Text = $numbers[$i].ToString()
    }
}
$endDataPopulationTime = Get-Date

 

 

# Main end time
$mainEnd = Get-Date

 

 

# Calculate elapsed times
$powerPointElapsedTime = $endPowerPointTime - $startPowerPointTime
$slideCreationElapsedTime = $endSlideCreationTime - $startSlideCreationTime
$dataPopulationElapsedTime = $endDataPopulationTime - $startDataPopulationTime
$mainElapsedTime = $mainEnd - $mainStart

 

 

# Display the elapsed times in the console and add to the PowerPoint slide
Write-Host "PowerPoint start time: $($powerPointElapsedTime.TotalSeconds) seconds"
Write-Host "Slide creation time: $($slideCreationElapsedTime.TotalSeconds) seconds"
Write-Host "Data population time: $($dataPopulationElapsedTime.TotalSeconds) seconds"
Write-Host "Total time: $($mainElapsedTime.TotalSeconds) seconds"

 

 

$slideTime = $presentation.Slides.Add($presentation.Slides.Count + 1, [Microsoft.Office.Interop.PowerPoint.PpSlideLayout]::ppLayoutText)
$shape = $slideTime.Shapes[1]
$shape.TextFrame.TextRange.Text = "Times (in seconds):`nPowerPoint start: $($powerPointElapsedTime.TotalSeconds)`nSlide creation: $($slideCreationElapsedTime.TotalSeconds)`nData population: $($dataPopulationElapsedTime.TotalSeconds)`nTotal: $($mainElapsedTime.TotalSeconds)"
