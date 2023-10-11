#Build out automate slide
# Define the PowerPoint application
$powerpoint = New-Object -ComObject PowerPoint.Application

# Get the active presentation
$presentation = $powerpoint.ActivePresentation

# Ensure the presentation has at least one slide
if ($presentation.Slides.Count -lt 1) {
    Write-Error "The presentation should have at least one slide."
    exit
}

# Copy the first slide
$presentation.Slides(1).Copy()

# Initialize the counter
$counter = 0

# This variable will capture the start time for each set of 10 slides
$batchStartTime = Get-Date

# Specify the output file for timings
$outputFile = "$env:USERPROFILE\Documents\SlideTimings.txt"

# Clear the file (or create it if it doesn't exist)
if (Test-Path $outputFile) {
    Clear-Content $outputFile
}

# Paste the first slide 100 times
1..100 | ForEach-Object {
    # Increment the counter
    $counter++

    # Paste the slide
    $presentation.Slides.Paste()

    # If counter is a multiple of 10, print the elapsed time for the batch of 10 slides and reset the start time
    if ($counter % 10 -eq 0) {
        $elapsedTime = (Get-Date) - $batchStartTime
        $message = "Time taken for slides $(${_}-9) to ${_}: $($elapsedTime.TotalSeconds) seconds"
        Write-Output $message
        Add-Content -Path $outputFile -Value $message
        
        # Reset the start time for the next batch
        $batchStartTime = Get-Date
    }
}

# Release the COM objects to free resources and prevent potential locks
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
