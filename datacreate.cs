using System;
using Microsoft.Office.Interop.PowerPoint;

class Program
{
    static void Main()
    {
        // Create a new PowerPoint application
        Application powerpoint = new Application();

        // Get the active presentation
        Presentation presentation = powerpoint.ActivePresentation;

        // Ensure the presentation has at least one slide
        if (presentation.Slides.Count < 1)
        {
            Console.WriteLine("The presentation should have at least one slide.");
            return;
        }

        // Copy the first slide
        presentation.Slides[1].Copy();

        // Initialize the counter
        int counter = 0;

        // This variable will capture the start time for each set of 10 slides
        DateTime batchStartTime = DateTime.Now;

        // Specify the output file for timings
        string outputFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\SlideTimings.txt";

        // Clear the file (or create it if it doesn't exist)
        if (System.IO.File.Exists(outputFile))
        {
            System.IO.File.WriteAllText(outputFile, string.Empty);
        }

        // Paste the first slide 100 times
        for (int i = 1; i <= 100; i++)
        {
            // Increment the counter
            counter++;

            // Paste the slide
            presentation.Slides.Paste();

            // If counter is a multiple of 10, print the elapsed time for the batch of 10 slides and reset the start time
            if (counter % 10 == 0)
            {
                TimeSpan elapsedTime = DateTime.Now - batchStartTime;
                string message = $"Time taken for slides {i - 9} to {i}: {elapsedTime.TotalSeconds} seconds";
                Console.WriteLine(message);
                System.IO.File.AppendAllText(outputFile, message + Environment.NewLine);
                
                // Reset the start time for the next batch
                batchStartTime = DateTime.Now;
            }
        }

        // Release the COM objects to free resources and prevent potential locks
        System.Runtime.InteropServices.Marshal.ReleaseComObject(presentation);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(powerpoint);
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
