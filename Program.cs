using ExcelCrack;
using OfficeOpenXml;
using System.Timers;

Console.WriteLine("Starting!");

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

//your password
var filePath = args.Length > 0 ? args[0] : "C:\\Users\\smartinmedia\\Desktop\\protected workbook.xlsx";
if (!File.Exists(filePath))
{
    Console.WriteLine($"File does not exist {filePath}");
    return;
}
var fileInfo = new FileInfo(filePath);

var generator = new Generator(1);
var threads = Environment.ProcessorCount;
object once = new object();
var attempts = 0;
var timer = new System.Timers.Timer(1000);
timer.Elapsed += OnTimedEvent;

void OnTimedEvent(object? sender, ElapsedEventArgs e)
{
    var attemptsPersecond = attempts;
    Console.WriteLine($"Attempts per second {attemptsPersecond}: {generator.getCurrentPassword()}");
    attempts = 0;
}

timer.Start();


var tasks = new List<Task>();
for (var thread = 1; thread <= threads; thread++)
{
    tasks.Add(Task.Run(() => CrackStream(filePath, generator, fileInfo)));
}  

await Task.WhenAny(tasks);

// Spire.XLS       80/s
// Aspose.Cells    65/s
// ExcelDataReader 8/s
// EPPLUS          88/s 
// IronXL          90/s - Cannot run outside of visual studio
// MinExcel password not supported
// ClosedXML password not supported
void CrackStream(string filePath, Generator generator, FileInfo fileInfo)
{
    if (fileInfo == null || fileInfo.DirectoryName == null)
    {
        throw new ArgumentNullException("fileInfo");
    }

    var ms = new MemoryStream();
    lock (once)
    {
        using var stream = File.Open(fileInfo.FullName, FileMode.Open, FileAccess.Read);
        stream.CopyTo(ms);
    }
    
    bool success = false;

    do
    {
        var password = generator.generateNextPassword();

        try
        {
            // Console.WriteLine($"Trying {password}");
            ms.Position = 0;

            using (var package = new ExcelPackage(ms, password));
            Console.WriteLine($"Success! Password is {password}");
            CreateResultFile(password, fileInfo.DirectoryName);
            success = true;
            Environment.Exit(0);
            return;
        }
        catch (Exception ex)
        {
        }

        attempts++;
        success = false;
    } while (!success);
    
}

void CreateResultFile(string password, string folderPath)
{
    var resultsPath = Path.Join(folderPath, "password.txt");
    if (File.Exists(resultsPath))
    {
        File.Delete(resultsPath);
    }

    File.WriteAllText(resultsPath, password);
}