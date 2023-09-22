// See https://aka.ms/new-console-template for more information
using ExcelCrack;
using ExcelDataReader;
using System.IO;
using System.Timers;

Console.WriteLine("Starting!");

// https://github.com/ExcelDataReader/ExcelDataReader#important-note-on-net-core
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

//your password
var filePath = args.Length > 0 ? args[0] : "C:\\Users\\smartinmedia\\Desktop\\protected workbook.xlsx";
if (!File.Exists(filePath))
{
    Console.WriteLine($"File does not exist {filePath}");
    return;
}

var generator = new Generator(1);
var threads = 1;

var attempts = 0;
var timer = new System.Timers.Timer(1000);
timer.Elapsed += OnTimedEvent;

void OnTimedEvent(object? sender, ElapsedEventArgs e)
{
    var attemptsPersecond = attempts;
    Console.WriteLine($"Attempts per second {attemptsPersecond}");
    attempts = 0;
}

timer.Start();

using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
{
    var tasks = new List<Task>();
    for (var thread = 1; thread <= threads; thread++)
    {
        MemoryStream ms = new MemoryStream();
        await stream.CopyToAsync(ms);

        tasks.Add(Task.Run(() => CrackStream(ms, generator)));

    }  

    await Task.WhenAny(tasks);
}

void CrackStream(MemoryStream ms, Generator generator)
{
    bool success = false;
    do
    {
        var password = generator.generateNextPassword();
        var config = new ExcelReaderConfiguration
        {
            Password = password
        };

        try
        {
            Console.WriteLine($"Trying {password}");
            using (var reader = ExcelReaderFactory.CreateReader(ms, config))
            {
                reader.Read();
            }


            Console.WriteLine($"Success! Password is {password}");
            success = true;
            return;
        }
        catch(Exception ex)
        {
            Console.WriteLine(ex.Message);
        }

        attempts++;
        success = false;
    } while (!success);    
}