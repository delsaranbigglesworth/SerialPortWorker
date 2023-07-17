using System;
using System.IO.Ports;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Management;
using System.Net.Http;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Text;
using System.Text.RegularExpressions;
using System.ServiceProcess;
using System.Runtime.InteropServices;

namespace SerialPortWorker;

public class Worker : BackgroundService
{
    private readonly ILogger<Worker> _logger;
    private readonly IHttpClientFactory _httpClientFactory;
    //private SerialPort _serialPort = new SerialPort("COM1", 9600);
    //private ManagementEventWatcher _insertWatcher;
    //private ManagementEventWatcher _removeWatcher;

    public Worker(ILogger<Worker> logger, IHttpClientFactory httpClientFactory)
    {
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }
    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Call the ExecuteCode method to run your code when the service first starts
        await ExecuteCode();



        // // Create event query to be notified within 1 second of an event
        // WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM _InstanceCreationEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_USBHub'");
        // _insertWatcher = new ManagementEventWatcher(insertQuery);
        // _insertWatcher.EventArrived += DeviceInsertedEventAsync;



        // WqlEventQuery removeQuery = new WqlEventQuery("SELECT * FROM _InstanceDeletionEvent WITHIN 1 WHERE TargetInstance ISA 'Win32_USBHub'");
        // _removeWatcher = new ManagementEventWatcher(removeQuery);
        // _removeWatcher.EventArrived += DeviceRemovedEvent;

        // // Start listening for events
        // _insertWatcher.Start();
        // _removeWatcher.Start();

        // Set up a ManagementEventWatcher to monitor for device insertion events
        WqlEventQuery insertQuery = new WqlEventQuery("SELECT * FROM __InstanceCreationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_SerialPort'");
        ManagementEventWatcher insertWatcher = new ManagementEventWatcher(insertQuery);
        insertWatcher.EventArrived += async (sender, e) =>
        {
            // When a device is inserted, call the ExecuteCode method to run your code
            await ExecuteCode();
        };
        insertWatcher.Start();

        while (!stoppingToken.IsCancellationRequested)
        {
            await Task.Delay(1000, stoppingToken);
        }
    }

    // private async void DeviceInsertedEventAsync(object sender, EventArrivedEventArgs e)
    // {
    //     // A device has been inserted!
    //     // Add your code here to handle the event
    //     _logger.LogInformation("A device has been plugged in!");
    //     Console.WriteLine("A device has been plugged in!");
    //     await ExecuteCode();
    // }

    // private void DeviceRemovedEvent(object sender, EventArrivedEventArgs e)
    // {
    //     // A device has been removed!
    //     // Add your code here to handle the event
    //     _logger.LogInformation("A device has been unplugged!");
    //     Console.WriteLine("A device has been unplugged!");
    // }

    // public override async Task StopAsync(CancellationToken cancellationToken)
    // {
    //     // Stop listening for events
    //     _insertWatcher.Stop();
    //     _removeWatcher.Stop();

    //     await base.StopAsync(cancellationToken);
    // }

    private async Task ExecuteCode()
    {

        string activeComPort = await FindActiveComPortAsync();
        SerialPort serialPort = new SerialPort(activeComPort, 9600);

        // Create a TaskCompletionSource object to synchronize the execution of your code
        TaskCompletionSource<bool> tcs = new TaskCompletionSource<bool>();

        // Attach an event handler to the DataReceived event
        serialPort.DataReceived += async (sender, e) =>
        {
            await OnDataReceived(sender, e);
            if (!tcs.Task.IsCompleted)
            {
                tcs.SetResult(true);
            }
        };

        // Attach an event handler to the ErrorReceived event
        serialPort.ErrorReceived += (sender, e) =>
        {
            // Handle the error here
            Console.WriteLine("An error occurred: " + e.EventType);
        };

        // Open the serial port
        if (!serialPort.IsOpen)
        {
            serialPort.Open();
        }

        // Wait for the OnDataReceived method to complete
        await tcs.Task;
    }

    public async Task<string> FindActiveComPortAsync()
    {
        string? activeComPort = null;

        while (string.IsNullOrWhiteSpace(activeComPort))
        {
            // Get a list of serial port names.
            string[] ports = SerialPort.GetPortNames();
            _logger.LogInformation("The following serial ports were found: " + string.Join(", ", ports));
            Console.WriteLine("The following serial ports were found: " + string.Join(", ", ports));


            //string activeComPort = "";
            using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0 AND Name LIKE '%CH340%'"))
            {
                ManagementObjectCollection collection = searcher.Get();
                foreach (var device in collection)
                {
                    string deviceName = (string)device.GetPropertyValue("Name");
                    string deviceDescription = (string)device.GetPropertyValue("Description");
                    if ((deviceName != null && deviceName.Contains("CH340")) || (deviceDescription != null && deviceDescription.Contains("CH340")))
                    {
                        string caption = (string)device.GetPropertyValue("Caption");
                        Match match = Regex.Match(caption, @"(COM\d+)");
                        if (match.Success)
                        {
                            activeComPort = match.Groups[1].Value;
                            //SerialPort serialPort = new SerialPort(comPort, 9600);




                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(activeComPort))
            {
                Console.WriteLine("COM Port: " + activeComPort);
                return activeComPort;
            }
            else
            {
                Console.WriteLine("Device not found");
            }

            // using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode = 0 AND Name LIKE '%CH340%'" ))
            // {
            //     ManagementObjectCollection collection = searcher.Get();
            //     foreach (var device in collection)
            //     {
            //         string deviceId = (string)device.GetPropertyValue("DeviceID");
            //         string deviceName = (string)device.GetPropertyValue("Name");
            //         if (deviceName.Contains("CH340"))
            //         {
            //             Console.WriteLine("COM Port: " + deviceId);
            //         }
            //     }
            // }

            await Task.Delay(1000 //stoppingToken
            );

            // Loop through each port name and check if the JSON string is there.
            // foreach (string port in ports)
            // {
            //     if (SerialPort.GetPortNames().Contains(port))
            //     {
            //         // Create a new SerialPort object with basic settings
            //         SerialPort serialPort = new SerialPort(port, 9600);
            //         serialPort.ReadTimeout = 5000;
            //         Console.WriteLine("Serial Port " + port + " IsOpen = " + serialPort.IsOpen);
            //         using (serialPort)
            //         {
            //             Console.WriteLine("Serial Port " + port + " IsOpen = " + serialPort.IsOpen);
            //             string? data = null;
            //             try
            //             {
            //                 // Open the port for communications
            //                 serialPort.Open();
            //                 Console.WriteLine("Serial Port " + port + " IsOpen = " + serialPort.IsOpen);
            //                 // Read data from the port
            //                 data = serialPort.ReadLine();
            //                 Console.WriteLine(data);
            //                 // Check if data contains any JSON string
            //                 if (data.StartsWith("{\"Id\":"))
            //                 {
            //                     _logger.LogInformation("Found JSON string on port {port}" + port);
            //                     Console.WriteLine("Found JSON string on port" + port);
            //                     Console.WriteLine(data);
            //                     activeComPort = port;
            //                     _logger.LogInformation("Setting " + port + " to active port");
            //                     Console.WriteLine("Setting " + port + " to active port");

            //                     break;
            //                 }
            //                 else
            //                 {
            //                     _logger.LogInformation("This is not the data we are looking for");
            //                 }
            //             }
            //             catch (Exception ex)
            //             {
            //                 _logger.LogInformation("No data received from port " + port + " within the specified time period. " + ex);
            //                 Console.WriteLine("No data received from port " + port + " within the specified time period." + ex);
            //             }
            //         }
            //     }
            //     else
            //     {
            //         _logger.LogInformation("Port " + port + " is not available.");
            //         Console.WriteLine("Port " + port + " is not available.");
            //     }
            // }
            await Task.Delay(10000);
        }

        return activeComPort;
    }
    private async Task OnDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        //string activeComPort = _serialPort.PortName;
        var serialPort = sender as SerialPort;
        // Create a new SerialPort object with basic settings
        //serialPort = new SerialPort(serialPort.PortName, 9600);
        if (serialPort != null)
        {
            serialPort.ReadTimeout = 5000;
            // ... Move try/catch into this if
            try
            {
                Console.WriteLine(serialPort.IsOpen + " OnDataReceived");
                string data = serialPort.ReadLine();
                _logger.LogInformation("Received: {data}" + data);
                Console.WriteLine("Received: " + data);

                // Validate the received data before parsing it as a JSON string
                if (data.StartsWith("{\"Id\":"))
                {
                    // Call your API here with the received data
                    // Parse the received JSON string
                    using JsonDocument jsonDoc = JsonDocument.Parse(data);
                    JsonElement root = jsonDoc.RootElement;

                    // Extract the properties from the JSON object
                    int id = root.GetProperty("Id").GetInt32();
                    // Getting the hostname of the computer hosting the sensor device
                    String deviceName = System.Environment.GetEnvironmentVariable("COMPUTERNAME") ?? string.Empty;
                    string sensorName = root.GetProperty("SensorName").GetString() ?? string.Empty;
                    string temperature = root.GetProperty("Temperature").GetString() ?? string.Empty;
                    string humidity = root.GetProperty("Humidity").GetString() ?? string.Empty;
                    DateTime whenCreated = root.GetProperty("WhenCreated").GetDateTime();

                    // Create a new object with the extracted properties
                    var extractedData = new
                    {
                        Id = id,
                        DeviceName = deviceName,
                        SensorName = sensorName,
                        Temperature = temperature,
                        Humidity = humidity,
                        WhenCreated = whenCreated
                    };

                    var json = JsonConvert.SerializeObject(extractedData);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    // Send the extracted JSON object to the API endpoint
                    HttpClient httpClient = _httpClientFactory.CreateClient();
                    string apiUrl = "https://sensorapi.densosystems.com/api/v1/production/Environmental/LattePanda";
                    HttpResponseMessage response = await httpClient.PostAsync(apiUrl, content);

                    if (response.IsSuccessStatusCode)
                    {
                        _logger.LogInformation("Data sent successfully to the API endpoint.");
                        Console.WriteLine("Data sent successfully to the API endpoint.");
                    }
                    else
                    {
                        _logger.LogError($"Error sending data to the API endpoint: {response.StatusCode}");
                        Console.WriteLine($"Error sending data to the API endpoint: {response.StatusCode}");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex + "Timeout while reading data from serial port");
                Console.WriteLine("Timeout while reading data from serial port" + ex);
            }
        }



    }
}
