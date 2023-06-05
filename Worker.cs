using System;
using System.IO.Ports;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;


namespace SerialPortWorker;

public class Worker : BackgroundService
{
    private readonly ILogger<Worker> _logger;
    private SerialPort _serialPort = new SerialPort();

    public Worker(ILogger<Worker> logger)
    {
        _logger = logger;
    }

    public override Task StartAsync(CancellationToken cancellationToken)
    {
        // Start the Http Client on start up
        // client = new HttpClient();
        return base.StartAsync(cancellationToken);
    }

    public override Task StopAsync(CancellationToken cancellationToken)
    {
        // Clean up client and Port on close
        // client.Dispose;
        return base.StopAsync(cancellationToken);
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _serialPort = new SerialPort("COM1", 9600);
        _serialPort.DataReceived += OnDataReceived;
        _serialPort.Open();

        while (!stoppingToken.IsCancellationRequested)
        {
            // _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
            _serialPort.ReadTimeout = 5000;
            await Task.Delay(60*1000, stoppingToken);
        }

        _serialPort.Close();
    }

    private void OnDataReceived(object sender, SerialDataReceivedEventArgs e)
    {
        var data = _serialPort.ReadExisting();
        _logger.LogInformation("Received: {data}", data);

        // Call your API here with the received data
    }
}
