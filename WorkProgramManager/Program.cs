//~~PRAISE THE OMNISSIAH~~

using Serilog;
using System.Diagnostics;
using System.Management;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace WorkProgramManager
{
    enum ProgramType { Exe, Uwp }

    class Configuration
    {
        public string LogDirectory { get; set; } = ".";
        public List<WorkProgram> Programs { get; set; } = [];

        public void Validate()
        {
            foreach (var program in Programs)
            {
                switch (program.Type)
                {
                    case ProgramType.Exe:
                        if (string.IsNullOrWhiteSpace(program.Path))
                            throw new ConfigurationException($"Program '{program.Name}' (EXE) must specify Path");
                        if (string.IsNullOrWhiteSpace(program.ProcessName))
                            throw new ConfigurationException($"Program '{program.Name}' (EXE) must specify ProcessName");
                        break;

                    case ProgramType.Uwp:
                        if (string.IsNullOrWhiteSpace(program.Aumid))
                            throw new ConfigurationException($"Program '{program.Name}' (UWP) must specify AUMID");
                        if (string.IsNullOrWhiteSpace(program.ProcessName))
                            throw new ConfigurationException($"Program '{program.Name}' (UWP) must specify ProcessName");
                    break;

                    default:
                        throw new ConfigurationException($"Program '{program.Name}' has unknown Type '{program.Type}'");
                }
            }
        }
    }

    public class ConfigurationException(string message) : Exception(message)
    {
    }

    class WorkProgram
    {
        public string? Name { get; set; }
        public ProgramType Type { get; set; }
        public string? Path { get; set; }
        public string? Aumid { get; set; }
        public string? ProcessName { get; set; }
    }

    partial class Program
    {
        static List<WorkProgram> workPrograms = [];
        static bool quietMode = false;
        private static string logDir = string.Empty;
        private static readonly string configFile = "work-programs.json";
        const string outputTemplate = "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}";

        private static JsonSerializerOptions GetOptions()
        {
            return new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
            };
        }

        static async Task Main(string[] args)
        {
            //check arguments
            if (args.Length == 0 ||
                !(args[0].Equals("start", StringComparison.OrdinalIgnoreCase)
                || args[0].Equals("stop", StringComparison.OrdinalIgnoreCase)))
            {
                Console.WriteLine("Usage: WorkProgramManager.exe [start|stop] [--quiet]");
                Console.ReadKey();
                return;
            }

            if (args.Contains("--quiet"))
            {
                quietMode = true;
            }

            if (!File.Exists(configFile))
            {
                Console.WriteLine($"Configuration file '{configFile}' not found.");
                Console.ReadKey();
                return;
            }

            //read config file
            try
            {
                var json = File.ReadAllText(configFile);
                var options = GetOptions();
                var config = JsonSerializer.Deserialize<Configuration>(json, options)
                    ?? throw new Exception("Failed to parse configuration");
                config.Validate();

                if (string.IsNullOrWhiteSpace(config.LogDirectory))
                    throw new InvalidOperationException($"LogDirectory must be set in '{configFile}");

                logDir = Path.GetFullPath(config.LogDirectory);
                Directory.CreateDirectory(logDir);
                workPrograms = config.Programs;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load configuration: {ex.Message}");
                Console.ReadKey();
                return;
            }

            //init logger
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Conditional(
                    _ => !quietMode,
                    wt => wt.Console(outputTemplate: outputTemplate)
                )
                .WriteTo.FileEx(
                    Path.Combine(logDir, "work-programs-log-.txt"),
                    periodFormat: "_yyyy-MM-dd",
                    rollingInterval: (Serilog.Sinks.FileEx.RollingInterval)RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: outputTemplate)
                .CreateLogger();

            //ACTION TREE
            var action = args[0].ToLowerInvariant();
            Log.Information("Action: {Action} | Timestamp: {Timestamp}", action, DateTime.Now);

            if (action == "start")
            {
                foreach (var wp in workPrograms)
                {
                    await StartProgram(wp);
                    await Task.Delay(1000);
                }
            }
            else if (action == "stop")
            {
                foreach (var wp in workPrograms)
                    await StopProgram(wp);
            }
            else    //need to specify option
            {
                Console.WriteLine("Usage: WorkProgramManager.exe [start|stop] [--quiet]");
                Console.ReadKey();
                return;
            }

            Log.Information("Action complete");
            Log.CloseAndFlush();
        }

        //START PROGRAMS
        static async Task StartProgram(WorkProgram program)
        {
            try
            {
                Log.Information("Attempting to start {Name}", program.Name);
                Process? proc = null;

                if (program.Type == ProgramType.Exe && !string.IsNullOrWhiteSpace(program.Path))
                {
                    proc = Process.Start(new ProcessStartInfo
                    {
                        FileName = program.Path!,
                        UseShellExecute = true
                    });
                    proc?.PriorityClass = ProcessPriorityClass.BelowNormal;
                }
                else if (program.Type == ProgramType.Uwp && !string.IsNullOrWhiteSpace(program.Aumid))
                {
                    proc = Process.Start("explorer.exe", $"shell:AppsFolder\\{program.Aumid}");
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Skipped {program.Name}: missing Path or AUMID in configuration");
                }

                if (proc == null)
                    throw new InvalidOperationException(
                        $"Failed to start {program.Name}: Process.Start returned null");

                await Task.Delay(3000);

                var running = Process.GetProcessesByName(program.ProcessName!);
                if (running.Length > 0)
                {
                    Log.Information("Started {Name} successfully (PID {Pid})", program.Name, running[0].Id);
                }
                else
                {
                    throw new InvalidOperationException(
                        $"Process '{program.ProcessName}' not found after launch; {program.Name} may have failed to start");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error starting {Name}", program.Name);
            }
        }

        //STOP PROGRAMS
        static async Task StopProgram(WorkProgram program)
        {
            try
            {
                //close UWP apps
                if (program.Type == ProgramType.Uwp)
                {
                    //runtime-guard UWP applications
                    if (!OperatingSystem.IsWindows())
                    {
                        Log.Warning("Skipping UWP shutdown: only supported on Windows");
                        return;
                    }

                    if (string.IsNullOrWhiteSpace(program.Aumid))
                    {
                        Log.Warning("Skipping UWP {Name}: No AUMID configured", program.Name);
                        return;
                    }

                    Log.Information("Closing UWP app: {Name}", program.Name);
                    bool closed = await TryCloseUwpAppAsync(program.Aumid!, 5000);
                    if (closed)
                    {
                        Log.Information("Closed {Name} gracefully", program.Name);
                        return;
                    }
                    
                    //attempt to kill process with PowerShell package
                    Log.Warning("Could not close {Name} gracefully; attempting to use PFN...", program.Name);
                    closed = KillUwpHostProcesses(program.Aumid!);
                    if (closed)
                    {
                        Log.Information("Closed {Name} using PFN", program.Name);
                        return;
                    }


                    //fallback to killing any matching Win32 processes
                    Log.Warning("Could not close {Name} using PFN, killing process...", program.Name);
                    KillExeProcesses(program.ProcessName!);
                    return;
                }

                //close EXE
                if (program.Type == ProgramType.Exe)
                {
                    try
                    {
                        //attempt to close gracefully
                        Process first = Process.GetProcessesByName(program.ProcessName!).FirstOrDefault()!;
                        if (first == null)
                        {
                            Log.Warning("Unable to get PID for process {Name}", program.ProcessName);
                        }
                        else
                        {
                            //attempt graceful close
                            bool closed = await CloseProcessGracefullyAsync(first.Id, 5000);
                            if (closed)
                            {
                                Log.Information("Closed {Name} gracefully", program.Name);
                                return;
                            }

                            //fallback to killing Win32 process
                            Log.Warning("Could not close {Name} gracefully; process will be terminated", program.Name);
                            closed = KillExeProcesses(program.ProcessName!);
                            if (closed)
                            {
                                Log.Information("Stopped {Name} via fallback kill", program.Name);
                                return;
                            }

                            throw new Exception($"Unable to kill process {program.Name}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Unable to kill EXE {Name}", program.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error shutting down {Name}", program.Name);
            }
        }

        [SupportedOSPlatform("windows")]
        static Task<bool> TryCloseUwpAppAsync(string aumid, int timeoutMs = 5_000)
        {
            try
            {
                //runtime-guard ManagementObjectSearcher
                if (!OperatingSystem.IsWindows())
                {
                    Log.Warning("UWP shutdown is only supported on Windows");
                    return Task.FromResult(false);          //must wrap return in Task
                }

                //Query all ApplicationFrameHost processes
                var searcher = new ManagementObjectSearcher(
                    "SELECT ProcessId, CommandLine FROM Win32_Process " +
                    "WHERE Name = 'ApplicationFrameHost.exe'");

                foreach (ManagementObject mo in searcher.Get().Cast<ManagementObject>())
                {
                    var cmdLine = (mo["CommandLine"] as string) ?? string.Empty;
                    if (!cmdLine.Contains(aumid, StringComparison.OrdinalIgnoreCase))
                        continue;

                    var pid = Convert.ToInt32(mo["ProcessId"]);
                    return CloseProcessGracefullyAsync(pid, timeoutMs);
                }

                //ApplicationFrameworkHost.exe did not contain cmdline with AUMID
                return Task.FromResult(false);
            }
            catch (TimeoutException tex)
            {
                Log.Warning(tex, "Timed out sending WM_Close to UWP app {Aumid}", aumid);
                return Task.FromResult(false);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Critical failure closing UWP app {Aumid}", aumid);
                throw;
            }
        }

        [SupportedOSPlatform("windows")]
        private static bool KillUwpHostProcesses(string aumid)
        {
            //runtime-guard ManagementObjectSearcher
            if (!OperatingSystem.IsWindows())
            {
                Log.Warning("UWP shutdown is only supported on Windows");
                return false;
            }

            //get PFN
            string packageFamilyName = aumid.Split('!')[0];

            //build hidden, non-elevated PowerShell call
            var startInfo = new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -WindowStyle Hidden " +
                    $"-Command \"Stop-Process -Package '{packageFamilyName}' " +
                    $"-ErrorAction SilentlyContinue\"",
                UseShellExecute = false,
                CreateNoWindow = true
            };

            try
            {
                using var ps = Process.Start(startInfo)!;
                ps.WaitForExit(5000);
                if (ps.HasExited)
                {
                    if (ps.ExitCode != 0)
                    {
                        Log.Warning("Stop-Process -Package {pfn} exited unsuccessfully (ExitCode {Code})", 
                            packageFamilyName, ps.ExitCode);
                        return false;
                    }
                    else
                    {
                        Log.Information("Terminated UWP package {Name}",  packageFamilyName);
                        return true;
                    }
                }
                else
                {
                    Log.Warning("Stop-Process -Package {pfn} timed out", packageFamilyName);
                    throw new Exception("Stop-Process timed out");
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Failed to terminate UWP package {name}", packageFamilyName);
                return false;
            }
        }

        private static bool KillExeProcesses(string processName)
        {
            var procs = Process.GetProcessesByName(processName);
            if (procs.Length == 0)
            {
                Log.Warning($"No running processes found for {processName}");
                return false;
            }

            foreach (var proc in procs)
            {
                Log.Information("Sending close to {Name} (PID {Pid})", processName, proc.Id);
                if (proc.MainWindowHandle != IntPtr.Zero && proc.CloseMainWindow() && proc.WaitForExit(5000))
                {
                    Log.Information("Closed {Name} cleanly (ExitCode {Code})", processName, proc.ExitCode);
                    return true;
                }
                else
                {
                    try
                    {
                        try
                        {
                            //try to kill subtree, if WorkProgramManager started it
                            proc.Kill(entireProcessTree: true);
                        }
                        catch (InvalidOperationException)
                        {
                            //process not started by WorkProgramManager, fallback to killing main process
                            proc.Kill();
                        }

                        Log.Information("Killed {Name} (PID {pid})", processName, proc.Id);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Error killing EXE {Name} (PID {pid})", processName, proc.Id);
                    }
                }
            }
            return false;
        }

        [LibraryImport("user32.dll", EntryPoint = "PostMessageA", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
        const uint WM_CLOSE = 0x0010;
        static async Task<bool> CloseProcessGracefullyAsync(int pid, int timeoutMs)
        {
            try
            {
                var proc = Process.GetProcessById(pid);
                //try CloseMainWindow (send WM_Close to main window)
                if (proc.CloseMainWindow() && await Task.Run(() => proc.WaitForExit(timeoutMs)))
                    return true;        //exited cleanly

                //if no main window or timed out, enumerate and send WM_Close
                foreach (ProcessThread thread in proc.Threads)
                {
                    EnumThreadWindows(thread.Id, (hWnd, lParam) =>
                    {
                        PostMessage(hWnd, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
                        return true;
                    }, IntPtr.Zero);
                }

                if (await Task.Run(() => proc.WaitForExit(timeoutMs)))
                    return true;        //exited cleanly

                //final option: kill
                proc.Kill(true);
                return true;
            }
            catch
            {
                return false;
            }
        }

        [LibraryImport("user32.dll", EntryPoint = "EnumThreadWindowsW")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static partial bool EnumThreadWindows(int threadId, EnumThreadWndProc callback, IntPtr lParam);
        delegate bool EnumThreadWndProc(IntPtr hWnd, IntPtr lParam);
    }
}
