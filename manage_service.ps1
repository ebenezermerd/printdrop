# Auto-Print Service Manager
# Ensures the service is running in background with proper monitoring

param(
    [string]$Action = "start"  # start, stop, status, restart
)

$ServiceName = "Auto-Print Service"
$WorkingDir = "C:\PrintDrop"
$PythonScript = "auto_print.py"

function Test-ServiceRunning {
    $processes = Get-Process -Name python -ErrorAction SilentlyContinue | Where-Object {
        $_.MainModule.FileName -like "*python*"
    }
    
    foreach ($proc in $processes) {
        try {
            $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.Id)").CommandLine
            if ($cmdLine -like "*auto_print.py*") {
                return $proc
            }
        } catch {
            # Continue checking other processes
        }
    }
    return $null
}

function Start-AutoPrintService {
    Write-Host "Starting Auto-Print Background Service..." -ForegroundColor Yellow
    
    # Kill any existing auto-print processes
    Stop-AutoPrintService -Silent
    
    # Start the service in background
    try {
        $process = Start-Process -FilePath "python" -ArgumentList $PythonScript -WorkingDirectory $WorkingDir -WindowStyle Hidden -PassThru
        
        # Wait a moment and verify it started
        Start-Sleep -Seconds 2
        $running = Test-ServiceRunning
        
        if ($running) {
            Write-Host "SUCCESS: Auto-Print Service started in background!" -ForegroundColor Green
            Write-Host "  Process ID: $($running.Id)" -ForegroundColor Cyan
            Write-Host "  Network Access: \\$env:COMPUTERNAME\PrintDrop" -ForegroundColor White
            Write-Host "  Target Printer: Canon iR2004/2204 UFRII LT" -ForegroundColor White
            Write-Host ""
            Write-Host "Service Features:" -ForegroundColor Yellow
            Write-Host "  - PDF files (SumatraPDF)" -ForegroundColor Green
            Write-Host "  - Word documents (COM automation)" -ForegroundColor Green
            Write-Host "  - Text/CSV files (PowerShell)" -ForegroundColor Green
            Write-Host "  - Print management (copies, settings)" -ForegroundColor Green
            Write-Host ""
            Write-Host "Service is now running in background and ready!" -ForegroundColor Green
            return $true
        } else {
            Write-Host "ERROR: Failed to start service properly" -ForegroundColor Red
            return $false
        }
    } catch {
        Write-Host "ERROR: Failed to start service - $_" -ForegroundColor Red
        return $false
    }
}

function Stop-AutoPrintService {
    param([switch]$Silent)
    
    if (-not $Silent) {
        Write-Host "Stopping Auto-Print Service..." -ForegroundColor Yellow
    }
    
    $stopped = 0
    $processes = Get-Process -Name python -ErrorAction SilentlyContinue
    
    foreach ($proc in $processes) {
        try {
            $cmdLine = (Get-CimInstance Win32_Process -Filter "ProcessId = $($proc.Id)").CommandLine
            if ($cmdLine -like "*auto_print.py*") {
                Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                $stopped++
            }
        } catch {
            # Continue
        }
    }
    
    if (-not $Silent) {
        if ($stopped -gt 0) {
            Write-Host "Stopped $stopped auto-print process(es)" -ForegroundColor Green
        } else {
            Write-Host "No auto-print processes were running" -ForegroundColor Cyan
        }
    }
}

function Show-ServiceStatus {
    Write-Host "Auto-Print Service Status Check" -ForegroundColor Cyan
    Write-Host "===============================" -ForegroundColor Cyan
    
    $running = Test-ServiceRunning
    
    if ($running) {
        Write-Host "STATUS: RUNNING" -ForegroundColor Green
        Write-Host "  Process ID: $($running.Id)" -ForegroundColor White
        Write-Host "  Started: $($running.StartTime)" -ForegroundColor White
        Write-Host "  Memory Usage: $([math]::Round($running.WorkingSet64/1MB, 2)) MB" -ForegroundColor White
        Write-Host ""
        Write-Host "Network Access: \\$env:COMPUTERNAME\PrintDrop" -ForegroundColor Cyan
        Write-Host "Service is monitoring for new files..." -ForegroundColor Green
    } else {
        Write-Host "STATUS: NOT RUNNING" -ForegroundColor Red
        Write-Host "Use 'manage_service.ps1 start' to start the service" -ForegroundColor Yellow
    }
}

# Main execution
switch ($Action.ToLower()) {
    "start" {
        $current = Test-ServiceRunning
        if ($current) {
            Write-Host "Service is already running (PID: $($current.Id))" -ForegroundColor Yellow
            Show-ServiceStatus
        } else {
            Start-AutoPrintService
        }
    }
    
    "stop" {
        Stop-AutoPrintService
    }
    
    "status" {
        Show-ServiceStatus
    }
    
    "restart" {
        Stop-AutoPrintService
        Start-Sleep -Seconds 1
        Start-AutoPrintService
    }
    
    default {
        Write-Host "Auto-Print Service Manager" -ForegroundColor Cyan
        Write-Host "Usage: manage_service.ps1 [start|stop|status|restart]" -ForegroundColor White
        Write-Host ""
        Write-Host "Commands:" -ForegroundColor Yellow
        Write-Host "  start   - Start the background service" -ForegroundColor White
        Write-Host "  stop    - Stop the background service" -ForegroundColor White
        Write-Host "  status  - Show current service status" -ForegroundColor White
        Write-Host "  restart - Restart the service" -ForegroundColor White
    }
}