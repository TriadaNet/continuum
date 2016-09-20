 get-service "WinHTTP Web Proxy Auto-Discovery Service" | Stop-Service -PassThru -Force | Set-Service -StartupType disabled
