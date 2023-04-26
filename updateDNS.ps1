Add-Content -Path $env:windir\System32\drivers\etc\hosts -Value "`n10.0.0.4`tzndemo.com" -Force
Set-DNSClientServerAddress "Ethernet" -ServerAddresses ("10.0.0.4","8.8.8.8")
