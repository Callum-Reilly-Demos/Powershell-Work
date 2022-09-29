
## sets user env variable for TNS_ADMIN
[System.Environment]::SetEnvironmentVariable('TNS_ADMIN','C:\TNSNames',
[System.EnvironmentVariableTarget]::User)

#needs to be filled in with info, but adds 32bit oracle odbc connection
Add-OdbcDsn -DsnType User -Platform '32-bit' -DriverName "Oracle in OraClient12Home1" -Name "" -SetPropertyValue @("Description=", "ServerName=", "DSN=")
