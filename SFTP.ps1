#
# Microsoft SQL.ps1 - IDM System PowerShell Script for Microsoft SQL Server.
#
# Any IDM System PowerShell Script is dot-sourced in a separate PowerShell context, after
# dot-sourcing the IDM Generic PowerShell Script '../Generic.ps1'.
#

$Log_MaskableKeys = @(
    'password'
)

#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log info "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(
            @{
                name = 'server'
                type = 'textbox'
                label = 'Server'
                description = 'Name or IP of SFTP server'
                value = ''
            }
            @{
                name = 'port'
                type = 'textbox'
                label = 'Port'
                description = 'Port the SFTP server is listening on'
                value = 22
            }
            @{
                name = 'username'
                type = 'textbox'
                label = 'Username'
                description = 'User account name to access Microsoft SQL server'
                value = ''
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                description = 'User account password to access Microsoft SQL server'
                value = ''
            }
            @{
                name = 'remote_directory'
                type = 'textbox'
                label = 'Remote Directory'
                description = 'The remote directory you wish to connect to on the SFTP server'
                value = ''
            }
            @{
                name = 'local_directory'
                type = 'textbox'
                label = 'Local Directory'
                description = 'The local directory where you wish to store the files from the SFTP server'
                value = ''
            }
            @{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                description = ''
                value = 5
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                description = ''
                value = 30
            }
        )
    }

    if ($TestConnection) {
        Open-SFTPConnection $ConnectionParams;
        Get-SFTPChildItem -SFTPSession $Global:SFTPSession | Out-Null;
        Close-SFTPConnection;
    }

    if ($Configuration) {
        Open-SFTPConnection $ConnectionParams;
        $files = Get-SFTPChildItem -SFTPSession $Global:SFTPSession;

        $filesComboDisplay = @( @{ id = ''; value = '' } );
        $filesComboDisplay += $files | Select-Object @{n="id"; e={$_.FullName}},@{n="value"; e={$_.FullName + " | " + $_.LastWriteTime.ToString() + " | (" + ([Math]::Round($_.Length / 1KB,3)).ToString() + "KB)" }}
        

        @(
            @{
                name = 'CSV1'
                type = 'combo'
                label = 'CSV 1'
                description = 'CSV you wish to load into the CSV1 table'
                table = @{
                    rows = $filesComboDisplay
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'value'
                    }
                }
                value = ''    # Default value of combo item
            }
            @{
                name = 'CSV2'
                type = 'combo'
                label = 'CSV 2'
                description = 'CSV you wish to load into the CSV2 table'
                table = @{
                    rows = $filesComboDisplay
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'value'
                    }
                }
                value = ''    # Default value of combo item
            }
            @{
                name = 'CSV3'
                type = 'combo'
                label = 'CSV 3'
                description = 'CSV you wish to load into the CSV3 table'
                table = @{
                    rows = $filesComboDisplay
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'value'
                    }
                }
                value = ''    # Default value of combo item
            }
            @{
                name = 'CSV4'
                type = 'combo'
                label = 'CSV 4'
                description = 'CSV you wish to load into the CSV4 table'
                table = @{
                    rows = $filesComboDisplay
                    settings_combo = @{
                        value_column = 'id'
                        display_column = 'value'
                    }
                }
                value = ''    # Default value of combo item
            }
        )
    }

    Log info "Done";
}

function Idm-CSV1Read {
    param (
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    $system_params = ConvertFrom-Json $SystemParams;

    if ($GetMeta) {
        @()
    }
    else {
        if ($system_params.CSV1) {
            Import-CSVFromSFTP -SystemParams $SystemParams -FullFilePath $system_params.CSV1;
        }
    }

    Log info "Done";
}

function Idm-CSV2Read {
    param (
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    $system_params = ConvertFrom-Json $SystemParams;

    if ($GetMeta) {
        @()
    }
    else {
        if ($system_params.CSV2) {
            Import-CSVFromSFTP -SystemParams $SystemParams -FullFilePath $system_params.CSV2;
        }        
    }

    Log info "Done";
}

function Idm-CSV3Read {
    param (
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    $system_params = ConvertFrom-Json $SystemParams;

    if ($GetMeta) {
        @()
    }
    else {
        if ($system_params.CSV3) {
            Import-CSVFromSFTP -SystemParams $SystemParams -FullFilePath $system_params.CSV3;
        }        
    }

    Log info "Done";
}

function Idm-CSV4Read {
    param (
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log info "-GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"

    $system_params = ConvertFrom-Json $SystemParams;

    if ($GetMeta) {
        @()
    }
    else {
        if ($system_params.CSV4) {
            Import-CSVFromSFTP -SystemParams $SystemParams -FullFilePath $system_params.CSV4;
        }        
    }

    Log info "Done";
}

function Import-CSVFromSFTP {
    param (
        # Parameters
        [string] $SystemParams,
        [string] $FullFilePath
    )

    Log info "-SystemParams='$SystemParams' -FullFilePath='$FullFilePath'"

    $system_params = ConvertFrom-Json $SystemParams;

    $FullFilePath = $fullFilePath.Replace(':', '');

    #Remove duplicate slashes from local directory path
    $local_directory = $system_params.local_directory.Replace('//', '/');  

    #Get the file name from the full file path
    $lastSlash = $FullFilePath.LastIndexOf('/') + 1;
    $fileNameLength = $FullFilePath.length - $lastSlash;
    $fileName = $FullFilePath.Substring($lastSlash, $fileNameLength);
    $localFilePath = $local_directory + "\" + $fileName;              
    
    Open-SFTPConnection $SystemParams;

    #Retrieve SFTP file from server
    Get-SFTPItem -SFTPSession $Global:SFTPSession -Path "$fileName" -Destination "$local_directory" -Force;

    Close-SFTPConnection;

    #Import the CSV file
    $csv = Import-CSV -Path $localFilePath;
    #Extract the CSV headers
    $csvHeaders = $csv | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name';
    
    #Construct NIM objects
    foreach ($line in $csv) {
        $output = @{};
        foreach ($header in $csvHeaders) {
            $output[$header] = $line."$header";
        }
        New-Object -TypeName PSCustomObject -Property $output;
    }

    Log info "Done";
}

function Idm-OnUnload {
    Close-SFTPConnection;
}

function Open-SFTPConnection {
    param (
        [string] $ConnectionParams
    )

    if (!(Get-Module Posh-SSH)) {
        if (Get-Module -ListAvailable | Where-Object {$_.Name -eq "Posh-SSH"}) {
            Log info "Loading Posh-SSH";
            Import-Module Posh-SSH;
        }
        else {
            throw "Cannot find Posh-SSH module in PowerShell please run 'Install-Module Posh-SSH' in PowerShell and try again";
        }
    }

    $connection_params = ConvertFrom-Json $ConnectionParams;

    # Convert password to secure string
    $secStringPassword = ConvertTo-SecureString $connection_params.password -AsPlainText -Force;

    # Create credential object from username and secure string password
    $credential = New-Object System.Management.Automation.PSCredential($connection_params.username, $secStringPassword);

    if ($Global:SFTPSession.connected) {
        #Log debug "Reusing SFTPSession"
    }
    else {
        Log info "Opening SFTP connection to '$($connection_params.server)' on port '$($connection_params.port)' with user account '$($connection_params.username)'";

        try {
            $Global:SFTPSession = New-SFTPSession -Credential $credential -Port $connection_params.port -ComputerName $connection_params.server -Force -WarningAction:SilentlyContinue;

            if ($connection_params.remote_directory) {
                if (!($connection_params.remote_directory.StartsWith('/'))) {
                    throw "Remote directory must start with a '/' and use '/' as a path seperator";
                }

                Set-SFTPLocation -SFTPSession $Global:SFTPSession -Path $connection_params.remote_directory;
            }

            if (!(Test-Path $connection_params.local_directory)) {
                throw "Local directory does not exist, please correct this and try again";
            }
        }
        catch {
            Log error "Failed: $_";
            Write-Error $_;
        }

        Log info "Done";
    }
}

function Close-SFTPConnection {
    Remove-SFTPSession -SFTPSession $Global:SFTPSession | Out-Null;
}