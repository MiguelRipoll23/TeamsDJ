# Define the path to the registry key
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"

# Define webhook URLs
$COMMAND_MIC_ACTIVE = "echo 'Microphone activated'"
$COMMAND_MIC_INACTIVE = "echo 'Microphone deactivated'"
$WEBHOOK_MIC_ACTIVE = ""
$WEBHOOK_MIC_INACTIVE = ""

# Utility function to log messages with timestamp
function Write-Console {
  param (
    [string]$message
  )

  Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $message"
}

# Function to execute a command before sending a webhook
function Send-Command {
  param (
    [string]$command
  )

  try {
    Invoke-Expression -Command $command
    Write-Console "Executed command: $command"
  }
  catch {
    Write-Console "Failed to execute command: $_"
  }
}

# Function to check if the microphone is being used
function Test-MicrophoneUsage {
  $microphoneInUse = $false

  try {
    # Get all subkeys under NonPackaged
    $keys = Get-ChildItem -Path $registryPath -ErrorAction SilentlyContinue

    # Check if any application has LastUsedTimeStop set to 0 (indicating active usage)
    foreach ($key in $keys) {
      $lastUsedTimeStop = (Get-ItemProperty -Path $key.PSPath -Name "LastUsedTimeStop" -ErrorAction SilentlyContinue).LastUsedTimeStop

      if ($lastUsedTimeStop -eq 0) {
        $microphoneInUse = $true
        break
      }
    }
  }
  catch {
    Write-Console "Error checking microphone usage: $_"
  }

  return $microphoneInUse
}

# Function to send webhook notification
function Send-Webhook {
  param (
    [string]$webhookUrl
  )

  try {
    Invoke-WebRequest -Uri $webhookUrl -Method Post -ErrorAction Stop
    Write-Console "Webhook sent to $webhookUrl"
  }
  catch {
    Write-Console "Failed to send webhook to $webhookUrl due to an error: $_"
  }
}

Write-Console "Teams DJ script started"

# Initialize the previous state (microphone is not in use initially)
$previousState = $false

# Loop to check microphone usage every 100ms
while ($true) {
  $isMicInUse = Test-MicrophoneUsage

  # Only send the webhook if the state has changed
  if ($isMicInUse -ne $previousState) {
    if ($isMicInUse) {
      Write-Console "Microphone is active"
      Send-Command -command $COMMAND_MIC_ACTIVE
      Send-Webhook -webhookUrl $WEBHOOK_MIC_ACTIVE
    }
    else {
      Write-Console "Microphone is inactive"
      Send-Command -command $COMMAND_MIC_INACTIVE
      Send-Webhook -webhookUrl $WEBHOOK_MIC_INACTIVE
    }

    # Update the previous state to the current state
    $previousState = $isMicInUse
  }

  Start-Sleep -Milliseconds 100
}
