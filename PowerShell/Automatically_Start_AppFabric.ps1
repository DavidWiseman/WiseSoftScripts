###########################Config###########################
$port = 22233 # default port of 22233 will be used to start AppFabric
$emailFrom = "noreply@wisesoft.local" # email notification sent from this address
$emailTo = "destination@wisesoft.local" #email notification sent to this address
$smtpServer = "MYSMTPHOST" # Name/IP of SMTP host used to send email notification
$MaxRetry = 3 #Max number of retry attempts if host fails to start
$RetryAfter = 10 #Retry after 10 seconds
############################################################
Import-Module DistributedCacheAdministration
Use-CacheCluster

#Get computer name
$hostname = gc env:computername 

$RetryCount = 0
# Keep trying to start AppFabric service until MaxRetry threshold is reached
while ($retryCount -le $MaxRetry) {
	if ($RetryCount -gt 0){
		"Host failed to start retry $RetryCount of $MaxRetry aftet $RetryAfter seconds..."
		# failed to start service so wait for a specified period of time before attempting retry
		Start-Sleep -s $RetryAfter
	}
	# Check if any hosts are running in the cluster
	$ActiveHostCount = (Get-CacheHost | where {$_.Status -eq "UP" -or $_.Status -eq "STARTING"} | Measure-Object).Count

	# Try to start the cache host/cluster
	try {
		IF ($ActiveHostCount -eq 0) {
			# No hosts are running in the cluster, start the cluster
			'Starting cache cluster...'
			Start-CacheCluster 
		}
		else {
			# Existing hosts are running in the cluster, start the cache host
			"Starting cache host $hostname..."
			Start-CacheHost -HostName $hostname -CachePort $port
		}
	}
	catch {
		$ErrorMsg = $_.Exception.ToString()
		$ErrorMsg
	}

	# Check the current host status (It should be "UP")
	$CurrentHostStatus = (Get-CacheHost -HostName $hostname -CachePort 22233).Status
	IF ($CurrentHostStatus -eq "UP" -or $CurrentHostStatus -eq "STARTING"){
		BREAK
	}
	$RetryCount +=1
}
# Finished starting AppFabric at this point. Generate email notification message
"Sending email notification..."

# Send either Success/Failed email notification based on status
 if ($CurrentHostStatus -ne 'UP') {
	$subject = "AppFabric Host '$HostName' failed to start"
	$NotificationMsg= "<div class='LargeErrorMsg'>AppFabric service on '$HostName' failed to start.  Status is '$CurrentHostStatus'</div><br/>
					<div class='ErrorMsg'>$ErrorMsg</div><br/>"
	$subject
 }
 else{
	$subject = "AppFabric Host '$HostName' started successfully"
	$NotificationMsg= "<div class='LargeNotification'>AppFabric service on '$HostName' started successfully.</div><br/>"
	$subject
 }

$smtp = new-object Net.Mail.SmtpClient($smtpServer)
$msg = New-Object Net.Mail.MailMessage($emailFrom,$EmailTo)
$msg.Subject = $subject

# Build HTML email body text
$body = "<html>
<head>
<head>
<style type='text/css'>
h1 {
color:#FFFFFF;
font:bold 16pt arial,sans-serif;
background-color:#204c7d;
text-align:center;
}
table {
font:8pt tahoma,arial,sans-serif;
}
body {
color:#000000;
font:8pt tahoma,arial,sans-serif;
margin:0px;
padding:0px;
}
th {
color:#FFFFFF;
font:bold 8pt tahoma,arial,sans-serif;
background-color:#204c7d;
padding-left:5px;
padding-right:5px;
}
td {
color:#000000;
font:8pt tahoma,arial,sans-serif;
border:1px solid #DCDCDC;
border-collapse:collapse;
padding-left:3px;
padding-right:3px;
}
.Warning {
background-color:#FFFF00; 
color:#2E2E2E;
}
.Critical {
background-color:#FF0000;
color:#FFFFFF;
}
.Healthy {
background-color:#458B00;
color:#FFFFFF;
}
.ErrorMsg{ 
	color:red
}
.LargeErrorMsg{
	color:red;
	font:bold 14pt arial,sans-serif;
}
.LargeNotification{
	color:Green;
	font:bold 14pt arial,sans-serif;
}
</head>
<body><h1>AppFabric service start notification</h1>
$NotificationMsg
<table><tr><th>Host</th><th>Status</th></tr>"

# HTML table generated to list each host in the cluster with the status (UP,DOWN etc)
ForEach($CacheHost in Get-CacheHost) {
	$class = switch ($CacheHost.Status) 