[Reflection.Assembly]::LoadFile("c:\program files\microsoft\exchange\web services\1.1\Microsoft.Exchange.WebServices.dll")
$s = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP1)
$s.Credentials = new-object net.networkcredential('xxxxxxxx', 'xxxxxxx', 'domainname.com')
$s.AutoDiscoverUrl("queues@peoplematter.com")
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($s, [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox)
$softdel = [Microsoft.Exchange.WebServices.Data.DeleteMode]::SoftDelete

$properties = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
$properties.RequestedBodyType = [Microsoft.Exchange.WebServices.Data.BodyType]::Text

while ($true) 
{
	$inbox.FindItems(1) | % {
	
		$_.Load($properties)
		
        #Grab queue number, verb, extension/cellnumber from e-mail subject
		if ($_.Subject -match "^(\d+)\s(\w+)\s(\d+)$")
		{
			# initialize strings
			$queueinfo = ""
			$query = ""

			# Grab queue number (5000/2000)
			$queuenum = $matches[1]

			# Grab queue choice (on/off)
			$queuechoice = $matches[2]

			# Grab extension/phone number
			$phone = $matches[3]

			$queryon = "asterisk -r -x 'queue add member Local/" + $phone + "@from-internal/n to " + $queuenum + "'"
			$queryoff = "asterisk -r -x 'queue remove member Local/" + $phone + "@from-internal/n from " + $queuenum + "'"

			if ($queuechoice -eq "on")
			{
				write-host "query on!"
				$msg = c:\scripts\plink.exe asteriskserver.domain.com -l root -i server.ppk $queryon
			}
			else
			{
				write-host "query off!"
				$msg = c:\scripts\plink.exe asteriskserver.domain.com -l root -i server.ppk $queryoff
			} 
				
			# Convert array into string (im sure there's a better way... no time :))

			foreach ($line in $msg)
			{
				$query += $line 
			}

			# Regex response to grab 'add/remove/already added' action + number + queue being modified
			if ($query -match "^(\w+).+Local/(\d+).+queue '(\d+)'")
			{
				$action = $matches[1]
				$number = $matches[2]
				$queue = $matches[3]
			}		

			# rewrite e-mail response if received unable response from asterisk
			if ($action -eq "unable")
			{
				$query = $number + " already on/off queue : " + $queue
			}
			else
			{
				$query = $action + " " + $number + " to/from queue: " + $queue
			}

			$checkqueue = "asterisk -r -x 'queue show " + $queuenum + "'" 
			
			$users = c:\scripts\plink.exe asteriskserver.domainname.com -l root -i server.ppk $checkqueue

			#loop through lines in queue
			foreach ($line in $users)
			{
				# Grab only the phone number on each line in queue
				if ($line -match "Local/(\d+)")
				{
					$queueinfo += "- " + $matches[1] + "<BR><BR>"
				}
			}
		
			$finished = "***** " + $query + " *****<BR><BR><BR>***** Updated Queue List *****<BR><BR>" + $queueinfo
			
			$_.reply($finished, $false)
		}
		elseif ($_.Subject -match "info")
		{
			$queueinfo = ""

			$checkqueue = "asterisk -r -x 'queue show 5000'"
			
			$users = c:\scripts\plink.exe asteriskserver.domainname.com -l root -i server.ppk $checkqueue

			#loop through lines in queue

			foreach ($line in $users)
			{
				# Grab only the phone number on each line in queue
				if ($line -match "Local/(\d+)")
				{
					$queueinfo += "- " + $matches[1] + "<BR><BR>"
				}
			}
		
			$finished = "***** Updated Queue 5000 Support List *****<BR><BR>" + $queueinfo
			
			$checkqueue = "asterisk -r -x 'queue show 2000'"

			$users = c:\scripts\plink.exe asteriskserver.domain.com -l root -i server.ppk $checkqueue

			#loop through lines in queue

			$queueinfo = ""

			foreach ($line in $users)
			{
				# Grab only the phone number on each line in queue
				if ($line -match "Local/(\d+)")
				{
					$queueinfo += "- " + $matches[1] + "<BR><BR>"
				}
			}


			$finished += "<BR><BR> ***** Updated Queue 2000 Sales List *****<BR><BR>" + $queueinfo
			
			$_.reply($finished, $false)
		}

	$_.Delete($softdel)

	}
			
	start-sleep 5
}
