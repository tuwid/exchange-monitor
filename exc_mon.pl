#!/usr/bin/perl

#An exchange monitoring set of subroutines 
#I'm not responsible for this, they made me write it againts my will as the client wanted it this way 

# the user should be in the format DOMAIN\USER
# if local privs are specified then its SERVERNAME\USER
# this works on windows as the API for the WMI its currently not available in linux

my $domain = '';
my $user = '';
my $password = '';
my $exchange_hostname = '';


sub exchange_stats{
	my $wmi_1 = create_wmi("$exchange_hostname","$domain\\$user","$password");

	my %properties;

	$properties{"Win32_PerfFormattedData_MSExchangeActiveSync_MSExchangeActiveSyncRequestsPersec"} = wmi_check("Win32_PerfFormattedData_MSExchangeActiveSync_MSExchangeActiveSync","RequestsPersec");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesAggregateDeliveryQueueLengthAllQueues"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","AggregateDeliveryQueueLengthAllQueues");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesMessagesQueuedForDelivery"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","MessagesQueuedForDelivery");
	#wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","AggregateDeliveryQueueLengthAllQueues");
	$properties{"Win32_PerfFormattedData_MSExchangeActiveSync_MSExchangeActiveSyncReuquestsPersec"} = wmi_check("Win32_PerfFormattedData_MSExchangeActiveSync_MSExchangeActiveSync","ReuquestsPersec");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportSmtpSend_MSExchangeTransportSmtpSendConnectionsTotal"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportSmtpSend_MSExchangeTransportSmtpSend","ConnectionsTotal");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportSMTPReceive_MSExchangeTransportSMTPReceiveConnectionsTotal"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportSMTPReceive_MSExchangeTransportSMTPReceive","ConnectionsTotal");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportSmtpSend_MSExchangeTransportSmtpSendMessagesSentPerSec"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportSmtpSend_MSExchangeTransportSmtpSend","MessagesSentPerSec");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesLargestDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","LargestDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesUnreachableQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","UnreachableQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesRetryNonSmtpDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","RetryNonSmtpDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesActiveNonSmtpDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","ActiveNonSmtpDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesSubmissionQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","SubmissionQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesRetryRemoteDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","RetryRemoteDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesActiveRemoteDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","ActiveRemoteDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesRetryMailboxDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","RetryMailboxDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueuesActiveMailboxDeliveryQueueLength"} = wmi_check("Win32_PerfFormattedData_MSExchangeTransportQueues_MSExchangeTransportQueues","ActiveMailboxDeliveryQueueLength");
	$properties{"Win32_PerfFormattedData_MSExchangeActiveUserCount"} = wmi_check("Win32_PerfFormattedData_MSExchange","ActiveUserCount"); # ??
	$properties{"Win32_PerfFormattedData_MSExchangeActiveConnectionCount"} = wmi_check("Win32_PerfFormattedData_MSExchange","ActiveConnectionCount"); # ??
	$properties{"Win32_PerfFormattedData_MSExchangeImap4_MSExchangeImap4CurrentConnections"} = wmi_check("Win32_PerfFormattedData_MSExchangeImap4_MSExchangeImap4","CurrentConnections");
	$properties{"Win32_PerfFormattedData_MSExchangePop3_MSExchangePop3ConnectionsCurrent"} = wmi_check("Win32_PerfFormattedData_MSExchangePop3_MSExchangePop3","ConnectionsCurrent");
	$properties{"Win32_PerfFormattedData_MSExchangeMailSubmission_MSExchangeMailSubmissionFailedSubmissionsPerSecond"} = wmi_check("Win32_PerfFormattedData_MSExchangeMailSubmission_MSExchangeMailSubmission","FailedSubmissionsPerSecond");
	$properties{"Win32_PerfFormattedData_MSExchangeIS_MSExchangeISMailboxActiveClientLogons"} = wmi_check("Win32_PerfFormattedData_MSExchangeIS_MSExchangeISMailbox","ActiveClientLogons");
	$properties{"Win32_PerfFormattedData_MSExchangeOWA_MSExchangeOWAAverageResponseTime"} = wmi_check("Win32_PerfFormattedData_MSExchangeOWA_MSExchangeOWA","AverageResponseTime");

	# inception of subrutines 
	sub wmi_check{
		my $inst = shift,
		my $elem = shift;
		my $tot = shift;

		my $conn = $wmi->InstancesOf("$inst") or die "snuk wmi-ja ne server\n";

			if(defined($tot)){
					foreach my $v (in $conn) {
					  #print "$inst\\$elem:".$v->{$elem}."\n" if $v->{Name} eq "_Total";
					  return($v->{$elem}) if $v->{Name} eq "_Total";

					}
				}
			else{
					foreach my $v (in $conn) {
					  #print "$inst\\$elem:".$v->{$elem}."\n";
					  return($v->{$elem});

					}
			}
		}
	}


	sub create_wmi {
		my ( $host, $user, $pass ) = @_;
		#X; 
		print "# ".$Thread." ".localtime()." trying wmi $host $user ".substr($pass,0,2)."****\n";
		my $wmi;

		my $loc = Win32::OLE->new('WbemScripting.SWbemLocator') || return "Cannot access WMI on local machine: ".Win32::OLE->LastError;

		if ( $host ne 'localhost' ) {
			$wmi = $loc->ConnectServer( $host, 'root/cimv2', $user, $pass ) || return "Cannot access WMI on remote machine: ".Win32::OLE->LastError;
		} else {
			$wmi = $loc->ConnectServer( $host, 'root/cimv2' ) || return "Cannot access WMI on remote machine: ".Win32::OLE->LastError;
		}
		#print "wmi OK\n";
		return $wmi;
	}


	sub wmi_listclasses {
		my ( $wmiService ) = @_;
			my $subDevices = $wmiService->SubclassesOf();
			my @classes;
			my %classes;
			foreach my $subDevProp ( in( $subDevices ) )
			{
				#print "Class $count: ".$subDevProp->{Path_}->{Path}."\n";
				if($subDevProp->{Path_}->{Path} =~ /.*:(.*)/){
					my $class = $1;
					if ($class =~ /Win32/) {
					#print "Detected class: $class\n";
					$classes{$class}=1;
					}
				}
			}

		foreach $class(sort keys %classes) {
			push(@classes,$class);
		}

		return @classes;
	}

	sub wmi_listproperties {
			my ( $wmiService, $class ) = @_;
			my @properties;
			my %properties;
			print "# ".$Thread." ".localtime()." SELECT * FROM $class\n";
			my $Computers = $wmiService->ExecQuery("SELECT * FROM $class");
			foreach my $pc (in ($Computers)) {
				%properties=wmi_properties($pc);
			}
		return %properties;
	}



	sub wmi_properties {
		my $node = $_[0];
		my @properties;
		my %properties;
		foreach my $object (in $node->{Properties_}) {
			if (ref($object->{Value}) eq "ARRAY") {
				#print " 1 $object->{Name} = { ";
				foreach my $value (in($object->{Value}) ) {
					#print "$value ";
					if (IsANumber($value)) {
						#print " INTEGER";
					}
					else {
						#print " STRING";
					}
				}
				#print "}\n";
			} 
			else {
			#print " 2 $object->{Name} = $object->{Value}";
			if (IsANumber($object->{Value})) {
				#	print " INTEGER\n";
				$properties{$object->{Name}}=$object->{Value};
			}
			else {
				#	print " STRING\n";
			}
			}
		}
	#print "----------------------------------\n";

	return %properties;
}
