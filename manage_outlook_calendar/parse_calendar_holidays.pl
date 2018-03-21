#!c:/perl/bin/perl.exe -w

use strict;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';
use Spreadsheet::XLSX;
use DateTime::Format::Excel;

# Current year
my $currentYear = "2015";
my $textVacBody = "Evento creado de forma automatica. Vacaciones. (NO MODIFICAR NADA)";
my $textVacPrefixSubject = "VAC: ";

# Connection to Outlook Calendar    
my $outlook = Win32::OLE->new('Outlook.Application') or die "Error!\n";
# Select calendar folder
my $ns = $outlook->GetNameSpace("MAPI") or die "Couldn't open MAPI namespace.\n";
my $rootFolder= $ns->Folders("diego.chaparro\@acens.com") or die "Couldn't open 'Root Folders'.\n";
my $calFolder = $rootFolder->Folders("Calendar") or die "Couldn't open 'Calendar Folder'.\n";

# Read excel file entries
my $excel = Spreadsheet::XLSX -> new ('plan-vacaciones.xlsx');
my $sheet = $excel->worksheet($currentYear);
foreach my $row ($sheet->{MinRow}+2 .. $sheet->{MaxRow}) {
	my $calName = ($sheet->{Cells}[$row][3])->{Val};
	if(! $calName) { next; }
	# Read Current Entries from Calendar for this employee
    my $empName = ($sheet->{Cells}[$row][0])->{Val};
	my @curCalEntList = readEntriesCal ($empName, $calName);
	my @curExcEntList = ();
	# Add entries to Calendar
	foreach my $col ($sheet->{MinCol}+7 ..  $sheet->{MaxCol}) {
		my $cell = $sheet->{Cells}[$row][$col];
		# If the cell has a value, the day should be registered as an event
		if ($cell) {
			my $date = ($sheet->{Cells}[0][$col])->value();
			push @curExcEntList, $date;
			# If the entry doesn't exist in calendar
			if ( ! ($date ~~ @curCalEntList)) {
				addEntryCal ($empName, $calName, $date);
			}
		}#cell
	}#col
	
	# Read Calendar entries and remove not valid
	
	# Select Calendar, and Items on this year
	my $sharedCalFolder = $calFolder->Folders($calName) or die "Couldn't open 'SharedCal Folder'.\n";
	my $calItems = $sharedCalFolder->{Items};
	$calItems->Sort("[Start]");
	$calItems = $calItems->Restrict("[End] >= '01/01/$currentYear 00:00'");
	my $numCalItems = $calItems->{Count};
	my $it = $calItems->GetFirst;
	for (my $i=0;$i<$numCalItems;$i++) {
		# If year, user name and body message is what expected
		my $eventYear = substr $it->{Start}->Date, 6, 4;
		if (($it->{Body} eq $textVacBody) && ($eventYear ge 2015 ) && ($it->{Subject} eq "$textVacPrefixSubject$empName")) {
			my $eventDay = substr $it->{Start}->Date, 0, 2;
			my $eventMonth = substr $it->{Start}->Date, 3, 2;
			my $dt = DateTime->new( year => $eventYear, month => $eventMonth, day => $eventDay );
			my $date = DateTime::Format::Excel->format_datetime($dt);
			print "Entrada leida: ", $it->{Subject}, ": ", $date, "(", $dt, ") \n";
			if ( ! ($date ~~ @curExcEntList)) {
				print "Borro la entrada: ", $empName, " ", $dt, "\n";
				$it->Delete;
				### Mejorable, si no empiezo desde el principio, no borra todas las entradas
				if ( $i+1 < $numCalItems ) {
					$calItems->Sort("[Start]");
					$numCalItems = $calItems->{Count};
					my $it = $calItems->GetFirst;
					$i=0;
				}
			} else {
				$it = $calItems->GetNext;
			}
		} else {
			$it = $calItems->GetNext;
		}
	}
	print "Salgo del bucle\n";
}#row

# Add entry to Calendar
sub addEntryCal {
	my ($empName, $calName, $dt) = @_;
	my $date = DateTime::Format::Excel->parse_datetime($dt); 
	my $year=$date->year;
	my $month=$date->month;
	my $day=$date->day; 
	my $sharedcalFolder = $calFolder->Folders($calName) or die "Couldn't open 'SharedCal Folder'.\n";

	#print "Comenzando a añadir a Calendario: Vacaciones ", $empName, ": ", $day, "/", $month, "/", $year, "\n";
	my $newappt = $outlook->CreateItem(olAppointmentItem);
	$newappt->{Subject} = "VAC: $empName";
	$newappt->{Start} = "$day/$month/$year 8:00:00 AM";
	$newappt->{Duration} = 660 ;  #1 minute
	$newappt->{Body} = $textVacBody;
	#$newappt->{ReminderMinutesBeforeStart} = 1;
	$newappt->{ReminderSet} = 0;
	$newappt->{BusyStatus} = olBusy;
	$newappt->Save();
	$newappt->Move($sharedcalFolder) or die "Couldn't move Calendar Entry \n";
	print "Entrada añadida a Calendario: Vacaciones ", $empName, ": ", $day, "/", $month, "/", $year, "\n";

}

# Read Calendar entries for a $empName in a $calName
sub readEntriesCal {
	my ($empName, $calName) = @_;
	my @entriesList = ();
	
	# Select Calendar, and Items on this year
	my $sharedCalFolder = $calFolder->Folders($calName) or die "Couldn't open 'SharedCal Folder'.\n";
	my $calItems = $sharedCalFolder->{Items};
	$calItems->Sort("[Start]");
	$calItems = $calItems->Restrict("[End] >= '01/01/$currentYear 00:00'");
	
	my $numCalItems = $calItems->{Count};
	my $it = $calItems->GetFirst;
	for (my $i=0;$i<$numCalItems;$i++) {
		# If year, user name and body message is what expected
		my $eventYear = substr $it->{Start}->Date, 6, 4;
		if (($it->{Body} eq $textVacBody) && ($eventYear ge $currentYear ) && ($it->{Subject} eq "$textVacPrefixSubject$empName")) {
			my $eventDay = substr $it->{Start}->Date, 0, 2;
			my $eventMonth = substr $it->{Start}->Date, 3, 2;
			my $dt = DateTime->new( year => $eventYear, month => $eventMonth, day => $eventDay );
			my $date = DateTime::Format::Excel->format_datetime($dt);
			#print "Entrada leida: ", $it->{Subject}, ": ", $date, "\n";
			push @entriesList, $date;
		}
		$it = $calItems->GetNext;
	}
	#print "LISTA FECHAS: ", @entriesList, "\n";
	return @entriesList;
}