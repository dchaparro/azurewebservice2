#!c:/perl/bin/perl.exe -w

use strict;
use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';
use DateTime::Format::Excel;

# Current year
my $current_year="2015";
my $textVacBody = "Evento creado de forma automatica. Vacaciones. (NO MODIFICAR NADA)";
my $empName = "Raul Moreno";
my $textVacPrefixSubject = "VAC: ";

# Connection to Outlook Calendar    
my $outlook = Win32::OLE->new('Outlook.Application') or die
"Error!\n";

# Select Calendar
my $ns = $outlook->GetNameSpace("MAPI") or die "Couldn't open MAPI namespace.\n";
my $rootFolder= $ns->Folders("diego.chaparro\@acens.com") or die "Couldn't open 'Root Folders'.\n";
my $calFolder = $rootFolder->Folders("Calendar") or die "Couldn't open 'Calendar Folder'.\n";
my $sharedCalFolder = $calFolder->Folders("pruebas") or die "Couldn't open 'SharedCal Folder'.\n";

my $calItems = $sharedCalFolder->{Items};
$calItems->Sort("[Start]");
$calItems = $calItems->Restrict("[End] >= '01/01/$current_year 00:00'");
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
		#print "Entrada leida: ", $it->{Subject}, ": ", $date, "(", $dt, ") \n";
		if ($it->{Subject} eq "$textVacPrefixSubject$empName") {
			print "Borro la entrada: ", $empName, " ", $dt, "\n";
			$it->Delete;
			### Mejorable, si no empiezo desde el principio, no borra todas las entradas
			$calItems->Sort("[Start]");
			my $numCalItems = $calItems->{Count};
			my $it = $calItems->GetFirst;
			$i=0;
		}
	}
	$it = $calItems->GetNext;
}

