#!c:/perl/bin/perl.exe -w

use strict;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';

my $oFolder;            # Folder object
my $folder;             # Folder name
my %Folders;            # Hash table of folder names -> EntryID

# Connection to Outlook Calendar    
my $outlook = Win32::OLE->new('Outlook.Application') or die
"Error!\n";

# Select Calendar
my $ns = $outlook->GetNameSpace("MAPI") or die "Couldn't open MAPI namespace.\n";
my $rootFolder= $ns->Folders("diego.chaparro\@acens.com") or die "Couldn't open 'Root Folders'.\n";
my $calFolder = $rootFolder->Folders("Calendar") or die "Couldn't open 'Calendar Folder'.\n";
my $sharedCalFolder = $calFolder->Folders("pruebas") or die "Couldn't open 'SharedCal Folder'.\n";

######################## Insertar evento #################
my $newappt = $outlook->CreateItem(olAppointmentItem);

$newappt->{Subject} = "Prueba (Ignorar)";

$newappt->{Start} = "12/07/2015 8:00:00 AM";
$newappt->{Duration} = 660 ;  #1 minute
#$newappt->{Location} = "Someplace";
$newappt->{Body} = "Test Stuff";
#$newappt->{ReminderMinutesBeforeStart} = 1;
$newappt->{BusyStatus} = olBusy;
$newappt->Save() or die "Couldn't save Calendar Entry \n";
$newappt->Move($sharedCalFolder);