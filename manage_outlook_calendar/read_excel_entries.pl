#!c:/perl/bin/perl.exe -w

use strict;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft Outlook';
use Spreadsheet::XLSX;
use DateTime::Format::Excel;

my $excel = Spreadsheet::XLSX -> new ('plan-vacaciones.xlsx');
my $sheet = $excel->worksheet("2015");



foreach my $row ($sheet->{MinRow}+2 .. $sheet->{MaxRow}) {
    my $cell_nombre = ($sheet->{Cells}[$row][0])->{Val};
	if( $cell_nombre !~ "Chaparro" ) { next; }
	foreach my $col ($sheet->{MinCol}+6 ..  $sheet->{MaxCol}) {
		my $cell = $sheet->{Cells}[$row][$col];
		if ($cell) {
			
			my $datecell = $sheet->{Cells}[0][$col];
			my $date = DateTime::Format::Excel->parse_datetime($datecell->value()); 
			my $year=$date->year; my $month=$date->month; my $day=$date->day; 
			printf("( %s , %s ) => %s : %s/%s/%s\n", $row, $col, $cell->{Val},$year,$month,$day);
		}#cell
	}#col
}#row

