#!c:/perl/bin/perl.exe -w

use strict;

my @list = ();

push @list, 332;

push @list, 335;

if ( 333 ~~ @list ) { print "332 Existe en la lista\n"; }

print @list;

if ( "2016" ge "2015" ) {
	print "Es mayor";
}