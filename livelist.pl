#! /usr/bin/perl

use feature qw/say/;
use strict;
use warnings;
use utf8;
use Encode;
use File::Basename;
use File::Path;
use File::Spec;
use File::Temp qw/tempfile tempdir/;
use HTTP::Request::Common;
use JSON;
use Time::Local;
use Getopt::Std;
use LWP;
use Time::Local;
use Win32::OLE 'CP_UTF8';
use Win32::OLE::Const;

$Win32::OLE::CP = CP_UTF8;
binmode STDOUT, ':raw:encoding(utf8)';
binmode STDERR, ':raw:encoding(utf8)';

$| = 1; select STDERR;
$| = 1; select STDOUT;
$, = "\t";

getopts 'o:pzl:', \my %opts;

my $targetdir = $opts{o} // 'i:/tmp';
my $userid = shift // exit;

my @uidlist = qw/711746 607591 6827278358 4344250372/;
my %uidlist = map { $_ => 1 } @uidlist;


sub ymd() {
    my ( $sec, $min, $hour, $day, $mon, $year ) = localtime;
    sprintf "%04d%02d%02d", $year + 1900, $mon + 1, $day;
}


sub ymdhms($) {
    my $ctime = shift;
    my ( $sec, $min, $hour, $day, $mon, $year ) = localtime int($ctime / 1000);
    sprintf "%04d-%02d-%02d %02d:%02d:%02d",
	$year + 1900, $mon + 1, $day, $hour, $min, $sec;
}


my $hostname = 'pocketapi.48.cn';
my $url = sprintf 'https://%s/live/api/v1/live/', $hostname;

my $ua = LWP::UserAgent->new;
$ua->agent( "PocketFans201807/6.0.13 (iPhone; iOS 10.3.3; Scale/2.00)" );
# $ua->proxy([qw/http https/], 'http://159.203.82.173:3128/');


sub getLiveList($$) {
    my $userid = shift;
    my $next = shift;
    my $params = { groupId => 0, 'next' => $next, loadMore => "true",
		   record => $opts{z} ? "false" : "true",
		   teamId => 0, userId => $userid };
    my $req = POST( $url . 'getLiveList' );
    say $req->method, $req->uri;
    my $paramstr = encode_json $params;
    $req->content_type( 'application/json; charset=utf-8' );
    $req->content( $paramstr );
    $req->header( 'Content-Length' => length $paramstr );
    $req->header( appInfo => '{"osType":"ios","vendor":"apple","os":"ios","appVersion":"6.0.13","osVersion":"10.3.3","deviceName":"iPhone 5","appBuild":"200513","deviceId":"DDDD-DDDD-DDDD-DDDD-DDDD"}' );

    my $res = $ua->request( $req );
    say $res->status_line;
    my $data = $res->content;
    unless ( $res->is_success ) {
	say $data;
	return undef, undef;
    }

    $data = decode_json $data;
    unless ( $data->{success} ) {
	say $data->{status}, $data->{message};
	return undef, undef;
    }

    my $content = $data->{content};
    $next = $content->{'next'};
    my $livelist = $content->{liveList};
    ( $next, $livelist );
}


sub getLiveOne($) {
    my $liveid = shift;
    my $params = { liveId => $liveid };
    my $req = POST( $url . 'getLiveOne' );
    my $paramstr = encode_json $params;
    $req->content_type( 'application/json; charset=utf-8' );
    $req->content( $paramstr );
    $req->header( 'Content-Length' => length $paramstr );

    my $res = $ua->request( $req );
    say $res->status_line;
    my $data = $res->content;
    unless ( $res->is_success ) {
	say $data;
	return;
    }

    $data = decode_json $data;
    unless ( $data->{success} ) {
	say $data->{status}, decode 'utf8', $data->{message};
	return;
    }

    $data->{content};
}


sub setcellvalue($$$) {
    my ( $sheet, $row, $data ) = @_;
    my $col = 'A';
    for my $value ( @$data ) {
	if ( $value and length( $value ) > 10 ) {
	    $sheet->Range( "$col$row" )->{NumberFormatLocal} = "@";
	}
	$sheet->Range( "$col$row" )->{Value} = $value;
	$col++;
    }
}


sub check_userid($$$) {
    unless ( $opts{l} or $opts{z} ) {
	return;
    }
    my ( $sheet, $row, $uid ) = @_;
    $uidlist{$uid} or return;
    my $range = $sheet->Range( "B$row" )->Interior;
    $range->{Pattern} = 1;
    $range->{ThemeColor} = 8;
    $range->{TintAndShade} = 0.8;
    $range->{PatternTintAndShade} = 0;
}


sub sleep2($) {
    my $len = shift;
    for ( 1 .. $len ) {
	print '.';
	sleep 2;
    }
    say '.';
}


my $ex = Win32::OLE->GetActiveObject( 'Excel.Application' );
$@ and die $@;
unless ( defined $ex ) {
    $ex = Win32::OLE->new( 'Excel.Application', 'Quit' ) or die $@;
}
$ex->{Visible} = 1;
my $book = $ex->Workbooks->Add;
my $sheet = $book->ActiveSheet;
$sheet->Activate;
$sheet->{Name} = ymd;

my @header1 = qw/liveId title liveType status ctime liveMode pictureOrientation/;
my @header2 = qw/userId nickname level isStar friends followers
    signature vip userRole/;
my @header3 = qw/roomId onlineNum type review needForward systemMsg
    msgFilePath playStreamPath mute/;

setcellvalue $sheet, 1, [ @header1, @header2, @header3 ];

$ex->ActiveWindow->{SplitColumn} = 0;
$ex->ActiveWindow->{SplitRow} = 1;
$ex->ActiveWindow->{FreezePanes} = 1;
$sheet->Columns("A:A")->{ColumnWidth} = 20;
$sheet->Columns("C:D")->{ColumnWidth} = 2;
$sheet->Columns("E:E")->{ColumnWidth} = 20;
$sheet->Columns("F:G")->{ColumnWidth} = 2;
$sheet->Columns("H:H")->{ColumnWidth} = 7;
$sheet->Columns("I:I")->{ColumnWidth} = 14;
$sheet->Columns("J:P")->{ColumnWidth} = 2;
$sheet->Columns("Q:R")->{ColumnWidth} = 8;
$sheet->Columns("S:V")->{ColumnWidth} = 2;
$sheet->Columns("W:W")->{ColumnWidth} = 3;
$sheet->UsedRange->AutoFilter;
$sheet->UsedRange->Interior->{Pattern} = 1;
$sheet->UsedRange->Interior->{ThemeColor} = 10;
$sheet->UsedRange->Interior->{TintAndShade} = 0.8;
$sheet->UsedRange->Interior->{PatternTintAndShade} = 0;

my $sortfunc = sub { $b->{liveId} cmp $a->{liveId} };
if ( $opts{l} or $opts{z} ) {
    $sortfunc = sub { $a->{liveId} cmp $b->{liveId} };
}

my %liveidlist;
for ( my( $next, $row ) = ( 0, 2 );;) {
    ( $next, my $livelist ) = getLiveList $userid, $next;
    $livelist or next;
    @$livelist or $opts{l} or last;
    for my $live ( sort $sortfunc @$livelist ) {
	my $liveid = $live->{liveId};
	say $liveid;
	$liveidlist{$liveid} and next;
	my $liveinfo = getLiveOne $liveid or next;
	$liveidlist{$liveid} = 1;
	$live->{ctime} = ymdhms $live->{ctime};
	setcellvalue $sheet, $row,
	    [ map( $live->{$_}, @header1 ),
	      map( $live->{userInfo}->{$_}, @header2 ),
	      map( $liveinfo->{$_}, @header3 ) ];
	check_userid $sheet, $row, $live->{userInfo}->{userId};
	$row++;
    }
    $opts{p} and last;
    $opts{l} and sleep2 $opts{l} and next;
    $opts{z} and last;
    sleep 1;
}

exit;
