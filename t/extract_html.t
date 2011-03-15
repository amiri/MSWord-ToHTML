#!/usr/bin/env perl

use strict;
use warnings;
use Devel::Dwarn;
use Test::Most qw/no_plan/;
use lib 'lib';
use MSWord::ToHTML5;
use Archive::Zip;
use IO::All;

my @docs  = glob('t/data/*.doc');
my @docxs = glob('t/data/*.docx');

my $converter;

my @type_methods     = qw/to_File/;
my @instance_methods = qw/validate_file/;

lives_ok { $converter = MSWord::ToHTML5->new } "My constructor works OK";

#can_ok( $converter, @type_methods );
can_ok( $converter, @instance_methods );

for my $doc (@docs) {
    my ( $response, $xml_dom,$html );
    lives_ok { $response = $converter->validate_file($doc) }
    "My recognize functions works OK for docs";
    lives_ok { $xml_dom = $converter->get_dom($response) } "My xml_dom is OK";
    lives_ok { $html = $converter->convert_to_html($xml_dom) } "My html is OK";
    #diag Dwarn $html;
}

for my $docx (@docxs) {
    my ( $response, $xml_dom,$html );
    lives_ok { $response = $converter->validate_file($docx) }
    "My recognize functions works OK for docxs";
    lives_ok { $xml_dom = $converter->get_dom($response) } "My xml_dom is OK";

    lives_ok { $html = $converter->convert_to_html($xml_dom) } "My html is OK";
    diag Dwarn $html;
}

#lives_ok { $converter->convert_doc($_) } "My converter works OK" for @docs;

#throws_ok { $converter->convert_doc($_) } qr/This is not a valid Word doc file/ for @docxs;

#lives_ok { $converter->convert_docx($_) } "My converter works OK" for @docxs;
#throws_ok { $converter->convert_docx($_) } qr/That does not appear to be a Word docx file/ for @docs;

#my @zips = map { Archive::Zip->new($_) } @docxs;
#my @files = map { $_->memberNamed('word/document.xml') } @zips;

#diag Dwarn \@files;

