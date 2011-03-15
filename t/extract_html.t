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
    my ( $response, $xml_dom, $html, $html5 );
    lives_ok { $response = $converter->validate_file($doc) }
    "My recognize functions works OK for docs";
    lives_ok { $xml_dom = $converter->get_dom($response) } "My xml_dom is OK";
    lives_ok { $html    = $converter->convert_doc_to_html($xml_dom) }
    "My html is OK";
    lives_ok { $html5 = $converter->convert_to_html5($html) }
    "My html5 is OK";
    diag Dwarn $html5;
}

for my $docx (@docxs) {
    my ( $response, $xml_dom, $html, $html5 );
    lives_ok { $response = $converter->validate_file($docx) }
    "My recognize functions works OK for docxs";
    lives_ok { $xml_dom = $converter->get_dom($response) } "My xml_dom is OK";
    lives_ok { $html    = $converter->convert_docx_to_html($xml_dom) }
    "My html is OK";
    lives_ok { $html5 = $converter->convert_to_html5($html) }
    "My html5 is OK";
    diag Dwarn $html5;
}
