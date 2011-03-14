#!/usr/bin/env perl

use strict;
use warnings;
use Devel::Dwarn;
use Test::Most qw/no_plan/;

use Text::Extract::Word;

my ( $footnotes, $text, $anno, $bookmarks, $body );

my @files = glob('t/data/*.doc');


for my $doc (@files) {
    my $file;
    lives_ok { $file = Text::Extract::Word->new($doc) }
    "I can extract text from $doc";

    lives_ok { $footnotes = $file->get_footnotes(':raw') }
    "I can extract footnotes from $doc";
    lives_ok { $text = $file->get_text } "I can extract text from $doc";
    lives_ok { $anno = $file->get_annotations }
    "I can extract annotations from $doc";
    lives_ok { $bookmarks = $file->get_bookmarks }
    "I can extract bookmarks from $doc";
    lives_ok { $body = $file->get_body(':raw') }
    "I can extract body from $doc";

    my $fn_count     = 1;
    my $fn_ref_count = 1;
    my $fn_link      = 'footnote_';
    my $fn_ref       = 'footnote_ref_';

    $body =~ s/\x02/$fn_link . $fn_count++/ge;
    $body = Text::Extract::Word::_filter( $body, undef );

    $footnotes =~ s/\x02/$fn_ref . $fn_ref_count++/ge;
    $footnotes = Text::Extract::Word::_filter( $footnotes, undef );
}
