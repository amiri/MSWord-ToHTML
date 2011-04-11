package MSWord::ToHTML5;

use Moose;
use namespace::autoclean;
use strictures 1;
use MooseX::Method::Signatures;
use MSWord::ToHTML5::Types::Library qw/:all/;
use Carp;
use MSWord::ToHTML5::Doc;
use MSWord::ToHTML5::DocX;
use Try::Tiny;

# ABSTRACT: Take old or new Word files and spit out HTML5

method validate_file( MyFile $file does coerce ) {
    try {
        return MSWord::ToHTML5::Doc->new( file => $file ) if is_MSDoc($file);
        return MSWord::ToHTML5::DocX->new( file => $file )
            if is_MSDocX($file);
    }
    catch {
        confess "I don't know what to do with this file: $_";
    };
}

__PACKAGE__->meta->make_immutable;

1;
