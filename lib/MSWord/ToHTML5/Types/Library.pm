package MSWord::ToHTML5::Types::Library;

use namespace::autoclean;
use strictures 1;
use MooseX::Types::Moose qw/Str/;
use MooseX::Types::IO::All qw/:all/;
use Moose::Util::TypeConstraints;
use MooseX::Types -declare => [qw/MyFile MSDoc MSDocX/];
use Try::Tiny;
use Text::Extract::Word;
use Archive::Zip qw/:ERROR_CODES :CONSTANTS/;
use File::Spec;

{
    local *STDERR;
    local *STDOUT;
    open( STDOUT, '>', File::Spec->devnull() );
    open( STDERR, '>', File::Spec->devnull() );

    subtype MyFile, as IO_All, where {
      length($_) > 0;
    }, message {
      "Did you pass a file?";
    };

    coerce MyFile, from Str, via { to_IO_All($_) };

    subtype MSDoc, as MyFile, where {
      try {
        Text::Extract::Word->new( $_->filepath . $_->filename );
      };
    }, message {
      "$_ does not appear to be a Word doc file";
    };

    coerce MSDoc, from MyFile | IO_All, via {$_};

    subtype MSDocX, as MyFile, where {
      my $unzip = Archive::Zip->new;
      Archive::Zip->new( $_->filepath . $_->filename )
        && Archive::Zip::MemberRead->new( $unzip, "word/document.xml" );
    }, message {
      "$_ does not appear to be a Word docx file";
    };

    coerce MSDocX, from MyFile | IO_All, via {$_};
}

__PACKAGE__->meta->make_immutable;

1;
