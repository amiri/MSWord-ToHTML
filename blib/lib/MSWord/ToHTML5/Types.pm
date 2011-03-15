package MSWord::ToHTML5::Types;

use MooseX::Types::Moose qw/:all/;
use Moose::Util::TypeConstraints;
use MooseX::Types::IO::All qw/:all/;
use MooseX::Types -declare =>
    [qw/File Doc DocX XML/];
use Archive::Zip qw/:ERROR_CODES :CONSTANTS/;
use Archive::Zip::MemberRead;
use autodie;
use File::Spec;

subtype File, as type IO_All, where {
    length($_) > 0 && -e "$_";
}, message {
    "I didn't understand the file you gave me";
};
coerce File, from Str, via { to_IO_All $_ };

subtype Doc, as File, where {

    open( STDOUT, '>', File::Spec->devnull );
    #open( STDERR, '>', File::Spec->devnull );
    my $text = system( 'antiword', $_ );
    ${^CHILD_ERROR_NATIVE} == 0;
}, message {
    "This is not a valid Word doc file.";
};

subtype XML, as IO_All;

subtype DocX, as File, where {

    #open( STDERR, '>', File::Spec->devnull );
    my $unzip = Archive::Zip->new;
    $unzip->read( $_ ) == AZ_OK;
}, message {
    "That does not appear to be a Word docx file";
};

coerce XML, from Doc, via {
    system( 'antiword', '-m', '8859-1', '-x', 'db', $_ )
}, from DocX, via {
    my $zip = Archive::Zip->new( $_ );
    $zip->memberNamed("word/document.xml")
};

__PACKAGE__->meta->make_immutable;

1;
