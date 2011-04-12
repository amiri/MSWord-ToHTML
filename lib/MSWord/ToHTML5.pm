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

our $VERSION = '0.001';

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

__END__

=head1 NAME

MSWord::ToHTML5

Take old or new format Word files and spit out HTML5.

=head1 NOTICE

Because of the PITA involved in processing Word files, I have punted most of
the work to L<Abiword|http://www.abisource.com/> and L<tidy|http://tidy.sourceforge.net/>.

Which means that you must have the binary programs tidy and abiword installed.

=head1 SYNOPSIS

    {
        package My::Word::Converter;

        use strict;
        use warnings;
        use MSWord::ToHTML5;

        my $converter = MSWord::ToHTML5->new;
        my $doc = $converter->validate_file("/home/myself/my_excellent_writing.doc");
          # This returns an instance of MSWord::ToHTML5::Doc
        my $docx = $converter->validate_file("/home/myself/my_excellent_notes.docx");
          # This returns an instance of MSWord::ToHTML5::DocX

        my $writing_html = $doc->get_html;
        my $notes_html = $docx->get_html;

    }

Those final two method calls return IO::All::File objects from files written to your
temp directory. I haven't tested this on Windows, but the module attempts to be
agnostic with regards to temporary directories.


=head1 METHODS

=over

=item new (constructor)

MSWord::ToHTML5 is a Moose class, so new is Moose's constructor.

=item validate_file

This gets you the only thing you need, which is an object ready to give you its HTML.

=item get_html

This is the other important method, that gives you an IO::All::File object with
your fresh HTML5 doc.

=back

=head1 AUTHOR

Amiri Barksdale, E<lt>amiri@arisdottle.netE<gt>

=head1 COPYRIGHT

Copyright (c) 2011 the MSWord::ToHTML5 L</AUTHOR> listed above.

=head1 LICENSE

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=head1 SEE ALSO

L<IO::All>

=cut
