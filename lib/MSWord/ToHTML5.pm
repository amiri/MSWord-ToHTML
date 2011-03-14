package MSWord::ToHTML5;

use Moose;
use feature "say";
use MooseX::Method::Signatures;
use IO::All;
use IO::All::File;
use Carp;
use Archive::Zip qw/:ERROR_CODES :CONSTANTS/;
use Archive::Zip::MemberRead;
use autodie;
use File::Spec;
use namespace::autoclean;
use Path::Class::Dir;
use File::Path qw/make_path/;
use File::Basename;
use XML::LibXML;
use XML::LibXSLT;

our $VERSION = '0.001';

# ABSTRACT: Take old or new Word files and spit out HTML5

has 'parser' => (
  is => 'ro',
  isa => 'XML::LibXML',
  lazy => 1,
  default => sub {
    XML::LibXML->new
  },
);
has 'transformer' => (
  is => 'ro',
  isa => 'XML::LibXSLT',
  lazy => 1,
  default => sub {
    XML::LibXSLT->new
  },
);

has 'style' => (
  is => 'ro',
  lazy => 1,
  default => sub {
    my $self = shift;
    $self->parser->load_xml(location=>'etc/docx2html-MS.xsl', no_cdata=>1)
  },
);

method validate_file($file) {
    return
      unless $file;

    my $dir_name = basename($file);
    say "My dir_name is: $dir_name";

    open( STDOUT, '>', File::Spec->devnull );
    my $path = Path::Class::Dir->new( 'tmp', $dir_name );
    $path->mkpath;

    system( 'antiword', $file );
    if ( ${^CHILD_ERROR_NATIVE} == 0 ) {
      system("antiword -m 8859-1 -x db $file > $path/document.xml");
      return io("$path/document.xml");
    }

    my $unzip = Archive::Zip->new;
    if ( $unzip->read($file) == AZ_OK ) {
      say "I will extract $file to $path";
      $unzip->extractTree( "", "$path/" );
      return io("$path/word/document.xml");
    }
    else {
      say "I could not read $file";
    }

    confess "I don't know what to do with this file";
}

method get_dom ($file) {
  return $self->parser->parse_file($file);
}

method convert_to_html ($dom) {
  my $stylesheet = $self->transformer->parse_stylesheet($self->style);
  my $results = $stylesheet->transform($dom);
  return $stylesheet->output_as_bytes($results);
}

__PACKAGE__->meta->make_immutable;

1;
