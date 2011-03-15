package MSWord::ToHTML5;

use Moose;
use feature "say";
use MooseX::Method::Signatures;
use IO::All;
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
use Digest::SHA1 qw/sha1_hex/;
use HTML::HTML5::Writer;

# ABSTRACT: Take old or new Word files and spit out HTML5

has 'parser' => (
  is      => 'ro',
  isa     => 'XML::LibXML',
  lazy    => 1,
  default => sub {
    XML::LibXML->new;
  },
);
has 'transformer' => (
  is      => 'ro',
  isa     => 'XML::LibXSLT',
  lazy    => 1,
  default => sub {
    XML::LibXSLT->new;
  },
);

has 'doc_style' => (
  is      => 'ro',
  lazy    => 1,
  default => sub {
    my $self = shift;
    $self->parser->load_xml(
      location => 'etc/docbook-xsl/xhtml-1_1/docbook.xsl',
      no_cdata => 1
    );
  },
);
has 'docx_style' => (
  is      => 'ro',
  lazy    => 1,
  default => sub {
    my $self = shift;
    $self->parser->load_xml(
      location => 'etc/docx2html-MS.xsl',
      no_cdata => 1
    );
  },
);

has 'writer' => (
  is      => 'ro',
  isa     => 'HTML::HTML5::Writer',
  lazy    => 1,
  default => sub {
    HTML::HTML5::Writer->new( markup => 'html' );
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

method convert_doc_to_html ($dom) {
  my $stylesheet = $self->transformer->parse_stylesheet( $self->doc_style );
  my $results    = $stylesheet->transform($dom);
  my $text       = $stylesheet->output_as_bytes($results);
  my $filename   = sha1_hex($text);
  $text > io("tmp/$filename");
  `tidy -f tmp/$filename.err -m --clean 1 --drop-empty-paras 1 --drop-font-tags 1 --char-encoding utf8 "tmp/$filename"`;
  return $self->parser->parse_html_fh( io("tmp/$filename") )
}

method convert_docx_to_html ($dom) {
  my $stylesheet = $self->transformer->parse_stylesheet( $self->docx_style );
  my $results    = $stylesheet->transform($dom);
  my $text       = $stylesheet->output_as_bytes($results);
  my $filename   = sha1_hex($text);
  $text > io("tmp/$filename");
  `tidy -f tmp/$filename.err -m --clean --word-2000 1 --drop-empty-paras 1 --drop-font-tags 1 --char-encoding utf8 "tmp/$filename"`;
  return $self->parser->parse_html_fh( io("tmp/$filename") )
}

method convert_to_html5 ($html) {
  return $self->writer->document($html)
}

__PACKAGE__->meta->make_immutable;

1;
