package MSWord::ToHTML5::Doc;

use Moose;
use namespace::autoclean;
use MSWord::ToHTML5::Types::Library qw/:all/;
with 'MSWord::ToHTML5::Roles::HasHTML';

has "file" => (
    is       => 'ro',
    isa      => MSDoc,
    required => 1,
);

__PACKAGE__->meta->make_immutable;

1;
