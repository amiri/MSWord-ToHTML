package MSWord::ToHTML5::DocX;

use Moose;
use namespace::autoclean;
use MSWord::ToHTML5::Types::Library qw/:all/;
with 'MSWord::ToHTML5::Roles::HasHTML';

has "file" => (
    is       => 'ro',
    isa      => MSDocX,
    required => 1,
);

__PACKAGE__->meta->make_immutable;

1;
