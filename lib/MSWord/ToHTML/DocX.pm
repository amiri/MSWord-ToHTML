package MSWord::ToHTML::DocX;

use Moose;
use namespace::autoclean;
use MSWord::ToHTML::Types::Library qw/:all/;
with 'MSWord::ToHTML::Roles::HasHTML';

has "file" => (
    is       => 'ro',
    isa      => MSDocX,
    required => 1,
);

__PACKAGE__->meta->make_immutable;

1;
