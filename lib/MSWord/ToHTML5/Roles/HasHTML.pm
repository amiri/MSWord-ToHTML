package MSWord::ToHTML5::Roles::HasHTML;

use Moose::Role;
use namespace::autoclean;
use feature 'say';
use MooseX::Method::Signatures;

method get_html {
    say $self;
    return @_;
}

1;
