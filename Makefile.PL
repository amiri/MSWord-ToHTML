#!/usr/bin/env perl
use inc::Module::Install;

auto_set_repository;
name 'MSWord-ToHTML5';
all_from 'lib/MSWord/ToHTML5.pm';
readme_from 'lib/MSWord/ToHTML5.pm';
abstract 'Take old or new Word files and spit out HTML5';
author 'Amiri Barksdale <amiri@arisdottle.net>';
license 'perl';
perl_version '5.010001';
version '0.001';

requires_external_bin 'abiword';
requires_external_bin 'tidy';

requires 'Archive::Zip' => '0';
requires 'Archive::Zip::MemberRead' => '0';
requires 'CSS' => '0';
requires 'Carp' => '0';
requires 'Devel::Dwarn' => '0';
requires 'Digest::SHA1' => '0';
requires 'Encode' => '0';
requires 'Encode::Guess' => '0';
requires 'File::Basename' => '0';
requires 'File::Path' => '0';
requires 'File::Spec' => '0';
requires 'HTML::Entities' => '0';
requires 'HTML::HTML5::Writer' => '0';
requires 'HTML::TreeBuilder' => '0';
requires 'IO::All' => '0';
requires 'IO::All::File' => '0';
requires 'List::MoreUtils' => '0';
requires 'Moose' => '0';
requires 'Moose::Role' => '0';
requires 'Moose::Util::TypeConstraints' => '0';
requires 'MooseX::Method::Signatures' => '0';
requires 'MooseX::Types' => '0';
requires 'MooseX::Types::IO::All' => '0';
requires 'MooseX::Types::Moose' => '0';
requires 'Path::Class::Dir' => '0';
requires 'Test::Most' => '0';
requires 'Text::Extract::Word' => '0';
requires 'Try::Tiny' => '0';
requires 'XML::LibXML' => '0';
requires 'XML::LibXSLT' => '0';
requires 'autodie' => '0';
requires 'namespace::autoclean' => '0';
requires 'strictures' => '0';

test_requires 'Test::Most' => '0';

auto_install;
WriteAll;
