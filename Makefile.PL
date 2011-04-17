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

requires 'Archive::Zip' => '1.30';
requires 'Archive::Zip::MemberRead' => '1.3';
requires 'CSS' => '1.09';
requires 'Carp' => '1.17';
requires 'Digest::SHA1' => '2.13';
requires 'Encode' => '2.42';
requires 'Encode::Guess' => '2.04';
requires 'File::Basename' => '2.78';
requires 'File::Path' => '2.08_01';
requires 'File::Spec' => '3.33';
requires 'HTML::Entities' => '3.68';
requires 'HTML::HTML5::Writer' => '0.102';
requires 'HTML::TreeBuilder' => '4.2';
requires 'IO::All' => '0.41';
requires 'IO::All::File' => '0';
requires 'List::MoreUtils' => '0.30';
requires 'Moose' => '2.0000';
requires 'Moose::Role' => '2.0000';
requires 'Moose::Util::TypeConstraints' => '2.0000';
requires 'MooseX::Method::Signatures' => '0.36';
requires 'MooseX::Types' => '0.25';
requires 'MooseX::Types::IO::All' => '0.03';
requires 'MooseX::Types::Moose' => '0.25';
requires 'Path::Class::Dir' => '0.23';
requires 'Text::Extract::Word' => '0.02';
requires 'Try::Tiny' => '0.09';
requires 'XML::LibXML' => '1.70';
requires 'XML::LibXSLT' => '1.70';
requires 'autodie' => '2.10';
requires 'namespace::autoclean' => '0.12';
requires 'strictures' => '0';

test_requires 'Test::Most' => '0.23';

auto_install;
WriteAll;
