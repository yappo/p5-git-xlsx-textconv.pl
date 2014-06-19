#!/usr/bin/env perl
use strict;
use warnings;
use Spreadsheet::XLSX;

my $xlsx_file = shift;
die "usage: $0 file.xlsx" unless $xlsx_file && $xlsx_file =~ /\.xlsx\z/;

my $xlsx = Spreadsheet::XLSX->new($xlsx_file);
for my $sheet (@{ $xlsx->{Worksheet} }) {
    sheet_walk($sheet);
}

sub sheet_walk {
    my $sheet = shift;
    my $sheet_name = $sheet->{Name};
    $sheet_name =~ s/([\[\]])/\\$1/g;

    for my $row ($sheet->{MinRow}..($sheet->{MaxRow} || $sheet->{MinRow})) {
        my @cells;
        if (defined $sheet->{Cells}[$row] && ref($sheet->{Cells}[$row]) eq 'ARRAY') {
            for my $cell (@{ $sheet->{Cells}[$row] }) {
            	if ($cell && defined $cell->{Val}) {
                    my $val = $cell->{Val};
                    $val =~ s/\\/\\\\/g;
                    $val =~ s/\r/\\r/g;
                    $val =~ s/\n/\\n/g;
                    $val =~ s/\t/\\t/g;
                    push @cells, $val;
            	} else {
                    push @cells, '';
            	}
            }
        }
        printf "[%s]: %s\n", $sheet_name, join("\t", @cells);
    }
}

__END__

=head1 name

git-xlsx-textconv.pl - git text converter for xlsx file

=head1 README

show README.md

=head1 LICENSE

Copyright (C) Kazuhiro Osawa 2014-

This library is free software; you can redistribute it and/or modify
it under the same terms as Perl itself.

=cut
