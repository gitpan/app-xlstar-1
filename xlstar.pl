#! /usr/bin/perl -w
#! /usr/bin/env perl -w

our($emacs_Time_stamp) = 'Time-stamp: <2010-06-14 18:06:53 johayek>' =~ m/<(.*)>/;

our     $rcs_Id=(join(' ',((split(/\s/,'$Id: xlstar.pl 1.36 2010/06/14 16:06:57 johayek Exp $'))[1..6])));
our   $rcs_Date=(join(' ',((split(/\s/,'$Date: 2010/06/14 16:06:57 $'))[1..2])));
our $rcs_Author=(join(' ',((split(/\s/,'$Author: johayek $'))[1])));
our    $RCSfile=(join(' ',((split(/\s/,'$RCSfile: xlstar.pl $'))[1])));
our $rcs_Source=(join(' ',((split(/\s/,'$Source: /media/_ARCHIVE/home/jochen_hayek-FROZEN-STUFF/usr/src/IDS_cronus_projects/200602--utility_xlstar/RCS/xlstar.pl $'))[1])));

use warnings;
use strict;

our $VERSION = '1.36';

{
##use English;
  use FileHandle;
  use strict;

  use Spreadsheet::ParseExcel;
  # -> ~johayek/usr/src/IDS_cronus_projects/___delivery_notes--CPAN.txt
  # -> ~johayek/Computers/Programming/Languages/Perl/README-CPAN
  
  &main;
}

sub main
{
  my($package,$filename,$line,$proc_name) = caller(0);

  my(%param) = @_;

  my($return_value) = 0;

  # described in:
  #	camel book / ch. 7: the std. perl lib. / lib. modules / Getopt::Long - ...

  use Getopt::Long;
  use Pod::Usage;
  %main::options = ();

  $main::options{debug} = 0;

  printf STDERR ">%s,%d,%s\n",__FILE__,__LINE__,$proc_name
    if 0 && $main::options{debug};
  printf STDERR "=%s,%d,%s: %s=>{%s}\n",__FILE__,__LINE__,$proc_name
    ,'$rcs_Id' => $rcs_Id
    if 0 && $main::options{debug};
  printf STDERR "=%s,%d,%s: %s=>{%s}\n",__FILE__,__LINE__,$proc_name
    ,'$emacs_Time_stamp' => $emacs_Time_stamp
    if 0 && $main::options{debug};

  {
    # defaults for the main::options;
    
    $main::options{dry_run}		       	= 0;
    $main::options{version}		       	= 0;
    $main::options{verbose}		       	= 0;

    $main::options{job_list} = 0;
    $main::options{job_extract} = 0;
  }

  my($result) =
    &GetOptions
      (\%main::options
     ##,'job_download_statement|jd!'
       ,'job_list|list!'
       ,'job_extract|extract|job_get|get!'

       ,'dry_run!'
       ,'version!'
       ,'help|?!'
       ,'man!'
       ,'debug!'
     ##,'verbose=s'		# sometimes we use it like this
       ,'verbose!'		# sometimes we use it like this 

       ,'file=s@'
       ,'to_stdout|to-stdout!'
       ,'multi_volume|multi-volume!'
       );
  $result || pod2usage(2);

  pod2usage(1) if $main::options{help};
  pod2usage(-exitstatus => 0, -verbose => 2) if $main::options{man};

  if   ($main::options{job_list}   ) { &App::XLSTar::job_list; }
  elsif($main::options{job_extract}) { &App::XLSTar::job_extract; }
  else
    {
      die "no job to be carried out";
    }

  printf STDERR "=%s,%d,%s: %s=>{%s}\n",__FILE__,__LINE__,$proc_name
    ,'$return_value' => $return_value
    if 0 && $main::options{debug};
  printf STDERR "<%s,%d,%s\n",__FILE__,__LINE__,$proc_name
    if 0 && $main::options{debug};
}
#
package App::XLSTar;

use warnings;
use strict;

sub job_list
{
  my($package,$filename,$line,$proc_name) = caller(0);

  my(%param) = @_;

  my($return_value) = 0;

  printf STDERR ">%s,%d,%s\n",__FILE__,__LINE__,$proc_name
    if 1 && $main::options{debug};

  if($#{$main::options{file}} < 0)
    {
      printf STDERR "=%s,%d,%s: %s=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
	,"\$#{$main::options{file}}",$#{$main::options{file}}
        ,'...'
	if 1 && $main::options{debug};

      die 'not even a single --file=... supplied';
    }
  else
    {
      foreach my $f (@{$main::options{file}})
	{
	  my $o_book = Spreadsheet::ParseExcel::Workbook->Parse($f);

	  for( my $i = 0 ; $i <= $#{ $o_book->{Worksheet} } ; $i++ )
	    {
	      printf STDERR "=%s,%d,%s: %s=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
		,"\$o_book->{Worksheet}[$i]{Name}",$o_book->{Worksheet}[$i]{Name}
		,'...'
		if 1 && $main::options{debug};

	      printf "%s\n"
		,$o_book->{Worksheet}[$i]{Name}
		,'...'
		;
	    }
	}
    }

  printf STDERR "=%s,%d,%s: %s=>{%s}\n",__FILE__,__LINE__,$proc_name
    ,'$return_value' => $return_value
    if 0 && $main::options{debug};
  printf STDERR "<%s,%d,%s\n",__FILE__,__LINE__,$proc_name
    if 1 && $main::options{debug};

  return $return_value;
}
#
sub job_extract
{
  my($package,$filename,$line,$proc_name) = caller(0);

  my(%param) = @_;

  my($return_value) = 0;

  printf STDERR ">%s,%d,%s\n",__FILE__,__LINE__,$proc_name
    if 1 && $main::options{debug};

  binmode(STDOUT, ":utf8");	# -> perluniintro -- w/o this we get a "Wide character in print at " from the printf below

  my(%files_to_extract);

  for( my $i = 0 ; $i <= $#{ ARGV } ; $i++ )
    {
      printf STDERR "=%s,%d,%s: %s=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
	,"\$ARGV[$i]",$ARGV[$i]
	,'...'
	if 1 && $main::options{debug};

      $files_to_extract{$ARGV[$i]} = 1;
    }

  my($files_to_extract__is_empty_p) = length(keys %files_to_extract) != -1;

  printf STDERR "=%s,%d,%s: %s=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
    ,"\$files_to_extract__is_empty_p",$files_to_extract__is_empty_p
    ,'...'
    if 1 && $main::options{debug};

  defined($main::options{to_stdout}) 	    || die 'missing: --to_stdout';

  if($#{$main::options{file}} < 0)
    {
      printf STDERR "=%s,%d,%s: %s=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
	,"\$#{$main::options{file}}",$#{$main::options{file}}
        ,'...'
	if 1 && $main::options{debug};

      die 'not even a single --file=... supplied';
    }
  else
    {
      foreach my $f (@{$main::options{file}})
	{
	  my $o_book = Spreadsheet::ParseExcel::Workbook->Parse($f);

	  for( my $i = 0 ; $i <= $#{ $o_book->{Worksheet} } ; $i++ )
	    {
	      printf STDERR "=%s,%d,%s: %s=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
		,"\$o_book->{Worksheet}[$i]{Name}",$o_book->{Worksheet}[$i]{Name}
		,'...'
		if 1 && $main::options{debug};

	      if(   $files_to_extract__is_empty_p
		 || exists($files_to_extract{ $o_book->{Worksheet}[$i]{Name} })
		)
		{
		  printf STDERR "%s\n"
		    ,$o_book->{Worksheet}[$i]{Name}
		    if 1 && $main::options{verbose};

		  ################################################################################
		  ################################################################################
		  ################################################################################
		  # this is, where this "Worksheet" will get extracted.
		  ################################################################################
		  ################################################################################
		  ################################################################################

		  for(  my $i_row = $o_book->{Worksheet}[$i]->{MinRow}
		     ;      defined($o_book->{Worksheet}[$i]->{MaxRow})
		       && $i_row <= $o_book->{Worksheet}[$i]->{MaxRow}
		     ;    $i_row++
		     )
		    {
		      my($separator) = '';

		      for(  my $i_column = $o_book->{Worksheet}[$i]->{MinCol}
			 ;         defined($o_book->{Worksheet}[$i]->{MaxCol})
			   && $i_column <= $o_book->{Worksheet}[$i]->{MaxCol}
			 ;    $i_column++
			 )
			{
			  my($o_work_cell) = $o_book->{Worksheet}[$i]->{Cells}[$i_row][$i_column];

			  printf STDERR "=%s,%d,%s: (%s,%s)=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
			    ,$i_row,$i_column
			    , defined( $o_work_cell )
			    ,'? defined( $o_work_cell ) ?'
			    if 1 && $main::options{debug};

			  printf STDERR "=%s,%d,%s: (%s,%s)=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
			    ,$i_row,$i_column
			    , defined( $o_work_cell->Value )
			    ,'? defined( $o_work_cell->Value ) ?'
			    if 1 && $main::options{debug} && defined( $o_work_cell );

			  printf STDERR "=%s,%d,%s: (%s,%s)=>{%s} // %s\n",__FILE__,__LINE__,$proc_name
			    ,$i_row,$i_column
			    , $o_work_cell->Value
			    ,'...'
			    if 1 && $main::options{debug} && defined( $o_work_cell ) && defined( $o_work_cell->Value );

			  my($v);
			  if( defined( $o_work_cell ) && defined( $o_work_cell->Value ) && ($o_work_cell->Value ne '') )
			    {
			      $v = $o_work_cell->Value;

			      $v =~ tr/\"/./; # this is a work-around, and probably far too simple -- TBD!!!

			      $v = '"' . $v . '"';
			    }
			  else
			    {
			      $v = '';
			    }

			  # w/o the "binmode(...)" above we get a "Wide character in print at " from this location:

			  printf "%s%s"
			    , $separator
			    , $v
			    ;

			  $separator = ',';
			}
		      printf "\n";
		    }

		}
	    }
	}
    }

  printf STDERR "=%s,%d,%s: %s=>{%s}\n",__FILE__,__LINE__,$proc_name
    ,'$return_value' => $return_value
    if 0 && $main::options{debug};
  printf STDERR "<%s,%d,%s\n",__FILE__,__LINE__,$proc_name
    if 1 && $main::options{debug};

  return $return_value;
}
__END__

=head1 NAME

xlstar.pl

=head1 SYNOPSIS

xlstar.pl [options] [file ...]

Options:
    --help
    --man

    --list
    --extract

    --file=...
    --to-stdout

    --...

=head1 OPTIONS

=over 8

=item B<--help>

Print a brief help message and exits.

=item B<--man>

Prints the manual page and exits.

=item B<--list>

list the contents of an archive

=item B<--extract>

extract files from an archive

=item B<--file=ARCHIVE>

use archive file ARCHIVE

(option is enforced, so that user is aware of semantics similar to those of the F<tar> utility)

=item B<--to-stdout>

extract files to standard output

(option is enforced, so that user is aware of semantics similar to those of the F<tar> utility)

=item B<--...>

...

=back

=head1 DESCRIPTION

This program tries to handle XLS files a little like I<tar> treats its archives.

It lists the names of the worksheets of an XLS file,
and it extracts worksheets one by one (i.e. per call of this program) to standard output.

=for comment section README gets extracted into a separate README file.

=head1 README

This program tries to handle XLS files a little like I<tar> treats its archives.

It lists the names of the worksheets of an XLS file,
and it extracts worksheets one by one (i.e. per call of this program) to standard output.

I personally keep this file as f<xlstar.pl>.

=head1 REQUIREMENTS

...

=head1 EXAMPLE

.../xlstar.pl --file=G300_20060130.xls --list

.../xlstar.pl --file=G300_20060130.xls --extract --to_stdout 'G300 Index Members' > G300.G300_Index_Members.20060130.csv

=head1 HISTORY

Q: Why is this script based on I<Spreadsheet::ParseExcel> ?

A: Simply because that module seems to be the mother of all XLS file parsers on CPAN, all the others depend on that one and make use of it.

Q: xls2csv.pl already existed, when this script got created in 2006, so why was it necessary to create something else?

A: xls2csv.pl's parameterisation was far too simplistic, so it did not fit into the planned usage style.

...

=head1 KNOWN PROBLEMS

=head2 UTF-8

We expect the XLS file to contain strings of UTF-8 code,
so the extracted CSV files may look weird, if not viewed as UTF-8 files.

=head2 pre-calculated values of formulas

We cannot resp. do not calculate the values of formulas,
and we entirely depend on printable values supplied for such cells.
We inherit that deficiency from I<Spreadsheet::ParseExcel>,
and there we can find a short descriptive note (slightly corrected by JH):

This module can not get the values of formulas in Excel files made with I<Spreadsheet::WriteExcel>.
Usually a formula has the result with it.
But e.g. I<Spreadsheet::WriteExcel> writes a formula with no result supplied.
If you set your Excel application I<Auto Calculation> off
(maybe [I<Tool>]-[I<Option>]-[I<Calculation>] or something),
you will see the same result from a file created by Excel itself.

=for comment "How to submit a script to CPAN" <http://www.cpan.org/scripts/submitting.html>

=head1 PREREQUISITES

Spreadsheet::ParseExcel

=for comment should actually be "=pod OSNAMES" according to f<Sample-0.1>, but that one gets simply ignored as a header, but not its body

=head1 OSNAMES

any

=head1 SCRIPT CATEGORIES

Win32
Win32/Utilities

=head1 AUTHOR

Jochen Hayek E<lt>Jochen+CPAN@Hayek.nameE<gt>

=head1 HISTORY

=over 8

=item B<xlstar_1_36.pl>

first CPAN upload

=cut
