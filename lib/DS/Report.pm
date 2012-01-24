package DS::Report;

use strict;
use warnings;
use Exporter;
use vars qw($VERSION @ISA @EXPORT @EXPORT_OK %EXPORT_TAGS);
use Carp;

$VERSION     = 0.1;
@ISA         = qw(Exporter);
@EXPORT      = qw(generate_excel generate_html generate_csv_output generate_text_table send_mail move_to_archive);
@EXPORT_OK   = qw(generate_excel generate_html generate_csv_output generate_text_table send_mail move_to_archive);
%EXPORT_TAGS = ( DEFAULT => \@EXPORT);

use Spreadsheet::WriteExcel;
use Spreadsheet::WriteExcel::Utility;
use DateTime;
use MIME::Lite;
use File::Copy;
use Text::ASCIITable;
use Data::Dump qw(dump);

use CGI qw/:standard :html3 -noDebug/;

use LWP;
use URI::Escape;

sub new {
    my $class = shift;
    my $self = {@_};
    bless($self,$class);
    $self->_validate;
    $self->_init;
    return $self;
}

sub _yyyymmddhhmmss() {
    my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
    # prints 2010-10-20_11-32-02
    # return sprintf("%04u-%02u-%02u_%02u-%02u-%02u", 1900 + $year, $mon + 1, $mday, $hour, $min, $sec);
    # prints 20101020_113202
    return sprintf("%04u%02u%02u_%02u%02u%02u", 1900 + $year, $mon + 1, $mday, $hour, $min, $sec);
}

sub _init {
    my $self = shift;

    $self->{mail_to} = join(",",@{$self->{mail_to}}) if ref($self->{mail_to}) eq "ARRAY";
    $self->{mail_cc} = join(",",@{$self->{mail_cc}}) if ref($self->{mail_cc}) eq "ARRAY";
    $self->{path} =~ s/\/$//; # remove trailing slash, if it is there
    $self->{callback} = \&callback unless defined $self->{callback};
    $self->{ignore_if_empty} = 0     unless defined $self->{ignore_if_empty};
    $self->{resultset_was_empty} = 0;

    $self->{pointintime} = _yyyymmddhhmmss();
    $self->{filetype} = 'csv' unless $self->{filetype};
}

sub callback {
    my $tmp = shift;
    # Do nothing
    return $tmp;
}

sub _validate {
    my $self = shift;

    croak "Path must be defined when calling new\nExample: path => 'my/path/to/dir/'\n"           unless defined($self->{path});
    croak "Filename must be defined when calling new\nExample: filename => 'filename.test'\n"     unless defined($self->{filename});
    croak "Datasource must be defined when calling new\nExample: datasource => \$sth\n"           unless defined($self->{datasource});

    croak "Path must exist when calling new\nExample: path => 'my/path/to/dir/'\n"                if (defined($self->{path}) && !(-d $self->{path}));
}

sub _placeInArray($$$){
# call with referance to array e.g.: inArray('element',1,\@array);
# returns array placement of supplied element
# second parameter decides whether to be case sensitive or note. A value higher than 0 equals case *IN*sensetive.
    my $element= shift;
    my $icase   = shift;
    my (@array) = @{$_[0]};
    my $ret     = undef;

    my $count = 0;
    foreach my $item (@array){
        if($item eq $element||($icase && lc $item eq lc $element)){
            $ret = $count;
	    last;
        }
        $count++;
    }
    return $ret;
}

sub _logstr(@){
    open LOG, ">>", $0.'.log' or warn $!;
    print LOG _yyyymmddhhmmss().' '.join(' ',@_)."\n";
    close LOG;
}

sub generate_excel{
    my $self = shift;

    # Add a worksheet
    my $workbook = Spreadsheet::WriteExcel->new($self->{path}."/".$self->{filename});
    my $sheet = $workbook->add_worksheet();

    # Add a Format
    my $format = $workbook->add_format();
    $format->set_bold();
    $format->set_size(12);
    #$format->set_color('blue');
    $format->set_align('center');

    $sheet->activate();

    if($self->{datasource}->isa("DBIx::ContextualFetch::st") || $self->{datasource}->isa("DBI::st")) {
        my $is_header = 1;
        # Save "resultset_was_empty" status for later use in send_mail
        $self->{resultset_was_empty} = 1 if $self->{datasource}->rows == 0;
        my $column_names = $self->{datasource}->{NAME};
        $column_names = $self->{callback}($column_names,$is_header);
        my $count = 0;
        foreach my $name (@$column_names){
                $sheet->set_column($count,$count, length($name) + 5);
                    $sheet->write(0, $count, $name, $format);
                    $count++;
            }
        my $row = 1;
        while(my $arrayref = $self->{datasource}->fetchrow_arrayref()){
                $arrayref = $self->{callback}($arrayref,!$is_header);
                    my $column = 0;
                foreach my $field (@$arrayref){
                    $sheet->write($row, $column, $field);
                    $column++;
                }
                    $row++;
            }
    }elsif(ref($self->{datasource}) eq "ARRAY") {
        # Feel free to implement ability to handle arrays
    }
    # Close workbook
    $workbook->close;
    return $self->{filename};
}

sub generate_html {
    my $self = shift;
    my @rows;
    my $column_names;

    my $newStyle=<<CSS;
<!-- 
    table caption {
        font-size: 14px;
        font-weight: bold;
    }
    table {
          /* width: 100%; */
          padding: 0px;
          border: none;
          border: 1px solid #789DB3;
      }
    table td {
        font-size: 12px;
	font-family: futura, helvetica, arial, sans-serif;
        border: none;
        background-color: #F4F4F4;
        vertical-align: middle;
        padding: 2px;
        /* font-weight: bold; */
    }
    table tr.special td {
        border-bottom: 1px solid #ff0000;
    }
-->
CSS

    open(OUTFILE,">$self->{path}/$self->{filename}");
    if($self->{datasource}->isa("DBIx::ContextualFetch::st") || $self->{datasource}->isa("DBI::st")) {
        my $is_header=1;
        # Save "resultset_was_empty" status for later use in send_mail
        $self->{resultset_was_empty} = 1 if $self->{datasource}->rows == 0;
        $column_names = $self->{datasource}->{NAME};
        $column_names = $self->{callback}($column_names,$is_header);
        while(my $arrayref = $self->{datasource}->fetchrow_arrayref()){
                no warnings 'uninitialized';
            $arrayref = $self->{callback}($arrayref,!$is_header);
                    my (@data) = @{$arrayref};
                    push(@rows,td(\@data));
            }
    }

    print OUTFILE start_html(-title => $self->{caption},
                             -head => meta({
                                            -http_equiv => 'Refresh',
                                            -content    => $self->{refresh},
                                            }),
                             -style => {
                                        -code => $newStyle
                                        }
                             );

    print OUTFILE table({-border=>'0'},
                        caption($self->{caption}),
                        TR([th(\@{$column_names}),@rows])
                        );
    print OUTFILE end_html;

    close(OUTFILE);
}

sub generate_csv {
    my $self = shift;

    my $separator = $self->{separator} || ";";

    open(OUTFILE,">$self->{path}/$self->{filename}");
    if($self->{datasource}->isa("DBIx::ContextualFetch::st") || $self->{datasource}->isa("DBI::st")) {
        print OUTFILE "$self->{caption}\n" if($self->{caption});
        my $is_header=1;

        # Save "resultset_was_empty" status for later use in send_mail
        $self->{resultset_was_empty} = 1 if $self->{datasource}->rows == 0;

        my $column_names = $self->{datasource}->{NAME};
        $column_names = $self->{callback}($column_names,$is_header);
        print OUTFILE join($separator, @{$column_names})."\n";
        while(my $arrayref = $self->{datasource}->fetchrow_arrayref()){
                no warnings 'uninitialized';
            $arrayref = $self->{callback}($arrayref,!$is_header);
                    print OUTFILE join($separator, @{$arrayref})."\n";
            }
    }

    close(OUTFILE);
}

sub http_post {
    my $self = shift;

    if($self->{ignore_if_empty} == 1 && $self->{resultset_was_empty} == 1) {
        print "Report empty - skipping send_mail\n";
        return;
    }

    if(!$self->{filename}){
	# LOG: something failed
	_logstr('Filename seems to have vanished.');
    } else {

	my $filestring;
	open FILE, "$self->{path}/$self->{filename}" or die "Couldn't open file: $!"; 
	while (<FILE>){
	    $filestring .= $_;
	}
	close FILE;

	# URI encoding
	$filestring = uri_escape($filestring) unless (defined $self->{nourlenc} && $self->{nourlenc} eq 1);

	print "$filestring\n";
	_logstr($filestring);
	
	my $ua = LWP::UserAgent->new;
        my $req = 0;
	
	if(uc($self->{httpmethod}) eq 'POST'){
	    $req = HTTP::Request->new( POST => $self->{httpurl});
	}elsif(uc($self->{httpmethod}) eq 'GET'){
	    $req = HTTP::Request->new( GET => $self->{httpurl});
	}
	$req->content_type('application/x-www-form-urlencoded');
	$req->content("data=$filestring");
	my $res = $ua->request($req);
	print $res->as_string;
    }
}

sub generate_csv_output {
    my $self = shift;

    my $separator    = $self->{separator} || ';';

    my $encapsulator = $self->{encapsulator} || '';
    my $nolineterm   = $self->{nolineterm};
    my $outputfile     .= $self->{filename};
    $outputfile     =~ s/\.csv$|\.txt$|\.exl$//;
    $outputfile     .= '_'.$self->{pointintime} if (defined $self->{filetimestamp} && $self->{filetimestamp} eq 1);
    $outputfile     .= '.'.$self->{filetype};

    $self->{filename} = $outputfile;

    $outputfile   = $self->{path}.'/'.$outputfile;
    open(OUTFILE,">$outputfile");

    if($self->{datasource}->isa("DBIx::ContextualFetch::st") || $self->{datasource}->isa("DBI::st")) {
        print "$self->{caption}\n" if($self->{caption});
        my $is_header=1;

        # Save "resultset_was_empty" status for later use in send_mail
        $self->{resultset_was_empty} = 1 if $self->{datasource}->rows == 0;

	my $column_names = $self->{datasource}->{NAME};
	$column_names = $self->{callback}($column_names,$is_header);

	# check wether or not to print column names
	unless(defined $self->{hidecolnames} && $self->{hidecolnames} eq 1){
	    print OUTFILE join($separator, @{$column_names})."\n";
	}

	# check which column-names are contained in ->dontencapsulate
	my @noencap;
	if($self->{dontencapsulate}){
	    my @donts = split(',',$self->{dontencapsulate});
	    foreach my $elem (@donts){
		my $place = _placeInArray($elem,1,\@$column_names);
		push(@noencap,$place);
	    }
	}

        while(my $arrayref = $self->{datasource}->fetchrow_arrayref()){
                no warnings 'uninitialized';
            $arrayref = $self->{callback}($arrayref,!$is_header);
		if(@noencap){
		    # print encapsulator unless requested not to
		    my $output = '';
		    my $elecount = 0;
		    foreach my $elem (@{$arrayref}){
			if(defined _placeInArray($elecount,1,\@noencap)){
			    $output .= $elem.$separator;
			} else {
			    $output .= $encapsulator.$elem.$encapsulator.$separator;
			}
			$elecount++;
		    }
		    $output =~ s/$separator$// if $nolineterm;
		    print OUTFILE $output."\n";
		} elsif($encapsulator){
		    # print ALL with encapsulator
		    my $line = $encapsulator.join($encapsulator.$separator.$encapsulator, @{$arrayref}).$encapsulator."\n";
		    $line =~ s/$separator$// if $nolineterm;
                    print OUTFILE $line;
		} else {
		    # print without encapsulator
		    my $line = join($separator, @{$arrayref})."\n";
		    $line =~ s/$separator$// if $nolineterm;
                    print OUTFILE $line;
		}
	}
    }
    close(OUTFILE);
}

sub generate_text_table {
    my $self = shift;

    my $t = Text::ASCIITable->new( { headingText => $self->{caption} });

    if($self->{datasource}->isa("DBIx::ContextualFetch::st") || $self->{datasource}->isa("DBI::st")) {
        my $is_header = 1;
  
        # Save "resultset_was_empty" status for later use in send_mail
        $self->{resultset_was_empty} = 1 if $self->{datasource}->rows == 0;
        
        my $column_names = $self->{datasource}->{NAME};
        $column_names = $self->{callback}($column_names,$is_header);
        while(my $arrayref = $self->{datasource}->fetchrow_arrayref()){
                $t->setCols(@{$column_names}) if $is_header == 1;
                    $is_header = 0;
                    $arrayref = $self->{callback}($arrayref,$is_header);
                    $t->addRow(@{$arrayref});
            }
    }
    print $t;
}    

sub send_mail {
    my $self = shift;

    # validate
    if(! defined $self->{mail_to}){
	_logstr("send_mail: No mail recipient defined.");
	print "send_mail: No mail recipient defined.\n";
	return;
    }

    if($self->{ignore_if_empty} == 1 && $self->{resultset_was_empty} == 1) {
        print "Report empty - skipping send_mail\n";
        _logstr("send_mail: Report empty - skipping send_mail");
        return;
    } else {
	my $tocc = $self->{mail_to};
	$tocc   .= ','.$self->{mail_cc} if defined $self->{mail_cc};
        _logstr("send_mail: Sending report to: ".$tocc);
    }

    my $msg = new MIME::Lite(From    => $self->{mail_from},
                             To      => $self->{mail_to},
                                  Cc      => $self->{mail_cc},
                             Subject => $self->{mail_subject},
                             Type    => 'multipart/mixed');
    $msg->attach(Type =>'TEXT',
                 Data => $self->{mail_text});

    $msg->attach(Type => 'BINARY',
                  Path => $self->{path}."/".$self->{filename},
                  Filename => $self->{filename},
                  Disposition => 'attachment',
                  Encoding    => 'base64');

    if($self->{mail_smtp_host}){
	$msg->send('smtp',$self->{mail_smtp_host}) || _logstr('Failed sending email via $self->{mail_smtp_host}');
    } else {
	$msg->send;
    }
}

sub move_to_archive {
    my $self = shift;
    move($self->{path}."/".$self->{filename},$self->{archive_path});
}

sub addColumn {
    my $self = shift;
    my ($arrayref,$content) = @_;
    # arrayref from DBI is readonly - therefore we must make a copy.
    my @local_copy = ();
    push @local_copy, @{$arrayref};
    push @local_copy, $content;
    return \@local_copy;
}

1;

__END__

=head1 NAME

    DS::Report - Perl module for generating reports in different formats and eventually sending them as mail.

=head1 SYNOPSIS

    use DS::Report;

my $report = DS::Report->new(
               path            => '/tmp/',
               filename        => 'test.xls',
               separator       => '|',
               encapsulator    => '"',
               dontencapsulate => 'column1,column3',
               hidecolnames    => 1,
               httpurl         => 'http://host.domain.tld/getpost.cgi',
               datasource      => $datasource,
               mail_from       => 'ds@domain.tld',
               mail_to         => @recipients,
               mail_cc         => 'ds@domain.tld',
               mail_subject    => 'The subject',
               mail_text       => 'The mail content',
               archive_path    => '/tmp/archive/',
               ignore_if_empty => 1,
               );
  $reportTool->generate_excel;
  $reportTool->send_mail;
  $reportTool->move_to_archive;

  #$reportTool->generate_text_table; # For debugging (prints to terminal)

=head1 DESCRIPTION

  The module generates reports, bases on input delivered by "new".

=head2 Methods

=over 4

=item * $object->generate_excel()

Generates an excel file in the "path"/"filename" location. 
The file is generated, based on "datasource" input.

=item * $object->generate_csv()

Generates a csv file in the "path"/"filename" location. 
The file is generated, based on "datasource" input and is separated by "separator". (defaults to ;)

=item * $object->generate_csv_output()

Generates csv output to STDOUT. Quite similar to generate_http() - just without the HTTP request part.
The content is generated, based on "datasource" input and is separated by "separator". (defaults to ;)

=item * $object->generate_http()

Generates CSV-like output and sends as HTTP request location. Useful for posting data directly to a webserver.

=item * $object->generate_html()

Generates a HTML file in the "path"/"filename" location. Useful for showing data directly on a webserver. 

=item * $object->generate_text_table()

Generates a table which is printed to the console. This is practical for debugging usage.

=item * $object->send_mail()

Sends email according to information passed to "new". It also attaches the file generated, if any.

=item * $object->move_to_archive()

Moves the generated file from "path" to "archive_path".

=item * The callback method

A function address can be passed to "new" via the "callback" parameter. The function will be called for every iteration of the resultset inside the class. The function receives an arrayref and a boolean. The arrayref contains the current row in the resultset. The callback method must return this arrayref. (modified or not) The boolean is a "is_header" value, that tells us if the current row is a header line or not. This line is used to decide whether to print the header or data.
This callback makes it posssible to add values/columns to whatever is printed to the final output (csv/xls/...) See callback example in the EXAMPLES section.

=item * $object->addColumn($arrayref,'ColumnText')

This is used inside the callback method to set the column content (header or content).
See callback example in the EXAMPLES section.

=back

=head1 NOTES

=head2 When using CRON.

    It's often easiest to create a shellscript that exports the necessary environment variables, changes to the folder that contains the DS folder (containing the Report.pm module), and executes the Perl script.
    Remember to export paths within you shellscript - check environment variables for correct paths.
    Also remember to point to the correct Perl binary on the first line of the Perl script, if using non-standard paths. E.G. #!/pack/perl/bin/perl

=head2 Shellscript example

    #!/bin/sh
    export PERL5LIB=/path/to/perl/lib
    export ORACLE_HOME=/path/to/oracle/product/9ir2

    cd ~/prod/cron/
    ./sqlreport.pl

=head1 EXAMPLES

=head2 Using standalone file to generate csv. (separated by |)

  use strict;
  use warnings;

  use DBI;
  use DS::Report;
  # Prepare and execute statement, and pass handle to Report->new
  my $dbh  = DBI->connect("dbi:Pg:dbname=preprod;host=databasehost",'username','password') or die DBI->errstr;
  my $sth = $dbh->prepare(qq[select id,status,message from orderflow limit 10 ]);
  $sth->execute();
  my $report = DS::Report->new(
                separator       => '|',
                path            => '/home/nas/data/testreport/',
                filename        => 'test.csv',
                datasource      => $sth,
                mail_from       => 'nas@domain.net',
                mail_to         => 'pm@domain.com',
                mail_cc         => '',
                mail_subject    => 'Testing the report',
                mail_text       => 'This is the test report from me',
                archive_path    => '/home/nas/data/archive/',
                );
  $reportTool->generate_csv;
  $reportTool->send_mail;
  $reportTool->move_to_archive;

=head2 Displaying data on a webserver.

  use strict;
  use warnings;
  use DBI;
  use DS::Report;

  # Prepare and execute statement, and pass handle to Report->new
  my $dbh = DBI->connect('dbi:Oracle:iaadan.world', 'username', 'password') or die DBI->errstr;

  my $sth = $dbh->prepare(qq[
                select count(1) as antal,ct.STATUS as STATUS,
                case when cd.PROCESSEDERROR is null then 'OK'
                when cd.PROCESSEDERROR is not null then 'Error'
                end as Error,
                ct.ID as Type, ct.VARIANT as Variant, ct.TITLE as Title
                from coupondata cd, COUPONTYPE ct
                where ct.ID = cd.COUPONTYPEID
                and (cd.processed !='Y' or cd.PROCESSEDERROR is not null)
                group by ct.status,cd.PROCESSEDERROR,ct.ID, ct.VARIANT,ct.TITLE,ct.TEST
    ]);

  $sth->execute();

  my $report = DS::Report->new(
                separator       => '|',
                path            => '/server/test/httpd/htdocs/CouponStatus/',
                filename        => 'index.html',
                datasource      => $sth,
                archive_path    => '/tmp/',
                caption         => 'Unprocessed coupons/coupons in error. - '.`date`,
                refresh         => '1800',
      );
  $report->generate_html;

=head2 Using the callback method

  use strict;
  use warnings;

  use DBI;
  use DS::Report;

  my $dbh  = DBI->connect("dbi:Pg:dbname=preprod;host=host.net",'user','pass') or die DBI->errstr;
  my $sth = $dbh->prepare(qq[SELECT salesman_id,order_id,coupon_serial FROM pb_express_scans pe
                             join orderflow o on o.id = pe.order_id
                             join flow f on f.id = o.starting_flow_id
                           where coupon_variant = '282_12'
                           and f.name = 'MobilePostpaidCouponNewCustomer'
                           order by order_id desc limit 10;
                           ]);
  $sth->execute;

  # Prepare Oracle statement for usage in callback sub
  my $dbh_ora = DBI->connect('dbi:Oracle:iaadan.world', 'coupon', 'coupon');
  my $sth_ora = $dbh_ora->prepare(qq[SELECT PROCESSEDDATE,TIFFLINK FROM COUPONDATA where ORDERFLOWID = ? ]);

  my @recipients = [ 'ds@domain.tld', 'ds@domain.tld' ];
  my $reportTool = DS::Report->new(
                                    path            => '/home/ds/work/tmp/',
                                    filename        => 'test.xls',
                                    separator       => "|",
                                    datasource      => $sth,
                                    datasource_type => 'dbi',
                                    mail_from       => 'ds@domain.tld',
                                    mail_to         => @recipients,
                                    mail_cc         => 'ds@domain.tld',
                                    mail_subject    => 'Testing the report',
                                    mail_text       => 'This is the testmail text from me',
                                    archive_path    => '/home/ds/work/tmp/archive/',
                                    callback        => \&callback_test,
                                    );
  #$reportTool->generate_excel;
  #$reportTool->generate_csv;
  #$reportTool->send_mail;
  #$reportTool->move_to_archive;
  $reportTool->generate_text_table; # For debugging. (printing to terminal)

  sub callback_test {
    my $arrayref = shift;
    my $is_header =shift;

    # Call addColumn with header string, if is_header is true
    #########################################################
    if($is_header) {
        $arrayref = $reportTool->addColumn($arrayref,'Tiff link');
        $arrayref = $reportTool->addColumn($arrayref,'Processed date');
    }else{
        # If is_header is false, then call addColumn with the data instead.
        ##################################################################
        # Fetch data by order_id
        $sth_ora->execute($arrayref->[1]);
        my ($processeddate,$tifflink) = $dbh_ora->selectrow_array($sth_ora);
        # Add column data
        $arrayref = $reportTool->addColumn($arrayref,$tifflink);
        $arrayref = $reportTool->addColumn($arrayref,$processeddate);
    }
    return $arrayref;
  }




=head1 AUTHOR

David K. Sundsskard (david@sundsskard.com)
Roi a Torkilsheyggi
=cut
