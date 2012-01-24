use inc::Module::Install;

perl_version '5.006';
license 'gpl';

name     'DS-Report';
all_from 'lib/DS/Report.pm';

include 'Module::AutoInstall';

requires 'DateTime'                         => 0;
requires 'MIME::Lite'                       => 0;
requires 'File::Copy'                       => 0;

feature 'Debug',
    -default => 1,
    'Data::Dump' => 0;

feature 'ASCII',
    -default => 1,
    'Text::ASCIITable' => 0;

feature 'Excel',
    -default => 1,
    'Spreadsheet::WriteExcel' => 0,
    'Spreadsheet::WriteExcel::Utility' => 0;

auto_install(
    force       => 1,           # pseudo-option to force install
);

WriteAll;
