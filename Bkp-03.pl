#!c:/Perl/perl.exe -w

use strict;
use warnings;

my $Lista = 'Arquivos.txt';
my $Caminho = 'Caminhos.txt';

BKP($Lista,$Caminho);

sub BKP{
	use File::Copy;
	open(my $file, q{<}, $_[0]);
	foreach my $arquivo ( <$file> ) {
	   #Remove o último caractere apenas se for igual a $/ "Separador de registro" 
	   chomp($arquivo);
	   # Verifica se o arquivo existe
	   if (-e $arquivo){
		open(my $caminho1, q{<}, $_[1]);
		foreach my $caminho2 ( <$caminho1> ) {
			chomp($caminho2);
			my $caminho3 = $caminho2."/".DataHora();
			mkdir $caminho3;
			copy($arquivo,$caminho3."/".$arquivo) or die "Copy failed: $!";
		}
	   }
	}
	}

# Criar layout de data
sub DataHora{
	my($dd,$mm,$yy,$day,$hh,$nn) = (localtime)[3,4,5,6,2,1];
	my $today =  join '', map sprintf("%02d", $_),($yy%100,$mm+1,$dd,);
	my $hr = join '', map sprintf("%02d", $_),($hh,$nn);

	# "Vasio" = Data e hora | "1" Apenas Data
	my $data = $_[0] ? $today:$today.'_'.$hr;
	my $datahora = $data;
	
	return ($datahora);
	}

sub CompactarArquivos{
# ARQVS, NomeDoPacote
	# Create a Zip file
	use Getopt::Std;
	use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
	my $zip = Archive::Zip->new();

	open(my $file, q{<}, $_[0]) or die "Can't open file $!\n";
	foreach my $arquivo ( <$file> ) {
	   #Remove o último caractere apenas se for igual a $/ "Separador de regsitro" 
	   chomp($arquivo);
	   # Verifica se o arquivo existe  
	   if (-e $arquivo){
	      # Add a file from disk
	      my $file_member = $zip->addFile($arquivo);
	   }
	}

	# Save the Zip file
	unless ( $zip->writeToFileNamed($_[1]) == AZ_OK ) {
	   die 'write error';
	}
	
	}	